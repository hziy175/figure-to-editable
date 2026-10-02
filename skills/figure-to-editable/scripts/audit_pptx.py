"""Audit named PPTX objects and final render fingerprints; not a visual scorer."""
import argparse
import hashlib
import json
import math
import re
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
KINDS = {'sp': 'shape', 'cxnSp': 'connector', 'pic': 'picture', 'grpSp': 'group'}

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def rect(node):
    if node is None:
        return None
    off, ext = node.find('a:off', NS), node.find('a:ext', NS)
    if off is None or ext is None:
        return None
    return [int(v) / 12700 for v in (off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy'))]

def collect(tree, slide, result, grouped=False):
    for node in tree:
        tag = node.tag.rsplit('}', 1)[-1]
        if tag not in KINDS and tag != 'graphicFrame':
            continue
        nv = next((e for e in node.iter() if e.tag == '{' + NS['p'] + '}cNvPr'), None)
        if nv is None:
            continue
        kind = KINDS.get(tag, 'chart' if node.find('.//c:chart', NS) is not None else 'graphic')
        xf = node.find('p:spPr/a:xfrm', NS)
        if xf is None:
            xf = node.find('p:xfrm', NS)
        if xf is None:
            xf = node.find('p:grpSpPr/a:xfrm', NS)
        text = '\n'.join(''.join(t.text or '' for t in p.iter('{' + NS['a'] + '}t'))
                         for p in node.findall('p:txBody/a:p', NS))
        key = (slide, nv.get('name', ''))
        result.setdefault(key, []).append({'kind': kind, 'text': text, 'bounds': rect(xf),
                                         'xf': xf, 'node': node, 'grouped': grouped})
        if tag == 'grpSp':
            collect(node, slide, result, grouped=True)

def endpoints(shape):
    xf, node, box = shape['xf'], shape['node'], shape['bounds']
    if xf is None or box is None or shape['grouped']:
        raise ValueError('grouped or missing transform; independent geometry review required')
    if int(xf.get('rot', '0')) or xf.get('flipH') in ('1', 'true') or xf.get('flipV') in ('1', 'true'):
        raise ValueError('rotated/flipped geometry requires independent review')
    x, y, w, h = box
    paths = node.findall('p:spPr/a:custGeom/a:pathLst/a:path', NS)
    if paths:
        if len(paths) != 1:
            raise ValueError('multiple paths are not a single straight line')
        p = paths[0]
        if [e.tag.rsplit('}', 1)[-1] for e in p] != ['moveTo', 'lnTo']:
            raise ValueError('only single-segment lines are supported')
        pw, ph = float(p.get('w')), float(p.get('h'))
        if pw <= 0 or ph <= 0:
            raise ValueError('nonpositive path dimensions')
        return [[x + float(e[0].get('x')) / pw * w, y + float(e[0].get('y')) / ph * h] for e in p]
    preset = node.find('p:spPr/a:prstGeom', NS)
    if preset is not None and preset.get('prst') == 'line':
        return [[x, y], [x + w, y + h]]
    raise ValueError('no supported native line geometry')

def run(args):
    manifest = read_json(args.manifest)
    errors = []
    if manifest.get('schema_version') != 1:
        errors.append('schema_version must be 1')
    if manifest.get('inventory_complete') is not True:
        errors.append('source inventory is not complete')
    if not re.fullmatch('[0-9a-fA-F]{64}', manifest.get('source_sha256', '')):
        errors.append('source_sha256 is missing or invalid')
    source_hash = sha(args.source)
    if source_hash != manifest.get('source_sha256', '').lower():
        errors.append('source fingerprint mismatch')
    size = manifest.get('slide_size_pt', [])
    if len(size) != 2 or any(not isinstance(v, (int, float)) or not math.isfinite(v) or v <= 0 for v in size):
        errors.append('invalid slide_size_pt')
    regions = manifest.get('regions', [])
    if not regions or len(set(regions)) != len(regions):
        errors.append('regions must be nonempty and unique')
    items = manifest.get('items', [])
    if not items:
        errors.append('empty item inventory')
    seen = set()
    for item in items:
        key = (item.get('slide'), item.get('name'))
        if key in seen or not isinstance(key[0], int) or key[0] < 1 or not isinstance(key[1], str) or not key[1]:
            errors.append('invalid or duplicate item key: ' + str(key))
        seen.add(key)
        if item.get('region') not in regions:
            errors.append('unknown region: ' + str(key))
        if item.get('kind') not in ('shape', 'connector', 'picture', 'chart', 'group'):
            errors.append('invalid object kind: ' + str(key))
        if item.get('kind') == 'picture' and not item.get('raster_reason'):
            errors.append('picture needs raster_reason: ' + str(key))
        if 'bounds_pt' in item and (len(item['bounds_pt']) != 4 or any(not math.isfinite(v) for v in item['bounds_pt'])):
            errors.append('invalid bounds: ' + str(key))
        if 'line_endpoints_pt' in item:
            pts = item['line_endpoints_pt']
            if len(pts) != 2 or any(len(p) != 2 or any(not math.isfinite(v) for v in p) for p in pts):
                errors.append('invalid endpoints: ' + str(key))
    report = {'phase': 'planning' if args.planning else 'final', 'expected_objects': len(items),
              'errors': errors, 'visual_fidelity_computed': False}
    if errors or args.planning:
        report['ok'] = not errors
        return report
    if not all((args.deck, args.preview, args.evidence)):
        raise ValueError('final audit requires --deck, --preview, --evidence')
    actual = {}
    with zipfile.ZipFile(args.deck) as z:
        if z.testzip() is not None:
            errors.append('PPTX ZIP integrity failure')
        pres = ET.fromstring(z.read('ppt/presentation.xml'))
        s = pres.find('p:sldSz', NS)
        actual_size = [int(s.get(k)) / 12700 for k in ('cx', 'cy')]
        if max(abs(a-b) for a,b in zip(size, actual_size)) > args.tolerance:
            errors.append('slide size mismatch')
        # Resolve display order via relationships; filenames alone do not define slide order.
        rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        targets = {r.get('Id'): r.get('Target') for r in rels}
        ids = pres.findall('p:sldIdLst/p:sldId', NS)
        if len(ids) != manifest.get('slide_count'):
            errors.append('slide count mismatch')
        import posixpath
        for i, sid in enumerate(ids, 1):
            target = targets[sid.get('{' + NS['r'] + '}id')]
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/' + target)
            tree = ET.fromstring(z.read(part)).find('p:cSld/p:spTree', NS)
            collect(tree, i, actual)
    for key in actual.keys() - seen:
        errors.append('unplanned object: ' + str(key))
    for item in items:
        key = (item['slide'], item['name']); matches = actual.get(key, [])
        if len(matches) != 1:
            errors.append('missing or ambiguous object: ' + str(key)); continue
        shape = matches[0]
        if shape['kind'] != item['kind']:
            errors.append('kind mismatch: ' + str(key))
        if 'text' in item and shape['text'] != item['text']:
            errors.append('text mismatch: ' + str(key))
        if 'bounds_pt' in item:
            if shape['grouped'] or shape['bounds'] is None:
                errors.append('bounds require supported ungrouped transform: ' + str(key))
            elif max(abs(a-b) for a,b in zip(item['bounds_pt'], shape['bounds'])) > args.tolerance:
                errors.append('bounds mismatch: ' + str(key))
        if 'line_endpoints_pt' in item:
            try:
                pts = endpoints(shape)
                if max(math.dist(a,b) for a,b in zip(pts,item['line_endpoints_pt'])) > args.tolerance:
                    errors.append('line endpoint mismatch: ' + str(key))
            except (ValueError, TypeError) as exc:
                errors.append(str(key) + ': ' + str(exc))
    evidence = read_json(args.evidence)
    for field, actual_hash in [('source_sha256', source_hash), ('deck_sha256', sha(args.deck)), ('preview_sha256', sha(args.preview))]:
        if evidence.get(field) != actual_hash:
            errors.append('final evidence fingerprint mismatch: ' + field)
    if not evidence.get('renderer'):
        errors.append('renderer not recorded')
    review = evidence.get('visual_review', {})
    reviewed = review.get('regions', [])
    if len(reviewed) != len(regions) or {v.get('id') for v in reviewed} != set(regions) or any(v.get('status') != 'pass' for v in reviewed):
        errors.append('visual region review incomplete or failed')
    if review.get('findings') != []:
        errors.append('visual findings missing or unresolved')
    report.update(ok=not errors, actual_objects=sum(map(len, actual.values())),
                  pictures=sum(v['kind']=='picture' for vs in actual.values() for v in vs),
                  visual_review_source='supplied review record; not inferred from object counts',
                  limitations=review.get('limitations', []))
    return report

def main():
    p=argparse.ArgumentParser(description=__doc__)
    for key in ('manifest','source'):p.add_argument('--'+key,required=True)
    for key in ('deck','preview','evidence','output'):p.add_argument('--'+key)
    p.add_argument('--planning',action='store_true');p.add_argument('--tolerance',type=float,default=.2)
    args=p.parse_args()
    try:
        if not math.isfinite(args.tolerance) or args.tolerance < 0:raise ValueError('invalid tolerance')
        result=run(args);status=0 if result['ok'] else 1
    except (OSError, ValueError, KeyError, TypeError, AttributeError, IndexError, zipfile.BadZipFile, ET.ParseError) as exc:
        result={'ok':False,'input_error':str(exc)};status=2
    output=json.dumps(result,ensure_ascii=False,indent=2)
    if args.output:Path(args.output).write_text(output+'\n',encoding='utf-8')
    print(output);return status

if __name__=='__main__':sys.exit(main())

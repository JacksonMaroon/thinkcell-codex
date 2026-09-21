"""Read-only native label fixture inspection. Requires olefile and lxml."""
from pathlib import Path
import hashlib
import io
import json
import zipfile
import olefile
from lxml import etree as E

BASE = Path(__file__).resolve().parent
NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

def value(ids, owner, field):
    ref = owner.find(field)
    if ref is None:
        return None
    source = ids.get(ref.get('idref')) if ref.get('idref') else ref
    if source is None:
        return None
    node = source.find('m_varval')
    if node is None:
        return None
    raw = node.text or node.get('val')
    try:
        return float(raw)
    except (TypeError, ValueError):
        return raw

def inspect(path):
    dest = BASE / 'extracted' / path.stem
    dest.mkdir(parents=True, exist_ok=True)
    result = {'file': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
              'models': [], 'charts': [], 'part_hashes': {}}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            data = package.read(name)
            result['part_hashes'][name] = hashlib.sha256(data).hexdigest()
            if name.startswith('ppt/charts/') and name.endswith('.xml'):
                root = E.fromstring(data)
                (dest / Path(name).name).write_bytes(E.tostring(root, pretty_print=True))
                series = []
                for ser in root.findall('.//c:ser', NS):
                    series.append({'values': ser.xpath('./c:val//c:pt/c:v/text()', namespaces=NS),
                                   'labels': ser.xpath('./c:dLbls/c:dLbl/c:tx//a:t/text()', namespaces=NS),
                                   'number_formats': ser.xpath('./c:dLbls/c:dLbl/c:numFmt/@formatCode', namespaces=NS),
                                   'show_values': ser.xpath('./c:dLbls/c:dLbl/c:showVal/@val', namespaces=NS),
                                   'show_percent': ser.xpath('./c:dLbls/c:dLbl/c:showPercent/@val', namespaces=NS)})
                result['charts'].append({'part': name, 'series': series,
                                        'field_count': len(root.findall('.//a:fld', NS))})
            if not (name.startswith('ppt/embeddings/') and name.endswith('.bin')):
                continue
            with olefile.OleFileIO(io.BytesIO(data)) as compound:
                if not compound.exists('think-cellXML'):
                    continue
                raw = compound.openstream('think-cellXML').read()
                root = E.fromstring(raw)
                ids = {node.get('id'): node for node in root if node.get('id')}
                (dest / (Path(name).stem + '.xml')).write_bytes(E.tostring(root, pretty_print=True))
                cols = []
                for table in root.findall('CSequenceChartDataTable'):
                    for ref in table.findall('./ocol/elem'):
                        vector = ids[ref.get('idref')]
                        cols.append({'category': value(ids, vector, 'm_varsrcCategory'),
                                     'values': [value(ids, ids[s.get('idref')], 'm_varsrcAbsolute')
                                                for s in vector.findall('./ocol/elem')]})
                labels = []
                for label in root.findall('CSequenceChartDataScalarLabel'):
                    labels.append({'id': label.get('id'),
                                   'native_shape_name': label.findtext('./m_ppttb/m_bstrShapeName'),
                                   'msgraph_rendering': label.find('m_bMSGraphRendering').get('val'),
                                   'prefix': label.findtext('./m_prec/m_strPrefix'),
                                   'suffix': label.findtext('./m_prec/m_strSuffix17909'),
                                   'digits': label.find('./m_prec/m_nDecimalDigits17909').get('val')})
                relative_counts = []
                for scalar in root.findall('CSequenceChartDataScalar'):
                    ref = scalar.find('m_varsrcRelative')
                    if ref is not None:
                        source = ids[ref.get('idref')]
                        relative_counts.append(len(source.findall('./m_ctextvar/*')))
                result['models'].append({'part': name, 'version': root.find('version').get('val'),
                                         'streams': ['/'.join(p) for p in compound.listdir()],
                                         'columns': cols, 'labels': labels,
                                         'relative_source_text_variable_counts': relative_counts,
                                         'chart_consistency': [x.get('val') for x in root.findall('./CSequenceChartSE/m_bConsistent')],
                                         'zero_width_space_in_model': '\u200b' in raw.decode('utf-8')})
    return result

results = [inspect(p) for p in sorted(BASE.glob('[0-9][0-9]-*.pptx'))]
(BASE / 'evidence.json').write_text(json.dumps(results, indent=2))
for result in results:
    print(result['file'])
    for model in result['models']:
        if model['labels']:
            print('  model', model['version'], 'first column', model['columns'][0]['values'],
                  'labels', len(model['labels']), 'suffixes', sorted(set(x['suffix'] or '' for x in model['labels'])),
                  'relative text variables', model['relative_source_text_variable_counts'])
    for chart in result['charts']:
        print('  chart labels', [s['labels'] for s in chart['series']], 'fields', chart['field_count'])

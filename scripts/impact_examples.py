"""Rebuild synthetic protocol examples. All hashes and replay claims are placeholders.

These examples demonstrate the contract; they do not assert actual Office execution.
Run this file to write examples/*.changes.json.
"""

from copy import deepcopy
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SECTIONS = ('entities', 'dependencies', 'computations', 'impacts', 'lineages')


def number(data):
    return {'type': 'number', 'data': str(data)}


def source(address=None, file_format='xlsx', index=0):
    if file_format == 'docx':
        return {'part': '/word/document.xml', 'locator': {'type': 'word', 'elementPath': [0, index]}}
    return {'part': '/xl/worksheets/sheet1.xml', 'locator': {'type': 'excel', 'sheetId': 1, **({'address': address} if address else {})}}


def observation(data, origin='literal', computation_id=None, property='/value'):
    result = {'property': property, 'status': 'known', 'origin': origin,
              'freshness': 'notApplicable' if origin == 'literal' else 'verified',
              'value': number(data) if isinstance(data, (int, float)) else data}
    if computation_id:
        result['computationId'] = computation_id
    return result


def entity(entity_id, address, old, new, formula=None, origin='literal', kind='cell'):
    result = {'entityId': entity_id, 'kind': kind, 'label': entity_id}
    for side, data in (('before', old), ('after', new)):
        result[side] = {'source': source(address), 'observations': [] if data is None else [observation(data, origin, side + '-calc' if origin != 'literal' else None)]}
        if formula is not None:
            result[side]['formula'] = formula
    return result


def change(entity, change_id, old, new, property='/value', aspect='value'):
    return {'id': change_id, 'operation': 'update', 'aspect': aspect,
            'target': {key: entity[key] for key in ('entityId', 'kind', 'label')}, 'property': property,
            'before': {'value': number(old) if isinstance(old, (int, float)) else old, 'source': deepcopy(entity['before']['source'])},
            'after': {'value': number(new) if isinstance(new, (int, float)) else new, 'source': deepcopy(entity['after']['source'])},
            'match': {'method': 'manual', 'confidence': 1, 'evidence': ['Synthetic example identity; no Office file is supplied.']}}


def coverage(status='complete', code=None, entity_ids=None):
    return {'status': status, 'issues': [] if code is None else [{'code': code, 'detail': 'Synthetic coverage example.', 'entityIds': entity_ids or []}]}


def document(name, file_format='xlsx'):
    graph = {'version': '1.0.0', 'scope': {'entityIds': [], 'mechanisms': []}, **{section: [] for section in SECTIONS},
             'coverage': {category: coverage('notRequested') for category in ('dependencies', 'attribution', 'computation', 'structure')},
             'delivery': {'status': 'complete', 'omitted': dict.fromkeys(SECTIONS, 0)}}
    return {'schemaVersion': '1.0.0', 'id': name, 'format': file_format,
            'documents': {'before': {'name': 'synthetic-old.' + file_format, 'sha256': 'a' * 64}, 'after': {'name': 'synthetic-new.' + file_format, 'sha256': 'b' * 64}},
            'comparison': {'profile': 'office-semantic-v1', 'aspects': ['existence', 'value', 'formula', 'style'] if file_format != 'docx' else ['existence', 'text']},
            'coverage': {'status': 'complete', 'issues': []}, 'changes': [], 'extensions': {'ppt2wiki.impact': graph}}


def dependency(dependency_id, dependent_id, members, end=None):
    ref = {'kind': 'range' if end else 'cell', 'expression': 'A1:' + end if end else 'B1', 'resolution': 'resolved', 'members': members}
    if end:
        ref['areas'] = [{'sheetEntityId': 'sheet', 'start': 'A1', 'end': end}]
    return {'id': dependency_id, 'dependentEntityId': dependent_id, 'before': deepcopy(ref), 'after': deepcopy(ref)}


def cause(change_id, *dependency_ids):
    return {'kind': 'change', 'changeId': change_id, 'dependencyIds': list(dependency_ids)}


def impact(impact_id, entity_id, changes, causes, outcome='changed'):
    return {'id': impact_id, 'mechanism': 'formula', 'entityId': entity_id, 'property': '/value', 'outcome': outcome,
            'changeIds': changes, 'causes': causes,
            'evidence': {'level': 'verified', 'method': 'controlledReplay',
                         'detail': 'ILLUSTRATIVE ONLY: a real producer must attach an independently checked controlled replay report; these hashes are placeholders.',
                         'computationIds': ['before-calc', 'after-calc'], 'artifactSha256': 'f' * 64}}


def computations(doc):
    graph = doc['extensions']['ppt2wiki.impact']
    for side in ('before', 'after'):
        outputs = [{'entityId': entity['entityId'], 'property': item['property'], 'value': deepcopy(item['value'])}
                   for entity in graph['entities'] if entity[side]
                   for item in entity[side]['observations'] if item.get('computationId') == side + '-calc']
        graph['computations'].append({'id': side + '-calc', 'side': side, 'inputSha256': doc['documents'][side]['sha256'],
                                      'engine': {'name': 'SYNTHETIC EXAMPLE', 'version': '0'}, 'contextSha256': 'c' * 64,
                                      'artifactSha256': ('d' if side == 'before' else 'e') * 64, 'environment': [], 'outputs': outputs})
    graph['coverage'] = {category: coverage('notRequested' if category == 'structure' else 'complete') for category in graph['coverage']}


def range_example():
    doc = document('excel-range-impact')
    doc['comparison']['compareCachedFormulaValues'] = True
    graph = doc['extensions']['ppt2wiki.impact']
    sheet = entity('sheet', None, None, None, kind='sheet')
    a = entity('A5', 'A5', 10, 20)
    b = entity('B1', 'B1', 100, 110, 'SUM(A1:A10)', 'storedCache')
    c = entity('C1', 'C1', 200, 220, 'B1*2', 'storedCache')
    graph['entities'] = [sheet, a, b, c]
    doc['changes'] = [change(a, 'change-A5', 10, 20), change(b, 'change-B1', 100, 110), change(c, 'change-C1', 200, 220)]
    graph['dependencies'] = [dependency('range-to-B', 'B1', ['A5'], 'A10'), dependency('B-to-C', 'C1', ['B1'])]
    graph['impacts'] = [impact('impact-B', 'B1', ['change-B1'], [cause('change-A5', 'range-to-B')]),
                        impact('impact-C', 'C1', ['change-C1'], [cause('change-B1', 'B-to-C')])]
    computations(doc)
    return doc


def compensated_example():
    doc = range_example()
    doc['id'] = 'excel-compensated-impact'
    graph = doc['extensions']['ppt2wiki.impact']
    graph['entities'] = graph['entities'][:2] + [entity('A6', 'A6', 20, 10), entity('B1', 'B1', 100, 100, 'SUM(A1:A10)', 'storedCache')]
    doc['changes'] = [doc['changes'][0], change(graph['entities'][2], 'change-A6', 20, 10)]
    graph['dependencies'] = [dependency('range-to-B', 'B1', ['A5', 'A6'], 'A10')]
    graph['impacts'] = [impact('impact-B', 'B1', [], [cause('change-A5', 'range-to-B'), cause('change-A6', 'range-to-B')], 'unchanged')]
    graph['impacts'][0]['evidence']['detail'] = 'ILLUSTRATIVE: both changed inputs participate, but their effects cancel; no B value update is emitted.'
    graph['computations'] = []
    computations(doc)
    return doc


def reference_example():
    doc = range_example()
    doc['id'] = 'excel-reference-impact'
    graph = doc['extensions']['ppt2wiki.impact']
    graph['entities'] = [graph['entities'][0], entity('A6', 'A6', None, 5), entity('B1', 'B1', 100, 105, 'SUM(BudgetRange)', 'storedCache')]
    a, b = graph['entities'][1:]
    a['before'] = None
    add = {'id': 'add-A6', 'operation': 'add', 'aspect': 'existence', 'target': {key: a[key] for key in ('entityId', 'kind', 'label')}, 'property': '', 'before': None,
           'after': {'value': {'type': 'object', 'data': {'value': number(5)}}, 'source': deepcopy(a['after']['source'])}}
    doc['changes'] = [add, change(b, 'change-B1', 100, 105)]
    dep = dependency('name-to-B', 'B1', [], 'A5')
    for side in ('before', 'after'):
        dep[side].update(kind='name', expression='BudgetRange')
    dep['after']['areas'][0]['end'] = 'A6'
    dep['after']['members'] = ['A6']
    graph['dependencies'] = [dep]
    graph['impacts'] = [impact('impact-B', 'B1', ['change-B1'], [cause('add-A6', 'name-to-B'), {'kind': 'dependency', 'dependencyId': 'name-to-B', 'dependencyIds': []}])]
    graph['computations'] = []
    computations(doc)
    return doc


def derived_example():
    doc = range_example()
    doc['id'] = 'excel-derived-impact'
    doc['comparison']['compareCachedFormulaValues'] = False
    doc['changes'] = doc['changes'][:1]
    graph = doc['extensions']['ppt2wiki.impact']
    for entity in graph['entities'][2:]:
        for side in ('before', 'after'):
            entity[side]['observations'][0]['origin'] = 'recalculated'
    graph['impacts'][0]['changeIds'] = []
    graph['impacts'][1]['changeIds'] = []
    graph['impacts'][1]['causes'] = [cause('change-A5', 'range-to-B', 'B-to-C')]
    return doc


def unknown_example():
    doc = range_example()
    doc['id'] = 'excel-unknown-impact'
    doc['changes'] = doc['changes'][:1]
    doc['comparison']['compareCachedFormulaValues'] = False
    graph = doc['extensions']['ppt2wiki.impact']
    graph['entities'] = graph['entities'][:3]
    for side in ('before', 'after'):
        state = graph['entities'][2][side]
        state['formula'] = 'CUSTOM(A5)'
        state['observations'] = [{'property': '/value', 'status': 'unavailable', 'origin': 'storedCache', 'freshness': 'unverified'}]
    graph['dependencies'] = [dependency('dynamic-to-B', 'B1', ['A5'])]
    for side in ('before', 'after'):
        graph['dependencies'][0][side].update(kind='dynamic', expression='CUSTOM(A5)', resolution='partial')
    graph['computations'] = []
    graph['impacts'] = [impact('impact-B', 'B1', [], [cause('change-A5', 'dynamic-to-B')], 'unknown')]
    graph['impacts'][0]['evidence'] = {'level': 'potential', 'method': 'staticAnalysis', 'detail': 'CUSTOM is an illustrative unsupported function; no execution or causal claim.', 'computationIds': []}
    graph['coverage'].update(dependencies=coverage('partial', 'unsupportedReference', ['B1']), attribution=coverage('partial', 'unverifiedCause', ['B1']), computation=coverage('partial', 'unavailableResult', ['B1']))
    return doc


def style_example():
    doc = document('powerpoint-style-impact', 'pptx')
    doc['comparison']['aspects'] = ['style']
    graph = doc['extensions']['ppt2wiki.impact']
    theme_source = {'part': '/ppt/theme/theme1.xml', 'locator': {'type': 'xml', 'elementPath': []}}
    shape_source = {'part': '/ppt/slides/slide1.xml', 'locator': {'type': 'powerpoint', 'slideId': 256, 'shapePath': [2]}}
    theme = {'entityId': 'theme', 'kind': 'other', 'label': '主题', 'before': {'source': theme_source, 'observations': []}, 'after': {'source': deepcopy(theme_source), 'observations': []}}
    shape = {'entityId': 'shape', 'kind': 'shape', 'label': '标题', 'before': {'source': shape_source, 'observations': []}, 'after': {'source': deepcopy(shape_source), 'observations': []}}
    graph['entities'] = [theme, shape]
    old, new = {'type': 'text', 'data': '18293D'}, {'type': 'text', 'data': '245AC1'}
    doc['changes'] = [change(theme, 'theme-color', old, new, '/style/font/color', 'style'), change(shape, 'shape-color', old, new, '/style/font/color', 'style')]
    dep = dependency('theme-to-shape', 'shape', ['theme'])
    for side in ('before', 'after'):
        dep[side].update(kind='object', expression='theme accent color')
    graph['dependencies'] = [dep]
    item = impact('impact-shape', 'shape', ['shape-color'], [cause('theme-color', 'theme-to-shape')])
    item.update(mechanism='styleInheritance', entityId='shape', property='/style/font/color')
    item['evidence'] = {'level': 'supported', 'method': 'snapshotComparison', 'detail': 'Resolved effective style differs; no controlled replay has been supplied.', 'computationIds': []}
    graph['impacts'] = [item]
    graph['coverage'].update(dependencies=coverage(), attribution=coverage('partial', 'unverifiedCause', ['shape']))
    return doc


def split_example():
    doc = document('word-split-impact', 'docx')
    graph = doc['extensions']['ppt2wiki.impact']
    for entity_id, old, new, index in (('paragraph-old', '甲乙', None, 0), ('paragraph-new-1', None, '甲', 0), ('paragraph-new-2', None, '乙', 1)):
        item = {'entityId': entity_id, 'kind': 'paragraph', 'label': entity_id,
                'before': None if old is None else {'source': source(file_format='docx', index=index), 'observations': []},
                'after': None if new is None else {'source': source(file_format='docx', index=index), 'observations': []}}
        graph['entities'].append(item)
        side = 'after' if old is None else 'before'
        state = {'source': deepcopy(item[side]['source']), 'value': {'type': 'object', 'data': {'text': new if old is None else old}}}
        doc['changes'].append({'id': 'change-' + entity_id, 'operation': 'add' if old is None else 'remove', 'aspect': 'existence', 'property': '',
                               'target': {key: item[key] for key in ('entityId', 'kind', 'label')}, 'before': state if side == 'before' else None, 'after': state if side == 'after' else None})
    graph['lineages'] = [{'id': 'split-paragraph', 'kind': 'split', 'beforeEntityIds': ['paragraph-old'], 'afterEntityIds': ['paragraph-new-1', 'paragraph-new-2'], 'detail': 'Synthetic structural correspondence, not a recovered edit history; three core existence events belong to one split.'}]
    graph['coverage']['structure'] = coverage()
    return doc


def build_examples():
    examples = {item['id']: item for item in (range_example(), compensated_example(), reference_example(), derived_example(), unknown_example(), style_example(), split_example())}
    for item in examples.values():
        graph = item['extensions']['ppt2wiki.impact']
        graph['scope'] = {'entityIds': [entity['entityId'] for entity in graph['entities']], 'mechanisms': sorted({impact['mechanism'] for impact in graph['impacts']})}
    return examples


if __name__ == '__main__':
    for name, item in build_examples().items():
        destination = ROOT / 'examples' / (name + '.changes.json')
        destination.write_text(json.dumps(item, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        print('Wrote ' + destination.name)

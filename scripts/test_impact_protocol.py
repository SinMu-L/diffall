"""Contract regression cases, run by validate_change_protocol.py --self-test."""

from copy import deepcopy

from impact_examples import build_examples, change, coverage, number
from validate_change_protocol import validation_errors


def run_impact_self_tests(validator):
    fixtures = build_examples()
    accepted = rejected = 0

    def check(name, document, expected=None):
        nonlocal accepted, rejected
        errors = validation_errors(document, validator)
        if expected is None:
            if errors:
                raise AssertionError(f'{name} rejected: {errors}')
            accepted += 1
        else:
            if not errors or not any(expected in error for error in errors):
                raise AssertionError(f'{name}: expected {expected!r}, got {errors}')
            rejected += 1

    for name, document in fixtures.items():
        check(name, document)

    def case(name, mutate, expected, fixture='excel-range-impact'):
        document = deepcopy(fixtures[fixture])
        mutate(document, document['extensions']['ppt2wiki.impact'])
        check(name, document, expected)

    case('unknown extension version', lambda d, g: g.update(version='2.0.0'), '1.0.0')
    case('unknown extension core field', lambda d, g: g.update(causedBy='A5'), 'Additional properties')
    case('null extension is invalid', lambda d, g: d['extensions'].update({'ppt2wiki.impact': None}), 'not of type')
    case('duplicate registered entity', lambda d, g: g['entities'].append(deepcopy(g['entities'][1])), 'duplicate entityId')
    case('missing cause event', lambda d, g: g['impacts'][0]['causes'][0].update(changeId='missing'), 'unknown cause changeId')
    case('unregistered cause entity', lambda d, g: g['entities'].pop(1), 'unknown entityId')
    case('dangling dependency', lambda d, g: g['impacts'][0]['causes'][0].update(dependencyIds=['missing']), 'unknown dependencyId')
    case('wrong dependency path order', lambda d, g: g['impacts'][1].update(causes=[{'kind': 'change', 'changeId': 'change-A5', 'dependencyIds': ['B-to-C', 'range-to-B']}]), 'within one snapshot')
    case('source cell outside range', lambda d, g: g['dependencies'][0]['after']['areas'][0].update(end='A4'), 'outside the declared reference areas')
    case('range bounds reversed', lambda d, g: g['dependencies'][0]['before']['areas'][0].update(start='A10', end='A1'), 'bounds are reversed')
    case('range column overflow', lambda d, g: g['dependencies'][0]['after']['areas'][0].update(end='XFE10'), 'range exceeds')
    case('range row overflow', lambda d, g: g['dependencies'][0]['after']['areas'][0].update(end='A1048577'), 'range exceeds')
    case('range wrong sheet identity', lambda d, g: g['dependencies'][0]['after']['areas'][0].update(sheetEntityId='A5'), 'Excel sheet entity')
    case('wrong snapshot anchor', lambda d, g: g['entities'][1]['after']['source']['locator'].update(address='A6'), 'anchor disagrees')
    case('wrong computation input digest', lambda d, g: g['computations'][1].update(inputSha256='a' * 64), 'input digest')
    case('wrong computation side', lambda d, g: g['entities'][2]['after']['observations'][0].update(computationId='before-calc'), 'wrong snapshot side')
    case('stale cache cannot verify causality', lambda d, g: g['entities'][2]['after']['observations'][0].update(freshness='stale'), 'reliable results')
    case('cache comparison opt in', lambda d, g: d['comparison'].update(compareCachedFormulaValues=False), 'requires compareCachedFormulaValues')
    case('verified cache needs computation', lambda d, g: g['entities'][2]['after']['observations'][0].pop('computationId'), 'verified cache requires')
    case('computation output must match observation', lambda d, g: g['computations'][1]['outputs'][0].update(value=number(999)), 'does not match')
    case('core saved value must match saved observation', lambda d, g: d['changes'][1]['after'].update(value=number(999)), 'stored observation disagrees')
    case('computed values cannot replace snapshot values', lambda d, g: g['entities'][2]['after']['observations'][0].update(origin='recalculated'), 'cannot be substituted')
    case('static reference is not causal proof', lambda d, g: g['impacts'][0]['evidence'].update(method='staticAnalysis'), 'controlledReplay')
    case('verified attribution needs replay artifact', lambda d, g: g['impacts'][0]['evidence'].pop('artifactSha256'), 'artifactSha256')
    case('verified attribution needs both snapshots', lambda d, g: g['impacts'][0]['evidence'].update(computationIds=['after-calc']), 'one computation for each snapshot')
    case('verified cause cannot be unknown', lambda d, g: g['impacts'][0].update(causes=[{'kind': 'unresolved', 'detail': 'Unknown input.'}]), 'cannot contain an unresolved cause')
    case('confirmed result cannot say unchanged', lambda d, g: g['impacts'][0].update(outcome='unchanged'), 'requires changed outcome')
    case('equal computed values cannot say changed', lambda d, g: g['impacts'][0].update(outcome='changed'), 'equal observations', 'excel-compensated-impact')
    case('unavailable does not mean blank', lambda d, g: g['entities'][2]['after']['observations'][0].update(value={'type': 'null', 'data': None}), 'given schemas', 'excel-unknown-impact')
    case('unavailable cannot imply unchanged', lambda d, g: g['impacts'][0].update(outcome='unchanged'), 'equal reliable observations', 'excel-unknown-impact')
    case('confirmed change must be linked', lambda d, g: g['impacts'][0].update(changeIds=[]), 'exactly the delivered target')
    case('event cannot cause itself', lambda d, g: g['impacts'][0]['causes'][0].update(changeId='change-B1', dependencyIds=[]), 'its own cause')
    case('unrecorded engine drift', lambda d, g: g['computations'][1]['engine'].update(version='1'), 'engine drift')
    case('unrecorded context drift', lambda d, g: g['computations'][1].update(contextSha256='e' * 64), 'context drift')
    case('unresolved analysis must report coverage', lambda d, g: g['coverage'].update(dependencies=coverage()), 'partial dependency coverage', 'excel-unknown-impact')
    case('unknown cause must report coverage', lambda d, g: g['coverage'].update(attribution=coverage()), 'partial coverage', 'excel-unknown-impact')
    case('not requested cannot claim outputs', lambda d, g: g['coverage'].update(computation=coverage('notRequested')), 'notRequested cannot contain')
    case('complete delivery cannot omit', lambda d, g: g['delivery']['omitted'].update(impacts=10), 'omitted counts')
    case('truncation must report coverage', lambda d, g: g['delivery'].update(status='truncated', omitted={**g['delivery']['omitted'], 'impacts': 10}), 'truncatedContent')
    case('split cardinality', lambda d, g: g['lineages'][0].update(afterEntityIds=['paragraph-new-1']), 'multiple new entities', 'word-split-impact')
    case('split identities cannot overlap', lambda d, g: g['lineages'][0].update(afterEntityIds=['paragraph-old', 'paragraph-new-1']), 'distinct old and new identities', 'word-split-impact')
    case('impact scope must be explicit', lambda d, g: g['scope'].update(mechanisms=[]), 'outside declared analysis scope')
    case('wrong kind in registry', lambda d, g: g['entities'][1].update(kind='shape'), 'kinds disagree')
    case('unknown observation computation', lambda d, g: g['entities'][2]['after']['observations'][0].update(computationId='missing'), 'unknown computationId')
    case('duplicate computation output', lambda d, g: g['computations'][0]['outputs'].append(deepcopy(g['computations'][0]['outputs'][0])), 'duplicate entity/property output')
    case('unknown property result type', lambda d, g: g['entities'][2]['after']['observations'][0].update(value={'type': 'object', 'data': {}}), 'given schemas')
    case('formula context cannot disagree', lambda d, g: g['entities'][2]['after'].update(formula='=SUM(A1:A10)'), 'given schemas')
    case('reference observation is not reference change', lambda d, g: g['impacts'][0].update(causes=[{'kind': 'dependency', 'dependencyId': 'range-to-B', 'dependencyIds': []}]), 'changed reference')
    case('member bookkeeping is not reference change', lambda d, g: (g['dependencies'][0]['after'].update(members=[]), g['impacts'][0].update(causes=[{'kind': 'dependency', 'dependencyId': 'range-to-B', 'dependencyIds': []}])), 'changed reference')
    case('computation value type is checked', lambda d, g: g['computations'][0]['outputs'][0].update(value={'type': 'array', 'data': []}), 'not one of')
    case('analysis must be changeset level', lambda d, g: d['changes'][0].update(extensions={'ppt2wiki.impact': {}}), 'belongs on ChangeSet.extensions')

    mixed = deepcopy(fixtures['excel-range-impact'])
    g = mixed['extensions']['ppt2wiki.impact']
    g['dependencies'][0]['after'] = None
    g['dependencies'][1]['before'] = None
    g['impacts'][1]['causes'] = [{'kind': 'change', 'changeId': 'change-A5', 'dependencyIds': ['range-to-B', 'B-to-C']}]
    check('path cannot combine incompatible snapshot edges', mixed, 'within one snapshot')

    moved = deepcopy(fixtures['excel-range-impact'])
    g = moved['extensions']['ppt2wiki.impact']
    g['entities'][1]['after']['source']['locator']['address'] = 'A6'
    moved['changes'][0]['after']['source']['locator']['address'] = 'A6'
    check('same entity can move within the referenced range', moved)

    partial = deepcopy(fixtures['excel-range-impact'])
    g = partial['extensions']['ppt2wiki.impact']
    g['delivery']['status'] = 'truncated'
    g['delivery']['omitted']['impacts'] = 5
    g['coverage']['attribution'] = coverage('partial', 'truncatedContent')
    check('truncated but reference-closed graph', partial)

    cycle = deepcopy(fixtures['excel-range-impact'])
    g = cycle['extensions']['ppt2wiki.impact']
    edge = deepcopy(g['dependencies'][1])
    edge.update(id='C-to-B', dependentEntityId='B1')
    for side in ('before', 'after'):
        edge[side].update(expression='C1', members=['C1'])
    g['dependencies'].append(edge)
    check('cyclic dependencies are representable', cycle)

    engine = deepcopy(fixtures['excel-range-impact'])
    g = engine['extensions']['ppt2wiki.impact']
    g['computations'][1]['engine']['version'] = '1'
    for item in g['impacts']:
        item['causes'].append({'kind': 'environment', 'name': 'engine', 'before': {'type': 'text', 'data': '0'}, 'after': {'type': 'text', 'data': '1'}})
    check('explicit engine cause', engine)

    reordered = deepcopy(fixtures['excel-range-impact'])
    g = reordered['extensions']['ppt2wiki.impact']
    inputs = [{'name': 'clock', 'fingerprint': '1' * 64}, {'name': 'locale', 'fingerprint': '2' * 64}]
    g['computations'][0]['environment'] = inputs
    g['computations'][1]['environment'] = list(reversed(inputs))
    check('environment order is not drift', reordered)

    formula = deepcopy(fixtures['excel-range-impact'])
    g = formula['extensions']['ppt2wiki.impact']
    b = g['entities'][2]
    b['after']['formula'] = 'SUM(A1:A10)+1'
    for entity_index, change_index, value in ((2, 1, 111), (3, 2, 222)):
        g['entities'][entity_index]['after']['observations'][0]['value'] = number(value)
        formula['changes'][change_index]['after']['value'] = number(value)
        g['computations'][1]['outputs'][entity_index - 2]['value'] = number(value)
    formula['changes'].append(change(b, 'formula-B1', {'type': 'text', 'data': 'SUM(A1:A10)'}, {'type': 'text', 'data': 'SUM(A1:A10)+1'}, '/formula', 'formula'))
    g['impacts'][0]['causes'].append({'kind': 'change', 'changeId': 'formula-B1', 'dependencyIds': []})
    check('formula and referenced input can change together', formula)

    context = deepcopy(fixtures['excel-range-impact'])
    g = context['extensions']['ppt2wiki.impact']
    g['computations'][1]['contextSha256'] = 'e' * 64
    for item in g['impacts']:
        item['causes'].append({'kind': 'environment', 'name': 'context', 'before': {'type': 'text', 'data': 'c' * 64}, 'after': {'type': 'text', 'data': 'e' * 64}})
    check('explicit environment context change', context)

    cached = deepcopy(fixtures['excel-range-impact'])
    g = cached['extensions']['ppt2wiki.impact']
    g['computations'] = []
    for entity in g['entities'][2:]:
        for side in ('before', 'after'):
            item = entity[side]['observations'][0]
            item.pop('computationId')
            item['freshness'] = 'unverified'
    for item in g['impacts']:
        item['evidence'] = {'level': 'supported', 'method': 'snapshotComparison', 'detail': 'Only saved cache differences and static dependencies were observed.', 'computationIds': []}
    g['coverage']['attribution'] = coverage('partial', 'unverifiedCause')
    g['coverage']['computation'] = coverage('notRequested')
    check('saved cache differences without verified causality', cached)

    merged = deepcopy(fixtures['word-split-impact'])
    merged['documents']['before'], merged['documents']['after'] = merged['documents']['after'], merged['documents']['before']
    g = merged['extensions']['ppt2wiki.impact']
    for entity in g['entities']:
        entity['before'], entity['after'] = entity['after'], entity['before']
    for item in merged['changes']:
        item['before'], item['after'] = item['after'], item['before']
        item['operation'] = 'add' if item['operation'] == 'remove' else 'remove'
    lineage = g['lineages'][0]
    lineage['kind'] = 'merge'
    lineage['beforeEntityIds'], lineage['afterEntityIds'] = lineage['afterEntityIds'], lineage['beforeEntityIds']
    check('many to one structural merge', merged)

    print(f'PASS: impact extension {rejected} invalid cases rejected; {accepted} valid scenarios accepted')

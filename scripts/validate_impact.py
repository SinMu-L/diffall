"""Validate the formal ppt2wiki.impact extension without evaluating Office files."""

from functools import lru_cache
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
EXTENSION = 'ppt2wiki.impact'
SIDES = ('before', 'after')


@lru_cache(maxsize=1)
def impact_validator():
    from jsonschema import Draft202012Validator
    from referencing import Registry, Resource

    schemas = [json.loads((ROOT / 'schemas' / name).read_text(encoding='utf-8')) for name in (
        'office-change.v1.schema.json', 'office-impact.v1.schema.json',
    )]
    registry = Registry().with_resources((schema['$id'], Resource.from_contents(schema)) for schema in schemas)
    Draft202012Validator.check_schema(schemas[1])
    return Draft202012Validator(schemas[1], registry=registry)


def coordinates(address):
    letters = address.rstrip('0123456789')
    column = 0
    for letter in letters:
        column = column * 26 + ord(letter) - ord('A') + 1
    return column, int(address[len(letters):])


def impact_errors(document, equal_value, source_errors):
    extensions = document.get('extensions', {})
    if EXTENSION not in extensions:
        return []
    graph = extensions[EXTENSION]
    prefix = 'extensions.' + EXTENSION
    errors = [f'{prefix} {error.json_path}: {error.message}' for error in impact_validator().iter_errors(graph)]
    if errors:
        return errors

    def fail(label, message):
        errors.append(f'{prefix}.{label}: {message}')

    def index(items, field, label):
        result = {}
        for item in items:
            key = item[field]
            if key in result:
                fail(label, f'duplicate {field}: {key}')
            result[key] = item
        return result

    entities = index(graph['entities'], 'entityId', 'entities')
    dependencies = index(graph['dependencies'], 'id', 'dependencies')
    computations = index(graph['computations'], 'id', 'computations')
    index(graph['impacts'], 'id', 'impacts')
    index(graph['lineages'], 'id', 'lineages')
    changes = {change['id']: change for change in document['changes']}
    observations = {}
    for entity_id in graph['scope']['entityIds']:
        if entity_id not in entities:
            fail('scope', f'unknown entityId: {entity_id}')

    def entity_at(entity_id, side, label):
        entity = entities.get(entity_id)
        if entity is None:
            fail(label, f'unknown entityId: {entity_id}')
            return None
        if entity[side] is None:
            fail(label, f'{entity_id} does not exist on {side}')
        return entity[side]

    def equivalent(a, b, property):
        aspect = property.lstrip('/').split('/')[0]
        if aspect == 'geometry' and (a['type'] != 'object' or b['type'] != 'object'):
            return a == b
        return equal_value(a, b, aspect, document['comparison'].get('geometryToleranceEmu', 0))

    def reliable(observation):
        return observation is not None and observation['status'] == 'known' and observation['freshness'] in {'verified', 'notApplicable'}

    def coverage_issue(category, label, codes, message):
        coverage = graph['coverage'][category]
        if coverage['status'] != 'partial' or not any(issue['code'] in codes for issue in coverage['issues']):
            fail(label, message)

    for entity_id, entity in entities.items():
        if all(entity[side] is None for side in SIDES):
            fail(entity_id, 'entity must exist on at least one side')
        for side in SIDES:
            state = entity[side]
            if state is None:
                continue
            label = f'{entity_id}.{side}'
            errors.extend(source_errors(state['source'], document['format'], prefix + '.' + label))
            for observation in state['observations']:
                key = (entity_id, side, observation['property'])
                if key in observations:
                    fail(label, 'duplicate observation property')
                observations[key] = observation
                origin, freshness = observation['origin'], observation['freshness']
                if origin == 'literal' and freshness != 'notApplicable':
                    fail(label, 'literal observations require notApplicable freshness')
                if origin == 'storedCache' and freshness == 'notApplicable':
                    fail(label, 'cache freshness cannot be notApplicable')
                if observation['status'] == 'unavailable':
                    if freshness != 'unverified' or origin in {'recalculated', 'rendered'} or 'computationId' in observation:
                        fail(label, 'unavailable result must be unverified and cannot claim computation output')
                elif origin in {'recalculated', 'rendered'} and freshness != 'verified':
                    fail(label, 'computed observations require verified freshness')
                if origin == 'literal' and 'computationId' in observation:
                    fail(label, 'literal observation cannot cite a computation')
                if origin == 'storedCache' and freshness == 'verified' and 'computationId' not in observation:
                    fail(label, 'verified cache requires a matching computation')
                if graph['coverage']['computation']['status'] != 'notRequested' and (observation['status'] == 'unavailable' or freshness in {'stale', 'unverified'}):
                    coverage_issue('computation', label, {'unavailableResult', 'staleCache', 'externalUnavailable'}, 'unavailable or unverified result requires partial computation coverage')
                if observation.get('property') == '/value' and observation.get('value', {}).get('type') in {'object', 'array', 'asset'}:
                    fail(label, '/value observations must be scalar')
                computation_id = observation.get('computationId')
                if computation_id:
                    computation = computations.get(computation_id)
                    if computation is None:
                        fail(label, f'unknown computationId: {computation_id}')
                    elif computation['side'] != side:
                        fail(label, 'computation belongs to the wrong snapshot side')
                    elif not any(output['entityId'] == entity_id and output['property'] == observation['property'] and observation.get('value') is not None and equivalent(output['value'], observation['value'], observation['property']) for output in computation['outputs']):
                        fail(label, 'observation does not match the computation output')

    for change in document['changes']:
        entity = entities.get(change['target']['entityId'])
        if entity is None:
            continue
        label = change['id']
        if entity['kind'] != change['target']['kind']:
            fail(label, 'entity registry and Change target kinds disagree')
        for side in SIDES:
            state, anchor = change[side], entity[side]
            if state is None:
                if change['operation'] in {'add', 'remove'} and anchor is not None:
                    fail(label, 'whole-object existence disagrees with entity registry')
                continue
            if anchor is None or anchor['source'] != state['source']:
                fail(label, f'{side} Change anchor disagrees with entity registry')
            observation = observations.get((entity['entityId'], side, change['property']))
            if observation and observation['status'] == 'known' and observation['origin'] in {'literal', 'storedCache'}:
                if not equivalent(observation['value'], state['value'], change['property']):
                    fail(label, 'stored observation disagrees with core Change value')
                if change['aspect'] == 'value' and observation['origin'] == 'storedCache' and not document['comparison'].get('compareCachedFormulaValues', False):
                    fail(label, 'cached value update requires compareCachedFormulaValues')
            if change['aspect'] == 'formula' and anchor is not None and 'formula' in anchor:
                formula_value = {'type': 'null', 'data': None} if anchor['formula'] is None else {'type': 'text', 'data': anchor['formula']}
                if formula_value != state['value']:
                    fail(label, 'formula context disagrees with core formula Change')

    for computation_id, computation in computations.items():
        side = computation['side']
        if computation['inputSha256'] != document['documents'][side]['sha256']:
            fail(computation_id, 'input digest does not bind the declared snapshot')
        output_keys = set()
        for output in computation['outputs']:
            entity_at(output['entityId'], side, computation_id)
            key = output['entityId'], output['property']
            if key in output_keys:
                fail(computation_id, 'duplicate entity/property output')
            output_keys.add(key)
        names = [item['name'] for item in computation.get('environment', [])]
        if len(names) != len(set(names)):
            fail(computation_id, 'environment names must be unique')

    def in_area(entity_id, side, area):
        state = entities.get(entity_id, {}).get(side)
        sheet = entities.get(area['sheetEntityId'], {}).get(side)
        if state is None or sheet is None:
            return False
        loc, sheet_loc = state['source']['locator'], sheet['source']['locator']
        if loc['type'] != 'excel' or sheet_loc['type'] != 'excel' or 'address' not in loc:
            return False
        if state['source']['part'] != sheet['source']['part'] or loc.get('sheetId') != sheet_loc.get('sheetId'):
            return False
        x, y = coordinates(loc['address'])
        x1, y1 = coordinates(area['start'])
        x2, y2 = coordinates(area['end'])
        return x1 <= x <= x2 and y1 <= y <= y2

    for dependency_id, dependency in dependencies.items():
        if all(dependency[side] is None for side in SIDES):
            fail(dependency_id, 'dependency must exist on at least one side')
        for side in SIDES:
            ref = dependency[side]
            if ref is None:
                continue
            entity_at(dependency['dependentEntityId'], side, dependency_id)
            for member in ref['members']:
                entity_at(member, side, dependency_id)
            areas = ref.get('areas', [])
            for area in areas:
                sheet = entities.get(area['sheetEntityId'])
                sheet_state = entity_at(area['sheetEntityId'], side, dependency_id)
                if sheet and (sheet['kind'] != 'sheet' or (sheet_state and sheet_state['source']['locator']['type'] != 'excel')):
                    fail(dependency_id, 'area must refer to an Excel sheet entity')
                start, end = coordinates(area['start']), coordinates(area['end'])
                if any(x > 16384 or y > 1048576 for x, y in (start, end)):
                    fail(dependency_id, 'range exceeds XFD1048576')
                if start[0] > end[0] or start[1] > end[1]:
                    fail(dependency_id, 'range bounds are reversed')
            if areas:
                for member in ref['members']:
                    if entities.get(member, {}).get('kind') == 'cell' and not any(in_area(member, side, area) for area in areas):
                        fail(dependency_id, f'{member} is outside the declared reference areas on {side}')
            if ref['resolution'] != 'resolved':
                coverage_issue('dependencies', dependency_id, {'unsupportedReference', 'unresolvedReference', 'externalUnavailable'}, 'unresolved reference requires partial dependency coverage')

    def path_sides(start, end, path, label):
        possible = []
        for dependency_id in path:
            if dependency_id not in dependencies:
                fail(label, f'unknown dependencyId: {dependency_id}')
                return []
        for side in SIDES:
            if entities.get(start, {}).get(side) is None or entities.get(end, {}).get(side) is None:
                continue
            current, resolved = start, True
            for dependency_id in path:
                dependency = dependencies[dependency_id]
                ref = dependency[side]
                if ref is None or (current not in ref['members'] and not any(in_area(current, side, area) for area in ref.get('areas', []))):
                    break
                resolved = resolved and ref['resolution'] == 'resolved'
                current = dependency['dependentEntityId']
            else:
                if current == end:
                    possible.append((side, resolved))
        if not possible:
            fail(label, 'dependency path must connect the cause to the target within one snapshot')
        return possible

    for impact in graph['impacts']:
        label, entity_id, property = impact['id'], impact['entityId'], impact['property']
        entity = entities.get(entity_id)
        if entity is None:
            fail(label, f'unknown target entityId: {entity_id}')
        if entity_id not in graph['scope']['entityIds'] or impact['mechanism'] not in graph['scope']['mechanisms']:
            fail(label, 'impact is outside declared analysis scope')
        if impact['mechanism'] == 'formula' and entity:
            unresolved_formula = impact['evidence']['level'] != 'verified' and any(d['dependentEntityId'] == entity_id and any(d[side] is not None and d[side]['kind'] == 'dynamic' and d[side]['resolution'] == 'unresolved' for side in SIDES) for d in dependencies.values())
            if property != '/value' or not (unresolved_formula or any(entity[side] is not None and entity[side].get('formula') is not None for side in SIDES)):
                fail(label, 'formula impact requires /value and explicit formula context')
        targets = [change for change in document['changes'] if change['target']['entityId'] == entity_id and change['property'] == property and change['operation'] == 'update']
        if set(impact['changeIds']) != {change['id'] for change in targets}:
            fail(label, 'changeIds must name exactly the delivered target property updates')
        for change_id in impact['changeIds']:
            if change_id not in changes:
                fail(label, f'unknown target changeId: {change_id}')
        old, new = (observations.get((entity_id, side, property)) for side in SIDES)
        known = old is not None and new is not None and old['status'] == new['status'] == 'known'
        same = known and equivalent(old['value'], new['value'], property)
        outcome = impact['outcome']
        if targets and outcome != 'changed':
            fail(label, 'a confirmed property update requires changed outcome')
        if outcome == 'changed':
            if known and same:
                fail(label, 'changed outcome disagrees with equal observations')
            if not targets and not (known and not same and reliable(old) and reliable(new) and any(o['origin'] in {'recalculated', 'rendered'} for o in (old, new))):
                fail(label, 'changed outcome needs a core update or differing computed observations')
        elif outcome == 'unchanged' and not (same and reliable(old) and reliable(new)):
            fail(label, 'unchanged outcome requires equal reliable observations')
        for observation in (old, new):
            if targets and observation and observation['origin'] in {'recalculated', 'rendered'}:
                fail(label, 'derived observations cannot be substituted for original core Change values')
        if outcome == 'unknown' and known and reliable(old) and reliable(new):
            fail(label, 'known reliable observations cannot claim unknown outcome')

        evidence = impact['evidence']
        verified = evidence['level'] == 'verified'
        if evidence['level'] != 'verified' or outcome == 'unknown':
            coverage_issue('attribution', label, {'unverifiedCause', 'unavailableResult'}, 'unconfirmed attribution requires partial coverage')
        environment_causes = set()
        for cause in impact['causes']:
            if cause['kind'] == 'unresolved':
                if verified:
                    fail(label, 'verified causality cannot contain an unresolved cause')
            elif cause['kind'] == 'environment':
                environment_causes.add(cause['name'])
                if cause['before'] == cause['after']:
                    fail(label, 'environment cause must describe a change')
            else:
                if cause['kind'] == 'change':
                    change = changes.get(cause['changeId'])
                    if change is None:
                        fail(label, f'unknown cause changeId: {cause["changeId"]}')
                        continue
                    if cause['changeId'] in impact['changeIds']:
                        fail(label, 'a change cannot be its own cause')
                    start = change['target']['entityId']
                else:
                    dependency = dependencies.get(cause['dependencyId'])
                    if dependency is None:
                        fail(label, 'unknown changed dependency')
                        continue
                    def reference_signature(ref):
                        return None if ref is None else {key: value for key, value in ref.items() if key not in {'members', 'resolution'}}
                    if reference_signature(dependency['before']) == reference_signature(dependency['after']):
                        fail(label, 'dependency cause must describe a changed reference')
                    if verified and any(dependency[side] is not None and dependency[side]['resolution'] != 'resolved' for side in SIDES):
                        fail(label, 'verified reference change requires resolved references')
                    start = dependency['dependentEntityId']
                possible = path_sides(start, entity_id, cause['dependencyIds'], label)
                if verified and not any(resolved for _, resolved in possible):
                    fail(label, 'verified cause requires a fully resolved dependency path')
        proof_computations = []
        for computation_id in evidence['computationIds']:
            computation = computations.get(computation_id)
            if computation is None:
                fail(label, f'unknown evidence computationId: {computation_id}')
            else:
                proof_computations.append(computation)
        if verified:
            if outcome == 'unknown' or not reliable(old) or not reliable(new):
                fail(label, 'verified attribution requires reliable results on both sides')
            by_side = {}
            for computation in proof_computations:
                side = computation['side']
                by_side[side] = computation
                observation = observations.get((entity_id, side, property))
                if not observation or not any(output['entityId'] == entity_id and output['property'] == property and observation.get('value') is not None and equivalent(output['value'], observation['value'], property) for output in computation['outputs']):
                    fail(label, 'causal computation does not reproduce the assessed result')
            if len(proof_computations) != 2 or set(by_side) != set(SIDES):
                fail(label, 'verified attribution requires one computation for each snapshot')
            else:
                a, b = by_side['before'], by_side['after']
                if a['engine'] != b['engine'] and 'engine' not in environment_causes:
                    fail(label, 'engine drift requires an explicit environment cause')
                environment_a = {item['name']: item['fingerprint'] for item in a.get('environment', [])}
                environment_b = {item['name']: item['fingerprint'] for item in b.get('environment', [])}
                if (a['contextSha256'] != b['contextSha256'] or environment_a != environment_b) and 'context' not in environment_causes:
                    fail(label, 'context drift requires an explicit environment cause')

    for lineage in graph['lineages']:
        label = lineage['id']
        old, new = lineage['beforeEntityIds'], lineage['afterEntityIds']
        if set(old) & set(new):
            fail(label, 'split/merge must use distinct old and new identities')
        if lineage['kind'] == 'split' and not (len(old) == 1 and len(new) >= 2):
            fail(label, 'split requires one old entity and multiple new entities')
        if lineage['kind'] == 'merge' and not (len(old) >= 2 and len(new) == 1):
            fail(label, 'merge requires multiple old entities and one new entity')
        for side, ids in (('before', old), ('after', new)):
            for entity_id in ids:
                entity_at(entity_id, side, label)
                if entity_id not in graph['scope']['entityIds']:
                    fail(label, 'lineage is outside declared analysis scope')
                entity = entities.get(entity_id)
                if entity and entity['after' if side == 'before' else 'before'] is not None:
                    fail(label, 'split/merge identities must not survive on the opposite side')

    for category, items in (('dependencies', graph['dependencies']), ('attribution', graph['impacts']), ('computation', graph['computations']), ('structure', graph['lineages'])):
        if items and graph['coverage'][category]['status'] == 'notRequested':
            fail('coverage.' + category, 'notRequested cannot contain analysis records')
        for issue in graph['coverage'][category]['issues']:
            for entity_id in issue['entityIds']:
                if entity_id not in entities:
                    fail('coverage.' + category, f'unknown issue entityId: {entity_id}')
    omitted = graph['delivery']['omitted']
    if (graph['delivery']['status'] == 'complete') != (sum(omitted.values()) == 0):
        fail('delivery', 'delivery status disagrees with omitted counts')
    for section, category in (('entities', 'dependencies'), ('dependencies', 'dependencies'), ('computations', 'computation'), ('impacts', 'attribution'), ('lineages', 'structure')):
        if omitted[section]:
            coverage_issue(category, 'delivery', {'truncatedContent'}, 'omitted analysis requires matching partial coverage with truncatedContent')
    return errors

"""Validate ChangeSets without opening Office files.

Install requirements-change-protocol.txt, then run:
    python scripts/validate_change_protocol.py --self-test
An optional --dependency-dir supports an isolated jsonschema installation.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from functools import lru_cache
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from urllib.parse import unquote


ROOT = Path(__file__).resolve().parents[1]


def load_json(path: Path):
    def reject_constant(value):
        raise ValueError(f"Non-JSON numeric constant: {value}")

    return json.loads(path.read_text(encoding="utf-8"), parse_constant=reject_constant)


def equal_value(before, after, aspect, tolerance):
    if before["type"] != after["type"]:
        return False
    if before["type"] == "number":
        return Decimal(before["data"]) == Decimal(after["data"])
    if before["type"] == "asset":
        return before["data"]["sha256"] == after["data"]["sha256"]
    if aspect == "geometry":
        old, new = before["data"], after["data"]
        if old.keys() != new.keys():
            return False
        coordinates = {"x", "y", "width", "height"}
        for key in old:
            if key in coordinates:
                if old[key] is None or new[key] is None:
                    if old[key] != new[key]:
                        return False
                elif abs(old[key] - new[key]) > tolerance:
                    return False
            elif key == 'childFrame':
                if any(abs(old[key][field] - new[key][field]) > tolerance for field in coordinates):
                    return False
            elif old[key] != new[key]:
                return False
        return True
    return before["data"] == after["data"]


def source_errors(source, file_format, prefix):
    errors = []
    if file_format == 'pdf':
        locator = source['locator']
        if locator['type'] != 'pdf':
            return [f'{prefix}: PDF requires a PDF locator']
        if 'text' in locator:
            if not locator.get('itemIndices') or 'range' not in locator:
                errors.append(f'{prefix}: PDF text requires item indices and a UTF16 range')
            else:
                try:
                    length = len(locator['text'].encode('utf-16-le')) // 2
                except UnicodeEncodeError:
                    errors.append(f'{prefix}: PDF text contains an unpaired surrogate')
                else:
                    if locator['range'] != {'unit': 'utf16', 'start': 0, 'length': length}:
                        errors.append(f'{prefix}: PDF text and range disagree')
        return errors
    part = unquote(source["part"])
    if any(piece in {"", ".", ".."} for piece in part.split("/")[1:]):
        errors.append(f"{prefix}: part must be a canonical package part URI")
    locator = source["locator"]
    expected = {"docx": "word", "xlsx": "excel", "pptx": "powerpoint"}[file_format]
    if locator["type"] not in {expected, "xml"}:
        errors.append(f"{prefix}: {locator['type']} locator cannot anchor {file_format}")
    if locator["type"] == "excel" and "address" in locator:
        letters, row_string = re.fullmatch(r"([A-Z]+)([0-9]+)", locator["address"]).groups()
        column = 0
        for letter in letters:
            column = column * 26 + ord(letter) - ord("A") + 1
        row = int(row_string)
        if column > 16384 or row > 1048576:
            errors.append(f"{prefix}: Excel address exceeds XFD1048576")
        if "row" in locator and locator["row"] != row:
            errors.append(f"{prefix}: address and row disagree")
        if "column" in locator and locator["column"] != column:
            errors.append(f"{prefix}: address and column disagree")
    return errors


def semantic_errors(document):
    errors = []
    ids = set()
    target_kinds = {}
    properties = set()
    entity_operations = defaultdict(list)
    moves = defaultdict(list)
    aspects = set(document["comparison"]["aspects"])
    tolerance = document["comparison"].get("geometryTolerancePt" if document['format'] == 'pdf' else "geometryToleranceEmu", 0)

    for change in document["changes"]:
        label = change["id"]
        if label in ids:
            errors.append(f"{label}: duplicate Change ID")
        ids.add(label)
        target = change["target"]
        entity_id = target["entityId"]
        kind = target["kind"]
        previous_kind = target_kinds.setdefault(entity_id, kind)
        if previous_kind != kind:
            errors.append(f"{label}: entityId has inconsistent kinds")
        key = (entity_id, change["aspect"], change["property"])
        if key in properties:
            errors.append(f"{label}: duplicate entity/aspect/property")
        properties.add(key)
        entity_operations[entity_id].append(change["operation"])
        if change["aspect"] not in aspects:
            errors.append(f"{label}: aspect is outside the declared comparison scope")

        for side in ("before", "after"):
            state = change[side]
            if state is None:
                continue
            errors.extend(source_errors(state["source"], document["format"], f"{label}.{side}"))
            if state["value"]["type"] == "asset":
                errors.extend(source_errors(
                    {"part": state["value"]["data"]["part"], "locator": {"type": "xml", "elementPath": []}},
                    document["format"], f"{label}.{side}.asset",
                ))
            locator = state["source"]["locator"]
            if kind == "group" and change["aspect"] == "geometry":
                if "childFrame" not in state["value"]["data"]:
                    errors.append(f"{label}.{side}: group geometry requires childFrame")
            if kind == "textRange" and change["aspect"] == "text":
                if "range" not in locator:
                    errors.append(f"{label}.{side}: textRange needs an explicit source range")
                else:
                    try:
                        length = len(state["value"]["data"].encode("utf-16-le")) // 2
                    except UnicodeEncodeError:
                        errors.append(f"{label}.{side}: text contains an unpaired surrogate")
                    else:
                        if length != locator["range"]["length"]:
                            errors.append(f"{label}.{side}: text and UTF16 range lengths disagree")
            if change["aspect"] == "formula":
                value = state["value"]
                if value["type"] == "text" and (not value["data"] or value["data"].startswith("=")):
                    errors.append(f"{label}.{side}: formula must be nonempty OOXML text without leading =")

        if change["operation"] == "update" and equal_value(
            change["before"]["value"], change["after"]["value"], change["aspect"], tolerance
        ):
            errors.append(f"{label}: equal values must not produce update")

        if change["operation"] == "move":
            moves[entity_id].append(change["aspect"])
            old = change["before"]["value"]["data"]
            new = change["after"]["value"]["data"]
            same_parent = old["parentEntityId"] == new["parentEntityId"]
            if old["parentEntityId"] == entity_id or new["parentEntityId"] == entity_id:
                errors.append(f"{label}: an entity cannot be its own parent")
            if change["aspect"] == "order" and (not same_parent or old["index"] == new["index"]):
                errors.append(f"{label}: order move needs same parent and different index")
            if change["aspect"] == "parent" and same_parent:
                errors.append(f"{label}: parent move needs different parents")

    for entity_id, operations in entity_operations.items():
        if any(operation in {"add", "remove"} for operation in operations) and len(operations) > 1:
            errors.append(f"{entity_id}: whole-object add/remove cannot coexist with another event")
    for entity_id, move_aspects in moves.items():
        if "parent" in move_aspects and "order" in move_aspects:
            errors.append(f"{entity_id}: combine parent and order moves into one parent event")
    for index, issue in enumerate(document["coverage"]["issues"]):
        if "aspect" in issue and issue["aspect"] not in aspects:
            errors.append(f"coverage.issues[{index}]: aspect is outside comparison scope")
        if "source" in issue:
            if issue["side"] == "both":
                errors.append(f"coverage.issues[{index}]: source requires one explicit side")
            errors.extend(source_errors(issue["source"], document["format"], f"coverage.issues[{index}]"))
    return errors


@lru_cache(maxsize=1)
def character_validator():
    from jsonschema import Draft202012Validator
    schema = load_json(ROOT / 'schemas' / 'office-text-diff.v1.schema.json')
    Draft202012Validator.check_schema(schema)
    return Draft202012Validator(schema)


def character_errors(change):
    diff = change.get('extensions', {}).get('ppt2wiki.textDiff')
    if diff is None:
        return []
    prefix = change['id'] + '.extensions.ppt2wiki.textDiff'
    errors = [f'{prefix} {issue.json_path}: {issue.message}' for issue in character_validator().iter_errors(diff)]
    if errors:
        return errors
    if change['operation'] != 'update' or change['aspect'] not in {'text', 'value', 'formula'}:
        return [f'{prefix}: character ranges require a text/value/formula update']
    old = change['before']['value']['data']
    new = change['after']['value']['data']
    if not isinstance(old, str) or not isinstance(new, str):
        return [f'{prefix}: both compared values must have string representations']
    try:
        old_bytes, new_bytes = old.encode('utf-16-le'), new.encode('utf-16-le')
        if diff['beforeLength'] != len(old_bytes) // 2 or diff['afterLength'] != len(new_bytes) // 2:
            return [f'{prefix}: declared UTF16 text lengths disagree']
        old_cursor = new_cursor = 0
        for edit in diff['edits']:
            before, after = edit['before'], edit['after']
            for span, data, cursor in ((before, old_bytes, old_cursor), (after, new_bytes, new_cursor)):
                start, length = span['start'], span['length']
                if start < cursor or start + length > len(data) // 2:
                    return [f'{prefix}: character ranges overlap or exceed the text']
                for boundary in (start, start + length):
                    data[:boundary * 2].decode('utf-16-le')
                if data[start * 2:(start + length) * 2].decode('utf-16-le') != span['text']:
                    return [f'{prefix}: character range text disagrees with source']
            if old_bytes[old_cursor * 2:before['start'] * 2] != new_bytes[new_cursor * 2:after['start'] * 2]:
                return [f'{prefix}: unchanged text between edits differs']
            if before['text'] == after['text']:
                return [f'{prefix}: an edit must change characters']
            old_cursor = before['start'] + before['length']
            new_cursor = after['start'] + after['length']
        if old_bytes[old_cursor * 2:] != new_bytes[new_cursor * 2:]:
            return [f'{prefix}: character edits do not reconstruct the new text']
    except UnicodeError:
        return [f'{prefix}: character range splits an UTF16 surrogate pair']
    return []


def validation_errors(document, validator):
    errors = [
        f"schema {error.json_path}: {error.message}"
        for error in sorted(validator.iter_errors(document), key=lambda error: error.json_path)
    ]
    if errors:
        return errors
    from validate_impact import impact_errors
    misplaced = [f'{change["id"]}: ppt2wiki.impact belongs on ChangeSet.extensions'
                 for change in document['changes'] if 'ppt2wiki.impact' in change.get('extensions', {})]
    return (semantic_errors(document) + misplaced
            + [error for change in document['changes'] for error in character_errors(change)]
            + impact_errors(document, equal_value, source_errors))


def run_self_tests(examples, validator):
    word, excel, ppt = (examples[name] for name in ("word", "excel", "powerpoint"))
    cases = []

    def case(name, base, edits):
        value = copy.deepcopy(base)
        for path, replacement in edits:
            container = value
            for key in path[:-1]:
                container = container[key]
            container[path[-1]] = replacement
        cases.append((name, value))

    case("add requires absent before", word, [(('changes', 2, 'before'), word['changes'][0]['before'])])
    case("remove requires absent after", ppt, [(('changes', 3, 'after'), ppt['changes'][0]['after'])])
    case("update requires present before", excel, [(('changes', 0, 'before'), None)])
    missing_match = copy.deepcopy(word)
    del missing_match['changes'][0]['match']
    cases.append(("update requires match evidence", missing_match))
    case("geometry uses update", ppt, [(('changes', 1, 'operation'), 'move')])
    case("style requires a style property", excel, [(('changes', 1, 'property'), '/value')])
    case("decimal numbers travel as strings", excel, [(('changes', 0, 'before', 'value', 'data'), 1200000)])
    case("Change IDs are unique", word, [(('changes', 1, 'id'), 'word-001')])
    case("changes stay within declared scope", word, [(('comparison', 'aspects'), ['existence'])])
    case("format and locator agree", excel, [(('format',), 'docx')])
    case("partial coverage needs an issue", ppt, [(('coverage', 'issues'), [])])
    case("complete coverage has no issues", ppt, [(('coverage', 'status'), 'complete')])
    case("coverage source needs one side", ppt, [(('coverage', 'issues', 0, 'source'), ppt['changes'][0]['before']['source'])])
    case("Excel column boundary", excel, [(('changes', 0, 'after', 'source', 'locator', 'address'), 'XFE1')])
    case("Excel row boundary", excel, [(('changes', 0, 'after', 'source', 'locator', 'address'), 'C1048577')])
    case("address and row agree", excel, [(('changes', 0, 'after', 'source', 'locator', 'row'), 99)])
    case("decimal spelling is not a value change", excel, [(('changes', 0, 'after', 'value', 'data'), '1200000.0')])
    case("resource rename is not a media change", ppt, [(('changes', 5, 'after', 'value', 'data', 'sha256'), 'c' * 64)])
    duplicate_property = copy.deepcopy(word['changes'][0])
    duplicate_property['id'] = 'word-004'
    case("one event per entity aspect property", word, [(('changes',), word['changes'] + [duplicate_property])])
    case("matched entity keeps kind", word, [(('changes', 1, 'target', 'kind'), 'paragraph')])
    case("whole-object add excludes property changes", word, [
        (('changes', 2, 'target', 'entityId'), 'budget-text'),
        (('changes', 2, 'target', 'kind'), 'textRange'),
    ])
    case("order move needs a new position", ppt, [(('changes', 2, 'after', 'value', 'data', 'index'), 2)])
    case("order move keeps parent", ppt, [(('changes', 2, 'after', 'value', 'data', 'parentEntityId'), 'other-deck')])
    case("parent move changes parent", ppt, [
        (('comparison', 'aspects'), ppt['comparison']['aspects'] + ['parent']),
        (('changes', 2, 'aspect'), 'parent'),
    ])
    case("JSON Pointer escape is valid", excel, [(('changes', 1, 'property'), '/style/a~2b')])
    case("geometry respects configured tolerance", ppt, [(('comparison', 'geometryToleranceEmu'), 914400)])
    case("range uses UTF16 length", word, [(('changes', 0, 'before', 'source', 'locator', 'range', 'length'), 3)])
    case("part has no dot segments", word, [(('changes', 0, 'before', 'source', 'part'), '/word/../document.xml')])
    case("formula excludes leading equals sign", excel, [(('changes', 2, 'after', 'value', 'data'), '=SUM(B4:D4)')])
    case("null value is distinct from missing state", excel, [(('changes', 3, 'before'), None)])
    case("confidence stays within range", word, [(('changes', 0, 'match', 'confidence'), 1.1)])
    case("inline objects do not invent positions", ppt, [(('changes', 1, 'before', 'value', 'data', 'space'), 'inline')])
    case("groups preserve child coordinate frames", ppt, [(('changes', 1, 'target', 'kind'), 'group')])

    for name, document in cases:
        if not validation_errors(document, validator):
            raise AssertionError(f"Invalid case was accepted: {name}")

    positive_cases = []
    empty = copy.deepcopy(word)
    empty['changes'] = []
    positive_cases.append(("empty complete comparison", empty))
    emoji = copy.deepcopy(word)
    emoji['changes'][0]['before']['value']['data'] = '120😀'
    emoji['changes'][0]['before']['source']['locator']['range']['length'] = 5
    positive_cases.append(("non-BMP text uses two UTF16 units", emoji))
    formula_removed = copy.deepcopy(excel)
    formula_removed['changes'][2]['after']['value'] = {'type': 'null', 'data': None}
    positive_cases.append(("formula removal preserves cell", formula_removed))
    inline = copy.deepcopy(ppt)
    for side in ('before', 'after'):
        inline['changes'][1][side]['value']['data'].update({'space': 'inline', 'x': None, 'y': None})
    inline['changes'][1]['after']['value']['data']['width'] += 914400
    positive_cases.append(("inline dimensions need no absolute position", inline))
    group = copy.deepcopy(ppt)
    group['changes'][1]['target']['kind'] = 'group'
    for side in ('before', 'after'):
        group['changes'][1][side]['value']['data']['childFrame'] = {'x': 0, 'y': 0, 'width': 3657600, 'height': 914400}
    group['changes'][1]['after']['value']['data']['x'] = 914400
    group['changes'][1]['after']['value']['data']['childFrame']['width'] += 914400
    positive_cases.append(("child frame alone changes geometry", group))
    for name, document in positive_cases:
        errors = validation_errors(document, validator)
        if errors:
            raise AssertionError(f"Valid case was rejected: {name}: {errors}")
    print(f"PASS: {len(cases)} invalid cases rejected; {len(positive_cases)} valid edge cases accepted")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('files', nargs='*', type=Path)
    parser.add_argument('--dependency-dir', type=Path)
    parser.add_argument('--self-test', action='store_true')
    args = parser.parse_args()
    if args.dependency_dir:
        sys.path.insert(0, str(args.dependency_dir.resolve()))
    try:
        from jsonschema import Draft202012Validator
    except ImportError:
        parser.error('Install scripts/requirements-change-protocol.txt or provide --dependency-dir')

    schema = load_json(ROOT / 'schemas' / 'office-change.v1.schema.json')
    Draft202012Validator.check_schema(schema)
    validator = Draft202012Validator(schema)
    pdf_schema = load_json(ROOT / 'schemas' / 'pdf-change.v1.schema.json')
    Draft202012Validator.check_schema(pdf_schema)
    pdf_validator = Draft202012Validator(pdf_schema)
    example_paths = sorted((ROOT / 'examples').glob('*.changes.json'))
    paths = args.files or example_paths
    failed = False
    event_count = 0
    for path in paths:
        document = load_json(path)
        errors = validation_errors(document, pdf_validator if document.get('format') == 'pdf' else validator)
        if errors:
            failed = True
            print(f"FAIL: {path.name}")
            for error in errors:
                print(f"  {error}")
        else:
            event_count += len(document['changes'])
            print(f"PASS: {path.name} ({len(document['changes'])} changes)")
    if failed:
        return 1
    if args.self_test:
        examples = {path.name.split('.')[0]: load_json(path) for path in example_paths}
        run_self_tests(examples, validator)
        from test_impact_protocol import run_impact_self_tests
        run_impact_self_tests(validator)
    print(f"Validated {len(paths)} ChangeSets and {event_count} changes. Source file anchors were not resolved.")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

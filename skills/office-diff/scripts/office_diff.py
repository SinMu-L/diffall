"""Office diff for browser Python: standard-library IO + the original JS/WASM engine.

Use top-level await with compare_files()/self_test() in a Pyodide tool.
The input files are only read. No network calls, subprocesses, or package installs.
"""

import base64
from collections import Counter
import hashlib
import html
import io
import json
from pathlib import Path
import zipfile


SKILL_ROOT = Path(__file__).resolve().parent.parent
MAX_FILE_BYTES = 15 * 1024 * 1024
_bridge = None


def _json(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


async def _load_bridge():
    global _bridge
    if _bridge is not None:
        return _bridge
    try:
        import js
    except ImportError as error:
        raise RuntimeError(
            "此技能需要浏览器 Pyodide 的 js 桥接；纯 CPython 不能运行原 JS/WASM 引擎。"
        ) from error
    bundle_path = SKILL_ROOT / "assets" / "office-diff.bundle.js"
    manifest_path = SKILL_ROOT / "assets" / "engine-manifest.json"
    if not bundle_path.is_file() or not manifest_path.is_file():
        raise RuntimeError("技能资源未挂载：请导入整个 office-diff 文件夹并保留 assets/。")
    bundle = bundle_path.read_bytes()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if hashlib.sha256(bundle).hexdigest() != manifest["bundleSha256"]:
        raise RuntimeError("技能 bundle 与清单摘要不一致，请重新导入完整技能包。")
    try:
        candidate = js.eval(bundle.decode("utf-8"))
        await candidate.selfCheck()
    except Exception as error:
        raise RuntimeError("JS/WASM 执行自检失败：" + str(error)) from error
    _bridge = candidate
    return _bridge


def _check_inputs(before, after, before_name, after_name):
    formats = [Path(str(name)).suffix.lower() for name in (before_name, after_name)]
    if formats[0] not in (".docx", ".xlsx", ".pptx") or formats[0] != formats[1]:
        raise ValueError("需要同格式的 DOCX、XLSX 或 PPTX 两份文件；此包不支持 PDF。")
    if any(len(value) > MAX_FILE_BYTES for value in (before, after)):
        raise ValueError("单文件超过 15 MB，请缩小范围后比较。")


async def compare_bytes(before, after, before_name, after_name):
    """Compare two byte snapshots; return the complete, untrimmed ChangeSet."""
    _check_inputs(before, after, before_name, after_name)
    bridge = await _load_bridge()
    result = json.loads(await bridge.compareBase64(
        base64.b64encode(before).decode("ascii"),
        base64.b64encode(after).decode("ascii"),
        str(before_name), str(after_name),
    ))
    for side, content in (("before", before), ("after", after)):
        if result["documents"][side]["sha256"] != hashlib.sha256(content).hexdigest():
            raise RuntimeError("对比结果与输入快照摘要不一致。")
    return result


def _compact(value, string_limit):
    if isinstance(value, str) and len(value) > string_limit:
        return {"preview": value[:string_limit], "originalChars": len(value), "truncated": True}
    if isinstance(value, dict):
        return {key: _compact(item, string_limit) for key, item in value.items()}
    if isinstance(value, list):
        if len(value) > 32:
            return {
                "previewItems": [_compact(item, string_limit) for item in value[:32]],
                "originalItems": len(value), "truncated": True,
            }
        return [_compact(item, string_limit) for item in value]
    return value


def _has_truncation(value):
    if isinstance(value, dict):
        return value.get("truncated") is True or any(_has_truncation(v) for v in value.values())
    return isinstance(value, list) and any(_has_truncation(v) for v in value)


def _brief(change, string_limit=600):
    result = {key: change[key] for key in ("id", "operation", "aspect", "target", "property")}
    result["target"] = _compact(result["target"], string_limit)
    for side in ("before", "after"):
        state = change[side]
        result[side] = None if state is None else {
            "value": _compact(state["value"], string_limit), "source": state["source"],
        }
    if "match" in change:
        result["match"] = {
            "method": change["match"]["method"], "confidence": change["match"]["confidence"],
        }
    text_diff = change.get("extensions", {}).get("ppt2wiki.textDiff")
    if text_diff:
        result["characterDiff"] = {
            "unit": text_diff["unit"], "editCount": len(text_diff["edits"]),
            "detailOmitted": True,
        }
    return result


def _build_summary(change_set, *, offset=0, limit=12, max_chars=14000,
                   aspect=None, operation=None, change_ids=None, artifacts=None):
    if isinstance(offset, bool) or not isinstance(offset, int) or offset < 0:
        raise ValueError("offset 必须是非负整数。")
    if isinstance(limit, bool) or not isinstance(limit, int) or not 1 <= limit <= 100:
        raise ValueError("limit 必须在 1–100 之间。")
    if isinstance(max_chars, bool) or not isinstance(max_chars, int) or not 4000 <= max_chars <= 60000:
        raise ValueError("max_chars 必须在 4000–60000 之间。")
    if change_ids is not None and (isinstance(change_ids, str) or not isinstance(change_ids, (list, tuple, set))):
        raise ValueError("change_ids 应是变化 ID 的列表。")
    wanted = set(change_ids) if change_ids is not None else None
    all_changes = change_set["changes"]
    selected = [change for change in all_changes
                if (aspect is None or change["aspect"] == aspect)
                and (operation is None or change["operation"] == operation)
                and (wanted is None or change["id"] in wanted)]
    if offset > len(selected):
        raise ValueError("offset 超过筛选后的变化数。")
    summary = {
        "comparisonId": change_set["id"], "format": change_set["format"],
        "documents": change_set["documents"], "coverage": change_set["coverage"],
        "totalChanges": len(all_changes), "matchedChanges": len(selected),
        "counts": {
            "operation": dict(Counter(change["operation"] for change in all_changes)),
            "aspect": dict(Counter(change["aspect"] for change in all_changes)),
        },
        "offset": offset, "changes": [], "returnedChanges": 0,
        "hasMore": offset < len(selected), "nextOffset": None,
        "truncated": bool(selected),
        "interpretation": "差异只覆盖 coverage 声明的范围；缓存值不是公式重算证据。",
    }
    if artifacts is not None:
        summary["artifacts"] = artifacts
    impact = change_set.get("extensions", {}).get("ppt2wiki.impact")
    if impact:
        records = impact.get("impacts", [])
        summary["impactAnalysis"] = {
            "records": len(records), "detailOmitted": True,
            "coverage": impact.get("coverage"),
        }
    for change in selected[offset:offset + limit]:
        candidate = _brief(change)
        while True:
            trial = dict(summary, changes=summary["changes"] + [candidate])
            # Leave room for pagination fields and their integer representations.
            if len(_json(trial)) + 128 <= max_chars:
                summary["changes"].append(candidate)
                break
            if summary["changes"]:
                candidate = None
                break
            current_limit = candidate.pop("_stringLimit", 600)
            if current_limit <= 32:
                raise ValueError("单条变化的必要元数据超过摘要预算，请增加 max_chars。")
            next_limit = max(32, current_limit // 2)
            candidate = _brief(change, next_limit)
            candidate["_stringLimit"] = next_limit
        if candidate is None:
            break
    for change in summary["changes"]:
        change.pop("_stringLimit", None)
    summary["returnedChanges"] = len(summary["changes"])
    next_offset = offset + summary["returnedChanges"]
    summary["hasMore"] = next_offset < len(selected)
    summary["nextOffset"] = next_offset if summary["hasMore"] else None
    summary["truncated"] = (
        offset > 0 or summary["hasMore"] or _has_truncation(summary["changes"])
        or any(change.get("characterDiff", {}).get("detailOmitted") for change in summary["changes"])
        or bool(impact)
    )
    if len(_json(summary)) > max_chars:
        raise ValueError("必要元数据超过摘要预算，请增加 max_chars。")
    return summary


def _read_changes(path):
    result = json.loads(Path(path).read_text(encoding="utf-8"))
    if result.get("schemaVersion") != "1.0.0" or result.get("format") not in ("docx", "xlsx", "pptx"):
        raise ValueError("不是此技能支持的 Office Change v1 文件。")
    return result


def query_changes(changes_path, *, offset=0, limit=12, max_chars=14000,
                  aspect=None, operation=None, change_ids=None):
    return _build_summary(
        _read_changes(changes_path), offset=offset, limit=limit, max_chars=max_chars,
        aspect=aspect, operation=operation, change_ids=change_ids,
        artifacts={"changes": str(changes_path)},
    )


def get_change(changes_path, change_id):
    for change in _read_changes(changes_path)["changes"]:
        if change["id"] == change_id:
            return change
    raise KeyError("找不到变化 ID：" + str(change_id))


def _format_value(state):
    if state is None:
        return "（对象不存在）"
    value = state["value"]
    data = value.get("data")
    if isinstance(data, str):
        return data
    return json.dumps(data, ensure_ascii=False, indent=2)


def render_report(change_set):
    """Read-only standalone HTML. All document text is escaped; no scripts/URLs."""
    esc = lambda value: html.escape(str(value), quote=True)
    documents = change_set["documents"]
    cards = []
    for change in change_set["changes"]:
        location = change["target"]["label"]
        details = {key: change[key] for key in ("before", "after", "match", "extensions") if key in change}
        cards.append(
            '<article><header><b>' + esc(change["id"]) + '</b><span>'
            + esc(change["operation"] + " · " + change["aspect"]) + '</span></header>'
            + '<h2>' + esc(location) + '</h2><div class="values"><section><h3>修改前</h3><pre>'
            + esc(_format_value(change["before"])) + '</pre></section><section><h3>修改后</h3><pre>'
            + esc(_format_value(change["after"])) + '</pre></section></div>'
            + '<details><summary>定位、匹配与字符差异</summary><pre>'
            + esc(json.dumps(details, ensure_ascii=False, indent=2)) + '</pre></details></article>'
        )
    coverage = esc(json.dumps(change_set["coverage"], ensure_ascii=False, indent=2))
    impact = change_set.get("extensions", {}).get("ppt2wiki.impact")
    impact_html = '' if not impact else (
        '<details><summary>公式静态依赖与潜在影响（未重算）</summary><pre>'
        + esc(json.dumps(impact, ensure_ascii=False, indent=2)) + '</pre></details>'
    )
    return '''<!doctype html><html lang="zh-CN"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'">
<title>Office 文件对比</title><style>
*{box-sizing:border-box}body{font:15px/1.65 system-ui,"Microsoft YaHei",sans-serif;background:#f4f6f8;color:#233044;margin:0;padding:28px}main{max-width:1100px;margin:auto}h1{font-size:28px;margin:0 0 8px}h2{font-size:17px;overflow-wrap:anywhere}h3{font-size:13px;color:#627386;margin:0 0 8px}.note,article{background:white;border:1px solid #dbe3ea;border-radius:12px;padding:20px;margin:18px 0}header{display:flex;justify-content:space-between;gap:12px;color:#4b6580}.values{display:grid;grid-template-columns:1fr 1fr;gap:16px}.values section{min-width:0;background:#f7f9fb;padding:14px;border-radius:8px}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:13px/1.6 ui-monospace,monospace}small{color:#617184;overflow-wrap:anywhere}summary{cursor:pointer;color:#35627f;margin-top:12px}@media(max-width:650px){body{padding:12px}.values{grid-template-columns:1fr}}@media print{body{background:white}article{break-inside:avoid}}
</style><main><h1>Office 文件对比</h1><p>''' + esc(documents["before"]["name"]) + ' → ' + esc(documents["after"]["name"]) + '''</p>
<div class="note"><b>''' + str(len(change_set["changes"])) + ''' 处已识别变化</b><p>这是结构化差异审阅报告。部分覆盖不代表完整文档等价；未验证 Office 排版、公式重算或全部外观。</p><small>修改前 SHA-256：''' + esc(documents["before"]["sha256"]) + '<br>修改后 SHA-256：' + esc(documents["after"]["sha256"]) + '</small><details><summary>比较覆盖说明</summary><pre>' + coverage + '</pre></details></div>' + (''.join(cards) or '<p>已覆盖范围没有识别到变化。</p>') + impact_html + '</main></html>'


async def compare_files(before, after, output_dir, *, limit=12, max_chars=14000):
    before_path, after_path = Path(before).resolve(), Path(after).resolve()
    destination = Path(output_dir).resolve()
    if destination in (before_path, after_path) or destination.exists():
        raise FileExistsError("输出目录必须是本次新建的目录，不能覆盖输入或已有结果。")
    for file in (before_path, after_path):
        if not file.is_file():
            raise FileNotFoundError("找不到输入文件：" + str(file))
        if file.stat().st_size > MAX_FILE_BYTES:
            raise ValueError("单文件超过 15 MB：" + str(file))
    change_set = await compare_bytes(
        before_path.read_bytes(), after_path.read_bytes(), before_path.name, after_path.name,
    )
    paths = {"changes": str(destination / "changes.json"),
             "report": str(destination / "review.html"),
             "summary": str(destination / "ai-summary.json")}
    summary = _build_summary(change_set, limit=limit, max_chars=max_chars, artifacts=paths)
    full_json = json.dumps(change_set, ensure_ascii=False, indent=2) + "\n"
    report = render_report(change_set)
    destination.mkdir(parents=True, exist_ok=False)
    for name, content in (("changes", full_json), ("report", report), ("summary", _json(summary) + "\n")):
        with Path(paths[name]).open("x", encoding="utf-8") as stream:
            stream.write(content)
    return summary


def _fixture_files(format, revised):
    """Small synthetic files used only by self_test; no user files are accessed."""
    rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    office_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    namespaces = {
        "docx": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xlsx": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "pptx": "http://schemas.openxmlformats.org/presentationml/2006/main",
    }
    main = {"docx": "word/document.xml", "xlsx": "xl/workbook.xml", "pptx": "ppt/presentation.xml"}[format]
    amount = "150" if revised else "120"
    files = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="' + rel + '"><Relationship Id="rId1" Type="' + office_rel + '/officeDocument" Target="' + main + '"/></Relationships>',
    }
    if format == "docx":
        files[main] = '<w:document xmlns:w="' + namespaces[format] + '" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>预算' + amount + '万</w:t></w:r></w:p></w:body></w:document>'
    elif format == "xlsx":
        files[main] = '<workbook xmlns="' + namespaces[format] + '" xmlns:r="' + office_rel + '"><sheets><sheet name="预算" sheetId="1" r:id="rId1"/></sheets></workbook>'
        files["xl/_rels/workbook.xml.rels"] = '<Relationships xmlns="' + rel + '"><Relationship Id="rId1" Type="' + office_rel + '/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        files["xl/worksheets/sheet1.xml"] = '<worksheet xmlns="' + namespaces[format] + '"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>部门</t></is></c><c r="B1" t="inlineStr"><is><t>预算</t></is></c><c r="C1" t="inlineStr"><is><t>合计</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>研发</t></is></c><c r="B2"><v>' + amount + '</v></c><c r="C2"><f>SUM(B2:B2)</f><v>' + amount + '</v></c></row></sheetData></worksheet>'
    else:
        drawing = "http://schemas.openxmlformats.org/drawingml/2006/main"
        files[main] = '<p:presentation xmlns:p="' + namespaces[format] + '" xmlns:r="' + office_rel + '"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>'
        files["ppt/_rels/presentation.xml.rels"] = '<Relationships xmlns="' + rel + '"><Relationship Id="rId1" Type="' + office_rel + '/slide" Target="slides/slide1.xml"/></Relationships>'
        files["ppt/slides/slide1.xml"] = '<p:sld xmlns:p="' + namespaces[format] + '" xmlns:a="' + drawing + '"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="预算文本"/></p:nvSpPr><p:spPr/><p:txBody><a:p><a:r><a:t>预算' + amount + '万</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>'
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in files.items():
            archive.writestr(name, text.encode("utf-8"))
    return output.getvalue()


async def self_test():
    bridge = await _load_bridge()
    environment = json.loads(await bridge.selfCheck())
    cases = []
    for format, expected in (("docx", 1), ("xlsx", 2), ("pptx", 1)):
        before, after = _fixture_files(format, False), _fixture_files(format, True)
        changed = await compare_bytes(before, after, "before." + format, "after." + format)
        unchanged = await compare_bytes(before, before, "before." + format, "same." + format)
        if len(changed["changes"]) != expected or unchanged["changes"]:
            raise RuntimeError("内置 " + format + " 对比自检失败。")
        cases.append({"format": format, "changes": expected, "sameFileChanges": 0, "ok": True})
    return {"ok": True, "environment": environment, "cases": cases}

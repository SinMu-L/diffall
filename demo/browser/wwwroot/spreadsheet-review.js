/** A bounded window over either the full sheet or changed coordinates and their context. */
export function sheetAxis({max, changed = [], expanded = [], mode = 'changes', page = 0, pageSize = 80, target}) {
  max = Math.max(1, max);
  let candidates;
  if (mode === 'changes') {
    const visible = new Set([1]);
    for (const index of changed) for (let offset = -1; offset <= 1; offset++) {
      if (index + offset >= 1 && index + offset <= max) visible.add(index + offset);
    }
    for (const [start, end] of expanded) for (let index = start; index <= Math.min(end, max); index++) visible.add(index);
    // An unchanged sheet still offers a useful initial preview.
    if (!changed.length) for (let index = 1; index <= Math.min(max, 8); index++) visible.add(index);
    candidates = [...visible].sort((a, b) => a - b);
  }
  const total = candidates?.length ?? max;
  const pages = Math.max(1, Math.ceil(total / pageSize));
  if (target >= 1 && target <= max) {
    const index = candidates ? candidates.indexOf(target) : target - 1;
    if (index >= 0) page = Math.floor(index / pageSize);
  }
  page = Math.max(0, Math.min(page, pages - 1));
  const start = page * pageSize;
  const values = candidates ? candidates.slice(start, start + pageSize) : Array.from({length: Math.min(pageSize, total - start)}, (_, i) => start + i + 1);
  const entries = [];
  let previous = page === 0 ? 0 : values[0] - 1;
  for (const value of values) {
    if (value > previous + 1) entries.push({start: previous + 1, end: value - 1});
    entries.push(value);
    previous = value;
  }
  if (page === pages - 1 && previous < max) entries.push({start: previous + 1, end: max});
  return {values, entries, page, pages, total, hidden: max - total, start: start + 1, end: start + values.length};
}

export function columnName(column) {
  let name = '';
  while (column > 0) { column--; name = String.fromCharCode(65 + column % 26) + name; column = Math.floor(column / 26); }
  return name;
}

export function cellPosition(source) {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(source?.locator?.address ?? '');
  if (!match) return null;
  let column = 0;
  for (const letter of match[1]) column = column * 26 + letter.charCodeAt(0) - 64;
  return {row: Number(match[2]), column};
}

export const sourceKey = source => source ? JSON.stringify(source) : '';

/** Index once per comparison; never search every change for every visible cell. */
export function indexReview(result) {
  const changes = {before: new Map(), after: new Map()};
  const sheets = {before: new Map(), after: new Map()};
  const byId = new Map(result.changeSet.changes.map(change => [change.id, change]));
  for (const side of ['before', 'after']) {
    for (const change of result.changeSet.changes) {
      if (!change[side]) continue;
      const key = sourceKey(change[side].source);
      if (!changes[side].has(key)) changes[side].set(key, []);
      changes[side].get(key).push(change);
    }
    for (const node of result[side].nodes) {
      if (!sheets[side].has(node.pageKey)) sheets[side].set(node.pageKey, {cells: new Map(), maxRow: 1, maxColumn: 1});
      const sheet = sheets[side].get(node.pageKey);
      sheet.cells.set(`${node.row}:${node.column}`, node);
      sheet.maxRow = Math.max(sheet.maxRow, node.row || 1);
      sheet.maxColumn = Math.max(sheet.maxColumn, node.column || 1);
    }
  }
  return {changes, sheets, byId};
}

/** Center only this scroll container, keeping the surrounding page and detail list still. */
export function revealIn(container, target, behavior = 'auto') {
  if (!container || !target) return;
  const box = container.getBoundingClientRect(), item = target.getBoundingClientRect();
  container.scrollTo({
    top: Math.max(0, container.scrollTop + item.top - box.top - (container.clientHeight - item.height) / 2),
    left: Math.max(0, container.scrollLeft + item.left - box.left - (container.clientWidth - item.width) / 2),
    behavior
  });
}

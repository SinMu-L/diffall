export const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, character => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));

export const characterDiff = change => change?.before?.value.type === 'text' && change?.after?.value.type === 'text' ? change?.extensions?.['ppt2wiki.textDiff'] : undefined;

/** Ranges refer to the engine's UTF16 projection, even when a numeric display adds commas. */
export function renderCharacterText(display, diff, side, {rawText = display, enabled = true} = {}) {
  if (!diff || !enabled) return escapeHtml(display);
  let text = String(display), raw = String(rawText);
  let positions;
  if (text === raw) positions = Array.from({length: raw.length + 1}, (_, index) => index);
  else if (text.replaceAll(',', '') === raw) {
    positions = [];
    for (let index = 0; index < text.length; index++) if (text[index] !== ',') positions.push(index);
    positions.push(text.length);
  } else {
    // Retain all significant digits when the formatted preview would round or rewrite them.
    text = raw;
    positions = Array.from({length: raw.length + 1}, (_, index) => index);
  }
  let cursor = 0, output = '';
  for (const edit of diff.edits) {
    const span = edit[side];
    if (!span || span.start < 0 || span.length < 0 || span.start + span.length > raw.length || raw.slice(span.start, span.start + span.length) !== span.text) throw new Error('字符差异范围与预览文本不一致。');
    const start = positions[span.start];
    const end = span.length ? positions[span.start + span.length - 1] + 1 : start;
    output += escapeHtml(text.slice(cursor, start));
    if (span.length) {
      const tag = side === 'before' ? 'del' : 'ins';
      output += `<${tag} class="char-diff ${side}" data-start="${span.start}" data-length="${span.length}">${escapeHtml(text.slice(start, end))}</${tag}>`;
    } else output += `<span class="char-caret ${side}" aria-label="${side === 'before' ? '此处新增字符' : '此处删除字符'}"></span>`;
    cursor = end;
  }
  return output + escapeHtml(text.slice(cursor));
}

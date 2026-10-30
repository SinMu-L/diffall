/** Add decimal grouping without converting the exact source value to Number. */
export function formatNumberText(value) {
  const text = String(value);
  const decimal = /^(-?)(\d+)(\.\d+)?$/.exec(text);
  if (!decimal) return text;
  return decimal[1] + decimal[2].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (decimal[3] ?? '');
}

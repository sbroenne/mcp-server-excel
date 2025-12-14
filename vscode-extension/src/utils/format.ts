export function basename(p: string) {
  const ix = Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/'));
  return ix >= 0 ? p.substring(ix + 1) : p;
}

export function truncate(s: string, max = 80) {
  if (s.length <= max) return s;
  return s.substring(0, Math.max(0, max - 1)) + '…';
}

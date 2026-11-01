// Échappement HTML et Markdown minimal (titres, paragraphes, listes, gras, italique, code, liens)
// pour la revue de l'expert intégrée au rapport.
export const e = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export function enLigne(s) {
  return e(s)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[\s(])\*([^*\s][^*]*)\*/g, '$1<em>$2</em>')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '<a href="$2">$1</a>');
}

export function versHtml(md) {
  const out = [];
  let liste = null;
  let para = [];
  const fermer = () => {
    if (para.length) { out.push(`<p>${enLigne(para.join(' '))}</p>`); para = []; }
    if (liste) { out.push(`</${liste}>`); liste = null; }
  };
  for (const brut of String(md ?? '').split(/\r?\n/)) {
    const l = brut.trimEnd();
    const t = /^(#{2,4})\s+(.*)$/.exec(l);
    const u = /^\s*[-*]\s+(.*)$/.exec(l);
    const o = /^\s*\d+[.)]\s+(.*)$/.exec(l);
    if (!l.trim()) { fermer(); continue; }
    if (t) { fermer(); out.push(`<h${t[1].length + 1}>${enLigne(t[2])}</h${t[1].length + 1}>`); continue; }
    if (u || o) {
      if (para.length) { out.push(`<p>${enLigne(para.join(' '))}</p>`); para = []; }
      const type = u ? 'ul' : 'ol';
      if (liste !== type) { if (liste) out.push(`</${liste}>`); out.push(`<${type}>`); liste = type; }
      out.push(`<li>${enLigne((u ?? o)[1])}</li>`);
      continue;
    }
    if (liste) { out.push(`</${liste}>`); liste = null; }
    para.push(l.trim());
  }
  fermer();
  return out.join('\n');
}

// Découpe une revue Markdown en sections « ## Titre ».
export function sections(md) {
  const s = {};
  let cle = '_';
  for (const l of String(md ?? '').split(/\r?\n/)) {
    const t = /^##\s+(.*)$/.exec(l);
    if (t) { cle = t[1].trim().toLowerCase(); s[cle] = ''; continue; }
    s[cle] = (s[cle] ?? '') + l + '\n';
  }
  return s;
}

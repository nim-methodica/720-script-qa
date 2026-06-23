#!/usr/bin/env node
/*
 * md-to-html.js — ממיר report.md ל-report.html מעוצב (RTL, צבעי-חומרה, עוגני-שקף, טבלאות).
 * ללא תלויות חיצוניות (Node מובנה בלבד).
 * שימוש:  node md-to-html.js <report.md> [report.html]
 */
'use strict';
const fs = require('fs');
const path = require('path');

const inPath = process.argv[2];
if (!inPath) { console.error('usage: node md-to-html.js <report.md> [report.html]'); process.exit(1); }
const outPath = process.argv[3] || inPath.replace(/\.md$/i, '') + '.html';
const md = fs.readFileSync(inPath, 'utf8');

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const inline = (s) => esc(s)
  .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
  .replace(/`([^`]+)`/g, '<code>$1</code>');
const sevClass = (s) => s.includes('🔴') ? 'blocker' : s.includes('🟠') ? 'warning' : s.includes('🔵') ? 'polish' : '';

const lines = md.split(/\r?\n/);
const out = [];
let i = 0;
let inList = false;
const closeList = () => { if (inList) { out.push('</ul>'); inList = false; } };

const isTableRow = (l) => /^\s*\|.*\|\s*$/.test(l);
const cells = (l) => l.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());

while (i < lines.length) {
  const line = lines[i].replace(/\s+$/, '');
  let m;

  // table: header row + separator row + body rows
  if (isTableRow(line) && i + 1 < lines.length && /^\s*\|[\s:|-]+\|\s*$/.test(lines[i + 1])) {
    closeList();
    const header = cells(line);
    i += 2;
    const rows = [];
    while (i < lines.length && isTableRow(lines[i])) { rows.push(cells(lines[i])); i++; }
    out.push('<table>');
    out.push('<thead><tr>' + header.map((h) => `<th>${inline(h)}</th>`).join('') + '</tr></thead>');
    out.push('<tbody>');
    for (const r of rows) {
      const cls = sevClass(r.join(' '));
      out.push(`<tr class="${cls}">` + r.map((c) => `<td>${inline(c)}</td>`).join('') + '</tr>');
    }
    out.push('</tbody></table>');
    continue;
  }

  if ((m = line.match(/^#\s+(.*)/))) { closeList(); out.push(`<h1>${inline(m[1])}</h1>`); i++; continue; }
  if ((m = line.match(/^##\s+(.*)/))) { closeList(); out.push(`<h2>${inline(m[1])}</h2>`); i++; continue; }
  if ((m = line.match(/^###\s+(.*)/))) {
    closeList();
    const sm = m[1].match(/שקף\s+(\d+)/);
    out.push(`<h3${sm ? ` id="slide-${sm[1]}"` : ''}>${inline(m[1])}</h3>`); i++; continue;
  }
  if ((m = line.match(/^>\s?(.*)/))) { closeList(); out.push(`<blockquote>${inline(m[1])}</blockquote>`); i++; continue; }
  if ((m = line.match(/^[-*]\s+(.*)/))) {
    if (!inList) { out.push('<ul>'); inList = true; }
    out.push(`<li class="${sevClass(line)}">${inline(m[1])}</li>`); i++; continue;
  }
  if (line.trim() === '') { closeList(); i++; continue; }
  closeList();
  out.push(`<p class="${sevClass(line)}">${inline(line)}</p>`); i++;
}
closeList();

const title = (md.match(/^#\s+(.*)/m) || [, 'דוח בקרת איכות'])[1].replace(/[*`]/g, '');
const html = `<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<style>
  :root { --blocker:#d12; --warning:#e67e00; --polish:#1769d1; }
  body { font-family:"Assistant","Segoe UI",Arial,sans-serif; line-height:1.65; max-width:1000px;
         margin:0 auto; padding:24px; color:#222; background:#fafafa; }
  h1 { font-size:1.7rem; border-bottom:2px solid #ccc; padding-bottom:8px; }
  h2 { font-size:1.3rem; margin-top:30px; }
  h3 { font-size:1.1rem; margin-top:26px; background:#eef1f5; padding:6px 12px; border-radius:6px;
       border-inline-start:5px solid #6b7a90; }
  table { border-collapse:collapse; width:100%; margin:8px 0 18px; background:#fff; font-size:0.95rem; }
  th,td { border:1px solid #d6d6d6; padding:8px 10px; text-align:right; vertical-align:top; }
  th { background:#33405a; color:#fff; font-weight:600; }
  th:nth-child(1),td:nth-child(1) { width:16%; }
  th:nth-child(2),td:nth-child(2) { width:30%; }
  th:nth-child(3),td:nth-child(3) { width:30%; }
  th:nth-child(4),td:nth-child(4) { width:24%; }
  tr.blocker td:first-child { border-inline-start:5px solid var(--blocker); background:#fdecec; }
  tr.warning td:first-child { border-inline-start:5px solid var(--warning); background:#fff6e9; }
  tr.polish  td:first-child { border-inline-start:5px solid var(--polish);  background:#eaf2fd; }
  code { background:#eee; padding:1px 5px; border-radius:4px; font-family:Consolas,monospace; direction:ltr; display:inline-block; }
  blockquote { color:#555; border-inline-start:3px solid #bbb; margin:12px 0; padding:4px 12px; font-style:italic; }
  ul { padding-inline-start:20px; }
  strong { color:#000; }
</style>
</head>
<body>
${out.join('\n')}
</body>
</html>`;

fs.writeFileSync(outPath, html, 'utf8');
console.log('wrote ' + path.resolve(outPath));

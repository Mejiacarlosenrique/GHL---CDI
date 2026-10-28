// Puente Drive → GitHub: descarga los Excel de los trackers (ids en index.html: xlsx:"...")
// y guarda SOLO la pestaña diaria como CSV en data/trackers/<id>.csv, con el mismo formato
// que entrega Google Sheets (fechas dd/mm/aaaa, números enteros). El dashboard lee esos CSV.
// Las otras pestañas (clientes, cédulas, teléfonos) NO se publican.
import fs from 'fs';
import XLSX from 'xlsx';

const TABS = ['Kpi Diario Closer', 'DIARIO'];
const html = fs.readFileSync('index.html', 'utf8');
const ids = [...new Set([...html.matchAll(/xlsx:"([A-Za-z0-9_-]{20,})"/g)].map(m => m[1]))];
const p2 = n => String(n).padStart(2, '0');
const q = s => '"' + String(s).replace(/"/g, '""') + '"';

function tabToCSV(ws) {
  if (!ws['!ref']) return '';
  const R = XLSX.utils.decode_range(ws['!ref']);
  const lines = [];
  for (let r = R.s.r; r <= R.e.r; r++) {
    const row = [];
    for (let c = 0; c <= R.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      let v = '';
      if (cell && cell.t === 'n') {
        if (cell.z && XLSX.SSF.is_date(cell.z)) { const d = XLSX.SSF.parse_date_code(cell.v); v = p2(d.d) + '/' + p2(d.m) + '/' + d.y; }
        else v = String(Math.round(cell.z && cell.z.includes('%') ? cell.v * 100 : cell.v)); // formato % → como lo muestra Sheets ("1%" = 1)
      } else if (cell && cell.v != null && cell.t !== 'e') v = String(cell.v);
      row.push(v);
    }
    while (row.length && row[row.length - 1] === '') row.pop(); // sin celdas vacías al final (archivo más liviano)
    lines.push(row.map(q).join(','));
  }
  return lines.join('\n');
}

let failed = 0;
fs.mkdirSync('data/trackers', { recursive: true });
for (const id of ids) {
  try {
    const r = await fetch(`https://drive.usercontent.google.com/download?id=${id}&export=download&confirm=t`);
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const wb = XLSX.read(Buffer.from(await r.arrayBuffer()), { type: 'buffer', cellNF: true });
    const name = wb.SheetNames.find(n => TABS.some(t => n.trim().toLowerCase() === t.toLowerCase()));
    if (!name) throw new Error('sin pestaña diaria: ' + wb.SheetNames.join(', '));
    fs.writeFileSync(`data/trackers/${id}.csv`, tabToCSV(wb.Sheets[name]) + '\n');
    console.log('ok', id, name);
  } catch (e) {
    failed++;
    console.error('FALLÓ', id, e.message); // se conserva el CSV anterior
  }
}
fs.writeFileSync('data/trackers/updated.txt', new Date().toISOString() + '\n');
if (failed === ids.length) process.exit(1);

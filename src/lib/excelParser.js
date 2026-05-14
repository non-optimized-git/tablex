import * as XLSX from 'xlsx';

export async function parseExcel(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref']);

  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    headers.push(sheet[addr]?.v ?? null);
  }

  const rows = [];
  for (let r = 1; r <= range.e.r; r++) {
    const row = [];
    let hasValue = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const val = sheet[addr]?.v ?? null;
      row.push(val);
      if (val !== null && val !== '') hasValue = true;
    }
    if (hasValue) rows.push(row);
  }

  return { headers, rows };
}
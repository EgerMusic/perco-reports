function children(node, localName) { return [...node.children].filter(x => x.localName === localName); }
function descendants(node, localName) { return [...node.getElementsByTagName('*')].filter(x => x.localName === localName); }
function xml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Не удалось прочитать структуру Excel.');
  return doc;
}
function column(ref) {
  const letters = (ref.match(/^[A-Z]+/) || [''])[0];
  return [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
}
function cellText(cell, shared) {
  const type = cell.getAttribute('t');
  if (type === 'inlineStr') return descendants(cell, 't').map(x => x.textContent).join('');
  const value = children(cell, 'v')[0]?.textContent ?? '';
  if (type === 's') return shared[Number(value)] ?? '';
  return value;
}
export async function readXlsx(file) {
  if (!globalThis.JSZip) throw new Error('Не загружен модуль чтения Excel.');
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookFile = zip.file('xl/workbook.xml');
  if (!workbookFile) throw new Error('Файл не похож на Excel XLSX.');
  const workbook = xml(await workbookFile.async('string'));
  const relFile = zip.file('xl/_rels/workbook.xml.rels');
  const rels = relFile ? xml(await relFile.async('string')) : null;
  const sheets = descendants(workbook, 'sheet');
  const sharedFile = zip.file('xl/sharedStrings.xml');
  const shared = sharedFile ? descendants(xml(await sharedFile.async('string')), 'si').map(si => descendants(si, 't').map(t => t.textContent).join('')) : [];
  for (let index = 0; index < sheets.length; index++) {
    const relId = sheets[index].getAttribute('r:id') || sheets[index].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const target = rels && descendants(rels, 'Relationship').find(r => r.getAttribute('Id') === relId)?.getAttribute('Target');
    const path = target ? (target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\.\//, '')}`) : `xl/worksheets/sheet${index + 1}.xml`;
    const sheetFile = zip.file(path);
    if (!sheetFile) continue;
    const sheet = xml(await sheetFile.async('string'));
    const rows = [];
    for (const row of descendants(sheet, 'row')) {
      const rowIndex = Number(row.getAttribute('r')) - 1;
      const values = [];
      for (const cell of children(row, 'c')) {
        const col = column(cell.getAttribute('r') || '');
        if (col >= 0) values[col] = cellText(cell, shared);
      }
      rows[rowIndex] = values;
    }
    if (rows.some(r => r?.includes('Табельный номер') && r?.includes('Выход из'))) return rows.map(r => r || []);
  }
  throw new Error('В книге не найден лист с проходами PERCo.');
}

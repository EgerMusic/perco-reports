const xml = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
const col = index => { let result = ''; for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result; return result; };

// Style numbers are shared with the report builder in app.mjs.
export const excelStyle = Object.freeze({ plain: 0, header: 1, greenRow: 2, amberRow: 3, redRow: 4, blueRow: 5, grayRow: 6, green: 7, amber: 8, red: 9, blue: 10 });

const fill = rgb => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`;
const font = (rgb, bold = false) => `<font>${bold ? '<b/>' : ''}<sz val="11"/><color rgb="${rgb}"/><name val="Calibri"/></font>`;
const xf = (fontId, fillId, bold = false) => `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0" applyFont="${bold ? 1 : 0}" applyFill="${fillId ? 1 : 0}" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`;
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="6">${font('FF172326')}${font('FFFFFFFF', true)}${font('FF176C4D', true)}${font('FF95520D', true)}${font('FFAD3E30', true)}${font('FF275D83', true)}</fonts><fills count="8"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fill('FF20373D')}${fill('FFEAF6EE')}${fill('FFFFF3DF')}${fill('FFFFEAE7')}${fill('FFEAF3FA')}${fill('FFF2F5F4')}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="11">${xf(0, 0)}${xf(1, 2, true)}${xf(0, 3)}${xf(0, 4)}${xf(0, 5)}${xf(0, 6)}${xf(0, 7)}${xf(2, 3, true)}${xf(3, 4, true)}${xf(4, 5, true)}${xf(5, 6, true)}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

export async function workbookBytes(sheets, Zip = globalThis.JSZip) {
  if (!Zip) throw new Error('Не загружен модуль Excel.');
  const zip = new Zip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', stylesXml);
  sheets.forEach((sheet, i) => {
    const data = [sheet.headers, ...sheet.rows];
    const width = sheet.headers.length;
    const rows = data.map((values, rowIndex) => {
      const height = rowIndex ? Math.min(90, Math.max(20, Math.max(...values.map((value, j) => String(value ?? '').split('\n').reduce((lines, part) => lines + Math.max(1, Math.ceil(part.length / Math.max(10, (sheet.columnWidths?.[j] ?? 24) - 3))), 0))) * 16)) : 25;
      const cells = values.map((value, j) => {
        const style = rowIndex === 0 ? excelStyle.header : sheet.cellStyles?.[rowIndex - 1]?.[j] ?? sheet.rowStyles?.[rowIndex - 1] ?? excelStyle.plain;
        return `<c r="${col(j)}${rowIndex + 1}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xml(value)}</t></is></c>`;
      }).join('');
      return `<row r="${rowIndex + 1}" ht="${height}" customHeight="1">${cells}</row>`;
    }).join('');
    const cols = sheet.headers.map((h, j) => `<col min="${j + 1}" max="${j + 1}" width="${sheet.columnWidths?.[j] ?? Math.min(48, Math.max(14, String(h).length + 5))}" customWidth="1"/>`).join('');
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${col(width - 1)}${Math.max(1, data.length)}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>${cols}</cols><sheetData>${rows}</sheetData><autoFilter ref="A1:${col(width - 1)}${Math.max(1, data.length)}"/></worksheet>`);
  });
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}

export async function downloadWorkbook(sheets, filename) {
  const blob = await workbookBytes(sheets);
  const url = URL.createObjectURL(blob), link = document.createElement('a');
  link.href = url; link.download = filename; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

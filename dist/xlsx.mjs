function children(node, name) { return [...(node?.children || [])].filter(x => x.localName === name); }
function descendants(node, name) { return [...(node?.getElementsByTagName('*') || [])].filter(x => x.localName === name); }
function xml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Не удалось прочитать структуру Excel.');
  return doc;
}
function column(ref) { return [...(ref.match(/^[A-Z]+/i)?.[0] || '').toUpperCase()].reduce((n,c) => n*26+c.charCodeAt(0)-64,0)-1; }
function pathFrom(base, target) {
  const parts = (target.startsWith('/') ? target.slice(1) : `${base}/${target}`).split('/'), result = [];
  for (const part of parts) { if(part === '..')result.pop();else if(part && part !== '.')result.push(part); }
  return result.join('/');
}
const indexedColors = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCCC 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333').split(' ');
function dateText(value, formatId, code, epoch1904) {
  const n = Number(value);
  if (!Number.isFinite(n) || value === '') return value;
  const fmt = code.replace(/"[^"]*"|\\.|\[(?!h+\]|m+\]|s+\])[^\]]*\]/gi, '').toLowerCase();
  const hasDate = /[dy]/.test(fmt) || (formatId >= 14 && formatId <= 17) || (formatId >= 27 && formatId <= 36) || (formatId >= 50 && formatId <= 58) || formatId === 22;
  const hasTime = /[hs]/.test(fmt) || (formatId >= 18 && formatId <= 22) || (formatId >= 45 && formatId <= 47);
  if (!hasDate && !hasTime) return value;
  const seconds = Math.round(n * 86400);
  if (hasDate) {
    const origin = epoch1904 ? Date.UTC(1904,0,1) : Date.UTC(1899,11,n < 60 ? 31 : 30);
    const d = new Date(origin + seconds * 1000);
    if (!Number.isFinite(d.getTime())) return value;
    const pad = v => String(v).padStart(2,'0');
    const date = `${pad(d.getUTCDate())}-${pad(d.getUTCMonth()+1)}-${d.getUTCFullYear()}`;
    return hasTime ? `${date} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}` : date;
  }
  const time = ((seconds % 86400) + 86400) % 86400, pad=v=>String(v).padStart(2,'0');
  return `${pad(Math.floor(time/3600))}:${pad(Math.floor(time/60)%60)}:${pad(time%60)}`;
}
function cellText(cell, shared, style, formats, epoch1904) {
  const type = cell.getAttribute('t');
  if (type === 'inlineStr') return descendants(cell, 't').map(x=>x.textContent).join('');
  const value = children(cell, 'v')[0]?.textContent ?? '';
  if (type === 's') return shared[Number(value)] ?? '';
  if (type === 'e') throw new Error(`Ошибка Excel ${value} в ячейке ${cell.getAttribute('r')}.`);
  if (!value && children(cell,'f').length) throw new Error(`Формула в ячейке ${cell.getAttribute('r')} не имеет сохранённого результата. Пересчитайте и сохраните книгу в Excel.`);
  if (type === 'd') {
    const m = value.match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}:\d{2}:\d{2}))?/);
    if(m)return `${m[3]}-${m[2]}-${m[1]}${m[4]?` ${m[4]}`:''}`;
  }
  const formatId = Number(style?.getAttribute('numFmtId') || 0);
  return type === 'str' || type === 'b' ? value : dateText(value, formatId, formats.get(formatId) || '', epoch1904);
}
function inRange(ref, range) {
  const [a,b=a] = range.replace(/\$/g,'').split(':'), r=Number(ref.match(/\d+/)?.[0]), c=column(ref);
  return c>=column(a)&&c<=column(b)&&r>=Number(a.match(/\d+/)?.[0])&&r<=Number(b.match(/\d+/)?.[0]);
}

export async function readXlsx(file, type = 'events') {
  if (!globalThis.JSZip) throw new Error('Не загружен модуль чтения Excel.');
  const zip = await JSZip.loadAsync(await file.arrayBuffer()), workbookFile=zip.file('xl/workbook.xml');
  if(!workbookFile)throw new Error('Файл не похож на Excel XLSX.');
  const workbook=xml(await workbookFile.async('string')), epoch1904=['1','true'].includes(descendants(workbook,'workbookPr')[0]?.getAttribute('date1904'));
  const relFile=zip.file('xl/_rels/workbook.xml.rels'), rels=relFile?xml(await relFile.async('string')):null;
  const sharedFile=zip.file('xl/sharedStrings.xml'),shared=sharedFile?descendants(xml(await sharedFile.async('string')),'si').map(si=>descendants(si,'t').map(t=>t.textContent).join('')):[];
  const styleFile=zip.file('xl/styles.xml'),styleDoc=styleFile?xml(await styleFile.async('string')):null;
  const fills=children(descendants(styleDoc,'fills')[0],'fill'),styles=children(descendants(styleDoc,'cellXfs')[0],'xf');
  const numberFormats=new Map(descendants(styleDoc,'numFmt').map(f=>[Number(f.getAttribute('numFmtId')),f.getAttribute('formatCode')||'']));
  const themeFile=zip.file('xl/theme/theme1.xml'),theme=themeFile?descendants(xml(await themeFile.async('string')),'clrScheme')[0]:null;
  const themeColors=theme?[...theme.children].map(n=>n.firstElementChild?.getAttribute('lastClr')||n.firstElementChild?.getAttribute('val')):[];
  if(themeColors.length>=4)[themeColors[0],themeColors[1],themeColors[2],themeColors[3]]=[themeColors[1],themeColors[0],themeColors[3],themeColors[2]];
  const fillColors=fills.map(fill=>{
    const fg=descendants(fill,'fgColor')[0];if(!fg)return 'FFFFFF';
    let rgb=fg.getAttribute('rgb')?.slice(-6);
    if(!rgb&&fg.hasAttribute('theme'))rgb=themeColors[Number(fg.getAttribute('theme'))];
    if(!rgb&&fg.hasAttribute('indexed'))rgb=indexedColors[Number(fg.getAttribute('indexed'))];
    if(!rgb)return 'UNKNOWN';
    const tint=Number(fg.getAttribute('tint')||0);
    if(/^[0-9a-f]{6}$/i.test(rgb)&&tint)rgb=rgb.match(/../g).map(v=>{const c=parseInt(v,16);return Math.round(tint<0?c*(1+tint):c+(255-c)*tint).toString(16).padStart(2,'0');}).join('');
    return rgb.toUpperCase();
  });
  const result=[],cellStyles=[],rowSources=[],matchedSheets=[];
  for(const sheetInfo of descendants(workbook,'sheet')){
    const relId=sheetInfo.getAttribute('r:id')||sheetInfo.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    const target=descendants(rels,'Relationship').find(r=>r.getAttribute('Id')===relId)?.getAttribute('Target');
    if(!target)throw new Error('В книге отсутствует ссылка на лист.');
    const sheetFile=zip.file(pathFrom('xl',target));if(!sheetFile)throw new Error('В книге отсутствует файл листа.');
    const sheet=xml(await sheetFile.async('string')),rows=[],sheetStyles=[],cols=descendants(sheet,'col');
    const conditionals=descendants(sheet,'conditionalFormatting').flatMap(n=>(n.getAttribute('sqref')||'').split(/\s+/));
    for(const row of descendants(sheet,'row')){
      const ri=Number(row.getAttribute('r'))-1,values=[],rowStyles=[];
      for(const cell of children(row,'c')){
        const address=cell.getAttribute('r')||'',ci=column(address);if(ci<0)continue;
        const columnStyle=cols.find(c=>Number(c.getAttribute('min'))<=ci+1&&Number(c.getAttribute('max'))>=ci+1)?.getAttribute('style');
        const si=Number(cell.getAttribute('s')??row.getAttribute('s')??columnStyle??0),style=styles[si];
        values[ci]=cellText(cell,shared,style,numberFormats,epoch1904);
        rowStyles[ci]={fillColor:fillColors[Number(style?.getAttribute('fillId')||0)]||'FFFFFF',conditional:conditionals.some(range=>inRange(address,range))};
      }
      rows[ri]=values;sheetStyles[ri]=rowStyles;
    }
    const trimmed=rows.map(row=>(row||[]).map(v=>String(v??'').trim()));
    const matches=type==='schedule'?trimmed.some(r=>r.includes('ФИО')&&r.some(c=>/^Таб\./i.test(c)))
      :type==='presence'?trimmed.some(r=>r.includes('Табельный номер')&&r.includes('График работы')&&r.includes('Вход'))
      :trimmed.some(r=>r.includes('Табельный номер')&&r.includes('Выход из')&&r.includes('Вход в'));
    if(!matches)continue;
    const sheetName=sheetInfo.getAttribute('name');matchedSheets.push(sheetName);
    for(let ri=0;ri<rows.length;ri++){
      result.push(rows[ri]||[]);cellStyles.push(sheetStyles[ri]||[]);rowSources.push({sheetName,rowNumber:ri+1});
    }
  }
  if(!matchedSheets.length)throw new Error(type==='schedule'?'Не найден месячный график с ФИО и табельными номерами.':type==='presence'?'Не найден отчёт «Время присутствия» с графиком работы.':'Не найден отчёт о проходах с полями «Выход из» и «Вход в».');
  result.cellStyles=cellStyles;result.rowSources=rowSources;result.sheetName=matchedSheets.join(', ');
  return result;
}

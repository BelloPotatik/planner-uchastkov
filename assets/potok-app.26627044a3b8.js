'use strict';

/* Минимальный читатель XLSX для автономной сборки планировщика.
   Читаем только XML-книгу, связи, строки и листы; изображения и стили не распаковываем. */
(() => {
  const MAX_ARCHIVE_BYTES = 32 * 1024 * 1024;
  const MAX_PART_BYTES = 64 * 1024 * 1024;
  const MAX_UNPACKED_BYTES = 128 * 1024 * 1024;
  const MAX_CELLS = 1000000;
  const decoder = new TextDecoder('utf-8');

  function fail(message){ throw new Error('XLSX: ' + message); }
  function u16(view, offset){ return view.getUint16(offset, true); }
  function u32(view, offset){ return view.getUint32(offset, true); }

  function xmlDecode(value){
    const entities = { amp:'&', lt:'<', gt:'>', quot:'"', apos:"'" };
    return String(value || '').replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (all, entity) => {
      if(entity[0] === '#'){
        const hex = entity[1].toLowerCase() === 'x';
        const code = parseInt(entity.slice(hex ? 2 : 1), hex ? 16 : 10);
        if(Number.isFinite(code) && code >= 0 && code <= 0x10ffff){
          try{ return String.fromCodePoint(code); }catch(e){ return all; }
        }
        return all;
      }
      return entities[entity.toLowerCase()] ?? all;
    });
  }

  function xmlAttrs(source){
    const attrs = {};
    const re = /([:\w-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let match;
    while((match = re.exec(source))) attrs[match[1]] = xmlDecode(match[2] ?? match[3]);
    return attrs;
  }

  function partPath(owner, target){
    const raw = String(target || '').replace(/\\/g, '/');
    if(!raw) return '';
    const parts = raw.startsWith('/') ? [] : owner.split('/').slice(0, -1);
    for(const part of raw.split('/')){
      if(!part || part === '.') continue;
      if(part === '..') parts.pop();
      else parts.push(part);
    }
    return parts.join('/');
  }

  async function openZip(input){
    const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    if(bytes.length > MAX_ARCHIVE_BYTES) fail('архив больше 32 МБ');
    if(bytes.length < 22) fail('файл слишком мал');
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    const first = Math.max(0, bytes.length - 65557);
    for(let i = bytes.length - 22; i >= first; i--){
      if(u32(view, i) !== 0x06054b50) continue;
      if(i + 22 + u16(view, i + 20) === bytes.length){ eocd = i; break; }
    }
    if(eocd < 0) fail('не найден каталог ZIP');
    const count = u16(view, eocd + 10);
    const directorySize = u32(view, eocd + 12);
    const directoryOffset = u32(view, eocd + 16);
    if(count === 0xffff || directorySize === 0xffffffff || directoryOffset === 0xffffffff){
      fail('ZIP64 не поддерживается');
    }
    if(count > 4096 || directoryOffset + directorySize > bytes.length) fail('повреждён каталог ZIP');

    const entries = new Map();
    let offset = directoryOffset;
    for(let i = 0; i < count; i++){
      if(offset + 46 > bytes.length || u32(view, offset) !== 0x02014b50) fail('повреждена запись ZIP');
      const flags = u16(view, offset + 8);
      const method = u16(view, offset + 10);
      const compressedSize = u32(view, offset + 20);
      const uncompressedSize = u32(view, offset + 24);
      const nameLength = u16(view, offset + 28);
      const extraLength = u16(view, offset + 30);
      const commentLength = u16(view, offset + 32);
      const localOffset = u32(view, offset + 42);
      const end = offset + 46 + nameLength + extraLength + commentLength;
      if(end > bytes.length || compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localOffset === 0xffffffff){
        fail('повреждены размеры ZIP');
      }
      const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
      if(name && !name.endsWith('/')) entries.set(name, { flags, method, compressedSize, uncompressedSize, localOffset });
      offset = end;
    }

    let unpackedBytes = 0;
    async function read(name){
      const entry = entries.get(name);
      if(!entry) fail('нет части ' + name);
      if(entry.flags & 1) fail('зашифрованные архивы не поддерживаются');
      if(entry.uncompressedSize > MAX_PART_BYTES || unpackedBytes + entry.uncompressedSize > MAX_UNPACKED_BYTES){
        fail('превышен безопасный размер распаковки');
      }
      const at = entry.localOffset;
      if(at + 30 > bytes.length || u32(view, at) !== 0x04034b50) fail('повреждён локальный заголовок ZIP');
      const start = at + 30 + u16(view, at + 26) + u16(view, at + 28);
      const end = start + entry.compressedSize;
      if(start < at || end > bytes.length) fail('данные ZIP выходят за границы файла');
      const compressed = bytes.subarray(start, end);
      let result;
      if(entry.method === 0) result = compressed.slice();
      else if(entry.method === 8){
        if(typeof DecompressionStream === 'undefined') fail('браузер не поддерживает deflate');
        const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
        result = new Uint8Array(await new Response(stream).arrayBuffer());
      } else fail('неподдерживаемый метод сжатия ' + entry.method);
      if(result.length !== entry.uncompressedSize) fail('не совпал размер распакованной части ' + name);
      unpackedBytes += result.length;
      return result;
    }
    return { entries, read };
  }

  async function xml(archive, name){
    return decoder.decode(await archive.read(name));
  }

  function sharedStrings(source){
    if(!source) return [];
    const values = [];
    const si = /<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g;
    let item;
    while((item = si.exec(source))){
      let text = '';
      const tags = /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
      let tag;
      while((tag = tags.exec(item[1]))) text += xmlDecode(tag[1]);
      values.push(text);
    }
    return values;
  }

  function columnIndex(reference){
    const match = String(reference || '').match(/^([A-Z]+)/i);
    if(!match) return null;
    let index = 0;
    for(const letter of match[1].toUpperCase()) index = index * 26 + letter.charCodeAt(0) - 64;
    return index - 1;
  }

  function excelDate(serial, date1904){
    if(!Number.isFinite(serial) || serial < 1 || serial > 2958465) return null;
    const days = Math.floor(serial);
    const epoch = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30);
    const date = new Date(epoch + days * 86400000);
    return String(date.getUTCDate()).padStart(2, '0') + '.' +
      String(date.getUTCMonth() + 1).padStart(2, '0') + '.' + date.getUTCFullYear();
  }

  function worksheet(source, shared, date1904, cellBudget){
    const sparse = new Map();
    let maxRow = 0;
    let cells = 0;
    let lastColumn = -1;
    const re = /<(?:\w+:)?c\b([^>]*)\/>|<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g;
    let match;
    while((match = re.exec(source))){
      if(++cells > MAX_CELLS || ++cellBudget.count > MAX_CELLS) fail('слишком много ячеек');
      const attrs = xmlAttrs(match[1] ?? match[2]);
      const body = match[3] || '';
      const reference = attrs.r || '';
      const rowMatch = reference.match(/(\d+)$/);
      const rowNumber = rowMatch ? Number(rowMatch[1]) : maxRow + 1;
      if(!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > 1000000) fail('некорректная строка ' + reference);
      let column = columnIndex(reference);
      if(column === null) column = lastColumn + 1;
      if(column < 0 || column > 16383) fail('некорректная колонка ' + reference);
      lastColumn = column;
      const type = attrs.t || '';
      const valueMatch = body.match(/<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/);
      const rawValue = valueMatch ? xmlDecode(valueMatch[1]) : '';
      let value = rawValue;
      if(type === 's') value = shared[Number(rawValue)] ?? '';
      else if(type === 'inlineStr'){
        value = '';
        const texts = /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
        let text;
        while((text = texts.exec(body))) value += xmlDecode(text[1]);
      } else if(type === 'b') value = rawValue === '1' ? 'TRUE' : 'FALSE';
      if(column === 0 && type !== 's' && type !== 'inlineStr' && type !== 'str' && rawValue !== ''){
        const date = excelDate(Number(rawValue), date1904);
        if(date) value = date;
      }
      if(!sparse.has(rowNumber)) sparse.set(rowNumber, []);
      sparse.get(rowNumber)[column] = value;
      maxRow = Math.max(maxRow, rowNumber);
    }
    const grid = [];
    for(let row = 1; row <= maxRow; row++){
      const values = sparse.get(row) || [];
      let length = values.length;
      while(length > 0 && (values[length - 1] === undefined || values[length - 1] === '')) length--;
      values.length = length;
      grid.push(Array.from(values, value => value ?? ''));
    }
    return grid;
  }

  async function readWorkbookOnMainThread(input){
    const archive = await openZip(input);
    const workbookName = 'xl/workbook.xml';
    const relsName = 'xl/_rels/workbook.xml.rels';
    const book = await xml(archive, workbookName);
    const rels = await xml(archive, relsName);
    const relationships = new Map();
    const relRe = /<(?:\w+:)?Relationship\b([^>]*)\/?>(?:[\s\S]*?<\/(?:\w+:)?Relationship>)?/g;
    let rel;
    while((rel = relRe.exec(rels))){
      const attrs = xmlAttrs(rel[1]);
      if(attrs.Id && attrs.Target) relationships.set(attrs.Id, partPath(workbookName, attrs.Target));
    }
    const shared = archive.entries.has('xl/sharedStrings.xml') ? sharedStrings(await xml(archive, 'xl/sharedStrings.xml')) : [];
    const workbookPr = book.match(/<(?:\w+:)?workbookPr\b([^>]*)\/?>/);
    const date1904 = workbookPr ? /(?:date1904)\s*=\s*["'](?:1|true)["']/i.test(workbookPr[1]) : false;
    const sheetRe = /<(?:\w+:)?sheet\b([^>]*)\/?>(?:[\s\S]*?<\/(?:\w+:)?sheet>)?/g;
    const descriptors = [];
    let sheet;
    while((sheet = sheetRe.exec(book))){
      const attrs = xmlAttrs(sheet[1]);
      const relationId = attrs['r:id'] || attrs.id;
      descriptors.push({ name:attrs.name || 'Лист ' + (descriptors.length + 1), path:relationships.get(relationId) || '' });
    }
    if(!descriptors.length) fail('в книге нет листов');
    const sheets = [];
    const cellBudget = { count:0 };
    for(const descriptor of descriptors){
      if(!descriptor.path || !archive.entries.has(descriptor.path)){
        sheets.push({ name:descriptor.name, grid:[] });
        continue;
      }
      const source = await xml(archive, descriptor.path);
      sheets.push({ name:descriptor.name, grid:worksheet(source, shared, date1904, cellBudget) });
    }
    return { sheets, sheetNames:descriptors.map(item => item.name), sheetCount:descriptors.length };
  }

  const XLSX_WORKER_URL = 'assets/potok-xlsx-worker.7565d9778616.js';
  let workerRequests = 0;
  function processWorkbook(book, format){
    if(format === 'upkp') return parseUPKPSheets(book);
    if(format === 'ops-shared') return window.POTOK_SHARED.parseOpsBook(book);
    return book;
  }
  async function readWorkbook(input, format){
    const canUseWorker = typeof Worker === 'function' && typeof document !== 'undefined' &&
      (typeof location === 'undefined' || location.protocol !== 'file:');
    if(!canUseWorker) return processWorkbook(await readWorkbookOnMainThread(input), format);
    const id = ++workerRequests;
    return new Promise((resolve, reject) => {
      let worker, settled = false;
      const cleanup = () => {
        if(!worker) return;
        worker.onmessage = null; worker.onerror = null; worker.terminate();
      };
      const fallback = () => {
        if(settled) return;
        settled = true; cleanup();
        readWorkbookOnMainThread(input).then(book => processWorkbook(book, format)).then(resolve, reject);
      };
      try{
        worker = new Worker(XLSX_WORKER_URL);
        worker.onmessage = event => {
          if(settled || !event.data || event.data.id !== id) return;
          settled = true; cleanup();
          if(event.data.error) reject(new Error(event.data.error));
          else resolve(event.data.book);
        };
        worker.onerror = event => {
          if(event && event.preventDefault) event.preventDefault();
          fallback();
        };
        // Копия передаётся в worker, исходный ArrayBuffer остаётся доступным caller'у.
        const copy = (input instanceof Uint8Array ? input : new Uint8Array(input)).slice();
        worker.postMessage({ id, buffer:copy.buffer, format }, [copy.buffer]);
      }catch(error){ fallback(); }
    });
  }

  if(typeof window !== 'undefined') window.PLANNER_XLSX = { readWorkbook, get workerRequests(){ return workerRequests; } };
  if(typeof self !== 'undefined') self.POTOK_XLSX_PROCESS = processWorkbook;
})();


'use strict';

/* Разбор УПКП не зависит от DOM и выполняется также внутри XLSX Worker. */
function upkpNumber(value){
  const result = parseFloat(String(value ?? '').replace(/\s/g, '').replace(',', '.'));
  return Number.isFinite(result) ? result : 0;
}
function upkpDate(value){
  const match = String(value || '').trim().match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
  return match ? match[3] + '-' + match[2].padStart(2, '0') + '-' + match[1].padStart(2, '0') : null;
}
function parseUPKPGrid(grid){
  const out = [];
  let qtyCol = 12;
  for(const cells of grid.slice(0, 12)){
    const found = cells.findIndex(cell => /^N\s*секций/i.test(String(cell || '').trim()));
    if(found > 0){ qtyCol = found; break; }
  }
  const productCol = Math.max(0, qtyCol - 1), sumCol = qtyCol + 1;
  let curDate = null;
  for(const cells of grid){
    const c0 = String(cells[0] || '').trim();
    const match = c0.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if(match){ curDate = upkpDate(c0); continue; }
    if(!curDate || !c0) continue;
    const product = String(cells[productCol] || '').trim();
    const qty = upkpNumber(cells[qtyCol]);
    if(!product || !qty || /^[-+]?\d+(?:[.,]\d+)?$/.test(product.replace(/\s/g, ''))) continue;
    if(/неделя|^\s*(сб|вс)\s*$/.test(product)) continue;
    out.push({ custOrder:c0, prodOrder:'', prod:product, qty, date:curDate,
      minutes:qty, tdg:null, comment:'', sum:upkpNumber(cells[sumCol]) || 0 });
  }
  return out;
}
function parseUPKPSheets(book){
  const rows = [];
  // Повторы и исходный порядок листов сохраняются, дедупликации нет.
  for(const sheet of book.sheets) rows.push(...parseUPKPGrid(sheet.grid));
  return { rows, format:'upkp', sheetNames:book.sheetNames, sheetCount:book.sheetCount };
}


'use strict';

/* Общий лист «Общая загрузка»: чистые функции разбора и контекст операций.
   УПКП остаётся на отдельной книге в секциях. Здесь его строки используются
   только как доказанные даты сборки для сквозных связей автоплана. */
(() => {
  const SPECS = [
    { operation:'06 Сбока ОП', id:'s_op', sharedKey:'ОП', name:'ОП', group:'Сборка', category:'Сборка', workers:4, perWorker:450, resourceConfirmed:true },
    { operation:'09 Сборка САУ', id:'s_sau', sharedKey:'САУ', name:'Кипиа / САУ', group:'Сборка', category:'Сборка', workers:6, perWorker:450, resourceConfirmed:true },
    { operation:'08 Сборка БВ', id:'s_bv', sharedKey:'БВ', name:'БВ', group:'Сборка', category:'Сборка', workers:6, perWorker:450, resourceConfirmed:true },
    { operation:'03 Координатная пробивка', id:'s_coord', name:'Координатная пробивка', group:'Заготовка', category:'Заготовка' },
    { operation:'05 Гибка', id:'s_bending', name:'Гибка', group:'Заготовка', category:'Заготовка' },
    { operation:'Вальцовка', id:'s_rolling', name:'Вальцовка', group:'Заготовка', category:'Заготовка' },
    { operation:'ОТК', id:'s_qc', name:'ОТК', group:'Контроль', category:'Контроль' },
    { operation:'Комплектация приборов', id:'s_devices', name:'Комплектация приборов', group:'Комплектация', category:'Комплектация' },
    { operation:'Покрасочная', id:'s_painting', name:'Покрасочная', group:'Заготовка', category:'Заготовка' },
    { operation:'Балансировка', id:'s_balancing', name:'Балансировка', group:'Заготовка', category:'Заготовка' },
    { operation:'Производство ТО', id:'s_to', name:'Производство ТО', group:'Заготовка', category:'Заготовка' },
    { operation:'Критическая спецификация КО', id:'s_critical_ko', name:'Критическая спецификация КО', group:'Комплектация', category:'Комплектация' },
    { operation:'Критическая спецификация САУ', id:'s_critical_sau', name:'Критическая спецификация САУ', group:'Комплектация', category:'Комплектация' },
    { operation:'комплектация монтажного комплекта', id:'s_mounting', name:'Комплектация монтажного комплекта', group:'Комплектация', category:'Комплектация' },
    { operation:'МДГ', id:'s_mdg', name:'МДГ', group:'Заготовка', category:'Заготовка' },
    { operation:'Резка мембраны', id:'s_membrane', name:'Резка мембраны', group:'Заготовка', category:'Заготовка' },
    { operation:'Комплектация кабельной продукции', id:'s_cable', name:'Комплектация кабельной продукции', group:'Комплектация', category:'Комплектация' },
  ];

  const clean = value => String(value ?? '').replace(/\s+/g, ' ').trim();
  const normalizeOperation = value => clean(value).toLocaleLowerCase('ru-RU');
  const normalizeProduct = value => clean(value).toLocaleLowerCase('ru-RU');
  const isKdRb = operation => /(^|[^а-яa-z])(кд|рб)(?=$|[^а-яa-z])/i.test(clean(operation));
  const isSharedUpkp = operation => /^(?:07\s+)?с[бз]ока\s+упкп/i.test(clean(operation));
  const isAssembly = operation => /^(?:06\s+с[бз]ока\s+оп|07\s+с[бз]ока\s+упкп|08\s+сборка\s+бв|09\s+сборка\s+сау)/i.test(clean(operation));
  const sourceSiteKey = operation => 'ops:' + normalizeOperation(operation);
  const relationKey = row => {
    const order = clean(row && row.prodOrder).toLocaleUpperCase('ru-RU');
    const product = normalizeProduct(row && row.prod);
    return order && product ? order + '|' + product : '';
  };
  function hash(value){
    let result = 0;
    for(let index=0; index<value.length; index++) result = (result * 31 + value.charCodeAt(index)) | 0;
    return (result >>> 0).toString(36);
  }
  function specForOperation(operation){
    const normalized = normalizeOperation(operation);
    if(!normalized || normalized === 'операция' || isKdRb(normalized) || isSharedUpkp(normalized)) return null;
    const known = SPECS.find(spec => normalizeOperation(spec.operation) === normalized);
    if(known) return {...known, sourceSiteKey:sourceSiteKey(known.operation)};
    const name = clean(operation).replace(/^\d+\s+/, '');
    return {
      operation:clean(operation), id:'s_ops_' + hash(sourceSiteKey(operation)),
      sharedKey:'OPS_' + hash(sourceSiteKey(operation)), name:name || 'Новый участок',
      group:'Прочее', category:'Прочее', sourceSiteKey:sourceSiteKey(operation), discovered:true,
    };
  }
  function number(value){
    const result = Number(String(value ?? '').replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(result) ? result : 0;
  }
  function date(value){
    if(value === null || value === undefined || value === '') return null;
    const iso = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if(iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
    const ru = String(value).match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
    if(ru) return ru[3] + '-' + String(ru[2]).padStart(2, '0') + '-' + String(ru[1]).padStart(2, '0');
    const numeric = Number(value);
    if(Number.isFinite(numeric) && numeric > 20000 && numeric < 100000){
      const parsed = new Date(Date.UTC(1899, 11, 30) + numeric * 86400000);
      return parsed.toISOString().slice(0, 10);
    }
    return null;
  }
  function headerIndexes(grid){
    for(let rowIndex=0; rowIndex<Math.min(grid.length, 20); rowIndex++){
      const row = grid[rowIndex] || [];
      const headers = row.map(cell => normalizeOperation(cell));
      const operation = headers.findIndex(header => header === 'операция');
      const dateColumn = headers.findIndex(header => header === 'дата операции');
      const productionOrder = headers.findIndex(header => header === 'заказ на производство');
      if(operation >= 0 && dateColumn >= 0 && productionOrder >= 0){
        const find = text => headers.findIndex(header => header === text);
        return {
          rowIndex, operation, date:dateColumn, productionOrder,
          customerOrder:find('заказ покупателя'), product:find('продукция'), quantity:find('количество'),
          minutes:find('мин'), total:find('всего'), tdg:find('тдг'), comment:find('комментарий'),
          executor:find('исполнитель'), sum:find('сумма'),
        };
      }
    }
    throw new Error('На листе «Общая загрузка» не найдены обязательные колонки');
  }
  function value(row, index){ return index >= 0 ? row[index] : ''; }
  function isRepeatedHeader(cells, columns){
    return normalizeOperation(value(cells, columns.operation)) === 'операция' &&
      normalizeOperation(value(cells, columns.date)) === 'дата операции' &&
      normalizeOperation(value(cells, columns.productionOrder)) === 'заказ на производство';
  }
  function parseSharedGrid(grid){
    const columns = headerIndexes(grid);
    const rows = [], assemblyRows = [], operations = new Map();
    let kdRbExcluded = 0, upkpDelegated = 0, repeatedHeaders = 0;
    for(let index=columns.rowIndex+1; index<grid.length; index++){
      const cells = grid[index] || [];
      const operation = clean(value(cells, columns.operation));
      if(!operation) continue;
      // Google может повторять строку заголовков каждые 1000 строк. Это не участок
      // и не операция: пропускаем её до разбора дат и автоматического discovery.
      if(isRepeatedHeader(cells, columns)){ repeatedHeaders++; continue; }
      if(isKdRb(operation)){ kdRbExcluded++; continue; }
      const prodOrder = clean(value(cells, columns.productionOrder));
      const prod = clean(value(cells, columns.product));
      const qty = number(value(cells, columns.quantity));
      const minutes = columns.total >= 0
        ? number(value(cells, columns.total))
        : number(value(cells, columns.minutes)) * qty;
      const row = {
        custOrder:clean(value(cells, columns.customerOrder)), prodOrder, prod, qty,
        date:date(value(cells, columns.date)), minutes,
        tdg:date(value(cells, columns.tdg)), comment:clean(value(cells, columns.comment)),
        executor:clean(value(cells, columns.executor)), sum:number(value(cells, columns.sum)),
        operation, sourceRow:index + 1,
      };
      if(isAssembly(operation)) assemblyRows.push(row);
      if(isSharedUpkp(operation)){ upkpDelegated++; continue; }
      if(!specForOperation(operation)) continue;
      rows.push(row);
      if(!operations.has(normalizeOperation(operation))) operations.set(normalizeOperation(operation), operation);
    }
    return {rows, assemblyRows, operations:[...operations.values()], kdRbExcluded, upkpDelegated, repeatedHeaders};
  }
  function parseOpsBook(book){
    let sheet = book.sheets.find(item => normalizeOperation(item.name) === 'общая загрузка');
    if(!sheet){
      sheet = book.sheets.find(item => {
        try{ headerIndexes(item.grid); return true; }catch(error){ return false; }
      });
    }
    if(!sheet) throw new Error('В книге не найден лист «Общая загрузка»');
    return {
      ...parseSharedGrid(sheet.grid), format:'ops-shared', sheetName:sheet.name,
      sheetNames:[sheet.name], sheetCount:1, workbookSheetCount:book.sheetCount,
    };
  }
  async function parseOpsWorkbook(buffer){
    if(!window.PLANNER_XLSX) throw new Error('Модуль чтения XLSX не загружен');
    return window.PLANNER_XLSX.readWorkbook(buffer, 'ops-shared');
  }

  window.POTOK_SHARED = {
    SPECS, normalizeOperation, normalizeProduct, isKdRb, isSharedUpkp, isAssembly,
    sourceSiteKey, relationKey, specForOperation, parseSharedGrid, parseOpsBook, parseOpsWorkbook,
  };
})();


'use strict';
/* ================= data ================= */
const SNAP = window.PLANNER_SNAPSHOT || JSON.parse(document.getElementById('snapshot-data').textContent);
const CONTRACTS = SNAP.contracts || {};
const LS_KEY = 'planner_sites_v2';
// Менять при обновлении встроенного снимка УПКП: устаревший localStorage
// не должен маскироваться под актуальный план после выкладки новой версии.
const UPKP_SNAPSHOT_VERSION = '2026-10-29.2';
const UPKP_DAILY_CAPACITY = 30; // секций в день на бригаду УПКП целиком
const WD = ['Вс','Пн','Вт','Ср','Чт','Пт','Сб'];
const state = { sites: [], activeId: null, view: 'board', selectedDate: null, rangeFrom: null, rangeTo: null, _enterView: false };
let pendingPlan = null;
const reducedMotion = () => typeof matchMedia !== 'undefined' && matchMedia('(prefers-reduced-motion: reduce)').matches;
const raf = fn => (typeof requestAnimationFrame !== 'undefined' ? requestAnimationFrame(fn) : setTimeout(fn, 16));

/* ---------- тема оформления (светлая / тёмная) ---------- */
const THEME_KEY = 'planner.theme';
function currentTheme(){
  const saved = (typeof localStorage !== 'undefined' && localStorage.getItem(THEME_KEY)) || '';
  if(saved === 'dark' || saved === 'light') return saved;
  if(typeof matchMedia !== 'undefined' && matchMedia('(prefers-color-scheme: dark)').matches) return 'dark';
  return 'light';
}
function applyTheme(theme){
  if(typeof document !== 'undefined' && document.documentElement) document.documentElement.dataset.theme = theme;
  const b = typeof document !== 'undefined' && document.getElementById('btnTheme');
  if(b){
    b.dataset.mode = theme;
    b.title = theme === 'dark' ? 'Включить светлую тему' : 'Включить тёмную тему';
    if(b.setAttribute) b.setAttribute('aria-pressed', theme === 'dark' ? 'true' : 'false');
  }
  const label = typeof document !== 'undefined' && document.getElementById('themeLabel');
  if(label) label.textContent = theme === 'dark' ? 'Тёмная' : 'Светлая';
}
function toggleTheme(){
  const next = currentTheme() === 'dark' ? 'light' : 'dark';
  const apply = () => {
    try{ localStorage.setItem(THEME_KEY, next); }catch(e){}
    applyTheme(next);
  };
  if(typeof document !== 'undefined' && document.startViewTransition && !reducedMotion()){
    document.startViewTransition(apply); // плавный кроссфейд всей страницы
  } else apply();
}

/* ---------- журнал изменений ---------- */
const LOG_KEY = 'planner_log_v1';
function readLog(){
  try{ const a = JSON.parse(localStorage.getItem(LOG_KEY) || '[]'); return Array.isArray(a) ? a : []; }
  catch(e){ return []; }
}
function logChange(site, action, details){
  try{
    const arr = readLog();
    arr.push({ ts: Date.now(), site: site ? site.id : '', siteName: site ? site.name : '',
      action: String(action || ''), details: String(details || '') });
    if(arr.length > 3000) arr.splice(0, arr.length - 3000);
    localStorage.setItem(LOG_KEY, JSON.stringify(arr));
  }catch(e){}
}
function fmtTs(ts){
  const d = new Date(ts);
  return pad(d.getDate()) + '.' + pad(d.getMonth()+1) + '.' + d.getFullYear() + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
}

/* ---------- helpers ---------- */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pad = n => String(n).padStart(2,'0');
function fmtD(iso){ if(!iso) return '—'; const [y,m,d] = iso.split('-'); return d+'.'+m; }
function fmtDFull(iso){ if(!iso) return '—'; const [y,m,d] = iso.split('-'); return d+'.'+m+'.'+y; }
function parseDate(s){
  if(!s) return null;
  s = String(s).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return m[1]+'-'+m[2]+'-'+m[3];
  m = s.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
  if(m) return m[3]+'-'+pad(m[2])+'-'+pad(m[1]);
  return null;
}
function todayISO(){ const d = new Date(); return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate()); }
function addDays(iso, n){ const d = new Date(iso+'T00:00:00'); d.setDate(d.getDate()+n); return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate()); }
function wdOf(iso){ return new Date(iso+'T00:00:00').getDay(); }
function orderNum(s){
  s = String(s||'');
  // В УПКП номер идёт сразу после №. Не берём цифры из названий вроде «ЭМУ-4» / «СУ-10».
  const sheetOrder = s.match(/^\s*№\s*(\d+)/);
  if(sheetOrder) return Number(sheetOrder[1]);
  const customerOrder = s.match(/заказ\s+покупателя\s*(\d+)/i);
  if(customerOrder) return Number(customerOrder[1]);
  const groups = s.match(/\d+/g) || [];
  if(!groups.length) return null;
  return Number(s.includes('-') ? groups[groups.length-1] : groups[0]);
}
/* УПКП не имеет заказа на производство: договор — обязательная часть ключа позиции.
   Иначе одинаковые установки из разных договоров склеиваются в одну позицию. */
function positionKey(r, site){
  const order = site && site.format === 'upkp' ? (r.custOrder || r.prodOrder) : r.prodOrder;
  return (order || '') + '|' + (r.prod || '');
}
function upkpOrderLabel(value){
  const s = String(value || '').trim();
  const m = s.match(/^№\s*\d+/);
  return m ? m[0].replace(/^№\s*/, '№ ') : s;
}
function upkpCustomer(value){
  return String(value || '').replace(/^\s*№\s*\d+\s*/, '').trim();
}
function upkpProductLabel(value){
  let s = String(value || '').replace(/\s+/g, ' ').trim();
  s = s.split(/:{1,2}\s*приточная часть/i)[0];
  s = s.split(/\s+с\s+сист\.?\s*автоматики.*$/i)[0];
  s = s.split(/\s*,\s*(?=исполнение|исп-е|исп\.)/i)[0];
  s = s.split(/\s*,?\s*(?=РБ№|Система автоматического|система автоматического|Узел регулирования|ККБ\s)/)[0];
  return s.trim();
}
function displayProduct(r, site){ return site && site.format === 'upkp' ? upkpProductLabel(r.prod) : r.prod; }
function displayCustomer(r, site){ return site && site.format === 'upkp' ? (upkpCustomer(r.custOrder) || r.customer || '') : (r.customer || ''); }
function hashStr(s){ let h = 0; for(let i=0;i<s.length;i++) h = (h*31 + s.charCodeAt(i))|0; return (h>>>0).toString(36); }
function num(v){ const n = parseFloat(String(v??'').replace(/\s/g,'').replace(',','.')); return isNaN(n)?0:n; }
function fmtMin(m){ return (Math.round(m)||0).toLocaleString('ru-RU'); }
function uMin(site, v){ return fmtMin(v) + ' ' + ((site && site.unit) || 'мин'); }

/* ---------- csv ---------- */
const RU_MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
function curMonthSheet(){
  const d = new Date();
  return RU_MONTHS[d.getMonth()] + ' ' + d.getFullYear();
}
function csvToGrid(text){
  const rows = [];
  let cur = [''], q = false, row = cur;
  for(let i=0;i<text.length;i++){
    const c = text[i];
    if(q){
      if(c === '"'){ if(text[i+1] === '"'){ row[row.length-1]+='"'; i++; } else q = false; }
      else row[row.length-1]+=c;
    } else if(c === '"') q = true;
    else if(c === ',') row.push('');
    else if(c === '\n' || c === '\r'){
      if(c === '\r' && text[i+1] === '\n') i++;
      rows.push(row); row = ['']; cur = row;
    } else row[row.length-1]+=c;
  }
  if(row.length > 1 || row[0] !== '') rows.push(row);
  return rows;
}
function parseCSVText(text){
  const rows = csvToGrid(text);
  if(!rows.length) return [];
  const head = rows[0].map(h => h.trim().toLowerCase().replace(/^﻿/,''));
  const col = re => head.findIndex(h => re.test(h));
  const cCust = col(/заказ покуп/), cProd = col(/продукц/), cQty = col(/колич/),
        cDate = col(/дата операц/), cMin = col(/всего|мин.*итог|норма/i),
        cTdg = col(/тдг/), cComm = col(/коммент/), cPO = col(/заказ на произв/), cSum = col(/сумм/);
  if(cDate < 0 || cPO < 0) throw new Error('Не найдены колонки «Дата операции» / «Заказ на производство»');
  const out = [];
  for(let i=1;i<rows.length;i++){
    const r = rows[i];
    const date = parseDate(r[cDate]);
    const prodOrder = (r[cPO]||'').trim();
    if(!date && !prodOrder) continue;
    out.push({
      custOrder: cCust>=0?(r[cCust]||'').trim():'',
      prodOrder,
      prod: cProd>=0?(r[cProd]||'').trim():'',
      qty: cQty>=0?num(r[cQty]):0,
      date,
      minutes: cMin>=0?num(r[cMin]):0,
      tdg: cTdg>=0?parseDate(r[cTdg]):null,
      comment: cComm>=0?(r[cComm]||'').trim():'',
      sum: cSum>=0?num(r[cSum]):0
    });
  }
  return out;
}
/* ---------- формат «План УПКП»: дневные блоки, позиции со секциями ---------- */
function parseUPKP(text){ return parseUPKPGrid(csvToGrid(text)); }
async function parseUPKPWorkbook(buffer){
  if(!window.PLANNER_XLSX) throw new Error('Модуль чтения XLSX не загружен');
  return window.PLANNER_XLSX.readWorkbook(buffer, 'upkp');
}
function detectFormat(text){
  const first = (text || '').split('\n', 1)[0] || '';
  return /загрузка упкп/i.test(first) ? 'upkp' : 'ops';
}
function parseAny(text, format){
  if(format === 'upkp') return parseUPKP(text);
  if(format === 'ops') return parseCSVText(text);
  return detectFormat(text) === 'upkp' ? parseUPKP(text) : parseCSVText(text);
}
function rowIdentity(r){ return [r.prodOrder,r.prod,r.qty,r.date,r.minutes,r.custOrder].join('|'); }
function stableSharedRowIdentity(r){
  if(!r || !r.operation) return '';
  // Дата, количество и норма редактируются в Google Таблице. Их изменение не
  // превращает операцию в другую позицию и не должно отцеплять ручной перенос.
  return [r.prodOrder,r.prod,r.custOrder,r.operation].join('|');
}
function makeRows(raw, previousRows){
  const seen = {}, previous = new Map(), previousStable = new Map();
  const add = (map, key, id) => {
    if(!key) return;
    if(!map.has(key)) map.set(key, []);
    map.get(key).push(id);
  };
  const take = (map, key) => {
    const ids = map.get(key);
    return ids && ids.length ? ids.shift() : null;
  };
  const remove = (map, key, id) => {
    const ids = map.get(key);
    if(!ids) return;
    const index = ids.indexOf(id);
    if(index >= 0) ids.splice(index, 1);
  };
  for(const row of Array.isArray(previousRows) ? previousRows : []){
    add(previous, rowIdentity(row), row.id);
    add(previousStable, stableSharedRowIdentity(row), row.id);
  }
  return raw.map((r,i) => {
    const identity = rowIdentity(r), stableIdentity = stableSharedRowIdentity(r);
    // Сначала точное совпадение, затем устойчивый бизнес-ключ общей операции.
    // Так сохраняются и старые ID УПКП, и переносы после правки даты/нормы в Google.
    let id = take(previous, identity);
    if(!id && stableIdentity) id = take(previousStable, stableIdentity);
    if(id) remove(previousStable, stableIdentity, id);
    if(!id){
      id = 'r' + hashStr(identity);
      if(seen[id]) id += '_' + i;
    }
    seen[id] = 1;
    const c = CONTRACTS[String(orderNum(r.custOrder))] || {};
    return { id, custOrder:r.custOrder, prodOrder:r.prodOrder, prod:r.prod, qty:r.qty,
      date:r.date, minutes:r.minutes, tdg:r.tdg || c.tdg || null,
      contractDate:c.contractDate||null, mdg:c.mdg||null, customer:c.customer||'',
      comment:r.comment||'', executor:r.executor||'', operation:r.operation||'', sourceRow:r.sourceRow||null, sum:r.sum||0 };
  });
}

/* ---------- sites, persistent identities and user configuration ---------- */
const SHARED = window.POTOK_SHARED;
const SITE_CONFIG_KEY = 'planner.siteConfig.v1';
const SITE_CONFIG_PATH = 'shared/site-settings.json';
const DEFAULT_CATEGORIES = ['Заготовка','Сборка','Контроль','Комплектация','Прочее'];
function blankSiteConfig(){
  return { version:1, updatedAt:0, categories:DEFAULT_CATEGORIES.slice(), sites:{}, deleted:[] };
}
function readLocalSiteConfig(){
  try{
    const value = JSON.parse(localStorage.getItem(SITE_CONFIG_KEY) || 'null');
    if(!value || value.version !== 1) return blankSiteConfig();
    return {
      ...blankSiteConfig(), ...value,
      categories:Array.isArray(value.categories) ? value.categories : DEFAULT_CATEGORIES.slice(),
      sites:value.sites && typeof value.sites === 'object' ? value.sites : {},
      deleted:Array.isArray(value.deleted) ? value.deleted : [],
    };
  }catch(error){ return blankSiteConfig(); }
}
let siteConfig = readLocalSiteConfig();
function siteKeyOf(site){
  if(site.sourceSiteKey) return site.sourceSiteKey;
  if(site.id === 's_upkp' || site.format === 'upkp') return 'upkp:main';
  const spec = SHARED.SPECS.find(item => item.id === site.id);
  return spec ? SHARED.sourceSiteKey(spec.operation) : 'manual:' + site.id;
}
function ensureSiteIdentity(site){
  const spec = SHARED.SPECS.find(item => item.id === site.id);
  if(spec){
    site.sourceSiteKey = SHARED.sourceSiteKey(spec.operation);
    site.sourceOperation = spec.operation;
    site.sourceType = 'url'; site.sourceUrl = SNAP.opsUrl || '';
    site.format = 'ops-shared'; site.sharedKey = spec.sharedKey;
    if(!site.group) site.group = spec.group;
    if(!site.category) site.category = spec.category;
    if(!site.mappingNote) site.mappingNote = 'Общая загрузка → Операция = «' + spec.operation + '». Строки КД/РБ исключены.';
  }else if(site.id === 's_upkp' || site.format === 'upkp'){
    site.sourceSiteKey = 'upkp:main';
    if(!site.group) site.group = 'Сборка';
    if(!site.category) site.category = 'Сборка';
    if(!site.mappingNote) site.mappingNote = 'Отдельная книга УПКП: секции и договоры. Строки общей загрузки в минутах сюда не дублируются.';
  }else site.sourceSiteKey = siteKeyOf(site);
  return site;
}
function applySiteConfig(site){
  ensureSiteIdentity(site);
  const saved = siteConfig.sites[site.sourceSiteKey];
  if(!saved) return site;
  for(const key of ['name','group','category','workers','perWorker','fondMin','teamTotal','workdays','resourceConfirmed']){
    if(saved[key] !== undefined) site[key] = saved[key];
  }
  return site;
}
function configFromSite(site){
  ensureSiteIdentity(site);
  return {
    id:site.id, sourceOperation:site.sourceOperation || '', format:site.format || '', unit:site.unit || 'мин',
    mappingNote:site.mappingNote || '', name:site.name, group:site.group || 'Прочее', category:site.category || site.group || 'Прочее',
    workers:site.workers || 0, perWorker:site.perWorker || 0, fondMin:site.fondMin || 0,
    teamTotal:!!site.teamTotal, workdays:Array.isArray(site.workdays) ? site.workdays : [1,2,3,4,5],
    resourceConfirmed:site.resourceConfirmed !== false,
  };
}
let siteConfigPushTimer = 0;
function persistSiteConfiguration(remote = true){
  for(const site of state.sites) siteConfig.sites[siteKeyOf(site)] = configFromSite(site);
  siteConfig.categories = [...new Set([...DEFAULT_CATEGORIES, ...siteConfig.categories, ...state.sites.map(site => site.category).filter(Boolean)])];
  siteConfig.deleted = [...new Set(siteConfig.deleted)];
  siteConfig.updatedAt = Date.now();
  localStorage.setItem(SITE_CONFIG_KEY, JSON.stringify(siteConfig));
  if(remote && !applyingRemote){
    clearTimeout(siteConfigPushTimer);
    siteConfigPushTimer = setTimeout(pushSiteConfiguration, 900);
  }
}
function createSharedSite(spec){
  const snapshotKey = spec.id === 's_op' ? 'op' : spec.id === 's_sau' ? 'sau' : spec.id === 's_bv' ? 'bv' : null;
  const rows = makeRows((snapshotKey && SNAP[snapshotKey] || []).map(item => ({
    custOrder:item.order, prodOrder:item.prodOrder, prod:item.product, qty:item.qty,
    date:item.opDate, minutes:item.norm, tdg:item.tdg, comment:item.comment || '', sum:item.sum || 0,
    operation:spec.operation,
  })));
  const workers = spec.workers || 1, perWorker = spec.perWorker === undefined ? 450 : spec.perWorker;
  return applySiteConfig({
    id:spec.id, sharedKey:spec.sharedKey || 'OPS_' + hashStr(spec.sourceSiteKey || SHARED.sourceSiteKey(spec.operation)),
    name:spec.name, sourceType:'url', sourceUrl:SNAP.opsUrl || '', format:'ops-shared',
    sourceOperation:spec.operation, sourceSiteKey:SHARED.sourceSiteKey(spec.operation),
    group:spec.group || 'Прочее', category:spec.category || spec.group || 'Прочее',
    mappingNote:'Общая загрузка → Операция = «' + spec.operation + '». Строки КД/РБ исключены.',
    unit:'мин', workers, perWorker, fondMin:workers * perWorker, teamTotal:false,
    resourceConfirmed:spec.resourceConfirmed === true,
    workdays:[1,2,3,4,5], rows, overrides:{}, splits:{},
  });
}
function defaultSites(){
  const sites = SHARED.SPECS.map(createSharedSite);
  const upkpRows = makeRows((SNAP.upkp || []).map(item => ({
    custOrder:item.order, prodOrder:item.prodOrder, prod:item.product, qty:item.qty,
    date:item.opDate, minutes:item.norm, tdg:item.tdg, comment:item.comment || '', sum:item.sum || 0,
  })));
  const upkp = applySiteConfig({
    id:'s_upkp', sourceSiteKey:'upkp:main', name:'УПКП — план (live)', sourceType:'url', sourceUrl:SNAP.upkpUrl || '',
    autoMonth:true, format:'upkp', unit:'секц.', workers:11, perWorker:0, teamTotal:true,
    resourceConfirmed:true, group:'Сборка', category:'Сборка',
    mappingNote:'Отдельная книга УПКП: секции и договоры. Строки общей загрузки в минутах сюда не дублируются.',
    fondMin:UPKP_DAILY_CAPACITY, workdays:[1,2,3,4,5], rows:upkpRows, overrides:{}, splits:{},
    snapshotVersion:UPKP_SNAPSHOT_VERSION,
  });
  sites.push(upkp);
  return sites.filter(site => !siteConfig.deleted.includes(site.sourceSiteKey));
}
function migrateSavedSite(site, fallback){
  ensureSiteIdentity(site);
  if(site.id === 's_upkp'){
    // Одноразовая миграция старого ошибочного 330; последующие правки владельца не трогаем.
    if(!siteConfig.sites['upkp:main'] && site.fondMin === 330){
      site.fondMin = UPKP_DAILY_CAPACITY; site.perWorker = 0; site.teamTotal = true;
    }
    if(!site.workers) site.workers = 11;
    if(site.perWorker === undefined) site.perWorker = 0;
    if(site.teamTotal === undefined) site.teamTotal = true;
    site.sourceType = 'url'; site.sourceUrl = fallback.sourceUrl; site.autoMonth = true;
    site.format = 'upkp'; site.unit = 'секц.';
    const staleSnapshot = site.snapshotVersion !== fallback.snapshotVersion;
    if((staleSnapshot || !Array.isArray(site.rows) || !site.rows.length) && fallback.rows.length){
      site.rows = fallback.rows;
      if(staleSnapshot){
        const overrides = site.overrides || {}, splits = site.splits || {};
        if(Object.keys(overrides).length || Object.keys(splits).length){
          site.sourceMigrationBackup = { fromVersion:site.snapshotVersion || 'legacy', savedAt:new Date().toISOString(), overrides, splits };
        }
        site.overrides = {}; site.splits = {}; site.sharedTs = 0;
      }else{
        const validIds = new Set(fallback.rows.map(row => row.id));
        const validKeys = new Set(fallback.rows.map(row => positionKey(row, fallback)));
        site.overrides = Object.fromEntries(Object.entries(site.overrides || {}).filter(([id]) => validIds.has(id) || String(id).startsWith('sp')));
        site.splits = Object.fromEntries(Object.entries(site.splits || {}).filter(([key]) => validKeys.has(key)));
      }
      site.snapshotVersion = fallback.snapshotVersion;
      site.syncOk = false; site.syncAt = null;
      site.syncErr = staleSnapshot
        ? 'Снимок УПКП обновлён; прежние переносы архивированы. Нажмите «Обновить» для проверки живой таблицы.'
        : 'Показана резервная копия. Нажмите «Обновить» для связи с Google Sheets.';
    }
  }else if(fallback){
    if(site.workers === undefined) site.workers = fallback.workers;
    if(site.perWorker === undefined) site.perWorker = fallback.perWorker;
    if(site.fondMin === undefined) site.fondMin = fallback.fondMin;
    if(site.teamTotal === undefined) site.teamTotal = fallback.teamTotal;
  }
  return applySiteConfig(site);
}
function loadSites(){
  siteConfig = readLocalSiteConfig();
  try{
    const raw = localStorage.getItem(LS_KEY);
    const defs = defaultSites();
    if(raw){
      let arr = JSON.parse(raw);
      if(Array.isArray(arr) && arr.length){
        arr = arr.filter(site => site.id !== 's_snap').map(site => ensureSiteIdentity(site));
        arr = arr.filter(site => !siteConfig.deleted.includes(site.sourceSiteKey));
        defs.forEach(def => {
          if(!arr.some(site => site.id === def.id) && !siteConfig.deleted.includes(def.sourceSiteKey)) arr.push(def);
        });
        return arr.map(site => migrateSavedSite(site, defs.find(def => def.id === site.id))).filter(site => !siteConfig.deleted.includes(site.sourceSiteKey));
      }
    }
    return defs;
  }catch(error){ return defaultSites(); }
}
function saveSites(){ localStorage.setItem(LS_KEY, JSON.stringify(state.sites)); }

/* ---------- v7: кэш одного прохода рендера ----------
   effRows/buildPositions/computeDays вызываются по много раз за один рендер
   (KPI, доска, сменные, общие секции). Кэш живёт только внутри синхронного
   passBegin()/passEnd() — снаружи функции работают как раньше, без кэша,
   поэтому прямые мутации данных вне рендера всегда дают свежий результат. */
const _pass = { depth:0, eff:null, pos:null, days:null };
function passBegin(){
  if(_pass.depth++ === 0){ _pass.eff = new WeakMap(); _pass.pos = new WeakMap(); _pass.days = new WeakMap(); }
}
function passEnd(){
  if(--_pass.depth <= 0){ _pass.depth = 0; _pass.eff = _pass.pos = _pass.days = null; }
}

/* ---------- совместный доступ: общие правки через репозиторий GitHub ----------
   Общие переносы/разбивки/примечания/ресурс хранятся в файлах папки shared/
   репозитория планировщика (например shared/ОП.json). Планировщик сам читает
   и пишет их через GitHub API из браузера — скрипты в Google-таблицах не нужны.
   Каждому пользователю нужен один раз ввести персональный токен GitHub
   (кнопка ⚙ рядом с «🔗 Совместно»). Конфликты: кто сохранил последним —
   того правка (LWW), при коллизии sha — повторная запись поверх. */
const GH_TOKEN_KEY = 'planner.ghToken', GH_REPO_KEY = 'planner.ghRepo';
function ghCfg(){
  const token = (localStorage.getItem(GH_TOKEN_KEY) || '').trim();
  let repo = (localStorage.getItem(GH_REPO_KEY) || '').trim();
  if(!repo && typeof location !== 'undefined' && location.hostname){
    const m = location.hostname.match(/^([^.]+)\.github\.io$/);
    const seg = (location.pathname || '').split('/').filter(Boolean)[0];
    if(m && seg) repo = m[1] + '/' + seg;
  }
  if(!repo) repo = 'BelloPotatik/planner-uchastkov';
  return { token, repo };
}
function sharedKey(site){
  // УПКП не должен читать общий default: там оказались переносы другого плана.
  if(site.id === 's_upkp' || site.format === 'upkp') return 'УПКП';
  if(site.sharedKey) return site.sharedKey;
  const u = site.sourceUrl || '';
  let m = u.match(/[?&]sheet=([^&]+)/);
  if(m) return decodeURIComponent(m[1]);
  m = u.match(/[?&]gid=(\d+)/);
  if(m) return 'gid' + m[1];
  return 'default';
}
function ghSharedPath(site){ return 'shared/' + sharedKey(site).replace(/[^\wа-яёА-ЯЁ-]+/g, '_') + '.json'; }
function ghHeaders(token, extra){ return Object.assign({ 'Authorization':'Bearer ' + token, 'Accept':'application/vnd.github+json', 'X-GitHub-Api-Version':'2022-11-28' }, extra || {}); }
function b64utf8(s){ return btoa(unescape(encodeURIComponent(s))); }
function ghApi(cfg, path){ return 'https://api.github.com/repos/' + cfg.repo + '/contents/' + encodeURIComponent(path); }
async function ghGetPath(path, token){
  const cfg = ghCfg();
  const res = await fetch(ghApi(cfg, path), { headers:ghHeaders(token || cfg.token), cache:'no-store' });
  if(res.status === 404) return { sha:null, content:null };
  if(!res.ok) throw new Error('HTTP ' + res.status);
  const json = await res.json();
  let text = '';
  try{ text = decodeURIComponent(escape(atob((json.content || '').replace(/\n/g, '')))); }catch(error){}
  return { sha:json.sha, content:text };
}
function ghGetFile(site){ return ghGetPath(ghSharedPath(site), ghCfg().token); }
let siteConfigRepoStatus = { state:'local', message:'Настройки сохранены в этом браузере' };
async function pushSiteConfiguration(){
  if(applyingRemote) return false;
  const cfg = ghCfg();
  if(!cfg.token){
    siteConfigRepoStatus = { state:'local', message:'Локально сохранено. Для записи в репозиторий подключите GitHub-токен.' };
    return false;
  }
  siteConfigRepoStatus = { state:'saving', message:'Сохраняем настройки в репозиторий…' };
  try{
    const file = await ghGetPath(SITE_CONFIG_PATH, cfg.token);
    const payload = {...siteConfig, updatedAt:siteConfig.updatedAt || Date.now()};
    const put = sha => fetch(ghApi(cfg, SITE_CONFIG_PATH), {
      method:'PUT', headers:ghHeaders(cfg.token, {'Content-Type':'application/json'}),
      body:JSON.stringify({ message:'planner: настройки участков', content:b64utf8(JSON.stringify(payload, null, 2)), sha:sha || undefined, branch:'main' }),
    });
    let response = await put(file.sha);
    if(response.status === 422 || response.status === 409){
      const fresh = await ghGetPath(SITE_CONFIG_PATH, cfg.token);
      response = await put(fresh.sha);
    }
    if(!response.ok) throw new Error('HTTP ' + response.status);
    siteConfigRepoStatus = { state:'saved', message:'Настройки участков сохранены в репозиторий' };
    return true;
  }catch(error){
    siteConfigRepoStatus = { state:'error', message:'Локально сохранено; репозиторий: ' + String(error.message || error) };
    return false;
  }
}
function siteFromSavedConfig(key, saved){
  if(!saved.sourceOperation) return null;
  const spec = SHARED.specForOperation(saved.sourceOperation);
  if(!spec) return null;
  return applySiteConfig({
    ...createSharedSite({...spec, ...saved, operation:saved.sourceOperation}),
    id:saved.id || spec.id, sourceSiteKey:key, rows:[], overrides:{}, splits:{},
  });
}
async function pullSiteConfiguration(){
  const cfg = ghCfg();
  try{
    const url = 'https://raw.githubusercontent.com/' + cfg.repo + '/main/' + SITE_CONFIG_PATH + '?_plannerRefresh=' + Date.now();
    const response = await fetch(url, {cache:'no-store'});
    if(response.status === 404) return false;
    if(!response.ok) throw new Error('HTTP ' + response.status);
    const remote = await response.json();
    if(!remote || remote.version !== 1 || !(remote.updatedAt > siteConfig.updatedAt)) return false;
    applyingRemote = true;
    try{
      siteConfig = {
        ...blankSiteConfig(), ...remote,
        categories:Array.isArray(remote.categories) ? remote.categories : DEFAULT_CATEGORIES.slice(),
        sites:remote.sites || {}, deleted:Array.isArray(remote.deleted) ? remote.deleted : [],
      };
      state.sites = state.sites.filter(site => !siteConfig.deleted.includes(siteKeyOf(site)));
      for(const site of state.sites) applySiteConfig(site);
      const defaults = defaultSites();
      for(const fallback of defaults){
        if(!state.sites.some(site => site.id === fallback.id)) state.sites.push(fallback);
      }
      for(const [key, saved] of Object.entries(siteConfig.sites)){
        if(siteConfig.deleted.includes(key) || state.sites.some(site => siteKeyOf(site) === key)) continue;
        const restored = siteFromSavedConfig(key, saved);
        if(restored) state.sites.push(restored);
      }
      if(!state.sites.some(site => site.id === state.activeId)) state.activeId = state.sites[0] && state.sites[0].id;
      localStorage.setItem(SITE_CONFIG_KEY, JSON.stringify(siteConfig));
      saveSites();
      siteConfigRepoStatus = { state:'saved', message:'Настройки участков загружены из репозитория' };
      return true;
    }finally{ applyingRemote = false; }
  }catch(error){
    siteConfigRepoStatus = { state:'error', message:'Не удалось прочитать настройки репозитория: ' + String(error.message || error) };
    return false;
  }
}
let sharedPushTimers = {};
let applyingRemote = false;
function scheduleSharedPush(site){
  if(!site || !site.sharedOn || applyingRemote) return;
  clearTimeout(sharedPushTimers[site.id]);
  sharedPushTimers[site.id] = setTimeout(() => pushShared(site), 800);
}
async function ghPut(site, cfg, sha, payload){
  return fetch(ghApi(cfg, ghSharedPath(site)), { method:'PUT',
    headers: ghHeaders(cfg.token, { 'Content-Type':'application/json' }),
    body: JSON.stringify({ message:'planner: ' + sharedKey(site), content: b64utf8(JSON.stringify(payload)), sha: sha || undefined, branch:'main' }) });
}
async function pushShared(site){
  if(!site || !site.sharedOn) return;
  const cfg = ghCfg();
  if(!cfg.token){ site.sharedOk = false; site.sharedErr = 'нет токена GitHub'; renderSyncStatus(); return; }
  const payload = { site: sharedKey(site), overrides: site.overrides || {},
    splits: site.splits || {}, notes: site.notes || '',
    planStrict: site.planStrict !== false,
    workers:site.workers || 0, perWorker:site.perWorker || 0,
    fondMin:site.fondMin || 0, teamTotal:!!site.teamTotal, ts:Date.now() };
  try{
    const f = await ghGetFile(site);
    let res = await ghPut(site, cfg, f.sha, payload);
    if(res.status === 422 && f.sha){
      // коллизия sha (кто-то записал между чтением и записью) — перечитываем и пишем поверх (LWW)
      const f2 = await ghGetFile(site);
      res = await ghPut(site, cfg, f2.sha, payload);
    }
    if(!res.ok) throw new Error('HTTP ' + res.status);
    site.sharedTs = payload.ts;
    site.sharedOk = true; site.sharedErr = null;
  }catch(err){
    site.sharedOk = false; site.sharedErr = String(err && err.message || err);
  }
  renderSyncStatus();
}
async function pullShared(site){
  if(!site || !site.sharedOn) return false;
  const cfg = ghCfg();
  if(!cfg.token){ site.sharedOk = false; site.sharedErr = 'нет токен GitHub'; return false; }
  try{
    const f = await ghGetFile(site);
    if(!f.content) return false;
    const st = JSON.parse(f.content);
    if(!st || !st.ts) return false;
    if(st.ts <= (site.sharedTs || 0)) return false; // нет свежих чужих правок
    site.sharedTs = st.ts;
    site.overrides = st.overrides || {};
    site.splits = st.splits || {};
    if(typeof st.notes === 'string') site.notes = st.notes;
    site.planStrict = st.planStrict !== false;
    // Старый per-site файл применяется только под центральной пользовательской конфигурацией.
    if(st.workers){ site.workers = st.workers; }
    if(typeof st.perWorker === 'number'){ site.perWorker = st.perWorker; }
    if(st.fondMin){ site.fondMin = st.fondMin; }
    if(typeof st.teamTotal === 'boolean'){ site.teamTotal = st.teamTotal; }
    if(site.id === 's_upkp' && !siteConfig.sites['upkp:main'] && st.fondMin === 330 && st.teamTotal === false){
      site.workers = 11; site.perWorker = 0; site.fondMin = UPKP_DAILY_CAPACITY; site.teamTotal = true;
    }
    applySiteConfig(site);
    site.sharedOk = true; site.sharedErr = null;
    applyingRemote = true;
    try{ saveSites(); }finally{ applyingRemote = false; }
    return true;
  }catch(err){
    site.sharedOk = false; site.sharedErr = String(err && err.message || err);
    return false;
  }
}
function enableSharing(){
  const site = activeSite();
  if(site.sharedOn){
    if(confirm('Отключить совместный режим для участка «' + site.name + '»?\nОбщие правки перестанут синхронизироваться с другими пользователями.')){
      site.sharedOn = false; saveSites(); render();
      toast('Совместный режим отключён для «' + site.name + '»');
    }
    return;
  }
  if(!ghCfg().token){
    const tok = prompt('Совместный режим: общие правки хранятся в папке shared/ репозитория GitHub\n' + ghCfg().repo + ' — скрипты в Google-таблицах не нужны.\n\nСоздайте токен: github.com → Settings → Developer settings → Personal access tokens → Fine-grained:\n  • доступ только к репозиторию planner-uchastkov\n  • разрешение Contents: Read and write\n\nВставьте токен сюда (хранится только в этом браузере):');
    if(tok === null) return;
    if(!tok.trim()){ toast('Токен не введён — совместный режим не включён'); return; }
    localStorage.setItem(GH_TOKEN_KEY, tok.trim());
  }
  site.sharedOn = true; site.sharedTs = 0;
  pushShared(site).then(() => {
    pullShared(site).then(changed => {
      saveSites(); render();
      toast(changed ? 'Совместно включено для «' + site.name + '», применены общие правки' : 'Совместно включено для «' + site.name + '»: правки видны всем, у кого введён токен');
    });
  });
}
function shareSettings(){
  const tok = prompt('Токен GitHub (хранится только в этом браузере).\nОставьте «••••••», чтобы не менять; пустое значение — удалить токен:', ghCfg().token ? '••••••' : '');
  if(tok === null) return;
  if(tok.trim() && tok.trim() !== '••••••') localStorage.setItem(GH_TOKEN_KEY, tok.trim());
  else if(!tok.trim()) localStorage.removeItem(GH_TOKEN_KEY);
  const repo = prompt('Репозиторий для общих правок (owner/name):', ghCfg().repo);
  if(repo !== null && repo.trim()) localStorage.setItem(GH_REPO_KEY, repo.trim());
  renderSyncStatus();
  toast('Настройки совместного режима сохранены');
}
function activeSite(){ return state.sites.find(s => s.id === state.activeId) || state.sites[0]; }
function fondOf(site){
  if(site.workers && site.perWorker) return site.workers * site.perWorker;
  return site.fondMin || (site.unit === 'секц.' ? 20 : 480);
}
function suggestFond(site){
  const byDay = {};
  site.rows.forEach(r => { if(r.date) byDay[r.date] = (byDay[r.date]||0) + (r.minutes||0); });
  const floor = site.unit === 'секц.' ? 20 : 480;
  const step = site.unit === 'секц.' ? 5 : 60;
  const mx = Math.max(floor, ...Object.values(byDay));
  return Math.ceil(mx/step)*step;
}
function effRows(site){
  if(_pass.eff && _pass.eff.has(site)) return _pass.eff.get(site);
  const splits = site.splits || {};
  const keys = new Set(Object.keys(splits));
  const out = [];
  for(const r of site.rows){
    const k = positionKey(r, site);
    if(keys.has(k)) continue; // строки позиции заменены частями разбивки
    out.push({...r, dateEff: (site.overrides && site.overrides[r.id]) || r.date || null});
  }
  for(const k of keys){
    const s = splits[k];
    if(!s) continue;
    for(const p of (s.parts||[])){
      out.push({ id:p.id, custOrder:s.custOrder, prodOrder:s.prodOrder, prod:s.prod,
        qty:p.qty, minutes:p.minutes, date:p.date, tdg:s.tdg || null,
        contractDate:s.contractDate || null, customer:s.customer || '',
        comment:[s.sourceComment, s.commentBase, p.label].filter(Boolean).join(' · '),
        executor:s.executor || '', operation:s.operation || '', sourceRow:s.sourceRow || null,
        sum:0, splitPart:true, baseKey:k,
        dateEff: (site.overrides && site.overrides[p.id]) || p.date || null });
    }
  }
  const effRes = out.filter(r => r.prodOrder || r.custOrder);
  if(_pass.eff) _pass.eff.set(site, effRes);
  return effRes;
}
/* ---------- split: разбивка позиции на части (напр. 30 шт = 10+10+10) ---------- */
function findPart(site, id){
  for(const k of Object.keys(site.splits || {})){
    const p = (site.splits[k].parts || []).find(x => x.id === id);
    if(p) return { key:k, part:p };
  }
  return null;
}
function distributeQty(total, n){
  // равными долями, остаток распределяем по первым частям: 30/4 -> [8,8,7,7]
  const base = Math.floor(total / n), rem = total % n;
  return Array.from({length:n}, (_, i) => base + (i < rem ? 1 : 0));
}
function distributeMinutes(total, qtys, totalQty){
  // минуты пропорционально количеству, погрешность округления — в последнюю часть
  const parts = qtys.map(q => Math.round(total * q / totalQty));
  const drift = total - parts.reduce((a,b) => a + b, 0);
  parts[parts.length-1] += drift;
  return parts;
}
function buildSplitRecord(row, posRows, n, site){
  const totalQty = posRows.reduce((a,r) => a + (r.qty||0), 0);
  return buildSplitRecordQtys(row, posRows, distributeQty(totalQty, n), undefined, site);
}
/* разбивка с произвольными (не обязательно равными) количествами частей */
function buildSplitRecordQtys(row, posRows, qtys, commentBase, site){
  const totalQty = posRows.reduce((a,r) => a + (r.qty||0), 0);
  const totalMin = posRows.reduce((a,r) => a + (r.minutes||0), 0);
  const mins = distributeMinutes(totalMin, qtys, totalQty);
  const key = positionKey(row, site);
  const parts = qtys.map((q, i) => ({
    id:'sp' + hashStr(key + '|' + i + '|' + q + '|' + mins[i]),
    qty:q, minutes:mins[i], date:null,
    label:'часть ' + (i+1) + '/' + qtys.length + ' (' + q + ' шт)'
  }));
  return { key, custOrder:row.custOrder, prodOrder:row.prodOrder, prod:row.prod,
    tdg:row.tdg || null, contractDate:row.contractDate || null, customer:row.customer || '',
    sourceComment:row.comment || '', executor:row.executor || '', operation:row.operation || '', sourceRow:row.sourceRow || null,
    commentBase:commentBase || ('разбивка ' + totalQty + ' шт'), parts };
}
/* разумное дробление по фонду: день заполняется под ёмкость, остаток — следующие дни.
   Пачки получаются НЕ обязательно равными (напр. 20 шт при ёмкости 10 и занятости
   первого дня — 5 + 10 + 5). Если последняя пачка крошечная (< четверти дневной
   ёмкости) — части перебалансируются равномерно, чтобы не плодить «хвосты». */
function splitQtysByFond(totalQty, perUnit, fond, firstAvail){
  if(!totalQty || totalQty <= 0) return [];
  const per = perUnit > 0 ? perUnit : 1;
  const cap = Math.max(1, Math.floor(fond / per));
  if(totalQty * per <= (firstAvail != null ? firstAvail : fond)) return [totalQty]; // всё влезает в первый день
  const out = [];
  let left = totalQty;
  let avail = firstAvail != null ? firstAvail : fond;
  let guard = 0;
  while(left > 0 && guard++ < 500){
    let fit = Math.floor(avail / per);
    if(fit <= 0){ avail = fond; continue; } // день занят другими — следующий день полный
    const take = Math.min(left, fit);
    out.push(take);
    left -= take;
    avail = fond;
  }
  if(left > 0) out.push(left); // страховка от guard
  if(out.length > 1 && out[out.length-1] < Math.max(1, Math.floor(cap/4))){
    return distributeQty(totalQty, out.length); // крошечный хвост → равномерно
  }
  return out;
}
/* раскладка частей по дням: минимум подряд идущих рабочих дней от якорной даты */
function assignSplitDates(site, rec, anchorISO){
  const fond = fondOf(site);
  let d = nextWorkDay(site, anchorISO || todayISO());
  let used = 0;
  for(const p of rec.parts){
    if(used > 0 && used + p.minutes > fond){ d = nextWorkDay(site, addDays(d, 1)); used = 0; }
    p.date = d; used += p.minutes;
  }
  return rec;
}
function isWorkday(site, iso){ return site.workdays.includes(wdOf(iso)); }
function nextWorkDay(site, iso){
  let d = iso;
  for(let i=0;i<400;i++){
    if(isWorkday(site, d)) return d;
    d = addDays(d, 1);
  }
  return d;
}

/* ---------- positions (split detection) ---------- */
function buildPositions(site){
  if(_pass.pos && _pass.pos.has(site)) return _pass.pos.get(site);
  const m = new Map();
  for(const r of effRows(site)){
    const k = positionKey(r, site);
    if(!m.has(k)) m.set(k, {key:k, custOrder:r.custOrder, prodOrder:r.prodOrder,
      customer:r.customer || '', prod:r.prod, parts:[]});
    m.get(k).parts.push(r);
  }
  for(const p of m.values()){
    p.dates = [...new Set(p.parts.map(r => r.dateEff).filter(Boolean))].sort();
    p.split = p.dates.length > 1;
    p.totalQty = p.parts.reduce((a,r) => a + (r.qty||0), 0);
    p.totalMin = p.parts.reduce((a,r) => a + (r.minutes||0), 0);
  }
  if(_pass.pos) _pass.pos.set(site, m);
  return m;
}

/* ---------- days ---------- */
function computeDays(site){
  const ck = (state.rangeFrom||'') + '|' + (state.rangeTo||'') + '|' + todayISO();
  if(_pass.days && _pass.days.has(site)){
    const c = _pass.days.get(site);
    if(c.k === ck) return c.v;
  }
  const rows = effRows(site);
  const map = new Map();
  const nodate = [];
  for(const r of rows){
    if(!r.dateEff){ nodate.push(r); continue; }
    if(!map.has(r.dateEff)) map.set(r.dateEff, []);
    map.get(r.dateEff).push(r);
  }
  const dates = [...map.keys()].sort();
  const t = todayISO();
  // по умолчанию показываем только с актуальной (сегодняшней) даты;
  // более ранние даты открываются вручную через выбор периода (rangeFrom)
  const minD = state.rangeFrom || t;
  let to = dates.length ? dates[dates.length-1] : addDays(t, 10);
  if(to < minD) to = minD;
  if(state.rangeTo && state.rangeTo < to) to = state.rangeTo;
  if(state.rangeTo && state.rangeTo < minD) to = minD; // некорректный период — покажем хотя бы старт
  let all = [];
  for(let d = minD; d <= to; d = addDays(d,1)){
    const rowsD = map.get(d) || [];
    all.push({ date:d, rows:rowsD, total:rowsD.reduce((a,r)=>a+(r.minutes||0),0),
      work:isWorkday(site,d), has:map.has(d) });
  }
  all = all.filter(d => d.has || d.work); // рабочие дни + дни с данными
  const daysRes = { days:all, nodate };
  if(_pass.days) _pass.days.set(site, { k:ck, v:daysRes });
  return daysRes;
}

/* ---------- render ---------- */
function rowStatus(r){
  if(r.tdg && r.dateEff && r.dateEff > r.tdg) return 'red';
  if(r.contractDate && r.dateEff && r.dateEff > r.contractDate) return 'amber';
  return 'green';
}
function rowContextHTML(r){
  const parts = [];
  if(r.comment) parts.push('<span class="card-comment" title="Комментарий: ' + esc(r.comment) + '">💬 ' + esc(r.comment) + '</span>');
  if(r.executor) parts.push('<span class="card-executor" title="Исполнитель: ' + esc(r.executor) + '">Исп.: ' + esc(r.executor) + '</span>');
  return parts.length ? '<div class="card-context">' + parts.join('') + '</div>' : '';
}
function rowContextCells(r){
  return '<td class="l-context" title="' + esc(r.comment || '') + '">' + esc(r.comment || '—') + '</td>' +
    '<td class="l-context" title="' + esc(r.executor || '') + '">' + esc(r.executor || '—') + '</td>';
}
function assemblyDateHTML(r, site){
  if(!site || site.group !== 'Сборка') return '';
  return '<button type="button" class="mini assembly-date-edit" data-date-edit="' + esc(r.id) + '" title="Изменить дату сборки через календарь" aria-label="Изменить дату сборки ' + esc(displayProduct(r, site)) + '">📅 ' + fmtD(r.dateEff) + '</button>';
}
function cardHTML(r, posMap, site, compact){
  const p = posMap.get(positionKey(r, site));
  const st = rowStatus(r);
  const hide = filterHide(r);
  const isUPKP = site && site.format === 'upkp';
  const orderLabel = isUPKP ? upkpOrderLabel(r.custOrder) : (r.prodOrder || r.custOrder);
  const product = displayProduct(r, site);
  const meta = isUPKP ? displayCustomer(r, site) :
    (r.custOrder + (r.customer ? ' · ' + r.customer : ''));
  let badges = '';
  if(p && p.split){
    const idx = p.dates.indexOf(r.dateEff) + 1;
    badges += '<span class="badge split" title="' + esc(p.parts.map(x => fmtD(x.dateEff)+' — '+(x.qty||0)+' шт').join('; ')) + '">⤨ позиция на ' + p.dates.length + ' дн. (' + idx + '/' + p.dates.length + ')</span>';
  }
  if(r.tdg){
    const late = r.dateEff && r.dateEff > r.tdg;
    badges += '<span class="badge ' + (late?'tdg':'') + '">ТДГ ' + fmtD(r.tdg) + '</span>';
  }
  if(r.contractDate) badges += '<span class="badge">договор ' + fmtD(r.contractDate) + '</span>';
  if(compact){
    return '<div class="card compact st-' + st + (hide?'" style="display:none':'') + '">' +
      '<div class="card-top"><span>' + esc(orderLabel) + '</span><span class="card-min">' + uMin(site, r.minutes) + '</span></div>' +
      '<div class="card-prod" title="' + esc(r.prod) + '">' + esc(product) + '</div>' +
      '<div class="card-meta">' + (r.qty ? r.qty + ' шт' : '') + '</div>' + rowContextHTML(r) +
      (badges?'<div class="badges">' + badges + '</div>':'') +
    '</div>';
  }
  return '<div class="card st-' + st + (hide?'" style="display:none':'') + '" draggable="true" data-id="' + esc(r.id) + '">' +
    '<div class="card-top"><span>' + esc(orderLabel) + '</span><span class="card-min">' + uMin(site, r.minutes) +
    (r.splitPart ? '' : ' <button class="mini" data-split="' + esc(r.id) + '" title="Разбить позицию на части">✂</button>') +
    '</span></div>' +
    '<div class="card-prod" title="' + esc(r.prod) + '">' + esc(product) + '</div>' +
    '<div class="card-meta">' + esc(meta) + (r.qty?' · ' + r.qty + ' шт':'') + '</div>' + rowContextHTML(r) +
    (badges?'<div class="badges">' + badges + '</div>':'') +
    (site.group === 'Сборка' ? '<div class="assembly-date-row"><span>Дата сборки</span>' + assemblyDateHTML(r, site) + '</div>' : '') +
  '</div>';
}
function filterHide(r){
  const q = ($('searchInput').value || '').trim().toLowerCase();
  if(q){
    const hay = [r.prodOrder, r.custOrder, r.prod, r.customer, r.comment, r.executor, r.operation].join(' ').toLowerCase();
    if(!hay.includes(q)) return true;
  }
  if($('onlyBad').checked && rowStatus(r) === 'green') return true;
  return false;
}
function renderKPIs(site, daysInfo){
  const t = todayISO();
  const rows = effRows(site);
  const fond = fondOf(site);
  const todayRows = rows.filter(r => r.dateEff === t);
  const todayLoad = todayRows.reduce((a,r)=>a+(r.minutes||0),0);
  const allMin = rows.reduce((a,r)=>a+(r.minutes||0),0);
  const workDaysAll = daysInfo.days.filter(d => d.work);
  const periodCap = Math.max(1, workDaysAll.length) * fond;
  const futureDays = daysInfo.days.filter(d => d.work && d.date >= t);
  const futureRowsMin = rows.filter(r => r.dateEff >= t).reduce((a,r)=>a+(r.minutes||0),0);
  const avail = Math.max(0, futureDays.length * fond - futureRowsMin);
  const posCount = buildPositions(site).size;
  const chip = (label, val, cls='') => '<div class="kpi ' + cls + '"><b>' + val + '</b><span>' + label + '</span></div>';
  const unit = site.unit || 'мин';
  $('kpis').innerHTML =
    chip('позиций', posCount) +
    chip('операций', rows.length) +
    chip('сумма норм, ' + unit, fmtMin(allMin)) +
    chip('фонд, ' + unit + '/день', fmtMin(fond)) +
    chip('загрузка сегодня', (todayLoad?Math.round(todayLoad/fond*100):0) + '%' + (todayRows.length?' · ' + uMin(site, todayLoad):''), todayLoad>fond?'bad':'') +
    chip('загрузка периода', Math.round(allMin/periodCap*100) + '%') +
    chip('доступный ресурс, ' + unit, fmtMin(avail), avail < 0 ? 'bad' : (avail < fond*3 ? 'warn' : ''));
}
function renderOperationVisibility(site){
  const host = $('operationVisibility');
  if(!host) return;
  const rows = effRows(site), today = todayISO();
  const past = rows.filter(row => row.dateEff && row.dateEff < today);
  const future = rows.filter(row => row.dateEff && row.dateEff > today);
  const todayCount = rows.filter(row => row.dateEff === today).length;
  // При ручном периоде пользователь уже видит нужный архив. Подсказка нужна
  // именно на исходном виде «с сегодня», где пустой день легко принять за сбой.
  if(state.rangeFrom || state.rangeTo || todayCount || (!past.length && !future.length)){
    host.hidden = true; host.innerHTML = ''; return;
  }
  const nextDate = future.map(row => row.dateEff).sort()[0] || null;
  const nextCount = nextDate ? future.filter(row => row.dateEff === nextDate).length : 0;
  const firstPast = past.map(row => row.dateEff).sort()[0] || null;
  let html = '<div class="operation-visibility-copy"><b>Сегодня операций нет — данные не потеряны</b>' +
    '<span>В Google Таблице: ' + rows.length + ' операций · прошлых: ' + past.length + ' · предстоящих: ' + future.length + '</span></div>' +
    '<div class="operation-visibility-actions">';
  if(nextDate) html += '<button class="btn primary" data-nearest-operation="' + nextDate + '">Ближайшая: ' + fmtD(nextDate) + ' · ' + nextCount + ' опер.</button>';
  if(firstPast) html += '<button class="btn" data-show-past="' + firstPast + '">Прошлые</button>';
  html += '<button class="btn" data-show-all-operations>Весь период</button></div>';
  host.innerHTML = html; host.hidden = false;
}
function render(){
  passBegin();
  try{
  const site = activeSite();
  if(!site){ return; }
  const entering = state._enterView; state._enterView = false;
  const daysInfo = computeDays(site);
  const posMap = buildPositions(site);
  renderKPIs(site, daysInfo);
  renderOperationVisibility(site);
  renderSiteSelect();
  renderSyncStatus();
  const team = !!site.teamTotal;
  $('workersInput').value = site.workers || '';
  $('resMul').style.display = team ? 'none' : '';
  $('perWorkerWrap').style.display = team ? 'none' : '';
  if(!team) $('perWorkerInput').value = site.perWorker || '';
  $('fondTotalInput').readOnly = !team;
  $('fondTotalInput').value = fondOf(site);
  $('fondUnit').textContent = (site.unit || 'мин') + '/день' + (team ? ' (на бригаду)' : '');
  const pinBtn = $('btnPinBoard');
  if(pinBtn){
    pinBtn.classList.toggle('primary', !!site.onBoard);
    pinBtn.innerHTML = '<svg><use href="#i-pin"/></svg><span>' + (site.onBoard ? 'На общей доске ✓' : 'На общую доску') + '</span>';
    pinBtn.title = site.onBoard
      ? 'Участок выведен на общую доску: виден в обзорных секциях видов «Доска», «Список дней» и «Все участки». Нажмите, чтобы убрать'
      : 'Вывести этот участок на общую доску — его дни и загрузка (со своим ресурсом) будут видны в видах «Доска», «Список дней» и «Все участки»';
  }
  renderDayPanel(site, daysInfo);
  if(state.view === 'board') renderBoard(site, daysInfo, posMap);
  else if(state.view === 'list') renderList(site, daysInfo, posMap);
  else if(state.view === 'shifts') renderShifts();
  else if(state.view === 'multi') renderMulti();
  if(entering){
    const host = {board:$('boardLayout'), list:$('listView'), shifts:$('shiftsView'), multi:$('multiView')}[state.view];
    markEnter(host);
  }
  if(typeof renderWorkspace === 'function') renderWorkspace(site);
  }finally{ passEnd(); }
}
/* v7: пакетный рендер — серия событий за один кадр даёт один рендер, а не несколько */
let renderQueued = false;
function scheduleRender(){
  if(renderQueued) return;
  renderQueued = true;
  raf(() => { renderQueued = false; render(); });
}
function showOnly(name){
  $('boardLayout').style.display = name === 'board' ? '' : 'none';
  $('listView').hidden = name !== 'list';
  $('shiftsView').hidden = name !== 'shifts';
  $('multiView').hidden = name !== 'multi';
}
/* стаггер-появление: класс живёт ~1.1 с, потом снимается, чтобы не перезапускать анимации */
function markEnter(el){
  if(!el || reducedMotion()) return;
  el.classList.add('enter-anim');
  clearTimeout(el._et);
  el._et = setTimeout(() => el.classList.remove('enter-anim'), 1150);
}
/* скользящая плашка переключателя видов */
function segSync(){
  const seg = document.querySelector && document.querySelector('.seg');
  const pill = $('segPill');
  if(!seg || !pill) return;
  const on = seg.querySelector('button.on');
  if(!on){ pill.style.width = '0'; return; }
  pill.style.width = on.offsetWidth + 'px';
  pill.style.transform = 'translateX(' + (on.offsetLeft - 3) + 'px)';
}
function renderSiteSelect(){
  const sel = $('siteSelect');
  sel.innerHTML = state.sites.map(s => '<option value="' + esc(s.id) + '"' + (s.id === activeSite().id?' selected':'') + '>' + esc(s.name) + '</option>').join('');
  // кастомный выпадающий список поверх скрытого нативного
  const label = $('csLabel'), menu = $('csMenu');
  const act = activeSite();
  if(label) label.textContent = act ? act.name : 'Участок';
  if(menu) menu.innerHTML = state.sites.map((s, i) =>
    '<button type="button" class="cs-item' + (s.id === state.activeId ? ' on' : '') + '" data-site-id="' + esc(s.id) + '" style="--i:' + i + '" role="option">' + esc(s.name) + '</button>').join('');
}
function csOpen(open){
  const wrap = $('siteCSelect');
  const menu = $('csMenu');
  if(!wrap || !menu) return;
  if(open === undefined) open = !menu.classList.contains('open');
  if(open){
    const b = $('csBtn').getBoundingClientRect();
    menu.style.left = b.left + 'px';
    menu.style.top = (b.bottom + 8) + 'px';
    menu.style.minWidth = Math.max(b.width, 180) + 'px';
  }
  wrap.classList.toggle('open', open); // поворот шеврона
  menu.classList.toggle('open', open);
}
function renderDayPanel(site, daysInfo){
  const fond = fondOf(site);
  const el = $('dayList');
  el.innerHTML = '';
  for(const d of daysInfo.days){
    if(!d.has) continue;
    const pct = Math.round(d.total / fond * 100);
    const item = document.createElement('div');
    item.className = 'day-item' + (state.selectedDate === d.date ? ' sel' : '');
    item.innerHTML = '<div class="d1"><span>' + WD[wdOf(d.date)] + ', ' + fmtD(d.date) + '</span><span>' + pct + '%</span></div>' +
      '<div class="d2"><span>' + d.rows.length + ' опер.</span><span>' + uMin(site, d.total) + ' / ' + uMin(site, fond) + '</span></div>' +
      '<div class="mini-bar"><i class="' + (pct>100?'over':'') + '" style="width:' + Math.min(100,pct) + '%"></i></div>';
    item.addEventListener('click', () => jumpToDate(d.date));
    el.appendChild(item);
  }
  if(daysInfo.nodate.length){
    const item = document.createElement('div');
    item.className = 'day-item empty';
    item.innerHTML = '<div class="d1"><span>Без даты</span><span>' + daysInfo.nodate.length + '</span></div>';
    item.addEventListener('click', () => jumpToDate('nodate'));
    el.appendChild(item);
  }
}
function renderBoard(site, daysInfo, posMap){
  showOnly('board');
  const row = $('daysRow');
  row.innerHTML = '';
  const fond = fondOf(site);
  const t = todayISO();
  const cols = [];
  if(daysInfo.nodate.length) cols.push({date:'nodate'});
  cols.push(...daysInfo.days);
  let colIdx = 0;
  for(const d of cols){
    const col = document.createElement('div');
    if(col.style && col.style.setProperty) col.style.setProperty('--i', Math.min(colIdx++, 14));
    if(d.date === 'nodate'){
      col.className = 'day-col nodate';
      col.dataset.date = 'nodate';
      col.innerHTML = '<div class="day-head"><div class="dt">Без даты</div><div class="di">' + daysInfo.nodate.length + ' операций — распределите</div></div>';
      const body = document.createElement('div');
      body.className = 'day-body';
      daysInfo.nodate.forEach(r => body.insertAdjacentHTML('beforeend', cardHTML(r, posMap, site)));
      col.appendChild(body);
      row.appendChild(col);
      continue;
    }
    const pct = Math.round(d.total / fond * 100);
    const isToday = d.date === t;
    col.className = 'day-col' + (pct > 100 ? ' over' : '') + (isToday ? ' today' : '') + (state.selectedDate === d.date ? ' flash' : '');
    col.dataset.date = d.date;
    col.innerHTML = '<div class="day-head"><div class="dt">' + fmtD(d.date) + '<small>' + WD[wdOf(d.date)] + (d.work?'':' · вых') + '</small>' +
      (isToday ? '<span class="today-chip">сегодня</span>' : '') + '</div>' +
      '<div class="loadbar"><i style="width:' + Math.min(100,pct) + '%"></i></div>' +
      '<div class="di">' + uMin(site, d.total) + ' / ' + uMin(site, fond) + ' · ' + d.rows.length + ' опер.' + (pct>100?' · <b>перегруз +' + fmtMin(d.total-fond) + '</b>':'') + '</div></div>';
    if(isToday) col.insertAdjacentHTML('beforeend', '<i class="ring"></i>'); // v7: GPU-кольцо today
    const body = document.createElement('div');
    body.className = 'day-body';
    if(!d.rows.length) body.innerHTML = '<div class="day-empty">пусто — перетащите сюда операцию</div>';
    d.rows.slice().sort((a,b) => (a.prodOrder||'').localeCompare(b.prodOrder||'')).forEach(r => body.insertAdjacentHTML('beforeend', cardHTML(r, posMap, site)));
    col.appendChild(body);
    row.appendChild(col);
  }
  // общая доска: компактные секции выведенных участков — в формате доски
  const boardEl = $('board');
  if(boardEl.querySelectorAll) boardEl.querySelectorAll('.sub-board').forEach(el => { if(el.remove) el.remove(); });
  for(const ps of pinnedSites()) boardEl.insertAdjacentHTML('beforeend', boardSectionHTML(ps));
}
let listGeneration = 0;
function listDayHTML(site, d, posMap){
  const fond = fondOf(site);
  const pct = Math.round(d.total / fond * 100);
  const isUPKP = site.format === 'upkp';
  let html = '<div class="l-head' + (pct>100?' over':'') + '"><span>' + WD[wdOf(d.date)] + ', ' + fmtDFull(d.date) +
    (d.date === todayISO() ? ' <span class="today-chip">сегодня</span>' : '') + '</span>' +
    '<span class="lh-sub">' + d.rows.length + ' операций</span>' +
    '<span class="lh-sub">' + uMin(site, d.total) + ' / ' + uMin(site, fond) + ' (' + pct + '%)' + (pct>100?' — ПЕРЕГРУЗ':'') + '</span></div>';
  html += isUPKP
    ? '<table class="l-table"><thead><tr><th>Договор</th><th>Контрагент</th><th>Продукция</th><th>Кол-во</th><th>Секции</th><th>ТДГ</th><th>Комментарий</th><th>Исполнитель</th><th>Дата по договору</th><th>Статус</th></tr></thead><tbody>'
    : '<table class="l-table"><thead><tr><th>Заказ на пр-во</th><th>Заказ покупателя</th><th>Контрагент</th><th>Продукция</th><th>Кол-во</th><th>Норма</th><th>ТДГ</th><th>Комментарий</th><th>Исполнитель</th><th>Договор</th><th>Статус</th></tr></thead><tbody>';
  for(const r of d.rows.slice().sort((a,b)=>(a.prodOrder||'').localeCompare(b.prodOrder||''))){
    if(filterHide(r)) continue;
    const st = rowStatus(r);
    const p = posMap.get(positionKey(r, site));
    html += '<tr draggable="true" data-id="' + esc(r.id) + '" title="Перетащите строку на другой день; на телефоне удерживайте её" class="' + (st==='red'?'l-row-red':st==='amber'?'l-row-amber':'') + '">' +
      (isUPKP ? '' : '<td class="num">' + esc(r.prodOrder) + '</td>') +
      '<td class="num">' + esc(isUPKP ? upkpOrderLabel(r.custOrder) : r.custOrder) + '</td>' +
      '<td title="' + esc(r.custOrder) + '">' + esc(displayCustomer(r, site)) + '</td>' +
      '<td class="l-prod" title="' + esc(r.prod) + '">' + esc(displayProduct(r, site)) + '</td>' +
      '<td class="num">' + (r.qty||'') + (p && p.split?' <span class="badge split" title="Позиция разбита на ' + p.dates.length + ' дней: ' + esc(p.dates.map(x=>fmtD(x)).join(', ')) + '">⤨' + p.dates.length + '</span>':'') + '</td>' +
      '<td class="num">' + fmtMin(r.minutes) + '</td><td class="num">' + fmtD(r.tdg) + '</td>' + rowContextCells(r) +
      '<td class="num">' + fmtD(r.contractDate) + '</td><td>' +
      (st==='red'?'⚠ после ТДГ':st==='amber'?'⚠ после договора':'ок') + assemblyDateHTML(r, site) + '</td></tr>';
  }
  return html + '</tbody></table>';
}
function renderList(site, daysInfo, posMap){
  showOnly('list');
  const generation = ++listGeneration;
  const lv = $('listView');
  lv.innerHTML = '';
  // Пустой рабочий день тоже нужен: в него можно перенести операцию мышью или пальцем.
  const show = daysInfo.days.filter(d => d.has || d.work);
  let index = 0;
  const appendDay = () => {
    const d = show[index];
    const wrap = document.createElement('div');
    if(wrap.style && wrap.style.setProperty) wrap.style.setProperty('--i', Math.min(index, 14));
    wrap.className = 'l-day' + (state.selectedDate === d.date ? ' flash' : '');
    wrap.dataset.date = d.date;
    wrap.innerHTML = listDayHTML(site, d, posMap);
    lv.appendChild(wrap);
    index++;
  };
  const appendSections = () => {
    for(const ps of pinnedSites()) lv.insertAdjacentHTML('beforeend', listSectionHTML(ps));
  };
  if(!show.length){
    lv.innerHTML = '<div class="hint">Нет операций с датами в выбранном периоде.</div>';
    appendSections();
    return;
  }
  const paint = () => {
    if(generation !== listGeneration || state.view !== 'list' || activeSite() !== site) return;
    const started = typeof performance !== 'undefined' ? performance.now() : 0;
    do{ appendDay(); }while(index < show.length && (index < 4 || performance.now() - started < 6));
    if(index < show.length) raf(paint);
    else appendSections();
  };
  if(typeof requestAnimationFrame === 'undefined'){
    while(index < show.length) appendDay();
    appendSections();
  }else paint();
}

/* ---------- вид «Сменные»: задания всех участков по дням ---------- */
function renderShifts(){
  showOnly('shifts');
  const el = $('shiftsView');
  const t = todayISO();
  const from = state.rangeFrom || t;
  let html = '<div class="hint" style="margin-bottom:12px">Сменные задания всех участков' +
    (state.rangeFrom ? ' с ' + fmtDFull(state.rangeFrom) : ' с сегодняшнего дня (' + fmtDFull(t) + ')') +
    (state.rangeTo ? ' по ' + fmtDFull(state.rangeTo) : '') +
    '. Более ранние даты скрыты — откройте их через «Период» в верхней панели.</div>';
  // собираем объединённый список дат по всем участкам
  const perSite = state.sites.map(s => ({ site:s, info:computeDays(s) }));
  const dateSet = new Set();
  for(const ps of perSite) for(const d of ps.info.days) if(d.date >= from && (!state.rangeTo || d.date <= state.rangeTo) && (d.has || d.work)) dateSet.add(d.date);
  const dates = [...dateSet].sort();
  if(!dates.length){ el.innerHTML = html + '<div class="hint">В выбранном периоде сменных заданий нет.</div>'; return; }
  let shiftIdx = 0;
  for(const date of dates){
    const isToday = date === t;
    let dayHtml = '<div class="shift-day" style="--i:' + Math.min(shiftIdx++, 14) + '"><div class="l-head"><span>' + WD[wdOf(date)] + ', ' + fmtDFull(date) +
      (isToday ? ' <span class="today-chip">сегодня</span>' : '') + '</span></div>';
    let any = false;
    for(const ps of perSite){
      const d = ps.info.days.find(x => x.date === date);
      if(!d || !d.rows.length) continue;
      any = true;
      const fond = fondOf(ps.site);
      const pct = Math.round(d.total / fond * 100);
      dayHtml += '<div class="shift-site"><div class="ss-head"><span>' + esc(ps.site.name) + '</span>' +
        '<span class="ss-bar"><i class="' + (pct>100?'over':'') + '" style="width:' + Math.min(100,pct) + '%"></i></span>' +
        '<span class="ss-load">' + uMin(ps.site, d.total) + ' / ' + uMin(ps.site, fond) + ' · ' + pct + '%' + (pct>100?' — ПЕРЕГРУЗ':'') + '</span></div>' +
        '<table class="l-table"><thead><tr><th>Заказ на пр-во</th><th>Заказ покупателя</th><th>Контрагент</th><th>Продукция</th><th>Кол-во</th><th>Норма</th><th>ТДГ</th><th>Комментарий</th><th>Исполнитель</th></tr></thead><tbody>';
      for(const r of d.rows.slice().sort((a,b)=>(a.prodOrder||'').localeCompare(b.prodOrder||''))){
        const st = rowStatus(r);
        dayHtml += '<tr class="' + (st==='red'?'l-row-red':st==='amber'?'l-row-amber':'') + '">' +
          '<td class="num">' + esc(r.prodOrder) + '</td><td class="num">' + esc(r.custOrder) + '</td>' +
          '<td>' + esc(displayCustomer(r, ps.site)) + '</td><td class="l-prod" title="' + esc(r.prod) + '">' + esc(displayProduct(r, ps.site)) + '</td>' +
          '<td class="num">' + (r.qty||'') + '</td><td class="num">' + fmtMin(r.minutes) + '</td>' +
          '<td class="num">' + fmtD(r.tdg) + '</td>' + rowContextCells(r) + '</tr>';
      }
      dayHtml += '</tbody></table></div>';
    }
    if(!any) dayHtml += '<div class="hint" style="padding:10px 15px">заданий нет</div>';
    dayHtml += '</div>';
    html += dayHtml;
  }
  el.innerHTML = html;
}

/* ---------- вид «Все участки»: лента дней по каждому участку ---------- */
function renderMulti(){
  showOnly('multi');
  const el = $('multiView');
  const t = todayISO();
  const sites = state.sites.filter(s => s.onBoard);
  if(!sites.length){
    el.innerHTML = '<div class="hint" style="max-width:640px;margin:24px auto;text-align:center">На общую доску пока ничего не выведено.<br><br>' +
      'Откройте нужный участок и нажмите <b>«📌 На общую доску»</b> в верхней панели — ' +
      'участок со своим ресурсом появится здесь, а также в общих секциях видов «Доска» и «Список дней».</div>';
    return;
  }
  let html = '';
  let laneIdx = 0;
  for(const site of sites){
    const info = computeDays(site);
    const fond = fondOf(site);
    const res = siteResLine(site);
    const todayD = info.days.find(d => d.date === t);
    const todayPct = todayD ? Math.round(todayD.total / fond * 100) : 0;
    const workDays = info.days.filter(d => d.work);
    const totalMin = info.days.reduce((a,d) => a + d.total, 0);
    const periodPct = workDays.length ? Math.round(totalMin / (workDays.length * fond) * 100) : 0;
    const overDays = info.days.filter(d => d.total > fond).length;
    html += '<div class="lane" style="--i:' + Math.min(laneIdx++, 14) + '"><div class="lane-head">' +
      '<span class="ln-name">' + esc(site.name) + '</span>' +
      '<span class="ln-res" title="Ресурс участка">⚙ ' + esc(res) + '</span>' +
      '<span class="ln-kpis">' +
        '<span class="kpi' + (todayPct>100?' bad':'') + '"><b>' + todayPct + '%</b><span>сегодня</span></span>' +
        '<span class="kpi"><b>' + periodPct + '%</b><span>загрузка периода</span></span>' +
        (overDays ? '<span class="kpi bad"><b>' + overDays + '</b><span>дней перегруза</span></span>' : '') +
      '</span></div><div class="lane-cells">';
    for(const d of info.days){
      const pct = fond ? Math.round(d.total / fond * 100) : 0;
      const cls = !d.work ? 'off' : (d.total > fond ? 'l3' : (d.total > fond*0.85 ? 'l2' : (d.total > 0 ? 'l1' : 'l0')));
      const tip = fmtDFull(d.date) + ' · ' + uMin(site, d.total) + ' / ' + uMin(site, fond) + ' (' + pct + '%) · ' + d.rows.length + ' опер.';
      html += '<div class="lcell ' + cls + (d.date === t ? ' today' : '') + '" data-site="' + esc(site.id) + '" data-date="' + d.date + '" title="' + esc(tip) + '">' +
        '<div class="lc-d">' + fmtD(d.date) + '</div>' +
        '<div class="lc-p">' + (d.work ? pct + '%' : '—') + '</div>' +
        '<div class="lc-n">' + (d.rows.length || '') + '</div></div>';
    }
    html += '</div></div>';
  }
  el.innerHTML = html;
}

/* ---------- общая доска: подключаемые участки (site.onBoard) ---------- */
function siteResLine(site){
  return site.teamTotal
    ? 'бригада ' + (site.workers||'') + ' чел · норма ' + uMin(site, fondOf(site)) + ' на бригаду'
    : (site.workers||'') + ' × ' + uMin(site, site.perWorker||0) + ' = ' + uMin(site, fondOf(site)) + '/день';
}
function pinnedSites(){
  return state.sites.filter(s => s.onBoard && s.id !== state.activeId);
}
function sbTitleHTML(site){
  return '<div class="sb-title">📌 ' + esc(site.name) +
    '<span class="ln-res">⚙ ' + esc(siteResLine(site)) + '</span>' +
    '<button class="mini sb-unpin" data-unpin="' + esc(site.id) + '" title="Убрать участок с общей доски">✕ убрать</button></div>';
}
/* секция в формате доски: компактные колонки дней (без drag&drop — только обзор) */
function boardSectionHTML(site){
  const info = computeDays(site);
  const posMap = buildPositions(site);
  const fond = fondOf(site);
  const t = todayISO();
  const days = info.days.filter(d => d.has);
  let html = '<div class="sub-board">' + sbTitleHTML(site);
  if(!days.length) return html + '<div class="hint">Нет операций в выбранном периоде.</div></div>';
  html += '<div class="bcols">';
  for(const d of days){
    const pct = Math.round(d.total / fond * 100);
    html += '<div class="bcol' + (d.date === t ? ' today' : '') + '">' + (d.date === t ? '<i class="ring"></i>' : '') +
      '<div class="bc-head' + (pct > 100 ? ' over' : '') + '">' + WD[wdOf(d.date)] + ', ' + fmtD(d.date) +
      '<span class="bc-pct">' + pct + '%</span>' +
      '<div class="bc-load">' + uMin(site, d.total) + ' / ' + uMin(site, fond) + '</div></div>' +
      '<div class="bc-body">';
    d.rows.slice().sort((a,b)=>(a.prodOrder||'').localeCompare(b.prodOrder||'')).forEach(r => { html += cardHTML(r, posMap, site, true); });
    html += '</div></div>';
  }
  return html + '</div></div>';
}
/* секция в формате списка дней: таблицы по дням */
function listSectionHTML(site){
  const info = computeDays(site);
  const fond = fondOf(site);
  const t = todayISO();
  const show = info.days.filter(d => d.has);
  let html = '<div class="sub-board">' + sbTitleHTML(site);
  if(!show.length) return html + '<div class="hint">Нет операций в выбранном периоде.</div></div>';
  for(const d of show){
    const pct = Math.round(d.total / fond * 100);
    html += '<div class="l-day">' +
      '<div class="l-head' + (pct>100?' over':'') + '"><span>' + WD[wdOf(d.date)] + ', ' + fmtDFull(d.date) + (d.date === t ? ' <span class="today-chip">сегодня</span>' : '') + '</span>' +
      '<span class="lh-sub">' + d.rows.length + ' операций</span>' +
      '<span class="lh-sub">' + uMin(site, d.total) + ' / ' + uMin(site, fond) + ' (' + pct + '%)' + (pct>100?' — ПЕРЕГРУЗ':'') + '</span></div>' +
      '<table class="l-table"><thead><tr><th>Заказ на пр-во</th><th>Заказ покупателя</th><th>Контрагент</th><th>Продукция</th><th>Кол-во</th><th>Норма</th><th>ТДГ</th><th>Комментарий</th><th>Исполнитель</th><th>Статус</th></tr></thead><tbody>';
    for(const r of d.rows.slice().sort((a,b)=>(a.prodOrder||'').localeCompare(b.prodOrder||''))){
      const st = rowStatus(r);
      html += '<tr class="' + (st==='red'?'l-row-red':st==='amber'?'l-row-amber':'') + '">' +
        '<td class="num">' + esc(r.prodOrder) + '</td><td class="num">' + esc(r.custOrder) + '</td>' +
        '<td>' + esc(displayCustomer(r, site)) + '</td><td class="l-prod" title="' + esc(r.prod) + '">' + esc(displayProduct(r, site)) + '</td>' +
        '<td class="num">' + (r.qty||'') + '</td><td class="num">' + fmtMin(r.minutes) + '</td>' +
        '<td class="num">' + fmtD(r.tdg) + '</td>' + rowContextCells(r) + '' +
        '<td>' + (st==='red'?'⚠ после ТДГ':st==='amber'?'⚠ после договора':'ок') + '</td></tr>';
    }
    html += '</tbody></table></div>';
  }
  return html + '</div>';
}

/* ---------- drag & drop, touch and assembly date editor ---------- */
let dragId = null;
let dragoverRaf = 0;
let dateEditRowId = null;
function canEditPlan(){
  return !window.PLANNER_WORKSPACE || window.PLANNER_WORKSPACE.wCan('edit');
}
function clearDropHighlight(){
  document.querySelectorAll('.day-col.dragover, .l-day.dragover').forEach(el => el.classList.remove('dragover'));
}
function dropZoneOf(e){
  const zone = e.target.closest && e.target.closest('.day-col[data-date], .l-day[data-date]');
  return zone && zone.dataset.date ? zone : null;
}
document.addEventListener('dragstart', e => {
  const src = e.target.closest && e.target.closest('[draggable="true"][data-id]');
  if(!src || !canEditPlan()){ e.preventDefault(); return; }
  dragId = src.dataset.id;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', dragId);
});
document.addEventListener('dragover', e => {
  const zone = dropZoneOf(e);
  if(!zone || !dragId || !canEditPlan()) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  if(dragoverRaf) return;
  dragoverRaf = raf(() => {
    dragoverRaf = 0;
    clearDropHighlight();
    zone.classList.add('dragover');
  });
});
document.addEventListener('dragleave', e => {
  const zone = dropZoneOf(e);
  if(zone && !zone.contains(e.relatedTarget)) zone.classList.remove('dragover');
});
document.addEventListener('dragend', () => { dragId = null; clearDropHighlight(); });
document.addEventListener('drop', e => {
  const zone = dropZoneOf(e);
  clearDropHighlight();
  if(!zone || !dragId || !canEditPlan()) return;
  e.preventDefault();
  moveRowToDate(dragId, zone.dataset.date);
  dragId = null;
});

/* На телефоне движение остаётся прокруткой; перенос начинается после удержания. */
let touchDrag = null;
let touchFrame = 0;
function touchZone(x, y){
  const target = document.elementFromPoint && document.elementFromPoint(x, y);
  return target ? dropZoneOf({ target }) : null;
}
function stopTouchDrag(){
  if(!touchDrag) return;
  clearTimeout(touchDrag.timer);
  if(touchFrame){ cancelAnimationFrame(touchFrame); touchFrame = 0; }
  if(touchDrag.ghost) touchDrag.ghost.remove();
  touchDrag.source.classList.remove('touch-drag-source');
  document.body.classList.remove('touch-dragging');
  if(touchDrag.source.hasPointerCapture && touchDrag.source.hasPointerCapture(touchDrag.pointerId)){
    touchDrag.source.releasePointerCapture(touchDrag.pointerId);
  }
  clearDropHighlight();
  touchDrag = null;
}
function updateTouchDrag(){
  touchFrame = 0;
  if(!touchDrag || !touchDrag.active) return;
  const { x, y } = touchDrag;
  touchDrag.ghost.style.transform = 'translate3d(' + (x + 14) + 'px,' + (y + 14) + 'px,0)';
  clearDropHighlight();
  const zone = touchZone(x, y);
  if(zone) zone.classList.add('dragover');
  if(y < 100) window.scrollBy(0, -12);
  else if(y > window.innerHeight - 90) window.scrollBy(0, 12);
  if(state.view === 'board' && (x < 65 || x > window.innerWidth - 65)){
    $('daysRow').scrollLeft += x < 65 ? -14 : 14;
  }
}
function startPointerDrag(){
  if(!touchDrag || touchDrag.active) return;
  touchDrag.active = true;
  const ghost = document.createElement('div');
  ghost.className = 'touch-drag-ghost';
  const product = touchDrag.source.querySelector('.card-prod, .l-prod');
  ghost.textContent = product ? product.textContent : 'Операция';
  document.body.appendChild(ghost);
  touchDrag.ghost = ghost;
  touchDrag.source.classList.add('touch-drag-source');
  document.body.classList.add('touch-dragging');
  updateTouchDrag();
}
document.addEventListener('pointerdown', e => {
  const mouse = e.pointerType === 'mouse';
  if(!e.isPrimary || !['mouse', 'touch', 'pen'].includes(e.pointerType) || (mouse && e.button !== 0) || !canEditPlan()) return;
  const source = e.target.closest && e.target.closest('#listView > .l-day tr[data-id], .day-col .card[data-id]');
  if(!source || e.target.closest('button, input, select, a')) return;
  stopTouchDrag();
  touchDrag = { source, id:source.dataset.id, pointerId:e.pointerId, pointerType:e.pointerType,
    x:e.clientX, y:e.clientY, startX:e.clientX, startY:e.clientY, active:false, ghost:null, timer:0 };
  try{ source.setPointerCapture(e.pointerId); }catch(error){}
  if(mouse) e.preventDefault(); // не даём нативному drag таблицы оборвать pointer-перенос
  else touchDrag.timer = setTimeout(startPointerDrag, 320);
}, { passive:false });
document.addEventListener('pointermove', e => {
  if(!touchDrag || e.pointerId !== touchDrag.pointerId) return;
  touchDrag.x = e.clientX; touchDrag.y = e.clientY;
  if(!touchDrag.active){
    const distance = Math.hypot(e.clientX - touchDrag.startX, e.clientY - touchDrag.startY);
    if(touchDrag.pointerType === 'mouse' && (e.buttons & 1) && distance > 5) startPointerDrag();
    else if(touchDrag.pointerType !== 'mouse' && distance > 9) stopTouchDrag();
    if(!touchDrag || !touchDrag.active) return;
  }
  e.preventDefault();
  if(!touchFrame) touchFrame = raf(updateTouchDrag);
}, { passive:false });
document.addEventListener('pointerup', e => {
  if(!touchDrag || e.pointerId !== touchDrag.pointerId) return;
  const active = touchDrag.active, id = touchDrag.id;
  const zone = active ? touchZone(e.clientX, e.clientY) : null;
  if(active) e.preventDefault();
  stopTouchDrag();
  if(zone && canEditPlan()) moveRowToDate(id, zone.dataset.date);
}, { passive:false });
document.addEventListener('pointercancel', stopTouchDrag, { passive:true });
// touch-action нельзя надёжно менять посреди жеста: запрещаем scroll только
// после активации переноса, а до удержания оставляем обычную прокрутку страницы.
document.addEventListener('touchmove', e => {
  if(touchDrag && touchDrag.active) e.preventDefault();
}, { passive:false });

function openOperationDateEditor(rowId){
  if(!canEditPlan()) return;
  const site = activeSite();
  const row = effRows(site).find(item => item.id === rowId);
  if(!row || site.group !== 'Сборка') return;
  dateEditRowId = rowId;
  $('operationDateTitle').textContent = 'Дата сборки';
  $('operationDateProduct').textContent = (row.prodOrder || upkpOrderLabel(row.custOrder)) + ' · ' + displayProduct(row, site);
  $('operationDate').value = row.dateEff || todayISO();
  $('operationDateModal').classList.add('open');
}
function saveOperationDate(){
  const value = $('operationDate').value;
  if(!dateEditRowId || !/^\d{4}-\d{2}-\d{2}$/.test(value)){ toast('Выберите дату сборки'); return; }
  const id = dateEditRowId;
  dateEditRowId = null;
  $('operationDateModal').classList.remove('open');
  moveRowToDate(id, value);
}
document.addEventListener('click', e => {
  const button = e.target.closest && e.target.closest('[data-date-edit]');
  if(button) openOperationDateEditor(button.dataset.dateEdit);
});
$('operationDateCancel').addEventListener('click', () => {
  dateEditRowId = null;
  $('operationDateModal').classList.remove('open');
});
$('operationDateApply').addEventListener('click', saveOperationDate);
$('operationDate').addEventListener('change', () => {
  if($('operationDateModal').classList.contains('open')) saveOperationDate();
});
function moveRowToDate(rowId, dateISO){
  if(!canEditPlan()) return false;
  const site = activeSite();
  // часть разбивки?
  const fp = findPart(site, rowId);
  if(fp){
    const oldPartD = fp.part.date || 'nodate';
    fp.part.date = dateISO === 'nodate' ? null : dateISO;
    if(site.overrides) delete site.overrides[rowId];
    saveSites(); scheduleSharedPush(site); render();
    logChange(site, 'перенос', (site.splits[fp.key].prodOrder || '') + ' · ' + fp.part.label + ': ' +
      (oldPartD==='nodate'?'без даты':fmtDFull(oldPartD)) + ' → ' + (dateISO==='nodate'?'без даты':fmtDFull(dateISO)));
    toast('Перенесено: ' + fp.part.label + ' — ' + (dateISO==='nodate'?'без даты':fmtD(dateISO)));
    return true;
  }
  const r = site.rows.find(x => x.id === rowId);
  if(!r) return false;
  const oldD = ((site.overrides||{})[r.id]) || r.date || 'nodate';
  if(dateISO === 'nodate'){
    if(!r.date && !(site.overrides||{})[r.id]) return false;
    delete site.overrides[r.id];
    r._cleared = true;
    // чтобы сбросить дату окончательно — сохраняем override null как пустую строку
    site.overrides[r.id] = '';
    delete r._cleared;
  } else {
    site.overrides[r.id] = dateISO;
  }
  saveSites(); scheduleSharedPush(site);
  logChange(site, 'перенос', (r.prodOrder || r.custOrder) + ' · ' + r.prod + ': ' +
    (oldD==='nodate'?'без даты':fmtDFull(oldD)) + ' → ' + (dateISO==='nodate'?'без даты':fmtDFull(dateISO)));
  render();
  const label = r.prodOrder || r.custOrder;
  toast('Перенесено: ' + label + ' — ' + (oldD==='nodate'?'без даты':fmtD(oldD)) + ' → ' + (dateISO==='nodate'?'без даты':fmtD(dateISO)));
  return true;
}

/* ---------- split modal ---------- */
let splitCtx = null;
function openSplitModal(rowId){
  const site = activeSite();
  const r = site.rows.find(x => x.id === rowId);
  if(!r) return;
  const key = positionKey(r, site);
  const posRows = effRows(site).filter(x => positionKey(x, site) === key);
  splitCtx = { rowId, posRows };
  const totalQty = posRows.reduce((a,x) => a + (x.qty||0), 0);
  const totalMin = posRows.reduce((a,x) => a + (x.minutes||0), 0);
  const orderLabel = site.format === 'upkp' ? upkpOrderLabel(r.custOrder) : (r.prodOrder || r.custOrder);
  let html = '<div class="note">Позиция <b>' + esc(orderLabel) + '</b> — ' + esc(displayProduct(r, site)) +
    '<br>Всего: <b>' + totalQty + ' шт</b> · ' + uMin(site, totalMin) +
    (posRows.length > 1 ? '<br>Строк в позиции: ' + posRows.length + ' (уже разнесены по дням)' : '') + '</div>';
  const perUnit = totalQty > 0 ? totalMin / totalQty : 0;
  const capDay = perUnit > 0 ? Math.max(1, Math.floor(fondOf(site) / perUnit)) : totalQty;
  html += '<div class="fld"><label>Как дробить</label><select id="spMode">' +
    '<option value="smart">разумно — по фонду участка (день ≈ ' + capDay + ' шт, пачки могут быть неравными)</option>' +
    '<option value="equal">равными частями (указать число частей)</option></select></div>';
  html += '<div class="fld" id="spNWrap" style="display:none"><label>На сколько равных частей</label><input type="number" id="spN" value="3" min="2" max="60" step="1"></div>';
  html += '<div class="fld"><label>Даты частей</label><select id="spDates">' +
    '<option value="auto">подряд по дням с даты позиции (минимум дней, без превышения фонда)</option>' +
    '<option value="keep">оставить без дат — распределит автопланировщик</option></select></div>';
  html += '<div id="spPreview"></div>';
  // существующие разбивки
  const sk = Object.keys(site.splits || {});
  if(sk.length){
    html += '<div class="fld"><label>Созданные разбивки на участке</label></div>';
    for(const k of sk){
      const s = site.splits[k];
      html += '<div class="note">✂ ' + esc(s.prodOrder || s.custOrder) + ' — ' + (s.parts||[]).map(p => p.qty + ' шт').join(' + ') +
        ' <button class="mini" data-unsplit="' + esc(k) + '">вернуть как было</button></div>';
    }
  }
  $('splitBody').innerHTML = html;
  const upd = () => {
    const smart = $('spMode').value === 'smart';
    $('spNWrap').style.display = smart ? 'none' : '';
    const n = Math.max(2, Math.min(60, num($('spN').value) || 3));
    const qtys = smart ? splitQtysByFond(totalQty, perUnit, fondOf(site)) : distributeQty(totalQty, n);
    if(qtys.length < 2){ $('spPreview').innerHTML = '<div class="hint">Вся позиция влезает в один день — дробить не нужно.</div>'; return; }
    const mins = distributeMinutes(totalMin, qtys, totalQty);
    let t = '<table class="chg"><thead><tr><th>Часть</th><th>Кол-во</th><th>Норма</th></tr></thead><tbody>';
    qtys.forEach((q, i) => {
      t += '<tr><td>' + (i+1) + '</td><td><input type="number" data-spq="' + i + '" value="' + q + '" min="1" style="width:70px"> шт</td><td>' + uMin(site, mins[i]) + '</td></tr>';
    });
    t += '</tbody></table><div class="hint" style="margin-top:6px">Количества можно поправить вручную — сумма должна остаться ' + totalQty + ' шт.</div>';
    $('spPreview').innerHTML = t;
  };
  $('spMode').addEventListener('change', upd);
  $('spN').addEventListener('input', upd);
  upd();
  $('splitModal').classList.add('open');
}
function createSplitFromModal(){
  const site = activeSite();
  if(!splitCtx) return;
  const { rowId, posRows } = splitCtx;
  const row = site.rows.find(x => x.id === rowId);
  const totalQty = posRows.reduce((a,x) => a + (x.qty||0), 0);
  const totalMin = posRows.reduce((a,x) => a + (x.minutes||0), 0);
  const perUnit = totalQty > 0 ? totalMin / totalQty : 0;
  const smart = $('spMode').value === 'smart';
  const n = Math.max(2, Math.min(60, num($('spN').value) || 3));
  let qtys = smart ? splitQtysByFond(totalQty, perUnit, fondOf(site)) : distributeQty(totalQty, n);
  // ручная корректировка количеств (инпуты отражают текущий расчёт)
  const qInputs = document.querySelectorAll('[data-spq]');
  if(qInputs.length){
    let sum = 0;
    qtys = [...qInputs].map(inp => { const q = Math.max(1, num(inp.value) || 1); sum += q; return q; });
    if(sum !== totalQty){ toast('Сумма частей (' + sum + ') не равна количеству позиции (' + totalQty + ')'); return; }
  }
  if(qtys.length < 2){ toast('Вся позиция влезает в один день — дробить не нужно'); return; }
  const rec = buildSplitRecordQtys(row, posRows, qtys, undefined, site);
  const mode = $('spDates').value;
  const anchor = posRows.map(x => x.dateEff).filter(Boolean).sort()[0];
  if(mode === 'auto') assignSplitDates(site, rec, anchor || todayISO());
  site.splits = site.splits || {};
  site.splits[rec.key] = rec;
  saveSites(); scheduleSharedPush(site);
  logChange(site, 'разбивка', (rec.prodOrder || rec.custOrder) + ' · ' + rec.prod + ': ' + totalQty + ' шт → ' + rec.parts.map(p => p.qty).join(' + '));
  render();
  $('splitModal').classList.remove('open');
  splitCtx = null;
  toast('Разбито на ' + rec.parts.length + ' частей: ' + rec.parts.map(p => p.qty + ' шт').join(' + '));
}
function restoreSplit(key){
  const site = activeSite();
  if(!site.splits || !site.splits[key]) return;
  const s = site.splits[key];
  delete site.splits[key];
  saveSites(); scheduleSharedPush(site);
  logChange(site, 'отмена разбивки', (s.prodOrder || s.custOrder) + ' · ' + s.prod + ' — исходные строки возвращены');
  render();
  toast('Разбивка отменена, исходные строки возвращены');
}

/* ---------- navigation ---------- */
function jumpToDate(iso){
  state.selectedDate = iso === 'nodate' ? null : iso;
  if(state.view !== 'board') switchView('board');
  const t = todayISO();
  if(state.selectedDate){
    const site = activeSite();
    const { days } = computeDays(site);
    if(!days.some(d => d.date === state.selectedDate) && iso !== 'nodate'){
      // даты нет в диапазоне — просто подсветим ближайший день прокруткой в конец/начало
    }
  }
  render();
  setTimeout(() => {
    const cols = document.querySelectorAll('.day-col');
    let target = null;
    cols.forEach(c => { if(c.dataset.date === iso) target = c; });
    if(!target && cols.length) target = cols[cols.length-1];
    if(target){
      target.scrollIntoView({behavior:'smooth', inline:'center', block:'nearest'});
      target.classList.add('flash');
      setTimeout(() => target.classList.remove('flash'), 1300);
    }
  }, 60);
}
function switchView(v){
  state.view = v;
  $('viewBoard').classList.toggle('on', v === 'board');
  $('viewList').classList.toggle('on', v === 'list');
  $('viewShifts').classList.toggle('on', v === 'shifts');
  $('viewMulti').classList.toggle('on', v === 'multi');
  state._enterView = true; // стаггер-появление колонок нового вида
  render();
  raf(segSync);
}

/* ---------- google sheets ---------- */
function parseGUrl(u){
  const m = String(u).match(/\/d\/([a-zA-Z0-9-_]{20,})/);
  if(!m) return null;
  const gid = (String(u).match(/[#?&]gid=(\d+)/) || [])[1];
  return { id:m[1], gid };
}
function gvizUrl(id, gid, sheetName){
  let u = 'https://docs.google.com/spreadsheets/d/' + id + '/gviz/tq?tqx=out:csv';
  if(gid) u += '&gid=' + gid;
  else if(sheetName) u += '&sheet=' + encodeURIComponent(sheetName);
  return u;
}
let latestSharedWorkbook = null;
const sharedWorkbookFlights = new Map();
async function fetchSharedWorkbook(url){
  const parsed = parseGUrl(url);
  if(!parsed) throw new Error('Не получилось разобрать ссылку на общую книгу Google Sheets');
  const key = parsed.id;
  const previous = sharedWorkbookFlights.get(key);
  if(previous && Date.now() - previous.started < 15000) return previous.promise;
  const promise = (async () => {
    const source = 'https://docs.google.com/spreadsheets/d/' + key + '/export?format=xlsx&_plannerRefresh=' + Date.now();
    const response = await fetch(source, {mode:'cors', cache:'no-store'});
    if(!response.ok) throw new Error('HTTP ' + response.status);
    const buffer = await response.arrayBuffer();
    if(!buffer.byteLength) throw new Error('Google Sheets вернула пустую XLSX-книгу');
    const result = await SHARED.parseOpsWorkbook(buffer);
    result.fetchedAt = new Date().toISOString();
    latestSharedWorkbook = result;
    return result;
  })();
  sharedWorkbookFlights.set(key, {started:Date.now(), promise});
  try{ return await promise; }
  catch(error){ sharedWorkbookFlights.delete(key); throw error; }
}
function sharedRowsFor(result, operation){
  const normalized = SHARED.normalizeOperation(operation);
  return result.rows.filter(row => SHARED.normalizeOperation(row.operation) === normalized);
}
function discoverSharedSites(result){
  const added = [];
  for(const operation of result.operations || []){
    const spec = SHARED.specForOperation(operation);
    if(!spec || siteConfig.deleted.includes(spec.sourceSiteKey)) continue;
    if(state.sites.some(site => siteKeyOf(site) === spec.sourceSiteKey)) continue;
    const site = createSharedSite(spec);
    site.rows = makeRows(sharedRowsFor(result, operation));
    site.syncOk = true; site.syncAt = result.fetchedAt || new Date().toISOString();
    site.syncSheetNames = result.sheetNames || []; site.syncSheetCount = result.sheetCount || 1;
    state.sites.push(site); added.push(site);
    logChange(site, 'участок добавлен из общей загрузки', operation + ': ' + site.rows.length + ' операций');
  }
  if(added.length){ saveSites(); persistSiteConfiguration(); }
  return added;
}
async function fetchSiteRows(url, site){
  let u = String(url || '').trim();
  if(site && site.format === 'ops-shared'){
    const result = await fetchSharedWorkbook(u);
    return {...result, sharedRows:result.rows, rows:sharedRowsFor(result, site.sourceOperation)};
  }
  // autoMonth — сохранённый флаг формата УПКП. Теперь он означает всю книгу:
  // XLSX содержит все текущие и автоматически добавленные позже листы.
  if(site && site.autoMonth){
    const g = parseGUrl(u);
    if(!g) throw new Error('Не получилось разобрать ссылку на Google Sheets');
    u = 'https://docs.google.com/spreadsheets/d/' + g.id + '/export?format=xlsx';
    u += '&_plannerRefresh=' + Date.now();
    const res = await fetch(u, {mode:'cors', cache:'no-store'});
    if(!res.ok) throw new Error('HTTP ' + res.status);
    const buffer = await res.arrayBuffer();
    if(!buffer.byteLength) throw new Error('Google Sheets вернула пустую XLSX-книгу');
    return parseUPKPWorkbook(buffer);
  }
  if(!u.includes('/gviz/')){
    const g = parseGUrl(u);
    if(!g) throw new Error('Не получилось разобрать ссылку на Google Sheets');
    u = gvizUrl(g.id, g.gid, '');
  }
  // Google и браузер могут держать CSV в кэше: для ручного обновления нужен живой ответ.
  u += (u.includes('?') ? '&' : '?') + '_plannerRefresh=' + Date.now();
  const res = await fetch(u, {mode:'cors', cache:'no-store'});
  if(!res.ok) throw new Error('HTTP ' + res.status);
  const text = await res.text();
  if(/^\s*<html/i.test(text) || /войдите|sign in/i.test(text.slice(0, 500))) throw new Error('Таблица закрыта — откройте доступ «по ссылке» (просмотр)');
  const fmt = site && site.format ? site.format : detectFormat(text);
  return { rows: parseAny(text, fmt), format: fmt };
}

/* ---------- export ---------- */
function exportCSV(){
  const site = activeSite();
  const rows = effRows(site);
  const q = v => '"' + String(v ?? '').replace(/"/g,'""') + '"';
  const head = ['Заказ покупателя','Продукция','Количество','Дата операции','всего','ТДГ','Комментарий','Заказ на производство'];
  const lines = [head.join(',')];
  for(const r of rows){
    lines.push([r.custOrder, r.prod, r.qty, r.dateEff ? fmtDFull(r.dateEff) : '', r.minutes, r.tdg || '', r.comment, r.prodOrder].map(q).join(','));
  }
  const blob = new Blob(['﻿' + lines.join('\r\n')], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'plan_' + site.name.replace(/[^\wа-яА-ЯёЁ]+/gi,'_') + '_' + todayISO() + '.csv';
  a.click();
  URL.revokeObjectURL(a.href);
  toast('CSV сохранён');
}

/* ---------- autoplan v2 ----------
   Правила:
   1. Заказ не размазывается по дням — все его операции упакованы в минимум
      подряд идущих рабочих дней (день заполняется до фонда, остаток — на следующий).
   2. Договоры выполняются последовательно в порядке срочности (ТДГ/дата договора):
      при включённом «строго последовательно» следующий заказ начинается со следующего
      рабочего дня после полного закрытия предыдущего.
   3. Текстовые пожелания (site.notes) — директивы:
      «2615 -> 25.10» — заказ точно на дату; «после 22.10» — не раньше;
      «не позже 30.10» — дедлайн; «сначала: …» — приоритет вне очереди. */
function parseDateRu(s){
  const m = String(s || '').match(/(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?/);
  if(!m) return null;
  let y = m[3] ? (m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3])) : new Date().getFullYear();
  const d = new Date(y, Number(m[2]) - 1, Number(m[1]));
  return isNaN(d) ? null : d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate());
}
function parseNotes(site){
  const res = { pins: [], earliest: [], latest: [], first: [], after: [], unparsed: [] };
  const lines = String(site.notes || '').split('\n');
  const D = '(\\d{1,2}[.\\-/]\\d{1,2}(?:[.\\-/]\\d{2,4})?)';
  const reFirst = /^(?:сначала|приоритет|в первую очередь|первым делом)\s*:?\s*(.+)$/i;
  const rePinArrow = new RegExp('(.+?)\\s*(?:->|=)\\s*' + D + '\\s*$');
  const rePinWord = new RegExp('(.+?)\\s+(?:строго|точно|поставить|сделать|на)\\s+(?:на\\s+)?' + D + '\\s*$', 'i');
  const reEarly = new RegExp('(.+?)\\s+(?:после|не раньше|не ранее)\\s+' + D + '\\s*$', 'i');
  const reLatest = new RegExp('(?:не позже|не позднее|дедлайн|до|к)\\s+' + D + '\\s*:?\\s*(.+)$', 'i');
  const reAfter = /^(.+?)\s+(?:делать\s+)?после\s+(.+?)\s*$/i;
  for(let raw of lines){
    const line = raw.trim();
    if(!line) continue;
    let m;
    if((m = line.match(reFirst))){
      res.first.push(...m[1].split(/[,;]/).map(s => s.trim()).filter(Boolean));
    } else if((m = line.match(rePinArrow))){
      const d = parseDateRu(m[2]);
      if(d){ res.pins.push({ token: m[1].replace(/^!/, '').trim(), date: d }); continue; }
      res.unparsed.push(line);
    } else if((m = line.match(reEarly))){
      const d = parseDateRu(m[2]);
      if(d){ res.earliest.push({ token: m[1].trim(), date: d }); continue; }
      res.unparsed.push(line);
    } else if((m = line.match(reLatest))){
      const d = parseDateRu(m[1]);
      if(d){ res.latest.push({ token: m[2].trim(), date: d }); continue; }
      res.unparsed.push(line);
    } else if((m = line.match(rePinWord))){
      const d = parseDateRu(m[2]);
      if(d){ res.pins.push({ token: m[1].replace(/^!/, '').trim(), date: d }); continue; }
      res.unparsed.push(line);
    } else if((m = line.match(reAfter)) && !parseDateRu(m[2])){
      res.after.push({ token: m[1].trim(), ref: m[2].trim() });
    } else {
      res.unparsed.push(line);
    }
  }
  return res;
}
function posMatches(p, token){
  const t = String(token || '').toLowerCase();
  if(!t) return false;
  const numT = orderNum(t);
  if(numT !== null && numT !== undefined && !/[а-яa-z]{2,}/i.test(t)){
    return orderNum(p.custOrder) === numT || orderNum(p.prodOrder) === numT;
  }
  return (p.prodOrder || '').toLowerCase().includes(t)
    || (p.custOrder || '').toLowerCase().includes(t)
    || (p.customer || '').toLowerCase().includes(t)
    || p.parts.some(r => (r.prod || '').toLowerCase().includes(t));
}
/* Только точная пара «производственный заказ + нормализованная продукция».
   Для ОП/САУ/БВ учитываем актуальные effRows с ручными переносами. УПКП из
   общего XLSX даёт дату сборки, но не связывается с отдельной секционной книгой. */
function assemblyIndex(){
  const index = new Map();
  const add = (row, source) => {
    if(!row || !SHARED.isAssembly(row.operation)) return;
    const key = SHARED.relationKey(row), date = row.dateEff || row.date;
    if(!key || !date) return;
    const previous = index.get(key);
    if(previous && previous.date <= date) return;
    index.set(key, { key, date, operation:row.operation, sourceRow:row.sourceRow || null,
      siteId:source.id || '', siteName:source.name || '', source:source.source || 'Актуальный план сборки' });
  };
  for(const site of state.sites){
    if(site.format !== 'ops-shared' || !SHARED.isAssembly(site.sourceOperation)) continue;
    for(const row of effRows(site)) add(row, site);
  }
  if(latestSharedWorkbook){
    for(const row of latestSharedWorkbook.assemblyRows || []){
      if(SHARED.isSharedUpkp(row.operation)) add(row, { id:'s_upkp_shared', name:'УПКП · общая загрузка', source:'Общая загрузка → ' + row.operation });
    }
  }
  return index;
}
function assemblyConstraintsFor(site){
  const empty = { assemblyLinks:[], assemblyConstraints:[], unlinked:[] };
  if(!site || site.group !== 'Заготовка') return empty;
  const index = assemblyIndex();
  const result = { assemblyLinks:[], assemblyConstraints:[], unlinked:[] };
  for(const position of buildPositions(site).values()){
    const relationKey = SHARED.relationKey(position), link = index.get(relationKey);
    if(!link){
      result.unlinked.push({ key:position.key, prodOrder:position.prodOrder, prod:position.prod,
        reason:'Нет сборки с точным совпадением производственного заказа и продукции' });
      continue;
    }
    const constraint = { key:position.key, relationKey, prodOrder:position.prodOrder, prod:position.prod,
      assemblyDate:link.date, latest:addDays(link.date, -1), assemblyOperation:link.operation,
      assemblySiteId:link.siteId, assemblySiteName:link.siteName, source:link.source, sourceRow:link.sourceRow };
    result.assemblyLinks.push({...constraint});
    result.assemblyConstraints.push(constraint);
  }
  return result;
}
function autoplain(site){
  passBegin();
  try{
  const t = todayISO();
  const fond = fondOf(site);
  const strict = site.planStrict !== false;
  const notes = parseNotes(site);
  const rows = effRows(site);
  const posMap = buildPositions(site);
  const positions = [...posMap.values()];
  const assembly = assemblyConstraintsFor(site);
  const assemblyByKey = new Map(assembly.assemblyConstraints.map(item => [item.key, item]));
  for(const p of positions){
    p.totalMin = p.parts.reduce((a,r) => a + (r.minutes||0), 0);
    p.totalQty = p.parts.reduce((a,r) => a + (r.qty||0), 0);
    p.deadline = p.parts.map(r => r.tdg || r.contractDate).filter(Boolean).sort()[0] || null;
    p.assemblyConstraint = assemblyByKey.get(p.key) || null;
    if(p.assemblyConstraint && (!p.deadline || p.assemblyConstraint.latest < p.deadline)) p.deadline = p.assemblyConstraint.latest;
    p.pin = notes.pins.find(d => posMatches(p, d.token)) || null;
    p.early = notes.earliest.filter(d => posMatches(p, d.token)).map(d => d.date).sort().pop() || null;
    p.late = notes.latest.filter(d => posMatches(p, d.token)).map(d => d.date).sort()[0] || null;
    if(p.late && p.deadline) p.deadline = p.late < p.deadline ? p.late : p.deadline;
    else if(p.late) p.deadline = p.late;
    const aft = notes.after.find(a => posMatches(p, a.token));
    p.afterRefTok = aft ? aft.ref : null;
    p.afterRef = null;
  }
  // связи «X после Y»: ищем опорную позицию
  for(const p of positions){
    if(p.afterRefTok){
      const q = positions.find(x => x !== p && posMatches(x, p.afterRefTok));
      if(q) p.afterRef = q;
    }
  }
  const firstIdx = p => {
    const i = notes.first.findIndex(tok => posMatches(p, tok));
    return i === -1 ? notes.first.length : i;
  };
  positions.sort((a,b) =>
    (firstIdx(a) - firstIdx(b)) ||
    ((a.deadline||'9999').localeCompare(b.deadline||'9999')) ||
    ((orderNum(a.custOrder)||99999) - (orderNum(b.custOrder)||99999)));
  const plan = {};
  const load = {};
  const autoSplits = [];
  const plannedEnd = new Map(); // key позиции -> последняя запланированная дата
  const tooBig = [];
  let cursor = t; // строгая последовательность: следующий заказ не раньше этого дня
  const planOne = (p) => {
    const parts = p.parts.slice().sort((a,b) => (a.dateEff||'9999').localeCompare(b.dateEff||'9999'));
    if(p.pin){
      for(const r of parts){
        if((r.minutes||0) > fond) tooBig.push(r);
        plan[r.id] = p.pin.date;
      }
      plannedEnd.set(p.key, p.pin.date);
      return;
    }
    let start = t;
    if(p.early && p.early > start) start = p.early;
    if(p.afterRef && plannedEnd.has(p.afterRef.key)){
      const s2 = nextWorkDay(site, addDays(plannedEnd.get(p.afterRef.key), 1));
      if(s2 > start) start = s2;
    }
    if(strict && cursor > start) start = cursor;
    let d = nextWorkDay(site, start);
    let used = strict ? 0 : (load[d] || 0);
    // авто-разбивка: единственная строка позиции не влезает в дневной фонд и есть
    // количество — дробим разумно по фонду (пачки не обязательно равные), заказ идёт подряд.
    // Позиции из нескольких строк не дробим: их операции и так раскладываются по дням.
    const alreadySplit = p.parts.some(r => r.splitPart);
    if(!alreadySplit && p.parts.length === 1 && p.totalQty > 1 && p.totalMin > fond){
      const perUnit = p.totalMin / p.totalQty;
      if(!strict && used > 0 && used + perUnit > fond){ d = nextWorkDay(site, addDays(d, 1)); used = load[d] || 0; }
      const avail = strict ? fond : Math.max(0, fond - used);
      const qtys = splitQtysByFond(p.totalQty, perUnit, fond, avail);
      if(qtys.length > 1){
        const rec = buildSplitRecordQtys(p.parts[0], p.parts, qtys, 'авто-разбивка по фонду (' + p.totalQty + ' шт)', site);
        for(let i = 0; i < rec.parts.length; i++){
          if(i > 0){ d = nextWorkDay(site, addDays(d, 1)); used = strict ? 0 : (load[d] || 0); }
          rec.parts[i].date = d;
          plan[rec.parts[i].id] = d;
          used += rec.parts[i].minutes;
          if(!strict) load[d] = used;
        }
        autoSplits.push(rec);
        plannedEnd.set(p.key, d);
        cursor = nextWorkDay(site, addDays(d, 1));
        return;
      }
    }
    for(const r of parts){
      if((r.minutes||0) > fond) tooBig.push(r);
      let guard = 0;
      while(used + (r.minutes||0) > fond && guard++ < 800){
        d = nextWorkDay(site, addDays(d, 1));
        used = strict ? 0 : (load[d] || 0);
      }
      plan[r.id] = d;
      used += (r.minutes || 0);
      if(!strict) load[d] = (load[d] || 0) + (r.minutes || 0);
    }
    plannedEnd.set(p.key, d);
    cursor = nextWorkDay(site, addDays(d, 1));
  };
  // планируем последовательно; позиции со связью «после» ждут свою опору
  const pending = positions.slice();
  let guardPass = 0;
  while(pending.length){
    let progressed = false;
    for(let i = 0; i < pending.length; i++){
      const p = pending[i];
      if(p.afterRef && !plannedEnd.has(p.afterRef.key) && pending.includes(p.afterRef)) continue;
      pending.splice(i, 1);
      i--;
      planOne(p);
      progressed = true;
    }
    if(!progressed){ while(pending.length) planOne(pending.shift()); break; } // цикл в связях — планируем как есть
    if(++guardPass > positions.length + 2) break;
  }
  const changes = Object.keys(plan).filter(id => {
    const cur = rows.find(r => r.id === id);
    return cur && (cur.dateEff || null) !== plan[id];
  });
  // контроль сроков: последний день позиции позже дедлайна (ТДГ / «не позже»)
  const late = [];
  for(const p of positions){
    if(!p.deadline) continue;
    const end = plannedEnd.get(p.key);
    if(end && end > p.deadline) late.push({ prodOrder: p.prodOrder, prod: p.prod, end, deadline: p.deadline });
  }
  const dependencyLate = [];
  for(const p of positions){
    const constraint = p.assemblyConstraint, end = plannedEnd.get(p.key);
    if(constraint && end && end >= constraint.assemblyDate){
      dependencyLate.push({...constraint, end, reason:'Заготовка должна завершиться строго раньше сборки'});
    }
  }
  return { plan, changes, tooBig, total: rows.length, notes, autoSplits, late, ...assembly, dependencyLate };
  }finally{ passEnd(); }
}
function applyPlanToSite(site, plan, autoSplits){
  site.overrides = site.overrides || {};
  if(autoSplits && autoSplits.length){
    site.splits = site.splits || {};
    for(const rec of autoSplits) site.splits[rec.key] = rec;
  }
  for(const [id, d] of Object.entries(plan)){
    const fp = findPart(site, id);
    if(fp){ fp.part.date = d; delete site.overrides[id]; }
    else site.overrides[id] = d;
  }
}
function openAutoplan(){
  const site = activeSite();
  site.notes = site.notes || '';
  site.planStrict = site.planStrict !== false;
  $('planNotes').value = site.notes;
  $('planStrict').checked = site.planStrict;
  runAutoplanPreview();
  $('planModal').classList.add('open');
}
function runAutoplanPreview(){
  const site = activeSite();
  const res = autoplain(site);
  pendingPlan = res;
  const rows = effRows(site);
  const byId = id => rows.find(r => r.id === id);
  const n = res.notes;
  const dirCount = n.pins.length + n.earliest.length + n.latest.length + n.first.length + n.after.length;
  let html = '<div class="kpis" style="margin-bottom:10px">' +
    '<div class="kpi"><b>' + res.total + '</b><span>операций в плане</span></div>' +
    '<div class="kpi"><b>' + res.changes.length + '</b><span>дат изменится</span></div>' +
    '<div class="kpi"><b>' + (res.total - res.changes.length) + '</b><span>без изменений</span></div>' +
    (site.group === 'Заготовка' ? '<div class="kpi"><b>' + res.assemblyLinks.length + '</b><span>доказанных связей со сборкой</span></div>' : '') +
    (res.autoSplits.length?'<div class="kpi"><b>' + res.autoSplits.length + '</b><span>авто-разбивок по фонду</span></div>':'') +
    (res.unlinked.length?'<div class="kpi warn"><b>' + res.unlinked.length + '</b><span>без доказанной сборки</span></div>':'') +
    (res.late.length?'<div class="kpi bad"><b>' + res.late.length + '</b><span>не успевают к сроку</span></div>':'') +
    (res.tooBig.length?'<div class="kpi bad"><b>' + res.tooBig.length + '</b><span>норма больше фонда</span></div>':'') +
  '</div>';
  $('planHint').innerHTML = 'Правила: заказ идёт подряд и целиком · договоры по срочности (ТДГ)' +
    (site.planStrict !== false ? ' · строго: один договор за раз' : ' · компактно: договоры делят дни') +
    (site.group === 'Заготовка' ? ' · заготовка строго раньше доказанной сборки' : '') +
    (dirCount ? ' · директив из примечаний: ' + dirCount : '');
  if(n.unparsed && n.unparsed.length){
    html += '<div class="note">Не распознано строк: ' + n.unparsed.length + ' — ' +
      n.unparsed.slice(0, 5).map(esc).join('; ') + (n.unparsed.length > 5 ? '…' : '') +
      '<br>Попробуйте переформулировать (примеры — в подсказке поля) или попросите Kimi перевести текст в директивы.</div>';
  }
  if(site.format === 'upkp'){
    html += '<div class="note"><b>Ограничение связей УПКП.</b> В отдельной секционной книге нет производственного заказа, поэтому надёжно связать её строки с общей загрузкой нельзя. Даты УПКП из общего листа используются только как ограничения для заготовки.</div>';
  }
  if(res.dependencyLate.length){
    html += '<div class="note dependency-error"><b>План нельзя применить: заготовка не успевает к сборке.</b><br>' +
      res.dependencyLate.slice(0, 100).map(item => esc(item.prodOrder || '') + ' · ' + esc(item.prod) + ' — конец ' + fmtDFull(item.end) + ', сборка ' + fmtDFull(item.assemblyDate) + ' (' + esc(item.assemblyOperation) + ')').join('<br>') +
      (res.dependencyLate.length > 100 ? '<br>… и ещё ' + (res.dependencyLate.length - 100) : '') + '</div>';
  }
  if(res.unlinked.length){
    html += '<details class="note dependency-note"><summary>Без доказанной связи со сборкой: ' + res.unlinked.length + '. Гарантировать срок для них нельзя.</summary>' +
      res.unlinked.slice(0, 100).map(item => esc(item.prodOrder || '') + ' · ' + esc(item.prod) + ' — ' + esc(item.reason)).join('<br>') +
      (res.unlinked.length > 100 ? '<br>… и ещё ' + (res.unlinked.length - 100) : '') + '</details>';
  }
  if(res.assemblyLinks.length){
    html += '<details class="note dependency-ok"><summary>Доказанные связи по заказу и продукции: ' + res.assemblyLinks.length + '</summary>' +
      res.assemblyLinks.slice(0, 100).map(item => esc(item.prodOrder || '') + ' · ' + esc(item.prod) + ' — не позже ' + fmtDFull(item.latest) + '; сборка ' + fmtDFull(item.assemblyDate) + ' · ' + esc(item.assemblySiteName || item.assemblyOperation)).join('<br>') +
      (res.assemblyLinks.length > 100 ? '<br>… и ещё ' + (res.assemblyLinks.length - 100) : '') + '</details>';
  }
  if(res.late.length){
    html += '<div class="note" style="border-color:#f8c9d2;background:var(--red-light)"><b>Не укладываются в сроки ТДГ, договора или сборки:</b><br>' +
      res.late.slice(0, 100).map(l => esc(l.prodOrder || '') + ' · ' + esc(l.prod) + ' — конец ' + fmtDFull(l.end) + ' при сроке ' + fmtDFull(l.deadline)).join('<br>') +
      (res.late.length > 100 ? '<br>… и ещё ' + (res.late.length - 100) : '') + '</div>';
  }
  if(res.autoSplits.length){
    html += '<div class="note"><b>Авто-разбивка по фонду (пачки по ёмкости дня, не обязательно равные):</b><br>' +
      res.autoSplits.map(rec => esc(rec.prodOrder || rec.custOrder) + ' · ' + esc(rec.prod) + ' — ' +
        rec.parts.map(p => p.qty + ' шт на ' + fmtD(p.date)).join(' + ')).join('<br>') + '</div>';
  }
  if(res.changes.length){
    html += '<table class="chg"><thead><tr><th>Заказ на пр-во</th><th>Продукция</th><th>Кол-во</th><th>Было</th><th>Стало</th></tr></thead><tbody>';
    for(const id of res.changes.slice(0, 400)){
      const r = byId(id);
      if(!r) continue;
      html += '<tr><td>' + esc(r.prodOrder || r.custOrder) + '</td><td>' + esc(r.prod) + '</td><td>' + (r.qty||'') + '</td>' +
        '<td class="old">' + fmtD(r.dateEff) + '</td><td class="new">' + fmtD(res.plan[id]) + '</td></tr>';
    }
    if(res.changes.length > 400) html += '<tr><td colspan="5">… и ещё ' + (res.changes.length-400) + '</td></tr>';
    html += '</tbody></table>';
  } else html += '<div class="hint">Все даты уже соответствуют правилам.</div>';
  $('planBody').innerHTML = html;
  $('planApply').disabled = res.dependencyLate.length > 0;
  $('planApply').title = res.dependencyLate.length ? 'Сначала измените ресурс, дату сборки или ограничения: заготовка должна завершиться раньше сборки' : 'Применить рассчитанный план';
}

/* ---------- журнал изменений: UI и отчёт ---------- */
let logFilter = { from: '', to: '', site: '' };
function openLogModal(){
  const sel = $('logSite');
  sel.innerHTML = '<option value="">все участки</option>' +
    state.sites.map(s => '<option value="' + esc(s.id) + '"' + (logFilter.site === s.id ? ' selected' : '') + '>' + esc(s.name) + '</option>').join('');
  $('logFrom').value = logFilter.from || addDays(todayISO(), -7);
  $('logTo').value = logFilter.to || todayISO();
  renderLogTable();
  $('logModal').classList.add('open');
}
function logEntriesFiltered(){
  const from = $('logFrom').value || '';
  const to = $('logTo').value || '';
  const siteId = $('logSite').value || '';
  logFilter = { from, to, site: siteId };
  return readLog().filter(e => {
    const d = new Date(e.ts);
    const iso = d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate());
    if(from && iso < from) return false;
    if(to && iso > to) return false;
    if(siteId && e.site !== siteId) return false;
    return true;
  }).sort((a,b) => b.ts - a.ts);
}
function renderLogTable(){
  const arr = logEntriesFiltered();
  if(!arr.length){ $('logBody').innerHTML = '<div class="hint">За выбранный период изменений нет.</div>'; return; }
  let html = '<table class="log-table"><thead><tr><th>Время</th><th>Участок</th><th>Действие</th><th>Детали</th></tr></thead><tbody>';
  for(const e of arr.slice(0, 500)){
    html += '<tr><td class="lt-ts">' + fmtTs(e.ts) + '</td><td>' + esc(e.siteName) + '</td>' +
      '<td class="lt-act">' + esc(e.action) + '</td><td>' + esc(e.details) + '</td></tr>';
  }
  if(arr.length > 500) html += '<tr><td colspan="4">… и ещё ' + (arr.length - 500) + ' записей (сузьте период или выгрузите CSV)</td></tr>';
  html += '</tbody></table>';
  $('logBody').innerHTML = html;
}
function exportLogCsv(){
  const arr = logEntriesFiltered();
  const q = v => '"' + String(v ?? '').replace(/"/g,'""') + '"';
  const lines = [['Дата и время','Участок','Действие','Детали'].map(q).join(',')];
  for(const e of arr) lines.push([fmtTs(e.ts), e.siteName, e.action, e.details].map(q).join(','));
  const blob = new Blob(['﻿' + lines.join('\r\n')], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'planner_log_' + (logFilter.from || '') + '_' + (logFilter.to || todayISO()) + '.csv';
  a.click();
  URL.revokeObjectURL(a.href);
  toast('Отчёт по журналу сохранён: ' + arr.length + ' записей');
}

/* ---------- site settings: rename, grouping, categories and resources ---------- */
function fillSiteCategories(selected){
  const values = [...new Set([...DEFAULT_CATEGORIES, ...siteConfig.categories, selected].filter(Boolean))];
  $('ssCategory').innerHTML = values.map(value => '<option value="' + esc(value) + '">' + esc(value) + '</option>').join('');
  $('ssCategory').value = selected || 'Прочее';
}
function updateSiteSettingsResource(){
  const team = $('ssTeam').checked;
  $('ssPerWorkerWrap').style.display = team ? 'none' : '';
  $('ssFond').readOnly = !team;
  if(!team){
    const workers = Math.max(1, num($('ssWorkers').value) || 1);
    const perWorker = Math.max(0, num($('ssPerWorker').value));
    $('ssFond').value = workers * perWorker;
  }
}
function openSiteSettings(){
  const site = activeSite();
  if(!site) return;
  ensureSiteIdentity(site);
  $('ssName').value = site.name || '';
  $('ssGroup').value = site.group || 'Прочее';
  fillSiteCategories(site.category || site.group || 'Прочее');
  $('ssWorkers').value = site.workers || 1;
  $('ssPerWorker').value = site.perWorker || 0;
  $('ssFond').value = fondOf(site);
  $('ssTeam').checked = !!site.teamTotal;
  $('ssDays').innerHTML = [1,2,3,4,5,6,0].map(day => '<label><input type="checkbox" value="' + day + '"' + ((site.workdays || []).includes(day) ? ' checked' : '') + '> ' + WD[day] + '</label>').join('');
  $('ssSourceNote').innerHTML = '<b>Контекст источника.</b> ' + esc(site.mappingNote || 'Пользовательский участок.') +
    '<br><span>Стабильный ключ: ' + esc(site.sourceSiteKey) + '</span>';
  $('ssResourceHint').textContent = site.resourceConfirmed === false
    ? 'В общей таблице ресурса нет: сейчас используется предварительные 1 человек и 450 мин/день. После сохранения здесь ваши значения станут постоянными.'
    : 'Ресурс подтверждён вручную и имеет приоритет над системными значениями при обновлениях кода.';
  $('ssRepoStatus').textContent = siteConfigRepoStatus.message;
  $('ssNewCategory').value = '';
  updateSiteSettingsResource();
  $('siteSettingsModal').classList.add('open');
}
function addSiteCategory(){
  const input = $('ssNewCategory');
  const value = input.value.trim().replace(/\s+/g, ' ');
  if(!value){ toast('Введите название категории'); return; }
  if(siteConfig.categories.some(item => item.toLocaleLowerCase('ru-RU') === value.toLocaleLowerCase('ru-RU'))){
    fillSiteCategories(value); input.value = ''; return;
  }
  siteConfig.categories.push(value);
  fillSiteCategories(value);
  input.value = '';
  persistSiteConfiguration();
}
function saveSiteSettings(){
  const site = activeSite();
  if(!site) return;
  const before = site.name + ' · ' + (site.group || 'Прочее') + ' · ' + siteResLine(site);
  site.name = $('ssName').value.trim() || site.name;
  site.group = $('ssGroup').value || 'Прочее';
  site.category = $('ssCategory').value || site.group;
  site.workers = Math.max(1, num($('ssWorkers').value) || 1);
  site.teamTotal = $('ssTeam').checked;
  site.perWorker = site.teamTotal ? 0 : Math.max(1, num($('ssPerWorker').value) || 1);
  site.fondMin = site.teamTotal ? Math.max(1, num($('ssFond').value) || 1) : site.workers * site.perWorker;
  const workdays = [...document.querySelectorAll('#ssDays input:checked')].map(input => Number(input.value));
  site.workdays = workdays.length ? workdays : [1,2,3,4,5];
  site.resourceConfirmed = true;
  ensureSiteIdentity(site);
  saveSites(); persistSiteConfiguration(); scheduleSharedPush(site);
  logChange(site, 'настройки участка', before + ' → ' + site.name + ' · ' + site.group + ' · ' + siteResLine(site));
  $('siteSettingsModal').classList.remove('open');
  render();
  toast('Настройки «' + site.name + '» сохранены без изменения ключа источника');
}

/* ---------- ui wiring ---------- */
function toast(msg){
  const t = $('toast');
  t.textContent = msg;
  t.style.display = 'block';
  clearTimeout(t._h);
  t._h = setTimeout(() => t.style.display = 'none', 2600);
}
function wire(){
  $('siteSelect').addEventListener('change', e => { state.activeId = e.target.value; state.selectedDate = null; render(); });
  $('btnDelSite').addEventListener('click', () => {
    const site = activeSite();
    if(state.sites.length <= 1){ toast('Нельзя удалить последний участок'); return; }
    if(!confirm('Удалить участок «' + site.name + '»?')) return;
    const key = siteKeyOf(site);
    if(!siteConfig.deleted.includes(key)) siteConfig.deleted.push(key);
    state.sites = state.sites.filter(s => s.id !== site.id);
    state.activeId = state.sites[0].id;
    logChange(null, 'участок удалён без восстановления при синхронизации', site.name + ' · ' + key);
    saveSites(); persistSiteConfiguration(); render();
  });
  $('btnSiteSettings').addEventListener('click', openSiteSettings);
  $('ssCancel').addEventListener('click', () => $('siteSettingsModal').classList.remove('open'));
  $('ssSave').addEventListener('click', saveSiteSettings);
  $('ssAddCategory').addEventListener('click', addSiteCategory);
  $('ssTeam').addEventListener('change', updateSiteSettingsResource);
  $('ssWorkers').addEventListener('input', updateSiteSettingsResource);
  $('ssPerWorker').addEventListener('input', updateSiteSettingsResource);
  $('btnAddSite').addEventListener('click', () => { $('siteModal').classList.add('open'); });
  $('nsCancel').addEventListener('click', () => $('siteModal').classList.remove('open'));
  const dc = $('nsDays');
  dc.innerHTML = [1,2,3,4,5,6,0].map(d => '<label><input type="checkbox" value="' + d + '"' + (d>=1&&d<=5?' checked':'') + '> ' + WD[d] + '</label>').join('');
  $('nsCreate').addEventListener('click', createSiteFromModal);
  $('btnRefresh').addEventListener('click', refreshActive);
  $('btnExport').addEventListener('click', exportCSV);
  $('btnAutoplan').addEventListener('click', openAutoplan);
  $('btnShare').addEventListener('click', enableSharing);
  $('btnShareCfg').addEventListener('click', shareSettings);
  $('btnTheme').addEventListener('click', toggleTheme);
  // кастомный выпадающий список участка
  $('csBtn').addEventListener('click', e => { e.stopPropagation(); csOpen(); });
  $('csMenu').addEventListener('click', e => {
    const item = e.target.closest && e.target.closest('[data-site-id]');
    if(!item) return;
    const sel = $('siteSelect');
    sel.value = item.dataset.siteId;
    sel.dispatchEvent(new Event('change'));
    csOpen(false);
  });
  document.addEventListener('pointerdown', e => {
    const wrap = $('siteCSelect');
    if(wrap && wrap.classList.contains('open') && !(e.target.closest && e.target.closest('#siteCSelect')) && !(e.target.closest && e.target.closest('.cs-menu'))) csOpen(false);
  });
  document.addEventListener('keydown', e => { if(e.key === 'Escape') csOpen(false); });
  window.addEventListener('resize', () => { csOpen(false); segSync(); }, { passive:true });
  // v7: passive + ранний выход, если меню закрыто — скролл страницы не трогает DOM
  window.addEventListener('scroll', () => {
    const m = $('csMenu');
    if(m && m.classList.contains('open')) csOpen(false);
  }, { passive:true, capture:true });
  // журнал изменений
  $('btnLog').addEventListener('click', openLogModal);
  $('logClose').addEventListener('click', () => $('logModal').classList.remove('open'));
  $('logApply').addEventListener('click', renderLogTable);
  $('logFrom').addEventListener('change', renderLogTable);
  $('logTo').addEventListener('change', renderLogTable);
  $('logSite').addEventListener('change', renderLogTable);
  $('logCsv').addEventListener('click', exportLogCsv);
  $('logClear').addEventListener('click', () => {
    if(!confirm('Очистить журнал изменений полностью?')) return;
    localStorage.removeItem(LOG_KEY);
    renderLogTable();
    toast('Журнал очищен');
  });
  $('planCancel').addEventListener('click', () => { pendingPlan = null; $('planModal').classList.remove('open'); });
  $('planApply').addEventListener('click', () => {
    if(!pendingPlan) return;
    if(pendingPlan.dependencyLate.length){
      toast('План заблокирован: заготовка должна завершиться строго раньше сборки');
      return;
    }
    const site = activeSite();
    applyPlanToSite(site, pendingPlan.plan, pendingPlan.autoSplits);
    saveSites(); scheduleSharedPush(site);
    logChange(site, 'автоплан', 'изменено дат: ' + pendingPlan.changes.length +
      (pendingPlan.autoSplits.length ? ', авто-разбивок: ' + pendingPlan.autoSplits.length +
        ' (' + pendingPlan.autoSplits.map(r => (r.prodOrder || r.custOrder) + ' ' + r.parts.map(p => p.qty).join('+')).join('; ') + ')' : '') +
      (pendingPlan.assemblyLinks.length ? ', связей со сборкой: ' + pendingPlan.assemblyLinks.length : '') +
      (pendingPlan.unlinked.length ? ', без доказанной связи: ' + pendingPlan.unlinked.length : '') +
      (pendingPlan.late.length ? ', не успевают к сроку: ' + pendingPlan.late.length : ''));
    render();
    $('planModal').classList.remove('open');
    toast('План применён: изменено дат — ' + pendingPlan.changes.length +
      (pendingPlan.autoSplits.length ? ', авто-разбивок — ' + pendingPlan.autoSplits.length : ''));
    pendingPlan = null;
  });
  $('planNotes').addEventListener('change', () => {
    const site = activeSite();
    site.notes = $('planNotes').value;
    saveSites(); scheduleSharedPush(site); runAutoplanPreview();
  });
  $('planStrict').addEventListener('change', () => {
    const site = activeSite();
    site.planStrict = $('planStrict').checked;
    saveSites(); scheduleSharedPush(site); runAutoplanPreview();
  });
  $('splitCancel').addEventListener('click', () => { splitCtx = null; $('splitModal').classList.remove('open'); });
  $('splitCreate').addEventListener('click', createSplitFromModal);
  // делегированные клики: разбить / вернуть как было / ячейка сводного вида (и не дать начаться drag с кнопки)
  document.addEventListener('click', e => {
    const nearest = e.target.closest && e.target.closest('[data-nearest-operation]');
    if(nearest){ jumpToDate(nearest.dataset.nearestOperation); return; }
    const past = e.target.closest && e.target.closest('[data-show-past]');
    if(past){
      state.rangeFrom = past.dataset.showPast; state.rangeTo = addDays(todayISO(), -1); state.selectedDate = null;
      $('rangeFrom').value = state.rangeFrom; $('rangeTo').value = state.rangeTo; render(); return;
    }
    const all = e.target.closest && e.target.closest('[data-show-all-operations]');
    if(all){
      const dates = effRows(activeSite()).map(row => row.dateEff).filter(Boolean).sort();
      if(!dates.length) return;
      state.rangeFrom = dates[0]; state.rangeTo = dates[dates.length-1]; state.selectedDate = null;
      $('rangeFrom').value = state.rangeFrom; $('rangeTo').value = state.rangeTo; render(); return;
    }
    const uns = e.target.closest && e.target.closest('[data-unsplit]');
    if(uns){ restoreSplit(uns.dataset.unsplit); return; }
    const unp = e.target.closest && e.target.closest('[data-unpin]');
    if(unp){
      const s = state.sites.find(x => x.id === unp.dataset.unpin);
      if(s){ s.onBoard = false; saveSites(); logChange(s, 'с общей доски', s.name); render(); toast('«' + s.name + '» убран с общей доски'); }
      return;
    }
    const sp = e.target.closest && e.target.closest('[data-split]');
    if(sp){ openSplitModal(sp.dataset.split); return; }
    const lc = e.target.closest && e.target.closest('.lcell[data-site]');
    if(lc && !lc.classList.contains('off')){
      state.activeId = lc.dataset.site;
      state.selectedDate = null;
      switchView('board');
      jumpToDate(lc.dataset.date);
    }
  });
  document.addEventListener('mousedown', e => {
    if(e.target.closest && e.target.closest('.mini')) e.preventDefault();
  });
  const onResourceChange = () => {
    const site = activeSite();
    const before = fondOf(site);
    site.workers = Math.max(1, num($('workersInput').value) || 1);
    if(site.teamTotal){
      site.fondMin = Math.max(1, num($('fondTotalInput').value) || 1);
    } else {
      site.perWorker = Math.max(1, num($('perWorkerInput').value) || 1);
      site.fondMin = site.workers * site.perWorker;
    }
    site.resourceConfirmed = true;
    saveSites(); persistSiteConfiguration(); scheduleSharedPush(site);
    const after = fondOf(site);
    if(before !== after) logChange(site, 'ресурс', (site.teamTotal ? 'бригада ' + site.workers + ' чел' : site.workers + ' × ' + site.perWorker) +
      ': фонд ' + fmtMin(before) + ' → ' + fmtMin(after) + ' ' + (site.unit || 'мин') + '/день');
    scheduleRender();
  };
  $('workersInput').addEventListener('change', onResourceChange);
  $('perWorkerInput').addEventListener('change', onResourceChange);
  $('fondTotalInput').addEventListener('change', onResourceChange);
  // v7: поиск — дебаунс 110 мс и пропуск неизменного запроса; рендер пакетный
  let searchT = 0, lastQ = null;
  $('searchInput').addEventListener('input', () => {
    const q = $('searchInput').value;
    if(q === lastQ) return;
    lastQ = q;
    clearTimeout(searchT);
    searchT = setTimeout(scheduleRender, 110);
  });
  $('onlyBad').addEventListener('change', scheduleRender);
  $('btnPinBoard').addEventListener('click', () => {
    const site = activeSite();
    site.onBoard = !site.onBoard;
    saveSites();
    logChange(site, site.onBoard ? 'на общую доску' : 'с общей доски', site.name);
    render();
    toast(site.onBoard ? '«' + site.name + '» выведен на общую доску' : '«' + site.name + '» убран с общей доски');
  });
  // период отображения (сменные: по умолчанию — с актуальной даты)
  const onRange = () => {
    state.rangeFrom = $('rangeFrom').value || null;
    state.rangeTo = $('rangeTo').value || null;
    state.selectedDate = null;
    scheduleRender();
  };
  $('rangeFrom').addEventListener('change', onRange);
  $('rangeTo').addEventListener('change', onRange);
  $('rangeClear').addEventListener('click', () => {
    $('rangeFrom').value = ''; $('rangeTo').value = '';
    state.rangeFrom = null; state.rangeTo = null; state.selectedDate = null;
    scheduleRender();
  });
  $('btnToday').addEventListener('click', () => {
    $('rangeFrom').value = ''; $('rangeTo').value = '';
    state.rangeFrom = null; state.rangeTo = null;
    jumpToDate(todayISO());
  });
  $('viewBoard').addEventListener('click', () => switchView('board'));
  $('viewList').addEventListener('click', () => switchView('list'));
  $('viewShifts').addEventListener('click', () => switchView('shifts'));
  $('viewMulti').addEventListener('click', () => switchView('multi'));
}
async function createSiteFromModal(){
  const name = $('nsName').value.trim() || ('Участок ' + (state.sites.length + 1));
  const url = $('nsUrl').value.trim();
  const file = $('nsFile').files && $('nsFile').files[0];
  const workers = Math.max(1, num($('nsWorkers').value) || 1);
  const perWorker = Math.max(1, num($('nsPerWorker').value) || 1);
  const teamTotal = $('nsTeam').checked;
  const fond = teamTotal ? perWorker : workers * perWorker;
  const workdays = [...document.querySelectorAll('#nsDays input:checked')].map(i => Number(i.value));
  try{
    let raw = null, sourceType = 'manual', sourceUrl = '', fmt = null;
    let sheetNames = [], sheetCount = 0;
    if(file){
      const text = await file.text();
      fmt = detectFormat(text);
      raw = parseAny(text, fmt);
      sourceType = 'csv';
    } else if(url){
      const res = await fetchSiteRows(url);
      raw = res.rows; fmt = res.format;
      sourceType = 'url'; sourceUrl = url;
      const wholeBook = fmt === 'upkp' && !/[#?&]gid=/.test(url || '');
      if(wholeBook){
        // Первый CSV нужен только для определения формата; затем сразу читаем всю книгу.
        const full = await fetchSiteRows(url, {autoMonth:true, format:'upkp'});
        raw = full.rows; sheetNames = full.sheetNames || []; sheetCount = full.sheetCount || 0;
      }
    } else {
      toast('Укажите ссылку на Google Sheets или выберите CSV-файл');
      return;
    }
    if(!raw.length){ toast('Не удалось прочитать строки операций'); return; }
    const autoMonth = sourceType === 'url' && fmt === 'upkp' && !/[#?&]gid=/.test(url || '');
    const site = { id:'s' + Date.now().toString(36), name, sourceType, sourceUrl,
      workers, perWorker: teamTotal ? 0 : perWorker, teamTotal, fondMin:fond, workdays:workdays.length?workdays:[1,2,3,4,5], rows:makeRows(raw), overrides:{},
      format:fmt, unit:fmt === 'upkp' ? 'секц.' : 'мин', autoMonth,
      syncOk:sourceType === 'url', syncAt:sourceType === 'url' ? new Date().toISOString() : null,
      syncSheetNames:sheetNames, syncSheetCount:sheetCount };
    site.sourceSiteKey = 'manual:' + site.id; site.group = 'Прочее'; site.category = 'Прочее'; site.resourceConfirmed = true;
    state.sites.push(site);
    state.activeId = site.id;
    saveSites(); persistSiteConfiguration();
    logChange(site, 'участок создан', name + ': ' + raw.length + ' операций, фонд ' + fond);
    render();
    $('siteModal').classList.remove('open');
    $('nsName').value = ''; $('nsUrl').value = ''; $('nsFile').value = ''; $('nsTeam').checked = false;
    toast('Участок «' + name + '» создан: ' + raw.length + ' операций');
  }catch(err){
    alert('Ошибка загрузки: ' + err.message + '\n\nЕсли таблица закрыта — откройте доступ «по ссылке» (просмотр) или скачайте CSV вручную и загрузите файлом.');
  }
}
const siteSyncFlights = new WeakMap(), siteSyncAttempts = new WeakMap();
let syncApplyQueue = Promise.resolve();
function nextSyncApplyTurn(){
  // Общая книга пробуждает все участки одновременно. Очередь не даёт объединиться
  // тяжёлым обновлениям в один длинный кадр; в Node-проверках остаются микрозадачи.
  const turn = syncApplyQueue.then(() => new Promise(resolve => {
    if(typeof requestAnimationFrame === 'function') requestAnimationFrame(() => resolve());
    else resolve();
  }));
  syncApplyQueue = turn.catch(() => {});
  return turn;
}
async function refreshSite(site, options = {}){
  if(site.sourceType !== 'url' || !site.sourceUrl) return false;
  const active = siteSyncFlights.get(site);
  if(active) return active;
  siteSyncAttempts.set(site, Date.now());
  const flight = (async () => {
    try{
      const res = await fetchSiteRows(site.sourceUrl, site);
      if(!res.rows.length && site.rows.length){
        throw new Error('Google Sheets вернула 0 операций. Сохранённые данные оставлены без изменений.');
      }
      await nextSyncApplyTurn();
      site.rows = makeRows(res.rows, site.rows);
      site.format = res.format || site.format;
      site.syncSheetNames = res.sheetNames || [];
      site.syncSheetCount = res.sheetCount || 0;
      if(site.id === 's_upkp') site.snapshotVersion = UPKP_SNAPSHOT_VERSION;
      if(!site.unit) site.unit = res.format === 'upkp' ? 'секц.' : 'мин';
      site.syncOk = true;
      site.syncAt = new Date().toISOString();
      site.syncErr = null;
      if(site.format === 'ops-shared') discoverSharedSites({...res, rows:res.sharedRows || res.rows});
      if(!options.deferSave) saveSites();
      return true;
    }catch(err){
      site.syncOk = false;
      site.syncErr = String(err && err.message || err);
      if(!options.deferSave) saveSites();
      throw err;
    }
  })();
  siteSyncFlights.set(site, flight);
  try{ return await flight; }
  finally{ if(siteSyncFlights.get(site) === flight) siteSyncFlights.delete(site); }
}
function renderSyncStatus(){
  const el = $('syncStatus');
  const site = activeSite();
  if(!el || !site) return;
  if(site.sourceType !== 'url'){ el.textContent = ''; return; }
  const shared = site.sharedOn ? (site.sharedOk !== false ? ' · совместно ✓' : ' · совместно ✗') : '';
  if(site.syncOk){
    const t = site.syncAt ? new Date(site.syncAt) : null;
    const sheets = site.syncSheetCount ? ' · ' + site.syncSheetCount + ' листов' : '';
    el.textContent = '● ' + site.rows.length + ' операций' + sheets + (t ? ' · ' + pad(t.getHours()) + ':' + pad(t.getMinutes()) : '') + shared;
    el.style.color = site.sharedOk === false ? 'var(--red)' : 'var(--green)';
    el.title = 'Google Sheets синхронизирован: вся книга' + (site.syncSheetNames && site.syncSheetNames.length ? '\n' + site.syncSheetNames.join(', ') : '') + (site.syncErr ? '\n' + site.syncErr : '');
  } else if(site.rows.length){
    el.textContent = '◐ Резервная копия · ' + site.rows.length + ' операций' + shared;
    el.style.color = 'var(--amber)';
    el.title = (site.syncErr || 'Нет связи с Google Sheets') + '\nДанные не очищены. Нажмите «Обновить» после восстановления связи.';
  } else {
    el.textContent = '○ Нет данных · проверьте связь' + shared;
    el.style.color = 'var(--red)';
    el.title = site.syncErr || site.sharedErr || 'проверьте интернет и доступ к таблице';
  }
}
async function refreshActive(){
  const site = activeSite();
  if(site.sourceType !== 'url' || !site.sourceUrl){ toast('У активного участка нет ссылки на Google Sheets (снимок или CSV)'); return; }
  try{
    await refreshSite(site);
    render();
    toast('Обновлено из Google Sheets: ' + site.rows.length + ' операций (ручные переносы сохранены, где позиции не изменились)');
  }catch(err){
    alert('Не удалось обновить: ' + err.message);
  }
}

async function refreshSiteGroup(sites){
  const results = await Promise.allSettled(sites.map(site => refreshSite(site, {deferSave:true})));
  const okCount = results.filter(result => result.status === 'fulfilled' && result.value).length;
  const failCount = results.length - okCount;
  // Данные всех участков меняются атомарно относительно сохранения: одна запись
  // вместо повторной сериализации растущего плана для каждого из 17 участков.
  saveSites();
  scheduleRender();
  return { okCount, failCount };
}
async function autoRefreshSites(minGap = 60000){
  if(document.hidden) return { okCount:0, failCount:0 };
  const now = Date.now();
  const sites = state.sites.filter(site => site.sourceType === 'url' && site.sourceUrl &&
    now - (siteSyncAttempts.get(site) || 0) >= minGap);
  if(!sites.length) return { okCount:0, failCount:0 };
  return refreshSiteGroup(sites); // без toast: фоновая синхронизация не отвлекает
}

/* ---------- init ---------- */
function init(){
  if(init._done) return; // защита от двойного запуска (DOMContentLoaded + readyState)
  init._done = true;
  applyTheme(currentTheme());
  state.sites = loadSites();
  state.activeId = state.sites[0].id;
  wire();
  state._enterView = true;
  render();
  raf(segSync);
  // после загрузки шрифтов ширины кнопок меняются — пересчитываем плашку
  if(document.fonts && document.fonts.ready) document.fonts.ready.then(() => segSync());
  setTimeout(async () => {
    // Сначала пользовательская конфигурация, затем общий XLSX: переименование,
    // ресурсы и tombstone не должны перезаписываться системными defaults.
    if(await pullSiteConfiguration()) render();
    const sites = state.sites.filter(site => site.sourceType === 'url' && site.sourceUrl);
    const { okCount, failCount } = await refreshSiteGroup(sites);
    if(failCount && !okCount) toast('Нет связи с Google Sheets — показаны сохранённые данные. Проверьте интернет и откройте страницу по ссылке Kimi (localhost), а не как файл.');
    else if(okCount) toast('Синхронизировано участков: ' + okCount + (failCount ? ' (не удалось: ' + failCount + ')' : ''));
  }, 600);
  // Вся книга перечитывается каждые 10 минут на видимой вкладке. Новый лист
  // появится в плане без изменения настроек и без перечисления месяцев в коде.
  setInterval(() => { autoRefreshSites(60000); }, 10 * 60 * 1000);
  document.addEventListener('visibilitychange', () => {
    if(!document.hidden) autoRefreshSites(60000);
  });
  window.addEventListener('online', () => { autoRefreshSites(60000); });
  // совместный режим: опрос общих правок каждые 20 секунд
  setInterval(async () => {
    if(document.hidden) return;
    let changed = false;
    for(const s of state.sites){
      if(!s.sharedOn) continue;
      if(await pullShared(s)) changed = true;
    }
    if(changed) scheduleRender(); // v7: опрос нескольких участков — один рендер на кадр
  }, 20000);
}
if(typeof window !== 'undefined'){
  window.PLANNER = { parseCSVText, parseUPKP, parseUPKPGrid, parseUPKPWorkbook, parseAny, detectFormat, curMonthSheet, orderNum, buildPositions, computeDays, moveRowToDate, autoplain, makeRows,
    state, effRows, parseDate, jumpToDate, nextWorkDay, parseGUrl, suggestFond, fondOf, LS_KEY, loadSites, csvToGrid,
    parseNotes, distributeQty, distributeMinutes, buildSplitRecord, buildSplitRecordQtys, splitQtysByFond, assignSplitDates, applyPlanToSite, findPart,
    positionKey, upkpOrderLabel, upkpCustomer, upkpProductLabel, displayProduct, displayCustomer,
    logChange, readLog, LOG_KEY, fmtTs, currentTheme, toggleTheme,
    sharedKey, pullShared, pushShared, ghCfg, ghSharedPath, enableSharing, shareSettings,
    fetchSiteRows, refreshSite, refreshSiteGroup, autoRefreshSites, render, scheduleRender, saveSites, todayISO, addDays, isWorkday, siteResLine,
    defaultSites, discoverSharedSites, sharedRowsFor, fetchSharedWorkbook, siteKeyOf, configFromSite, applySiteConfig,
    persistSiteConfiguration, pullSiteConfiguration, pushSiteConfiguration, saveSiteSettings, openSiteSettings,
    assemblyIndex, assemblyConstraintsFor, cardHTML, rowContextCells,
    get pendingAutoplan(){ return pendingPlan; },
    get siteConfiguration(){ return JSON.parse(JSON.stringify(siteConfig)); } };
  if(typeof document !== 'undefined' && document.getElementById && typeof window.addEventListener !== 'undefined'){
    window.addEventListener('DOMContentLoaded', init);
    if(document.readyState !== 'loading') init();
  }
}


'use strict';
/* ==========================================================================
   ПОТОК / workspace extensions
   Карусель панелей, виджеты, расчёт загрузки, профили и нативный WebMCP.
   Файл выполняется после app.js в том же standalone-бандле.
   ========================================================================== */
var workspaceReady = false;

/* ---------- small utilities -------------------------------------------- */
function w$(id){ return document.getElementById(id); }
function wEsc(value){
  return String(value == null ? '' : value).replace(/[&<>"']/g, ch => ({
    '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'
  })[ch]);
}
function wStoreGet(key, fallback){
  try{
    const value = JSON.parse(localStorage.getItem(key));
    return value == null ? fallback : value;
  }catch(error){ return fallback; }
}
function wStoreSet(key, value){
  try{ localStorage.setItem(key, JSON.stringify(value)); }catch(error){}
}
function wPlanner(){ return window.PLANNER; }
function wActiveSite(){
  const planner = wPlanner();
  return planner.state.sites.find(site => site.id === planner.state.activeId) || planner.state.sites[0];
}
function wFmt(value){ return Math.round(value || 0).toLocaleString('ru-RU'); }
function wIsIsoDate(value){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return false;
  const date = new Date(value + 'T00:00:00Z');
  return !Number.isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

/* ---------- workload calculations -------------------------------------- */
function isoWeekInfo(value){
  const date = new Date(value + 'T00:00:00Z');
  if(Number.isNaN(date.getTime())) return { week:0, year:0, key:'' };
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const year = date.getUTCFullYear();
  const yearStart = new Date(Date.UTC(year, 0, 1));
  const week = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return { week, year, key:year + '-W' + String(week).padStart(2, '0') };
}
function isoWeekStart(key){
  const match = String(key || '').match(/^(\d{4})-W(\d{2})$/);
  if(!match) return null;
  const year = Number(match[1]), week = Number(match[2]);
  const date = new Date(Date.UTC(year, 0, 4));
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() - day + 1 + (week - 1) * 7);
  return date.toISOString().slice(0, 10);
}
function currentLoadSelection(){
  const today = wPlanner().todayISO();
  return { day:today, week:isoWeekInfo(today).key, month:today.slice(0, 7) };
}
function workloadPeriod(site, period, anchor, fullPeriod){
  const planner = wPlanner();
  const reference = anchor || planner.todayISO();
  const partial = (period === 'week' || period === 'month') && !fullPeriod;
  let from = reference;
  let to = reference;
  if(period === 'week'){
    const day = new Date(reference + 'T00:00:00').getDay();
    const weekStart = planner.addDays(reference, day === 0 ? -6 : 1 - day);
    // Текущая неделя — только остаток; другую неделю можно открыть полностью.
    from = fullPeriod ? weekStart : reference;
    to = planner.addDays(weekStart, 6);
  }else if(period === 'month'){
    const monthStart = reference.slice(0, 8) + '01';
    // По умолчанию месяц означает остаток: с выбранного/сегодняшнего дня.
    from = fullPeriod ? monthStart : reference;
    const next = new Date(monthStart + 'T00:00:00');
    next.setMonth(next.getMonth() + 1);
    next.setDate(0);
    to = next.getFullYear() + '-' + String(next.getMonth() + 1).padStart(2, '0') + '-' + String(next.getDate()).padStart(2, '0');
  }
  const rows = planner.effRows(site).filter(row => row.dateEff && row.dateEff >= from && row.dateEff <= to);
  const load = rows.reduce((sum, row) => sum + (row.minutes || 0), 0);
  let workdays = period === 'day' ? (planner.isWorkday(site, from) ? 1 : 0) : 0;
  if(period !== 'day'){
    for(let date = from; date <= to; date = planner.addDays(date, 1)){
      if(planner.isWorkday(site, date)) workdays++;
    }
  }
  const capacity = planner.fondOf(site) * workdays;
  const week = period === 'week' ? isoWeekInfo(from) : null;
  return {
    period, reference, from, to, load, capacity, workdays, rows:rows.length, partial,
    pct:capacity > 0 ? Math.round(load / capacity * 100) : (load > 0 ? 100 : 0), unit:site.unit || 'мин', siteId:site.id,
    weekNumber:week ? week.week : null, weekYear:week ? week.year : null,
    weekKey:week ? week.key : null, monthKey:period === 'month' ? reference.slice(0, 7) : null
  };
}
function selectedWorkload(site){
  const period = loadPeriodState;
  const selection = loadSelectionState[period];
  if(period === 'week'){
    const today = wPlanner().todayISO(), current = selection === isoWeekInfo(today).key;
    return workloadPeriod(site, period, current ? today : (isoWeekStart(selection) || today), !current);
  }
  if(period === 'month'){
    const today = wPlanner().todayISO();
    const current = selection === today.slice(0, 7);
    return workloadPeriod(site, period, current ? today : selection + '-01', !current);
  }
  return workloadPeriod(site, 'day', wIsIsoDate(selection) ? selection : wPlanner().todayISO(), true);
}
function workloadLabel(period){ return period === 'day' ? 'День' : period === 'week' ? 'Неделя' : 'Месяц'; }
function workloadPeriodNote(load){
  const origin = load.reference === wPlanner().todayISO() ? 'с сегодня' : 'с выбранной даты';
  if(load.period === 'week') return 'Неделя № ' + load.weekNumber + ' · ' + load.from + ' — ' + load.to + (load.partial ? ' · ' + origin : ' · полная неделя');
  if(load.period === 'month') return 'Месяц · ' + load.from + ' — ' + load.to + (load.partial ? ' · ' + origin : ' · полный месяц');
  return 'День · ' + load.from;
}
function workloadBar(load, compact){
  const over = load.pct > 100;
  return '<div class="load-track"><i class="' + (over ? 'over' : '') + '" style="width:' + Math.min(100, Math.max(0, load.pct)) + '%"></i></div>' +
    (compact ? '' : '<small>' + wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + wEsc(load.unit) + ' · ' + load.rows + ' операций</small>');
}

/* ---------- panel carousel --------------------------------------------- */
const DOCK_PANELS = [
  ['Дни загрузки', 'Навигация по сменам'],
  ['Загрузка', 'Ресурс и фактическая норма'],
  ['Автоплан с Kimi', 'Явный запрос · без API-ключа'],
  ['Мои виджеты', 'Персональная сводка']
];
let dockIndexState = Math.max(0, Math.min(3, Number(localStorage.getItem('planner.dockPanel')) || 0));
let dockMoveTimer = 0;
function setDockPanel(index){
  dockIndexState = (index + DOCK_PANELS.length) % DOCK_PANELS.length;
  const track = w$('dockTrack');
  if(!track) return;
  track.classList.add('moving');
  track.style.transform = 'translateX(' + (-dockIndexState * 100) + '%)';
  clearTimeout(dockMoveTimer);
  dockMoveTimer = setTimeout(() => track.classList.remove('moving'), 420);
  w$('dockTitle').textContent = DOCK_PANELS[dockIndexState][0];
  w$('dockFooterName').textContent = DOCK_PANELS[dockIndexState][1];
  w$('dockIndex').textContent = String(dockIndexState + 1).padStart(2, '0');
  document.querySelectorAll('.dock-dot').forEach((dot, i) => {
    dot.classList.toggle('on', i === dockIndexState);
    dot.setAttribute('aria-current', i === dockIndexState ? 'true' : 'false');
  });
  document.querySelectorAll('.dock-panel').forEach((panel, i) => {
    panel.setAttribute('aria-hidden', i === dockIndexState ? 'false' : 'true');
    panel.inert = i !== dockIndexState;
  });
  if(workspaceReady){
    const site = wActiveSite();
    if(site && dockIndexState === 1) renderLoadPanel();
    else if(site && dockIndexState === 2) syncKimiFields(site);
    else if(site && dockIndexState === 3) renderCustomWidgets(site);
  }
  try{ localStorage.setItem('planner.dockPanel', String(dockIndexState)); }catch(error){}
}

/* ---------- widgets ----------------------------------------------------- */
const WIDGET_KEY = 'planner.widgets.v1';
const WIDGET_COLLAPSE_KEY = 'planner.widgetCollapsed.v1';
const WIDGET_CATALOG = {};
const WIDGET_GROUPS = ['Загрузка', 'Операции', 'Сроки', 'Заказы', 'Ресурсы', 'Участки', 'Синхронизация', 'Журнал', 'Действия'];
function addWidget(id, group, title, description, render){
  WIDGET_CATALOG[id] = { group, title, description, render };
}
function widgetMetric(value, caption, detail, tone){
  return '<strong class="widget-big"' + (tone ? ' style="color:var(--' + tone + ')"' : '') + '>' + wEsc(value) + '</strong>' +
    '<span class="widget-caption">' + wEsc(caption) + '</span>' + (detail ? '<span class="widget-caption">' + wEsc(detail) + '</span>' : '');
}
function widgetRows(site){ return wPlanner().effRows(site); }
function widgetSum(rows){ return rows.reduce((sum, row) => sum + (row.minutes || 0), 0); }
function widgetRowsIn(site, load){
  return widgetRows(site).filter(row => row.dateEff && row.dateEff >= load.from && row.dateEff <= load.to);
}
function widgetRowList(rows, site, limit){
  const planner = wPlanner();
  if(!rows.length) return '<span class="widget-caption">Нет операций для отображения</span>';
  return '<ul class="widget-list">' + rows.slice(0, limit || 4).map(row => {
    const order = site.format === 'upkp' ? planner.upkpOrderLabel(row.custOrder) : (row.prodOrder || row.custOrder);
    return '<li><b>' + wEsc(order) + '</b> · ' + wEsc(planner.displayProduct(row, site)) + '<br>' + wEsc(row.dateEff || 'без даты') + ' · ' + wFmt(row.minutes || 0) + ' ' + wEsc(site.unit || 'мин') + '</li>';
  }).join('') + '</ul>';
}
function widgetGroups(rows, key){
  const groups = new Map();
  rows.forEach(row => {
    const name = typeof key === 'function' ? key(row) : row[key];
    const label = String(name || 'Не указано');
    if(!groups.has(label)) groups.set(label, { name:label, rows:0, total:0 });
    const group = groups.get(label);
    group.rows++; group.total += row.minutes || 0;
  });
  return [...groups.values()].sort((a, b) => b.total - a.total);
}
function widgetSiteLogs(site){ return wPlanner().readLog().filter(entry => entry.site === site.id); }

/* Загрузка: все показатели вычисляются из реальных операций и фонда. */
addWidget('load', 'Загрузка', 'Выбранный период', 'Загрузка с учётом выбранного дня, недели или месяца', site => {
  const load = selectedWorkload(site);
  return widgetMetric(load.pct + '%', workloadPeriodNote(load), wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + load.unit) + workloadBar(load);
});
addWidget('free', 'Загрузка', 'Свободный ресурс', 'Остаток фонда в выбранном периоде', site => {
  const load = selectedWorkload(site), free = load.capacity - load.load;
  return widgetMetric(wFmt(Math.max(0, free)) + ' ' + load.unit, free >= 0 ? 'свободно в выбранном периоде' : 'перегруз сверх фонда', free < 0 ? 'Не хватает: ' + wFmt(-free) + ' ' + load.unit : workloadPeriodNote(load), free < 0 ? 'red' : '');
});
addWidget('workdays', 'Загрузка', 'Рабочие дни', 'Число рабочих дней и операций периода', site => {
  const load = selectedWorkload(site);
  return widgetMetric(load.workdays, 'рабочих дней', load.rows + ' операций · ' + workloadPeriodNote(load));
});
addWidget('average', 'Загрузка', 'Средняя загрузка', 'Средняя норма на рабочий день периода', site => {
  const load = selectedWorkload(site), average = Math.round(load.load / Math.max(1, load.workdays));
  return widgetMetric(wFmt(average) + ' ' + load.unit, 'в среднем на рабочий день', 'Фонд дня: ' + wFmt(wPlanner().fondOf(site)) + ' ' + load.unit);
});
addWidget('peak', 'Загрузка', 'Пиковый день', 'Самый загруженный день выбранного периода', site => {
  const planner = wPlanner(), load = selectedWorkload(site), rows = widgetRowsIn(site, load);
  const days = widgetGroups(rows, 'dateEff').sort((a, b) => b.total - a.total);
  if(!days.length) return widgetMetric('—', 'операций в периоде нет', workloadPeriodNote(load));
  const peak = days[0], pct = Math.round(peak.total / planner.fondOf(site) * 100);
  return widgetMetric(wFmt(peak.total) + ' ' + load.unit, peak.name + ' · ' + pct + '% фонда', peak.rows + ' операций', pct > 100 ? 'red' : '');
});
addWidget('overload', 'Загрузка', 'Дни перегруза', 'Дни, где норма превышает ресурс участка', site => {
  const planner = wPlanner(), load = selectedWorkload(site), fond = planner.fondOf(site);
  const days = widgetGroups(widgetRowsIn(site, load), 'dateEff').filter(day => day.total > fond);
  return widgetMetric(days.length, 'дней выше фонда', days.length ? days.map(day => day.name + ' · +' + wFmt(day.total - fond)).join('\n') : 'Перегруза в выбранном периоде нет', days.length ? 'red' : '');
});
addWidget('today', 'Загрузка', 'Сегодня', 'Загрузка текущего календарного дня', site => {
  const load = workloadPeriod(site, 'day');
  return widgetMetric(load.pct + '%', 'сегодня · ' + load.from, wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + load.unit) + workloadBar(load);
});
addWidget('tomorrow', 'Загрузка', 'Завтра', 'Плановая загрузка следующего календарного дня', site => {
  const date = wPlanner().addDays(wPlanner().todayISO(), 1), load = workloadPeriod(site, 'day', date);
  return widgetMetric(load.pct + '%', 'завтра · ' + date, wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + load.unit) + workloadBar(load);
});
addWidget('month', 'Загрузка', 'Остаток месяца', 'С сегодняшнего дня до конца текущего месяца', site => {
  const load = workloadPeriod(site, 'month');
  return widgetMetric(load.pct + '%', 'с сегодня до конца месяца', wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + load.unit) + workloadBar(load);
});
addWidget('week', 'Загрузка', 'Текущая неделя', 'Календарная неделя с понедельника по воскресенье', site => {
  const load = workloadPeriod(site, 'week');
  return widgetMetric(load.pct + '%', 'неделя № ' + load.weekNumber, load.from + ' — ' + load.to) + workloadBar(load);
});

/* Операции и состав производственного плана. */
addWidget('operations', 'Операции', 'Все операции', 'Строки, позиции и суммарная норма', site => {
  const rows = widgetRows(site);
  return widgetMetric(rows.length, 'операций · ' + wPlanner().buildPositions(site).size + ' позиций', 'Сумма норм: ' + wFmt(widgetSum(rows)) + ' ' + (site.unit || 'мин'));
});
addWidget('positions', 'Операции', 'Позиции', 'Уникальные производственные позиции', site => widgetMetric(wPlanner().buildPositions(site).size, 'уникальных позиций', 'Без склеивания разных договоров'));
addWidget('quantity', 'Операции', 'Количество изделий', 'Сумма количества по всем операциям', site => {
  const rows = widgetRows(site), qty = rows.reduce((sum, row) => sum + (row.qty || 0), 0);
  return widgetMetric(wFmt(qty), 'единиц продукции', rows.length + ' операций');
});
addWidget('unscheduled', 'Операции', 'Без даты', 'Операции, которые ещё нужно распределить', site => {
  const rows = widgetRows(site).filter(row => !row.dateEff);
  return widgetMetric(rows.length, 'операций без даты', wFmt(widgetSum(rows)) + ' ' + (site.unit || 'мин'), rows.length ? 'amber' : '');
});
addWidget('split', 'Операции', 'Разбитые позиции', 'Позиции, разнесённые на несколько дней', site => {
  const positions = [...wPlanner().buildPositions(site).values()].filter(position => position.split);
  return widgetMetric(positions.length, 'позиций на нескольких днях', positions.length ? positions.slice(0, 3).map(p => (p.prodOrder || p.custOrder) + ': ' + p.dates.length + ' дн.').join('\n') : 'Разбитых позиций нет');
});
addWidget('todayOperations', 'Операции', 'Операции сегодня', 'Карточки текущей смены', site => {
  const rows = widgetRows(site).filter(row => row.dateEff === wPlanner().todayISO());
  return widgetMetric(rows.length, 'операций сегодня', wFmt(widgetSum(rows)) + ' ' + (site.unit || 'мин')) + widgetRowList(rows, site, 3);
});
addWidget('largestOperation', 'Операции', 'Крупнейшая операция', 'Строка с максимальной нормой', site => {
  const rows = widgetRows(site).slice().sort((a, b) => (b.minutes || 0) - (a.minutes || 0));
  return rows.length ? widgetRowList(rows, site, 1) : widgetMetric('—', 'операций нет');
});
addWidget('productMix', 'Операции', 'Номенклатура', 'Уникальные изделия и основные группы', site => {
  const groups = widgetGroups(widgetRows(site), row => wPlanner().displayProduct(row, site));
  return widgetMetric(groups.length, 'уникальных наименований', groups.slice(0, 3).map(group => group.name + ' · ' + group.rows).join('\n'));
});
addWidget('problem', 'Операции', 'Проблемные операции', 'Без даты или с датой позже ТДГ', site => {
  const rows = widgetRows(site).filter(row => !row.dateEff || (row.tdg && row.dateEff > row.tdg));
  return widgetMetric(rows.length, 'требуют внимания', rows.length ? 'Проверьте даты и сроки' : 'Проблемных операций нет', rows.length ? 'red' : '') + widgetRowList(rows, site, 2);
});

/* Сроки: раздельно ТДГ и даты договоров. */
addWidget('deadlines', 'Сроки', 'Сроки и просрочка', 'Просроченные и ближайшие ТДГ', site => {
  const rows = widgetRows(site), late = rows.filter(row => row.tdg && row.dateEff && row.dateEff > row.tdg);
  const next = rows.filter(row => row.tdg && row.dateEff && row.dateEff <= row.tdg).sort((a, b) => a.tdg.localeCompare(b.tdg));
  return widgetMetric(late.length, 'операций после ТДГ', late.length ? 'Нужна корректировка плана' : 'Просроченных сроков нет', late.length ? 'red' : '') + widgetRowList(next, site, 3);
});
addWidget('overdue', 'Сроки', 'Просрочка ТДГ', 'Только операции позже требуемой даты', site => {
  const rows = widgetRows(site).filter(row => row.tdg && row.dateEff && row.dateEff > row.tdg);
  return widgetMetric(rows.length, 'операций после ТДГ', wFmt(widgetSum(rows)) + ' ' + (site.unit || 'мин'), rows.length ? 'red' : '');
});
addWidget('dueToday', 'Сроки', 'ТДГ сегодня', 'Операции с требуемой датой сегодня', site => {
  const today = wPlanner().todayISO(), rows = widgetRows(site).filter(row => row.tdg === today);
  return widgetMetric(rows.length, 'операций с ТДГ сегодня', today) + widgetRowList(rows, site, 3);
});
addWidget('dueWeek', 'Сроки', 'ТДГ на этой неделе', 'Сроки текущей календарной недели', site => {
  const load = workloadPeriod(site, 'week'), rows = widgetRows(site).filter(row => row.tdg && row.tdg >= load.from && row.tdg <= load.to);
  return widgetMetric(rows.length, 'операций · неделя № ' + load.weekNumber, load.from + ' — ' + load.to) + widgetRowList(rows, site, 3);
});
addWidget('nextDeadline', 'Сроки', 'Ближайший ТДГ', 'Следующий срок начиная с сегодня', site => {
  const today = wPlanner().todayISO(), rows = widgetRows(site).filter(row => row.tdg && row.tdg >= today).sort((a, b) => a.tdg.localeCompare(b.tdg));
  return rows.length ? widgetMetric(rows[0].tdg, 'ближайший ТДГ', rows[0].prodOrder || rows[0].custOrder) + widgetRowList(rows, site, 2) : widgetMetric('—', 'будущих ТДГ нет');
});
addWidget('contractLate', 'Сроки', 'После даты договора', 'Операции позже договорной даты', site => {
  const rows = widgetRows(site).filter(row => row.contractDate && row.dateEff && row.dateEff > row.contractDate);
  return widgetMetric(rows.length, 'операций после даты договора', rows.length ? 'Проверьте договорные ограничения' : 'Нарушений нет', rows.length ? 'amber' : '');
});

/* Договоры, заказчики и продукция. */
addWidget('contracts', 'Заказы', 'Договоры', 'Число уникальных договоров в плане', site => {
  const groups = widgetGroups(widgetRows(site), row => wPlanner().orderNum(row.custOrder) || row.custOrder);
  return widgetMetric(groups.length, 'договоров', groups.slice(0, 3).map(group => '№ ' + group.name + ' · ' + group.rows + ' оп.').join('\n'));
});
addWidget('customers', 'Заказы', 'Контрагенты', 'Число уникальных контрагентов', site => {
  const groups = widgetGroups(widgetRows(site), row => wPlanner().displayCustomer(row, site) || row.customer || 'Не указан');
  return widgetMetric(groups.length, 'контрагентов', groups.slice(0, 3).map(group => group.name).join('\n'));
});
addWidget('topCustomer', 'Заказы', 'Крупный контрагент', 'Контрагент с наибольшей суммой норм', site => {
  const groups = widgetGroups(widgetRows(site), row => wPlanner().displayCustomer(row, site) || row.customer || 'Не указан');
  return groups.length ? widgetMetric(groups[0].name, wFmt(groups[0].total) + ' ' + (site.unit || 'мин'), groups[0].rows + ' операций') : widgetMetric('—', 'контрагентов нет');
});
addWidget('topProduct', 'Заказы', 'Основное изделие', 'Продукция с наибольшей суммой норм', site => {
  const groups = widgetGroups(widgetRows(site), row => wPlanner().displayProduct(row, site));
  return groups.length ? widgetMetric(groups[0].name, wFmt(groups[0].total) + ' ' + (site.unit || 'мин'), groups[0].rows + ' операций') : widgetMetric('—', 'продукции нет');
});
addWidget('largestContract', 'Заказы', 'Крупный договор', 'Договор с наибольшей суммой норм', site => {
  const groups = widgetGroups(widgetRows(site), 'custOrder');
  return groups.length ? widgetMetric(site.format === 'upkp' ? wPlanner().upkpOrderLabel(groups[0].name) : groups[0].name, wFmt(groups[0].total) + ' ' + (site.unit || 'мин'), groups[0].rows + ' операций') : widgetMetric('—', 'договоров нет');
});

/* Ресурс участка и календарь. */
addWidget('resource', 'Ресурсы', 'Фонд участка', 'Состав и дневной ресурс', site => widgetMetric(wFmt(wPlanner().fondOf(site)) + ' ' + (site.unit || 'мин'), 'фонд в день', wPlanner().siteResLine(site)));
addWidget('team', 'Ресурсы', 'Состав бригады', 'Люди и способ расчёта нормы', site => widgetMetric(site.workers || '—', 'человек', site.teamTotal ? 'Норма задана на бригаду целиком' : 'Норма на человека: ' + wFmt(site.perWorker || 0) + ' ' + (site.unit || 'мин')));
addWidget('workweek', 'Ресурсы', 'Рабочая неделя', 'Настроенные рабочие дни участка', site => {
  const names = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
  return widgetMetric((site.workdays || []).length, 'рабочих дней в неделю', (site.workdays || []).map(day => names[day]).join(', '));
});
addWidget('unit', 'Ресурсы', 'Единица планирования', 'Минуты, секции или другая норма', site => widgetMetric(site.unit || 'мин', 'единица загрузки', site.format === 'upkp' ? 'Количество секций равно норме' : 'Норма из исходной таблицы'));
addWidget('capacity', 'Ресурсы', 'Ресурс периода', 'Фонд всех рабочих дней выбранного периода', site => {
  const load = selectedWorkload(site);
  return widgetMetric(wFmt(load.capacity) + ' ' + load.unit, 'доступный ресурс', load.workdays + ' рабочих дней · ' + workloadPeriodNote(load));
});

/* Сравнение участков и общая доска. */
addWidget('sites', 'Участки', 'Все участки', 'Число участков и операций в системе', site => {
  const sites = wPlanner().state.sites, operations = sites.reduce((sum, item) => sum + item.rows.length, 0);
  return widgetMetric(sites.length, 'участков', operations + ' исходных операций');
});
addWidget('compare', 'Участки', 'Сравнение загрузки', 'Все участки за один выбранный период', () => {
  return '<ul class="widget-list">' + wPlanner().state.sites.map(item => {
    const load = selectedWorkload(item);
    return '<li><b>' + wEsc(item.name) + '</b><br>' + load.pct + '% · ' + wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + wEsc(load.unit) + '</li>';
  }).join('') + '</ul>';
});
addWidget('busiestSite', 'Участки', 'Самый загруженный участок', 'Максимальный процент за выбранный период', () => {
  const items = wPlanner().state.sites.map(site => ({ site, load:selectedWorkload(site) })).sort((a, b) => b.load.pct - a.load.pct);
  return items.length ? widgetMetric(items[0].load.pct + '%', items[0].site.name, workloadPeriodNote(items[0].load), items[0].load.pct > 100 ? 'red' : '') : widgetMetric('—', 'участков нет');
});
addWidget('boards', 'Участки', 'Общая доска', 'Участки, вынесенные на общую доску', () => {
  const pinned = wPlanner().state.sites.filter(item => item.onBoard);
  return widgetMetric(pinned.length, 'выведено на общую доску', pinned.length ? pinned.map(item => item.name).join('\n') : 'Используйте кнопку «На общую доску»');
});

/* Источник и синхронизация. */
addWidget('sync', 'Синхронизация', 'Состояние данных', 'Live-источник или резервная копия', site => {
  const status = site.syncOk ? 'Синхронизировано' : site.rows.length ? 'Резервная копия' : 'Нет данных';
  return widgetMetric(site.rows.length, 'операций в источнике', status + (site.syncAt ? ' · ' + new Date(site.syncAt).toLocaleTimeString('ru-RU', {hour:'2-digit', minute:'2-digit'}) : ''), site.syncOk ? '' : 'amber');
});
addWidget('source', 'Синхронизация', 'Источник участка', 'Тип и адрес источника данных', site => widgetMetric(site.sourceType === 'url' ? 'Google Sheets' : 'CSV / снимок', site.autoMonth ? 'вся книга' + (site.syncSheetCount ? ' · ' + site.syncSheetCount + ' листов' : '') : 'фиксированный источник', site.sourceUrl || 'Локальный источник'));
addWidget('shared', 'Синхронизация', 'Совместный режим', 'Состояние общих правок GitHub', site => widgetMetric(site.sharedOn ? 'Включён' : 'Выключен', 'совместные правки', site.sharedOn ? (site.sharedOk === false ? 'Есть ошибка синхронизации' : 'Общие данные активны') : 'Ручные правки хранятся локально', site.sharedOn && site.sharedOk === false ? 'red' : ''));
addWidget('syncTime', 'Синхронизация', 'Время обновления', 'Когда последний раз получены живые данные', site => widgetMetric(site.syncAt ? new Date(site.syncAt).toLocaleTimeString('ru-RU', {hour:'2-digit', minute:'2-digit'}) : '—', site.syncAt ? new Date(site.syncAt).toLocaleDateString('ru-RU') : 'живая синхронизация ещё не выполнена', site.syncOk ? 'Данные синхронизированы' : 'Показана сохранённая копия'));

/* Журнал изменений. */
addWidget('activity', 'Журнал', 'Последние изменения', 'Три последние записи по участку', site => {
  const entries = widgetSiteLogs(site).slice(-3).reverse();
  return entries.length ? '<ul class="widget-list">' + entries.map(entry => '<li><b>' + wEsc(entry.action) + '</b><br>' + wEsc(entry.details).slice(0, 140) + '</li>').join('') + '</ul>' : widgetMetric('—', 'изменений пока нет');
});
addWidget('activityCount', 'Журнал', 'Всего изменений', 'Количество записей журнала по участку', site => widgetMetric(widgetSiteLogs(site).length, 'записей журнала', 'Лимит хранения: 3000 записей'));
addWidget('lastActivity', 'Журнал', 'Последнее действие', 'Что менялось в последний раз', site => {
  const entries = widgetSiteLogs(site), last = entries[entries.length - 1];
  return last ? widgetMetric(last.action, new Date(last.ts).toLocaleString('ru-RU'), last.details) : widgetMetric('—', 'изменений пока нет');
});

/* Действия не дублируют логику, а вызывают существующие кнопки приложения. */
addWidget('actions', 'Действия', 'Быстрые действия', 'Автоплан, отчёт, экспорт и переход к сегодня', () => '<div class="widget-actions"><button class="btn" data-widget-action="autoplan">Автоплан</button><button class="btn" data-widget-action="log">Отчёт</button><button class="btn" data-widget-action="export">CSV</button><button class="btn" data-widget-action="today">Сегодня</button></div>');
addWidget('navigation', 'Действия', 'Форматы плана', 'Быстрое переключение представления', () => '<div class="widget-actions"><button class="btn" data-widget-view="viewBoard">Доска</button><button class="btn" data-widget-view="viewList">Список</button><button class="btn" data-widget-view="viewShifts">Сменные</button><button class="btn" data-widget-view="viewMulti">Все участки</button></div>');

let customWidgets = wStoreGet(WIDGET_KEY, ['load', 'resource', 'actions']);
if(!Array.isArray(customWidgets)) customWidgets = ['load', 'resource', 'actions'];
customWidgets = customWidgets.filter((id, index, all) => WIDGET_CATALOG[id] && all.indexOf(id) === index);
let collapsedWidgets = wStoreGet(WIDGET_COLLAPSE_KEY, []);
if(!Array.isArray(collapsedWidgets)) collapsedWidgets = [];
let loadPeriodState = localStorage.getItem('planner.loadPeriod') || 'day';
if(!['day', 'week', 'month'].includes(loadPeriodState)) loadPeriodState = 'day';
const loadDefaults = currentLoadSelection();
let loadSelectionState = Object.assign({}, loadDefaults, wStoreGet('planner.loadSelection.v1', {}));
if(!wIsIsoDate(loadSelectionState.day)) loadSelectionState.day = loadDefaults.day;
if(!isoWeekStart(loadSelectionState.week)) loadSelectionState.week = loadDefaults.week;
if(!/^\d{4}-\d{2}$/.test(loadSelectionState.month || '')) loadSelectionState.month = loadDefaults.month;

function syncLoadControls(){
  if(!w$('loadPeriod')) return;
  w$('loadPeriod').value = loadPeriodState;
  ['day', 'week', 'month'].forEach(period => {
    const wrap = w$('load' + period.charAt(0).toUpperCase() + period.slice(1) + 'Wrap');
    if(wrap) wrap.hidden = period !== loadPeriodState;
  });
  w$('loadDay').value = loadSelectionState.day;
  w$('loadWeek').value = loadSelectionState.week;
  w$('loadMonth').value = loadSelectionState.month;
  const info = isoWeekInfo(isoWeekStart(loadSelectionState.week));
  w$('loadWeekNumber').textContent = info.week ? 'Неделя № ' + info.week + ' · ' + info.year : '';
}
function renderLoadStrip(site){
  const host = w$('siteLoadStrip');
  if(!host) return;
  host.innerHTML = ['day', 'week', 'month'].map(period => {
    const load = workloadPeriod(site, period);
    const label = period === 'week' ? 'Неделя № ' + load.weekNumber : workloadLabel(period);
    return '<div class="top-load' + (load.pct > 100 ? ' over' : '') + '" title="' + wEsc(workloadPeriodNote(load)) + '">' +
      '<span>' + label + '</span><b>' + load.pct + '%</b><small>' + wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + wEsc(load.unit) + '</small></div>';
  }).join('');
}
function renderLoadPanel(){
  const host = w$('loadWidgets');
  if(!host) return;
  syncLoadControls();
  const planner = wPlanner();
  const active = wActiveSite();
  const hero = selectedWorkload(active);
  let html = '<div class="load-hero' + (hero.pct > 100 ? ' over' : '') + '"><div class="load-hero-head"><b>' + wEsc(active.name) + '</b><strong>' + hero.pct + '%</strong></div>' +
    '<div class="load-numbers">' + wFmt(hero.load) + ' / ' + wFmt(hero.capacity) + ' ' + wEsc(hero.unit) + '</div>' + workloadBar(hero) +
    '<small>' + wEsc(workloadPeriodNote(hero)) + ' · рабочих дней: ' + hero.workdays + '</small></div>';
  html += '<div class="section-label">Все участки · один выбранный период</div><div class="load-site-list">';
  planner.state.sites.forEach(site => {
    const load = selectedWorkload(site);
    html += '<div class="load-row" data-load-site="' + wEsc(site.id) + '"><div class="load-row-head"><b>' + wEsc(site.name) + '</b><span>' + load.pct + '%</span></div>' + workloadBar(load, true) +
      '<small>' + wFmt(load.load) + ' / ' + wFmt(load.capacity) + ' ' + wEsc(load.unit) + '</small></div>';
  });
  host.innerHTML = html + '</div>';
}
function widgetBody(id, site){ return WIDGET_CATALOG[id] ? WIDGET_CATALOG[id].render(site) : ''; }
function renderWidgetCatalog(){
  const host = w$('widgetCatalogGrid');
  if(!host) return;
  const query = (w$('widgetSearch').value || '').trim().toLowerCase();
  const group = w$('widgetGroup').value || '';
  const entries = Object.entries(WIDGET_CATALOG).filter(([, widget]) =>
    (!group || widget.group === group) && (!query || (widget.title + ' ' + widget.description + ' ' + widget.group).toLowerCase().includes(query)));
  w$('widgetCatalogCount').textContent = entries.length + ' из ' + Object.keys(WIDGET_CATALOG).length;
  host.innerHTML = entries.length ? entries.map(([id, widget]) =>
    '<article class="catalog-item"><div class="catalog-item-group">' + wEsc(widget.group) + '</div><b>' + wEsc(widget.title) + '</b><p>' + wEsc(widget.description) + '</p>' +
    '<button class="btn" data-catalog-add="' + id + '" ' + (customWidgets.includes(id) || !wCan('widgets') ? 'disabled' : '') + '>' + (customWidgets.includes(id) ? 'Добавлен' : '+ Добавить') + '</button></article>'
  ).join('') : '<div class="hint">По этому запросу виджетов нет. Попробуйте другое название или категорию.</div>';
}
function renderCustomWidgets(site){
  const host = w$('customWidgets');
  if(!host) return;
  w$('btnWidgetCatalog').disabled = !wCan('widgets');
  w$('widgetCount').textContent = customWidgets.length + ' из ' + Object.keys(WIDGET_CATALOG).length;
  host.innerHTML = customWidgets.length ? customWidgets.map((id, index) => {
    const collapsed = collapsedWidgets.includes(id);
    return '<article class="widget-card" data-widget="' + id + '"><div class="widget-head"><b>' + wEsc(WIDGET_CATALOG[id].title) + '</b>' +
      '<button data-widget-collapse="' + id + '" title="' + (collapsed ? 'Развернуть' : 'Свернуть') + '">' + (collapsed ? '+' : '−') + '</button>' +
      '<button data-widget-up="' + index + '" title="Выше" ' + (index === 0 ? 'disabled' : '') + '>↑</button>' +
      '<button data-widget-down="' + index + '" title="Ниже" ' + (index === customWidgets.length - 1 ? 'disabled' : '') + '>↓</button>' +
      '<button data-widget-remove="' + id + '" title="Убрать">×</button></div><div class="widget-body" ' + (collapsed ? 'hidden' : '') + '>' + widgetBody(id, site) + '</div></article>';
  }).join('') : '<div class="hint">Откройте каталог и добавьте первый виджет. Можно вывести любой состав из 48 вариантов.</div>';
  const modal = w$('widgetCatalogModal');
  if(modal && modal.classList.contains && modal.classList.contains('open')) renderWidgetCatalog();
}
function addWidgetToPanel(id){
  if(!WIDGET_CATALOG[id] || customWidgets.includes(id) || !wCan('widgets')) return;
  customWidgets.push(id);
  wStoreSet(WIDGET_KEY, customWidgets);
  renderCustomWidgets(wActiveSite());
}

/* ---------- profiles: UI permissions, not server authentication -------- */
const PROFILE_KEY = 'planner.profiles.v1';
const ACTIVE_PROFILE_KEY = 'planner.activeProfile';
const OWNER_PROFILE = { id:'owner', name:'Владелец', role:'admin' };
const ROLE_NAMES = { admin:'Администратор · полный доступ', planner:'Планировщик · план и ресурсы', viewer:'Наблюдатель · только просмотр' };
let profiles = wStoreGet(PROFILE_KEY, [OWNER_PROFILE]);
if(!Array.isArray(profiles)) profiles = [OWNER_PROFILE];
profiles = profiles.filter(profile => profile && ROLE_NAMES[profile.role] && profile.id && profile.name);
if(!profiles.some(profile => profile.id === 'owner')) profiles.unshift(OWNER_PROFILE);
profiles = profiles.map(profile => profile.id === 'owner' ? OWNER_PROFILE : profile);
let currentProfile = OWNER_PROFILE;

function encodeProfile(profile){
  const json = JSON.stringify({ v:1, id:profile.id, name:profile.name, role:profile.role });
  return btoa(unescape(encodeURIComponent(json))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}
function decodeProfile(value){
  try{
    const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
    const data = JSON.parse(decodeURIComponent(escape(atob(normalized))));
    if(data.v !== 1 || !ROLE_NAMES[data.role] || !data.id || !data.name) return null;
    return { id:String(data.id).slice(0, 80), name:String(data.name).slice(0, 80), role:data.role };
  }catch(error){ return null; }
}
function initProfile(){
  if(typeof location === 'undefined') return;
  const params = new URLSearchParams(location.search || '');
  if(params.has('access')){
    const linked = decodeProfile(params.get('access') || '');
    currentProfile = linked || { id:'guest', name:'Гость по ссылке', role:'viewer' };
    if(linked && !profiles.some(profile => profile.id === linked.id)) profiles.push(linked);
  }else{
    const savedId = localStorage.getItem(ACTIVE_PROFILE_KEY);
    currentProfile = profiles.find(profile => profile.id === savedId) || OWNER_PROFILE;
  }
  wStoreSet(PROFILE_KEY, profiles);
  try{ localStorage.setItem(ACTIVE_PROFILE_KEY, currentProfile.id); }catch(error){}
}
function wCan(permission){
  if(currentProfile.role === 'admin') return true;
  if(currentProfile.role === 'viewer') return permission === 'export';
  return ['edit', 'resource', 'sync', 'widgets', 'export'].includes(permission);
}
function activateProfile(id){
  const profile = profiles.find(item => item.id === id);
  if(!profile) return;
  currentProfile = profile;
  try{ localStorage.setItem(ACTIVE_PROFILE_KEY, id); }catch(error){}
  if(typeof history !== 'undefined' && history.replaceState && typeof location !== 'undefined'){
    const url = new URL(location.href);
    if(id === 'owner') url.searchParams.delete('access');
    else url.searchParams.set('access', encodeProfile(profile));
    history.replaceState(null, '', url);
  }
  renderProfiles();
  renderWorkspace(wActiveSite());
  w$('profilesModal').classList.remove('open');
  if(window.PLANNER) window.PLANNER.render();
}
function profileLink(profile){
  const url = new URL(location.href);
  url.hash = '';
  url.search = '';
  url.searchParams.set('access', encodeProfile(profile));
  return url.toString();
}
async function copyProfileLink(profile){
  const link = profileLink(profile);
  try{
    await navigator.clipboard.writeText(link);
    toast('Ссылка профиля скопирована — секретов в ней нет');
  }catch(error){ prompt('Скопируйте ссылку профиля:', link); }
}
function renderProfiles(){
  document.body.dataset.role = currentProfile.role;
  w$('profileName').textContent = currentProfile.name;
  w$('profileAvatar').textContent = currentProfile.name.trim().charAt(0).toUpperCase() || 'П';
  const current = w$('profileCurrent');
  if(current) current.innerHTML = '<b>' + wEsc(currentProfile.name) + '</b><span>' + wEsc(ROLE_NAMES[currentProfile.role]) + '</span>';
  const list = w$('profilesList');
  if(!list) return;
  list.innerHTML = profiles.map(profile =>
    '<div class="profile-item"><div><b>' + wEsc(profile.name) + '<span class="profile-role">' + wEsc(profile.role) + '</span></b><small>' + wEsc(ROLE_NAMES[profile.role]) + '</small></div>' +
    '<button class="btn profile-use" data-profile-use="' + wEsc(profile.id) + '">' + (profile.id === currentProfile.id ? 'Активен' : 'Открыть') + '</button>' +
    '<button class="btn" data-profile-copy="' + wEsc(profile.id) + '" title="Скопировать ссылку">Ссылка</button></div>'
  ).join('');
  w$('btnCreateProfile').disabled = currentProfile.role !== 'admin';
}
function applyPermissions(){
  const editControls = ['btnAutoplan', 'planApply', 'splitCreate', 'btnPinBoard', 'btnKimiRequest', 'operationDateApply'];
  const adminControls = ['btnDelSite', 'btnAddSite', 'nsCreate', 'btnShare', 'btnShareCfg', 'logClear', 'btnSiteSettings', 'ssSave', 'ssAddCategory'];
  editControls.forEach(id => { if(w$(id)) w$(id).disabled = !wCan('edit'); });
  adminControls.forEach(id => { if(w$(id)) w$(id).disabled = currentProfile.role !== 'admin'; });
  if(w$('planApply') && wCan('edit') && wPlanner().pendingAutoplan) w$('planApply').disabled = wPlanner().pendingAutoplan.dependencyLate.length > 0;
  ['workersInput', 'perWorkerInput', 'fondTotalInput'].forEach(id => { if(w$(id)) w$(id).disabled = !wCan('resource'); });
  if(w$('btnRefresh')) w$('btnRefresh').disabled = !wCan('sync');
  if(w$('kimiNotes')) w$('kimiNotes').disabled = !wCan('edit');
  if(w$('kimiStrict')) w$('kimiStrict').disabled = !wCan('edit');
  if(w$('planNotes')) w$('planNotes').disabled = !wCan('edit');
  if(w$('planStrict')) w$('planStrict').disabled = !wCan('edit');
}

/* ---------- Kimi Work / native WebMCP ---------------------------------- */
let kimiContext = null;
let kimiPending = null;
let kimiSequence = 0;
let kimiWaitResolve = null;
let kimiUiState = 'idle';
function kimiToolResult(data){
  return { content:[{ type:'text', text:JSON.stringify(data) }], structuredContent:data };
}
function kimiStatus(message, state){
  const host = w$('kimiStatus');
  if(!host) return;
  host.className = 'connection-card' + (state ? ' ' + state : '');
  host.textContent = message;
}
function kimiPublicRequest(pending){
  return {
    requestId:pending.id,
    sequence:pending.sequence,
    generatedAt:pending.generatedAt,
    site:pending.site,
    instructions:pending.instructions,
    operations:pending.operations,
    baseline:pending.baseline,
    responseContract:{
      requestId:'string — повторить requestId',
      assignments:'массив {rowId, date}; вернуть ровно все ID из baseline.assignments, даты YYYY-MM-DD',
      rationale:'короткое объяснение на русском',
      warnings:'массив предупреждений о сроках и перегрузе'
    }
  };
}
function prepareKimiRequest(){
  if(!wCan('edit')) throw new Error('В этом профиле автоплан доступен только для просмотра.');
  const planner = wPlanner();
  const site = wActiveSite();
  site.notes = w$('kimiNotes').value.trim();
  site.planStrict = w$('kimiStrict').checked;
  planner.saveSites();
  const result = planner.autoplain(site);
  const operations = planner.effRows(site).map(row => ({
    rowId:row.id,
    positionKey:planner.positionKey(row, site),
    contract:row.custOrder || '',
    productionOrder:row.prodOrder || '',
    customer:row.customer || '',
    product:row.prod || '',
    quantity:row.qty || 0,
    norm:row.minutes || 0,
    unit:site.unit || 'мин',
    currentDate:row.dateEff || null,
    deadline:row.tdg || null,
    contractDate:row.contractDate || null,
    splitPart:!!row.splitPart
  }));
  const id = 'plan-' + Date.now().toString(36) + '-' + (++kimiSequence);
  const pending = {
    id, sequence:kimiSequence, generatedAt:new Date().toISOString(), siteId:site.id,
    sourceSignature:site.rows.map(row => row.id).join('|'), result,
    site:{
      id:site.id, name:site.name, unit:site.unit || 'мин', format:site.format,
      resourcePerDay:planner.fondOf(site), workdays:site.workdays,
      strictSequence:site.planStrict, today:planner.todayISO()
    },
    instructions:site.notes || 'Дополнительных пожеланий нет. Соблюдать сроки, последовательность и ресурс участка.',
    operations,
    baseline:{
      assignments:Object.entries(result.plan).map(([rowId, date]) => ({ rowId, date })),
      autoSplits:result.autoSplits.map(record => ({
        key:record.key, product:record.prod,
        parts:record.parts.map(part => ({ rowId:part.id, quantity:part.qty, norm:part.minutes, date:part.date }))
      })),
      late:result.late,
      assemblyLinks:result.assemblyLinks,
      assemblyConstraints:result.assemblyConstraints,
      unlinked:result.unlinked,
      dependencyLate:result.dependencyLate,
      explanation:'Базовый локальный расчёт. Сохранить все ID, упаковать позиции последовательно. Для каждой assemblyConstraints позиция обязана завершиться строго раньше assemblyDate; для unlinked связь не доказана и гарантия срока сборки невозможна.'
    }
  };
  kimiPending = pending;
  kimiUiState = 'waiting';
  w$('planNotes').value = site.notes;
  w$('planStrict').checked = site.planStrict;
  const payload = kimiPublicRequest(pending);
  if(kimiWaitResolve){
    const resolve = kimiWaitResolve;
    kimiWaitResolve = null;
    resolve(kimiToolResult(payload));
  }
  kimiStatus('Запрос № ' + pending.sequence + ' готов. Ожидающий сеанс Kimi получил данные; если сеанс ещё не запущен, попросите Kimi в активном чате ожидать автоплан и подготовьте запрос ещё раз.', 'pending');
  return payload;
}
function applyKimiPlan(input){
  if(!wCan('edit')) throw new Error('Профиль не имеет права изменять план.');
  if(!kimiPending) throw new Error('Нет активного запроса. Сначала нажмите «Подготовить запрос для Kimi».');
  if(!input || input.requestId !== kimiPending.id) throw new Error('requestId не совпадает с активным запросом.');
  if(!Array.isArray(input.assignments)) throw new Error('Нужен массив assignments.');
  const planner = wPlanner();
  const site = planner.state.sites.find(item => item.id === kimiPending.siteId);
  if(!site) throw new Error('Участок запроса больше не существует.');
  if(site.rows.map(row => row.id).join('|') !== kimiPending.sourceSignature) throw new Error('Операции участка изменились после запроса. Подготовьте новый запрос.');
  const expected = Object.keys(kimiPending.result.plan);
  const beforeRows = planner.effRows(site);
  const beforeDates = new Map(beforeRows.map(row => [row.id, row.dateEff || null]));
  const meta = new Map();
  beforeRows.forEach(row => meta.set(row.id, { key:planner.positionKey(row, site), deadline:row.tdg || row.contractDate || null }));
  kimiPending.result.autoSplits.forEach(record => record.parts.forEach(part => {
    meta.set(part.id, { key:record.key, deadline:record.tdg || record.contractDate || null });
  }));
  const baselineGroups = new Map();
  const nextGroups = new Map();
  expected.forEach(id => {
    const key = meta.get(id) ? meta.get(id).key : id;
    if(!baselineGroups.has(key)) baselineGroups.set(key, []);
    baselineGroups.get(key).push(kimiPending.result.plan[id]);
  });
  const seen = new Set();
  const plan = {};
  input.assignments.forEach(item => {
    if(!item || typeof item.rowId !== 'string' || seen.has(item.rowId)) throw new Error('Пустой или повторяющийся rowId.');
    if(!expected.includes(item.rowId)) throw new Error('Неизвестный rowId: ' + item.rowId);
    if(!wIsIsoDate(item.date)) throw new Error('Некорректная ISO-дата для ' + item.rowId);
    if(item.date < planner.todayISO()) throw new Error('Автоплан не должен ставить операции в прошлое: ' + item.rowId);
    const info = meta.get(item.rowId);
    const baselineDate = kimiPending.result.plan[item.rowId];
    if(info && info.deadline && item.date > info.deadline && item.date > baselineDate) throw new Error('План увеличивает просрочку для ' + item.rowId + '.');
    seen.add(item.rowId);
    plan[item.rowId] = item.date;
    const key = info ? info.key : item.rowId;
    if(!nextGroups.has(key)) nextGroups.set(key, []);
    nextGroups.get(key).push(item.date);
  });
  const missing = expected.filter(id => !seen.has(id));
  if(missing.length) throw new Error('План неполный: отсутствует операций ' + missing.length + '.');
  const span = dates => {
    const values = dates.map(date => Date.parse(date + 'T00:00:00'));
    return Math.max(...values) - Math.min(...values);
  };
  nextGroups.forEach((dates, key) => {
    const baseline = baselineGroups.get(key) || [];
    if(new Set(dates).size > new Set(baseline).size || span(dates) > span(baseline)) throw new Error('Позиция разбросана по дням сильнее базового расчёта: ' + key + '. Упакуйте её последовательно.');
  });
  const assemblyByKey = new Map((kimiPending.result.assemblyConstraints || []).map(item => [item.key, item]));
  nextGroups.forEach((dates, key) => {
    const constraint = assemblyByKey.get(key);
    if(!constraint) return;
    const end = dates.slice().sort().pop();
    if(end >= constraint.assemblyDate){
      throw new Error('Нарушена связь со сборкой для ' + (constraint.prodOrder || key) + ': заготовка ' + end + ', сборка ' + constraint.assemblyDate + '. Нужна дата не позже ' + constraint.latest + '. План не применён.');
    }
  });
  const changedCount = expected.filter(id => (beforeDates.get(id) || null) !== plan[id]).length;
  planner.applyPlanToSite(site, plan, kimiPending.result.autoSplits);
  planner.saveSites();
  const rationale = String(input.rationale || 'План проверен и возвращён Kimi.').slice(0, 3000);
  const warnings = Array.isArray(input.warnings) ? input.warnings.slice(0, 30).map(item => String(item).slice(0, 300)) : [];
  planner.logChange(site, 'автоплан · Kimi', 'изменено дат: ' + changedCount + '; операций: ' + expected.length + '; ' + rationale);
  kimiPending = null;
  kimiUiState = 'done';
  if(site.sharedOn) planner.pushShared(site).catch(() => {});
  planner.render();
  w$('kimiResult').innerHTML = '<div class="kimi-result-card"><b>План Kimi применён · ' + expected.length + ' операций</b>' + wEsc(rationale) +
    (warnings.length ? '<ul class="widget-list">' + warnings.map(item => '<li>' + wEsc(item) + '</li>').join('') + '</ul>' : '') + '</div>';
  kimiStatus('Готово. Изменения записаны в журнал участка; API-запросов и списания кредитов через эту интеграцию нет.', 'ok');
  return kimiToolResult({ ok:true, siteId:site.id, operations:expected.length, rationale, warnings });
}
function registerKimiTools(){
  const host = document.modelContext || navigator.modelContext;
  if(!host || typeof host.registerTool !== 'function'){
    kimiStatus('Канал Kimi Work недоступен в этом браузере. Откройте страницу во встроенном браузере Kimi Work; локальный автоплан продолжит работать.', 'error');
    return;
  }
  kimiContext = host;
  try{
    host.registerTool({
      name:'planner_get_autoplan_request',
      description:'Прочитать текущий явно подготовленный запрос автоплана производственного участка. Ничего не изменяет.',
      inputSchema:{ type:'object', properties:{}, additionalProperties:false },
      annotations:{ readOnlyHint:true },
      execute:async () => kimiToolResult(kimiPending ? { pending:true, ...kimiPublicRequest(kimiPending) } : { pending:false })
    });
    host.registerTool({
      name:'planner_wait_for_autoplan',
      description:'Ожидать нажатие пользователем кнопки «Подготовить запрос для Kimi». Один pending-вызов; не опрашивать страницу. Вернуть запрос с sequence больше afterSequence.',
      inputSchema:{ type:'object', properties:{ afterSequence:{ type:'integer', minimum:0, default:0 } }, additionalProperties:false },
      annotations:{ readOnlyHint:true },
      execute:async args => {
        const after = args && Number.isInteger(args.afterSequence) ? args.afterSequence : 0;
        if(kimiPending && kimiPending.sequence > after) return kimiToolResult(kimiPublicRequest(kimiPending));
        if(kimiWaitResolve) throw new Error('Уже есть ожидающий вызов planner_wait_for_autoplan.');
        return new Promise(resolve => { kimiWaitResolve = resolve; });
      }
    });
    host.registerTool({
      name:'planner_apply_autoplan',
      description:'Проверить и применить полный план Kimi. Изменяет даты участка и пишет журнал. Все rowId из baseline.assignments обязательны ровно один раз.',
      inputSchema:{
        type:'object', required:['requestId', 'assignments'], additionalProperties:false,
        properties:{
          requestId:{ type:'string' },
          assignments:{ type:'array', minItems:1, items:{ type:'object', required:['rowId', 'date'], additionalProperties:false, properties:{ rowId:{type:'string'}, date:{type:'string', pattern:'^\\d{4}-\\d{2}-\\d{2}$'} } } },
          rationale:{ type:'string' }, warnings:{ type:'array', items:{type:'string'} }
        }
      },
      annotations:{ readOnlyHint:false },
      execute:async input => applyKimiPlan(input)
    });
    kimiStatus('Канал Kimi Work доступен. Нужен активный сеанс агента, ожидающий запрос; страница сама ничего не отправляет в фоне.', 'ok');
  }catch(error){
    kimiContext = null;
    kimiStatus('Не удалось зарегистрировать канал Kimi: ' + error.message, 'error');
  }
}

/* ---------- rendering and event wiring --------------------------------- */
function renderSiteRail(){
  const host = w$('siteRailList');
  if(!host) return;
  const planner = wPlanner();
  w$('siteCount').textContent = planner.state.sites.length;
  host.innerHTML = planner.state.sites.map(site => {
    const load = workloadPeriod(site, 'day');
    const short = site.name.replace(/\s*\(live\)/gi, '').replace(/—.*$/, '').trim();
    const letters = short.split(/[\s/]+/).filter(Boolean).slice(0, 2).map(part => part.charAt(0)).join('').toUpperCase();
    return '<button class="site-rail-item' + (site.id === planner.state.activeId ? ' on' : '') + '" data-rail-site="' + wEsc(site.id) + '"><span class="site-monogram">' + wEsc(letters) + '</span><span class="site-rail-copy"><b>' + wEsc(short) + '</b><small>' + site.rows.length + ' операций · ' + wFmt(planner.fondOf(site)) + ' ' + wEsc(site.unit || 'мин') + '/день</small></span><span class="site-mini-load' + (load.pct > 100 ? ' over' : '') + '">' + load.pct + '%</span></button>';
  }).join('');
}
function syncKimiFields(site){
  const notes = w$('kimiNotes');
  if(!notes) return;
  if(notes.dataset.siteId !== site.id && document.activeElement !== notes){
    notes.value = site.notes || '';
    notes.dataset.siteId = site.id;
  }
  if(document.activeElement !== w$('kimiStrict')) w$('kimiStrict').checked = site.planStrict !== false;
}
function renderWorkspace(site){
  if(!workspaceReady || !site || !document.querySelector) return;
  const planner = wPlanner();
  const date = new Date();
  w$('headerDate').textContent = new Intl.DateTimeFormat('ru-RU', { weekday:'long', day:'numeric', month:'long' }).format(date);
  w$('stageSiteName').textContent = site.name.replace(/\s*\(live\)/gi, '');
  renderSiteRail();
  renderLoadStrip(site);
  // Скрытые панели карусели не тратят кадр на вычисления и innerHTML.
  // setDockPanel() отрисовывает выбранную панель в момент открытия.
  if(dockIndexState === 1) renderLoadPanel();
  else if(dockIndexState === 2) syncKimiFields(site);
  else if(dockIndexState === 3) renderCustomWidgets(site);
  applyPermissions();
}
function guardPermissions(event){
  const target = event.target.closest ? event.target.closest('button, [data-split], [data-unsplit], [data-unpin]') : null;
  if(!target) return;
  const id = target.id;
  let permission = null;
  if(['btnAutoplan', 'planApply', 'splitCreate', 'btnPinBoard', 'btnKimiRequest', 'operationDateApply'].includes(id) || target.hasAttribute('data-split') || target.hasAttribute('data-unsplit') || target.hasAttribute('data-unpin') || target.hasAttribute('data-date-edit')) permission = 'edit';
  if(['btnDelSite', 'btnAddSite', 'nsCreate', 'btnShare', 'btnShareCfg', 'logClear', 'btnCreateProfile', 'btnSiteSettings', 'ssSave', 'ssAddCategory'].includes(id)) permission = 'admin';
  if(permission && (permission === 'admin' ? currentProfile.role !== 'admin' : !wCan(permission))){
    event.preventDefault();
    event.stopImmediatePropagation();
    alert('Профиль «' + currentProfile.name + '» не имеет права на это действие. Роли на этой статической странице ограничивают интерфейс.');
  }
}
/* ---------- resizable dock ---------------------------------------------- */
const DOCK_WIDTH_KEY = 'planner.dockWidth';
let dockWidthState = Number(localStorage.getItem(DOCK_WIDTH_KEY)) || 304;
function dockWidthLimits(){
  const viewport = window.innerWidth || 1280;
  return { min:280, max:Math.max(280, Math.min(620, viewport - 420)) };
}
function setDockWidth(value, persist){
  const limits = dockWidthLimits();
  dockWidthState = Math.round(Math.max(limits.min, Math.min(limits.max, Number(value) || 304)));
  document.documentElement.style.setProperty('--dock-width', dockWidthState + 'px');
  const handle = w$('dockResizeHandle');
  if(handle){
    handle.setAttribute('aria-valuemin', limits.min);
    handle.setAttribute('aria-valuemax', limits.max);
    handle.setAttribute('aria-valuenow', dockWidthState);
  }
  if(persist){ try{ localStorage.setItem(DOCK_WIDTH_KEY, String(dockWidthState)); }catch(error){} }
}
function wireDockResize(){
  setDockWidth(dockWidthState, false);
  w$('dockNarrow').addEventListener('click', () => setDockWidth(dockWidthState - 32, true));
  w$('dockWide').addEventListener('click', () => setDockWidth(dockWidthState + 32, true));
  const handle = w$('dockResizeHandle');
  handle.addEventListener('dblclick', () => setDockWidth(304, true));
  handle.addEventListener('keydown', event => {
    if(!['ArrowLeft', 'ArrowRight', 'Home'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    if(event.key === 'Home') setDockWidth(304, true);
    else setDockWidth(dockWidthState + (event.key === 'ArrowLeft' ? 24 : -24), true);
  });
  let drag = null, frame = 0;
  handle.addEventListener('pointerdown', event => {
    if(event.button !== 0) return;
    event.preventDefault();
    drag = { x:event.clientX, width:dockWidthState, latest:event.clientX, pointerId:event.pointerId };
    handle.setPointerCapture(event.pointerId);
    document.body.classList.add('dock-resizing');
  });
  handle.addEventListener('pointermove', event => {
    if(!drag) return;
    drag.latest = event.clientX;
    if(frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      if(drag) setDockWidth(drag.width + drag.x - drag.latest, false);
    });
  });
  const finish = event => {
    if(!drag) return;
    if(event.clientX != null) setDockWidth(drag.width + drag.x - event.clientX, false);
    setDockWidth(dockWidthState, true);
    if(handle.hasPointerCapture && handle.hasPointerCapture(drag.pointerId)) handle.releasePointerCapture(drag.pointerId);
    drag = null;
    if(frame){ cancelAnimationFrame(frame); frame = 0; }
    document.body.classList.remove('dock-resizing');
  };
  handle.addEventListener('pointerup', finish);
  handle.addEventListener('pointercancel', finish);
  window.addEventListener('resize', () => setDockWidth(dockWidthState, false), { passive:true });
}
function wireWorkspace(){
  w$('dockPrev').addEventListener('click', () => setDockPanel(dockIndexState - 1));
  w$('dockNext').addEventListener('click', () => setDockPanel(dockIndexState + 1));
  w$('dockDots').innerHTML = DOCK_PANELS.map((panel, index) => '<button class="dock-dot" data-dock="' + index + '" title="' + wEsc(panel[0]) + '" aria-label="' + wEsc(panel[0]) + '"></button>').join('');
  w$('dockDots').addEventListener('click', event => {
    const dot = event.target.closest('[data-dock]');
    if(dot) setDockPanel(Number(dot.dataset.dock));
  });
  w$('controlDock').addEventListener('keydown', event => {
    if(event.key === 'ArrowLeft') setDockPanel(dockIndexState - 1);
    if(event.key === 'ArrowRight') setDockPanel(dockIndexState + 1);
  });
  let touchX = null;
  w$('controlDock').addEventListener('touchstart', event => { touchX = event.touches[0].clientX; }, { passive:true });
  w$('controlDock').addEventListener('touchend', event => {
    if(touchX == null) return;
    const delta = event.changedTouches[0].clientX - touchX;
    if(Math.abs(delta) > 60) setDockPanel(dockIndexState + (delta < 0 ? 1 : -1));
    touchX = null;
  }, { passive:true });
  w$('btnDock').addEventListener('click', () => document.body.classList.toggle('dock-open'));
  wireDockResize();
  w$('loadPeriod').addEventListener('change', event => {
    loadPeriodState = event.target.value;
    try{ localStorage.setItem('planner.loadPeriod', loadPeriodState); }catch(error){}
    renderWorkspace(wActiveSite());
  });
  const changeLoadSelection = (period, value) => {
    loadSelectionState[period] = value;
    wStoreSet('planner.loadSelection.v1', loadSelectionState);
    renderWorkspace(wActiveSite());
  };
  w$('loadDay').addEventListener('change', event => { if(wIsIsoDate(event.target.value)) changeLoadSelection('day', event.target.value); });
  w$('loadWeek').addEventListener('change', event => { if(isoWeekStart(event.target.value)) changeLoadSelection('week', event.target.value); });
  w$('loadMonth').addEventListener('change', event => { if(/^\d{4}-\d{2}$/.test(event.target.value)) changeLoadSelection('month', event.target.value); });
  w$('loadCurrent').addEventListener('click', () => {
    loadSelectionState = currentLoadSelection();
    wStoreSet('planner.loadSelection.v1', loadSelectionState);
    renderWorkspace(wActiveSite());
  });
  w$('widgetGroup').innerHTML = '<option value="">Все категории</option>' + WIDGET_GROUPS.map(group => '<option value="' + wEsc(group) + '">' + wEsc(group) + '</option>').join('');
  w$('btnWidgetCatalog').addEventListener('click', () => {
    renderWidgetCatalog();
    w$('widgetCatalogModal').classList.add('open');
  });
  w$('widgetCatalogClose').addEventListener('click', () => w$('widgetCatalogModal').classList.remove('open'));
  w$('widgetCatalogModal').addEventListener('click', event => { if(event.target === w$('widgetCatalogModal')) w$('widgetCatalogModal').classList.remove('open'); });
  w$('widgetSearch').addEventListener('input', renderWidgetCatalog);
  w$('widgetGroup').addEventListener('change', renderWidgetCatalog);
  w$('widgetCatalogGrid').addEventListener('click', event => {
    const button = event.target.closest('[data-catalog-add]');
    if(button) addWidgetToPanel(button.dataset.catalogAdd);
  });
  w$('customWidgets').addEventListener('click', event => {
    const target = event.target.closest('button');
    if(!target) return;
    if(target.dataset.widgetAction){
      const map = { autoplan:'btnAutoplan', log:'btnLog', export:'btnExport', today:'btnToday' };
      w$(map[target.dataset.widgetAction]).click();
      return;
    }
    if(target.dataset.widgetView){ w$(target.dataset.widgetView).click(); return; }
    if(!wCan('widgets')) return;
    if(target.dataset.widgetRemove){
      customWidgets = customWidgets.filter(id => id !== target.dataset.widgetRemove);
      collapsedWidgets = collapsedWidgets.filter(id => id !== target.dataset.widgetRemove);
      wStoreSet(WIDGET_COLLAPSE_KEY, collapsedWidgets);
    }else if(target.dataset.widgetCollapse){
      const id = target.dataset.widgetCollapse;
      collapsedWidgets = collapsedWidgets.includes(id) ? collapsedWidgets.filter(item => item !== id) : collapsedWidgets.concat(id);
      wStoreSet(WIDGET_COLLAPSE_KEY, collapsedWidgets);
    }else if(target.dataset.widgetUp !== undefined){
      const index = Number(target.dataset.widgetUp);
      if(index > 0) [customWidgets[index - 1], customWidgets[index]] = [customWidgets[index], customWidgets[index - 1]];
    }else if(target.dataset.widgetDown !== undefined){
      const index = Number(target.dataset.widgetDown);
      if(index < customWidgets.length - 1) [customWidgets[index + 1], customWidgets[index]] = [customWidgets[index], customWidgets[index + 1]];
    }else return;
    wStoreSet(WIDGET_KEY, customWidgets);
    renderCustomWidgets(wActiveSite());
  });
  w$('btnKimiRequest').addEventListener('click', () => {
    try{ prepareKimiRequest(); }catch(error){ kimiStatus(error.message, 'error'); }
  });
  w$('kimiNotes').addEventListener('change', () => {
    const site = wActiveSite();
    site.notes = w$('kimiNotes').value;
    wPlanner().saveSites();
  });
  w$('kimiStrict').addEventListener('change', () => {
    const site = wActiveSite();
    site.planStrict = w$('kimiStrict').checked;
    wPlanner().saveSites();
  });
  w$('btnProfiles').addEventListener('click', () => { renderProfiles(); w$('profilesModal').classList.add('open'); });
  w$('profilesClose').addEventListener('click', () => w$('profilesModal').classList.remove('open'));
  w$('btnCreateProfile').addEventListener('click', () => {
    if(currentProfile.role !== 'admin') return;
    const name = w$('newProfileName').value.trim();
    const role = w$('newProfileRole').value;
    if(!name || !ROLE_NAMES[role]) return;
    const id = 'p-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
    profiles.push({ id, name:name.slice(0, 80), role });
    wStoreSet(PROFILE_KEY, profiles);
    w$('newProfileName').value = '';
    renderProfiles();
  });
  w$('profilesList').addEventListener('click', event => {
    const use = event.target.closest('[data-profile-use]');
    const copy = event.target.closest('[data-profile-copy]');
    if(use) activateProfile(use.dataset.profileUse);
    if(copy){
      const profile = profiles.find(item => item.id === copy.dataset.profileCopy);
      if(profile) copyProfileLink(profile);
    }
  });
  w$('siteRailList').addEventListener('click', event => {
    const item = event.target.closest('[data-rail-site]');
    if(!item) return;
    w$('siteSelect').value = item.dataset.railSite;
    w$('siteSelect').dispatchEvent(new Event('change'));
  });
  w$('loadWidgets').addEventListener('click', event => {
    const item = event.target.closest('[data-load-site]');
    if(!item) return;
    w$('siteSelect').value = item.dataset.loadSite;
    w$('siteSelect').dispatchEvent(new Event('change'));
  });
  document.addEventListener('click', guardPermissions, true);
  document.addEventListener('dragstart', event => {
    if(!wCan('edit') && event.target.closest('[draggable="true"][data-id]')){
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  }, true);
  /* FLIP: одна анимация только для фактически перенесённой карточки. */
  let flip = null;
  document.addEventListener('drop', event => {
    const activeDragId = typeof dragId !== 'undefined' ? dragId : '';
    const source = activeDragId ? document.querySelector('[data-id="' + activeDragId + '"]') : null;
    if(source && source.getBoundingClientRect){
      const rect = source.getBoundingClientRect();
      flip = { id:source.dataset.id, x:rect.left, y:rect.top };
    }
  }, true);
  document.addEventListener('drop', () => {
    if(!flip || !document.querySelector) return;
    const moved = flip;
    flip = null;
    requestAnimationFrame(() => {
      const target = document.querySelector('[data-id="' + moved.id + '"]');
      if(!target || !target.animate || !target.getBoundingClientRect) return;
      const rect = target.getBoundingClientRect();
      target.animate([
        { transform:'translate(' + (moved.x - rect.left) + 'px,' + (moved.y - rect.top) + 'px)', opacity:.75 },
        { transform:'translate(0,0)', opacity:1 }
      ], { duration:300, easing:'cubic-bezier(.16,1,.3,1)' });
    });
  });
}
function initWorkspace(){
  if(initWorkspace._done || !document.querySelector || !window.PLANNER) return;
  initWorkspace._done = true;
  initProfile();
  wireWorkspace();
  workspaceReady = true;
  setDockPanel(dockIndexState);
  renderProfiles();
  renderWorkspace(wActiveSite());
  registerKimiTools();
}
if(typeof window !== 'undefined'){
  window.PLANNER_WORKSPACE = {
    workloadPeriod, selectedWorkload, isoWeekInfo, isoWeekStart, setDockPanel, setDockWidth,
    prepareKimiRequest, applyKimiPlan, renderWorkspace, initWorkspace, wCan,
    widgetBody, addWidgetToPanel, catalog:WIDGET_CATALOG,
    get profiles(){ return profiles; }, get currentProfile(){ return currentProfile; },
    get pendingKimiRequest(){ return kimiPending; },
    get customWidgets(){ return customWidgets.slice(); },
    get loadPeriod(){ return loadPeriodState; },
    get loadSelection(){ return {...loadSelectionState}; }
  };
  if(typeof document !== 'undefined' && document.addEventListener){
    document.addEventListener('DOMContentLoaded', initWorkspace);
    if(document.readyState !== 'loading') initWorkspace();
  }
}


'use strict';
/* ==========================================================================
   ПОТОК / календари
   Единый popover для date, week и month. Сами input остаются на месте:
   приложение получает обычные value, input и change, а браузерный нативный
   picker больше не выпадает поверх фирменного интерфейса.
   ========================================================================== */
(() => {
  const DAY = 86400000;
  const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const WEEKDAYS = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];
  let panel = null;
  let activeInput = null;
  let cursor = null;
  let closeTimer = 0;
  let openFrame = 0;
  let positionFrame = 0;

  function iso(date){ return date.toISOString().slice(0, 10); }
  function fromISO(value){ return new Date(value + 'T00:00:00Z'); }
  function validDate(value){
    if(!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return false;
    const date = fromISO(value);
    return !Number.isNaN(date.getTime()) && iso(date) === value;
  }
  function addDays(value, amount){ const date = fromISO(value); date.setUTCDate(date.getUTCDate() + amount); return iso(date); }
  function today(){
    if(window.PLANNER && typeof window.PLANNER.todayISO === 'function') return window.PLANNER.todayISO();
    const now = new Date();
    return iso(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())));
  }
  function weekStart(value){
    const date = fromISO(value);
    const day = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 1 - day);
    return iso(date);
  }
  function weekKey(value){
    const date = fromISO(value);
    const day = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - day);
    const year = date.getUTCFullYear();
    const first = new Date(Date.UTC(year, 0, 1));
    const week = Math.ceil(((date - first) / DAY + 1) / 7);
    return year + '-W' + String(week).padStart(2, '0');
  }
  function weekFromKey(key){
    const match = String(key || '').match(/^(\d{4})-W(\d{2})$/);
    if(!match) return null;
    const date = new Date(Date.UTC(Number(match[1]), 0, 4));
    const day = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 1 - day + (Number(match[2]) - 1) * 7);
    return weekKey(iso(date)) === key ? iso(date) : null;
  }
  function validMonth(value){
    const match = String(value || '').match(/^(\d{4})-(\d{2})$/);
    return !!match && Number(match[2]) >= 1 && Number(match[2]) <= 12;
  }
  function currentValue(mode){
    const date = today();
    return mode === 'week' ? weekKey(date) : mode === 'month' ? date.slice(0, 7) : date;
  }
  function initialCursor(input){
    if(input.type === 'week') return weekFromKey(input.value) || today();
    if(input.type === 'month' && validMonth(input.value)) return input.value + '-01';
    return validDate(input.value) ? input.value : today();
  }
  function monthTitle(){
    const month = MONTHS[cursor.getUTCMonth()];
    return month + ' ' + cursor.getUTCFullYear() + (activeInput.type === 'month' ? '' : '');
  }
  function dateGrid(){
    const first = new Date(Date.UTC(cursor.getUTCFullYear(), cursor.getUTCMonth(), 1));
    const start = weekStart(iso(first));
    const selected = activeInput.value;
    const current = today();
    let html = '<div class="pc-weekdays" aria-hidden="true">' + WEEKDAYS.map(day => '<span>' + day + '</span>').join('') + '</div><div class="pc-date-grid">';
    for(let index = 0; index < 42; index++){
      const value = addDays(start, index);
      const date = fromISO(value);
      const outside = date.getUTCMonth() !== cursor.getUTCMonth();
      const classes = ['pc-day'];
      if(outside) classes.push('outside');
      if(value === current) classes.push('today');
      if(value === selected) classes.push('selected');
      const label = new Intl.DateTimeFormat('ru-RU', { day:'numeric', month:'long', year:'numeric', weekday:'long', timeZone:'UTC' }).format(date);
      html += '<button type="button" class="' + classes.join(' ') + '" data-date="' + value + '" aria-label="' + label + '" aria-pressed="' + (value === selected) + '"' + (outside ? ' tabindex="-1"' : '') + '>' + date.getUTCDate() + '</button>';
    }
    return html + '</div>';
  }
  function weekGrid(){
    const first = new Date(Date.UTC(cursor.getUTCFullYear(), cursor.getUTCMonth(), 1));
    const start = weekStart(iso(first));
    const current = today();
    const selected = activeInput.value;
    let html = '<div class="pc-weekdays pc-week-head" aria-hidden="true"><span>№</span>' + WEEKDAYS.map(day => '<span>' + day + '</span>').join('') + '</div><div class="pc-week-list">';
    for(let row = 0; row < 6; row++){
      const monday = addDays(start, row * 7);
      const key = weekKey(monday);
      const number = Number(key.slice(-2));
      html += '<button type="button" class="pc-week-row' + (key === selected ? ' selected' : '') + '" data-week="' + key + '" aria-label="Неделя № ' + number + ', с ' + monday + '" aria-pressed="' + (key === selected) + '"><span class="pc-week-number">' + number + '</span>';
      for(let column = 0; column < 7; column++){
        const value = addDays(monday, column);
        const date = fromISO(value);
        const classes = [];
        if(date.getUTCMonth() !== cursor.getUTCMonth()) classes.push('outside');
        if(value === current) classes.push('today');
        html += '<span class="' + classes.join(' ') + '">' + date.getUTCDate() + '</span>';
      }
      html += '</button>';
    }
    return html + '</div>';
  }
  function monthGrid(){
    const year = cursor.getUTCFullYear();
    return '<div class="pc-month-grid">' + MONTHS.map((month, index) => {
      const value = year + '-' + String(index + 1).padStart(2, '0');
      const classes = ['pc-month'];
      if(value === today().slice(0, 7)) classes.push('today');
      if(value === activeInput.value) classes.push('selected');
      return '<button type="button" class="' + classes.join(' ') + '" data-month="' + value + '" aria-pressed="' + (value === activeInput.value) + '">' + month + '</button>';
    }).join('') + '</div>';
  }
  function position(){
    positionFrame = 0;
    if(!activeInput || !panel || panel.hidden) return;
    const rect = activeInput.getBoundingClientRect();
    if(!rect.width && !rect.height){ closeCalendar(false); return; }
    const margin = 8;
    const width = Math.min(318, window.innerWidth - margin * 2);
    const left = Math.max(margin, Math.min(window.innerWidth - width - margin, rect.left));
    panel.style.width = width + 'px';
    const height = panel.offsetHeight || 360;
    const below = window.innerHeight - rect.bottom - margin;
    const above = rect.top - margin;
    const top = below >= height + 6 || below >= above
      ? Math.min(window.innerHeight - height - margin, rect.bottom + 6)
      : Math.max(margin, rect.top - height - 6);
    panel.style.left = Math.round(left) + 'px';
    panel.style.top = Math.round(top) + 'px';
    panel.dataset.placement = top >= rect.bottom ? 'bottom' : 'top';
  }
  function schedulePosition(){
    if(positionFrame || !activeInput) return;
    positionFrame = requestAnimationFrame(position);
  }
  function render(){
    if(!activeInput || !panel) return;
    const mode = activeInput.type;
    const title = mode === 'month' ? String(cursor.getUTCFullYear()) : monthTitle();
    const currentLabel = mode === 'week' ? 'Текущая неделя' : mode === 'month' ? 'Текущий месяц' : 'Сегодня';
    const hint = mode === 'week' ? 'Выберите всю календарную неделю' : mode === 'month' ? 'Выберите месяц' : 'Выберите дату';
    panel.innerHTML = '<div class="pc-accent"></div><header class="pc-header"><button type="button" class="pc-nav" data-nav="-1" aria-label="Назад">‹</button><div class="pc-heading"><strong>' + title + '</strong><span>' + hint + '</span></div><button type="button" class="pc-nav" data-nav="1" aria-label="Вперёд">›</button></header><div class="pc-body">' + (mode === 'week' ? weekGrid() : mode === 'month' ? monthGrid() : dateGrid()) + '</div><footer class="pc-footer"><span>' + (mode === 'week' ? 'Пн — Вс · ISO' : 'PlanFlow · период') + '</span><button type="button" data-current>' + currentLabel + '</button></footer>';
    panel.setAttribute('aria-label', mode === 'week' ? 'Выбор календарной недели' : mode === 'month' ? 'Выбор месяца' : 'Выбор даты');
    position();
  }
  function closeCalendar(restoreFocus){
    clearTimeout(closeTimer);
    if(openFrame){ cancelAnimationFrame(openFrame); openFrame = 0; }
    if(positionFrame){ cancelAnimationFrame(positionFrame); positionFrame = 0; }
    const input = activeInput;
    activeInput = null;
    if(input) input.setAttribute('aria-expanded', 'false');
    if(!panel) return;
    panel.classList.remove('open');
    closeTimer = setTimeout(() => { if(!activeInput) panel.hidden = true; }, 180);
    if(restoreFocus !== false && input && input.isConnected) input.focus({ preventScroll:true });
  }
  function openCalendar(target, focusGrid){
    const input = typeof target === 'string' ? document.getElementById(target) : target;
    if(!input || !['date', 'week', 'month'].includes(input.type) || !panel) return;
    clearTimeout(closeTimer);
    if(activeInput && activeInput !== input) activeInput.setAttribute('aria-expanded', 'false');
    activeInput = input;
    cursor = fromISO(initialCursor(input));
    if(input.type === 'month') cursor.setUTCDate(1);
    input.setAttribute('aria-expanded', 'true');
    panel.hidden = false;
    panel.classList.remove('open');
    render();
    openFrame = requestAnimationFrame(() => {
      openFrame = 0;
      if(!activeInput) return;
      panel.classList.add('open');
      if(focusGrid){
        const selected = panel.querySelector('.selected') || panel.querySelector('[data-date], [data-week], [data-month]');
        if(selected) selected.focus({ preventScroll:true });
      }
    });
  }
  function moveCursor(direction){
    if(!activeInput) return;
    // С 31-го числа обычный setUTCMonth(+1) мог бы перескочить месяц.
    cursor.setUTCDate(1);
    if(activeInput.type === 'month') cursor.setUTCFullYear(cursor.getUTCFullYear() + direction);
    else cursor.setUTCMonth(cursor.getUTCMonth() + direction);
    render();
    const selected = panel.querySelector('.selected');
    if(selected) selected.focus({ preventScroll:true });
    else panel.querySelector('[data-nav="-1"]').focus({ preventScroll:true });
  }
  function choose(value){
    const input = activeInput;
    if(!input) return;
    input.value = value;
    closeCalendar();
    input.dispatchEvent(new Event('input', { bubbles:true }));
    input.dispatchEvent(new Event('change', { bubbles:true }));
  }
  function onPanelClick(event){
    const button = event.target.closest('button');
    if(!button || !panel.contains(button)) return;
    if(button.dataset.nav) moveCursor(Number(button.dataset.nav));
    else if(button.dataset.date) choose(button.dataset.date);
    else if(button.dataset.week) choose(button.dataset.week);
    else if(button.dataset.month) choose(button.dataset.month);
    else if(button.hasAttribute('data-current')) choose(currentValue(activeInput.type));
  }
  function onGridKeys(event){
    if(!activeInput || !panel || panel.hidden) return;
    if(event.key === 'Escape'){
      event.preventDefault(); event.stopPropagation(); closeCalendar(); return;
    }
    if(!panel.contains(event.target)) return;
    if(event.target.dataset.nav && ['ArrowLeft', 'ArrowRight'].includes(event.key)){
      event.preventDefault(); moveCursor(event.key === 'ArrowLeft' ? -1 : 1); return;
    }
    const buttons = [...panel.querySelectorAll('[data-date], [data-week], [data-month]')];
    const index = buttons.indexOf(event.target.closest('button'));
    if(index < 0) return;
    const columns = activeInput.type === 'month' ? 3 : activeInput.type === 'week' ? 1 : 7;
    const moves = { ArrowLeft:-1, ArrowRight:1, ArrowUp:-columns, ArrowDown:columns };
    if(!(event.key in moves)) return;
    event.preventDefault();
    const next = Math.max(0, Math.min(buttons.length - 1, index + moves[event.key]));
    buttons[next].focus({ preventScroll:true });
  }
  function enhance(input){
    if(input.dataset.potokCalendar) return;
    input.dataset.potokCalendar = 'true';
    input.classList.add('pl-calendar-input');
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('aria-haspopup', 'dialog');
    input.setAttribute('aria-expanded', 'false');
    input.setAttribute('aria-controls', 'potokCalendar');
    input.addEventListener('click', event => { event.preventDefault(); openCalendar(input, false); });
    input.addEventListener('keydown', event => {
      if(['ArrowDown', 'Enter', ' '].includes(event.key)){
        event.preventDefault(); openCalendar(input, true);
      }else if(event.key === 'Escape' && activeInput === input){
        event.preventDefault(); closeCalendar(false);
      }
    });
  }
  function initCalendars(){
    if(initCalendars.done || !document.body || !document.querySelectorAll) return;
    initCalendars.done = true;
    panel = document.createElement('section');
    panel.id = 'potokCalendar';
    panel.className = 'potok-calendar';
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-modal', 'false');
    panel.hidden = true;
    document.body.appendChild(panel);
    panel.addEventListener('click', onPanelClick);
    document.querySelectorAll('input[type="date"], input[type="week"], input[type="month"]').forEach(enhance);
    document.addEventListener('pointerdown', event => {
      if(activeInput && event.target !== activeInput && !panel.contains(event.target)) closeCalendar(false);
    }, true);
    document.addEventListener('keydown', onGridKeys, true);
    window.addEventListener('resize', schedulePosition, { passive:true });
    window.addEventListener('scroll', schedulePosition, { passive:true, capture:true });
  }

  window.PLANNER_CALENDAR = {
    open:openCalendar, close:closeCalendar, init:initCalendars,
    isoWeek:weekKey, isoWeekStart:weekFromKey,
    get input(){ return activeInput; },
    get enhancedCount(){ return document.querySelectorAll ? document.querySelectorAll('[data-potok-calendar]').length : 0; }
  };
  if(typeof document !== 'undefined' && document.body && document.addEventListener){
    if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initCalendars);
    else initCalendars();
  }
})();

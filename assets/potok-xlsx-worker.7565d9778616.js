self.window = self;
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

  const XLSX_WORKER_URL = '__XLSX_WORKER_ASSET__';
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

/* ZIP, XML и преобразование ячеек в операции выполняются вне главного потока. */
self.onmessage = async event => {
  const { id, buffer, format } = event.data || {};
  try{
    const book = self.POTOK_XLSX_PROCESS(await self.PLANNER_XLSX.readWorkbook(buffer), format);
    self.postMessage({ id, book });
  }catch(error){
    self.postMessage({ id, error:String(error && error.message || error) });
  }finally{
    self.close();
  }
};

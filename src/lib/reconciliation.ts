import Papa from 'papaparse';
import ExcelJS from 'exceljs';

export function normalizeGrz(grz: string): string {
  if (!grz) return "";
  let grzStr = String(grz).toUpperCase();
  let grzClean = grzStr.replace(/[^A-Z0-9А-Я]/g, '');

  const transliterationMap: Record<string, string> = {
    'А': 'A', 'В': 'B', 'Е': 'E', 'К': 'K', 'М': 'M', 'Н': 'H',
    'О': 'O', 'Р': 'P', 'С': 'C', 'Т': 'T', 'У': 'Y', 'Х': 'X'
  };

  return grzClean.split('').map(char => transliterationMap[char] || char).join('');
}

export interface PrilRow {
  date_str: string;
  route: string;
  start_time_str: string;
  grz_raw: string;
  grz_norm: string;
  actual_work_km: number;
  direction: string;
  start_datetime: Date | null;
}

export interface TransactionRow {
  DATE: string;
  TIME: string;
  VREG_NUM: string;
  ROUTE_NUM: string;
  TRIP_NO: string;
  CR_TIME: string;
  IN_NAME: string;
  CONDUCTOR: string;
  tran_datetime: Date | null;
  vreg_norm: string;
}

export interface KrcRow {
  route: string;
  conductor: string;
  time: string;
  datetime: Date | null;
}

export interface ReconciliationResult {
  date: string;
  route: string;
  startTime: string;
  grz: string;
  status: string;
  tripNo: string;
  mileage: number;
  direction: string;
  transCount: number;
  conductor: string;
  openTimes: string;
  closeTimes: string;
  krcStatus: string;
}

export interface ReconciliationMetadata {
  route: string;
  month: string;
  year: string;
}

export interface ReconciliationResponse {
  results: ReconciliationResult[];
  stats: { confirmed: number; unconfirmed: number; krcChecks: number; totalMileage: number };
  metadata: ReconciliationMetadata;
}

export async function parseCsvLocal(content: string, separator: string = ';'): Promise<any[]> {
  return new Promise((resolve) => {
    // First pass: parse without headers to find the actual header row
    const firstPass = Papa.parse(content, {
      delimiter: separator,
      header: false,
      skipEmptyLines: true,
    });

    const rows = firstPass.data as string[][];
    if (!rows || rows.length === 0) {
      resolve([]);
      return;
    }

    // Find the row that likely contains headers by checking for keywords
    const keywords = ['дата', 'маршрут', 'грз', 'госномер', 'время', 'рейс', 'date', 'vreg', 'time', 'route', 'сутки', 'кондуктор', 'провод'];
    let headerIndex = 0;
    let maxMatches = 0;

    // Check first 20 rows
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const row = rows[i];
      if (!Array.isArray(row)) continue;
      
      const matches = row.filter(cell => 
        cell && keywords.some(kw => String(cell).toLowerCase().includes(kw))
      ).length;
      
      if (matches > maxMatches) {
        maxMatches = matches;
        headerIndex = i;
      }
    }

    // If we found a header row, use it
    const headers = rows[headerIndex];
    const dataRows = rows.slice(headerIndex + 1);
    
    const result = dataRows.map(row => {
      const obj: any = {};
      headers.forEach((h, idx) => {
        if (h !== undefined && h !== null) {
          obj[String(h).trim()] = row[idx];
        }
      });
      return obj;
    });

    console.log(`Parsed CSV with ${result.length} rows. Header found at index ${headerIndex}.`);
    resolve(result);
  });
}

async function readFileAsText(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  
  // Try UTF-8 first
  const utf8Decoder = new TextDecoder('utf-8');
  const text = utf8Decoder.decode(buffer);
  
  // Check for common Russian keywords (more exhaustive list)
  const cyrillic = /[а-яёА-ЯЁ]/;
  const commonKeywords = ['дата', 'маршрут', 'грз', 'рейс', 'время', 'кондуктор', 'фио', 'проверка'];
  const hasCyrillic = cyrillic.test(text);
  const foundKeywords = commonKeywords.some(h => text.toLowerCase().includes(h));
  
  if (!hasCyrillic || (!foundKeywords && text.includes('\uFFFD'))) {
    try {
      // Try Windows-1251 if UTF-8 doesn't seem right
      const win1251Decoder = new TextDecoder('windows-1251');
      const winText = win1251Decoder.decode(buffer);
      if (cyrillic.test(winText)) {
        return winText;
      }
    } catch (e) {
      // fallback
    }
  }
  return text;
}

export function detectSeparator(text: string): string {
  const sample = text.slice(0, 10000);
  const semiCount = (sample.match(/;/g) || []).length;
  const commaCount = (sample.match(/,/g) || []).length;
  const tabCount = (sample.match(/\t/g) || []).length;
  
  if (tabCount > semiCount && tabCount > commaCount) return '\t';
  return semiCount >= commaCount ? ';' : ',';
}

export async function reconcileFiles(
  prilFile: File, 
  transFile: File, 
  tripDurationMinutes: number = 120,
  krcFile?: File | null
): Promise<ReconciliationResponse> {
  const findKey = (obj: any, target: string | string[], excludes: string[] = []) => {
    if (!obj) return undefined;
    const keys = Object.keys(obj);
    const targets = Array.isArray(target) ? target.map(t => t.toLowerCase().trim()) : [target.toLowerCase().trim()];
    const exList = excludes.map(ex => ex.toLowerCase().trim());
    
    // 1. Strict equality (case-insensitive)
    let found = keys.find(k => {
      const kl = k.toLowerCase().trim();
      return targets.includes(kl) && !exList.some(ex => kl.includes(ex));
    });
    if (found) return found;

    // 2. Word-based match (case-insensitive)
    found = keys.find(k => {
      const keyLow = k.toLowerCase();
      if (exList.some(ex => keyLow.includes(ex))) return false;
      return targets.some(t => {
        const regex = new RegExp(`(^|[^a-zа-яё0-9])${t}($|[^a-zа-яё0-9])`, 'i');
        return regex.test(keyLow);
      });
    });
    if (found) return found;

    // 3. Soft inclusion (priority to shortest key to avoid over-matching partials)
    const candidates = keys.filter(k => {
      const keyLow = k.toLowerCase();
      if (exList.some(ex => keyLow.includes(ex))) return false;
      return targets.some(t => keyLow.includes(t));
    }).sort((a, b) => a.length - b.length);
    
    return candidates[0];
  };

  const parseDate = (dStr: any, tStr: any = ''): Date | null => {
    if (!dStr) return null;
    
    let y: number, m: number, d: number;
    let h = 0, min = 0, sec = 0;

    // 1. Extract Date Components
    if (dStr instanceof Date) {
      y = dStr.getFullYear();
      m = dStr.getMonth();
      d = dStr.getDate();
    } else if (typeof dStr === 'number') {
      const date = new Date(Math.round((dStr - 25569) * 86400 * 1000));
      y = date.getFullYear();
      m = date.getMonth();
      d = date.getDate();
    } else {
      let cleanD = String(dStr).trim();
      // Handle "01/03/2026 10:12 - 01/03/2026 20:03" or "01.03.2026 06:00-18:00"
      if (cleanD.includes('-')) {
        const parts = cleanD.split('-');
        cleanD = parts[0].trim();
      }
      
      // Remove any trailing time if we just want the date part, e.g. "01.03.2026 19:24" -> "01.03.2026"

      const nums = cleanD.split(/[^0-9]+/).filter(Boolean).map(Number);
      if (nums.length >= 3) {
        // Date part
        if (nums[0] > 1000) { 
          y = nums[0]; m = nums[1] - 1; d = nums[2]; 
        } else if (nums[2] > 1000) { 
          d = nums[0]; m = nums[1] - 1; y = nums[2]; 
        } else {
          // YY
          d = nums[0]; m = nums[1] - 1; y = 2000 + nums[2];
        }
        
        // If no separate tStr, try to get time from the dStr itself
        if (!tStr && nums.length >= 5) {
          h = nums[3]; min = nums[4]; sec = nums[5] || 0;
        }
      } else {
        const nd = new Date(cleanD.replace(/\./g, '-'));
        if (isNaN(nd.getTime())) return null;
        y = nd.getFullYear(); m = nd.getMonth(); d = nd.getDate();
      }
    }

    // 2. Extract Time Components if tStr provided
    if (tStr) {
      if (tStr instanceof Date) {
        h = tStr.getHours();
        min = tStr.getMinutes();
        sec = tStr.getSeconds();
      } else if (typeof tStr === 'number') {
        const totalSec = Math.round(tStr * 86400);
        h = Math.floor(totalSec / 3600);
        min = Math.floor((totalSec % 3600) / 60);
        sec = totalSec % 60;
      } else {
        const tNums = String(tStr).split(/[^0-9]+/).filter(Boolean).map(Number);
        if (tNums.length >= 2) {
          // If time string starts with a date like "01.03.2026 19:24:10", 
          // we need to identify which indices are H and M.
          // Usually time is at the end.
          if (tNums.length >= 5) {
             // Look for HH:MM(:SS) at the end
             // If we have 5 parts (D, M, Y, H, Min)
             if (tNums.length === 5) {
                h = tNums[3];
                min = tNums[4];
             } else {
                // If 6 parts (D, M, Y, H, Min, Sec)
                h = tNums[3];
                min = tNums[4];
                sec = tNums[5];
             }
          } else {
            h = tNums[0]; 
            min = tNums[1]; 
            sec = tNums[2] || 0;
          }
        }
      }
    }

    if (y! === undefined || m! === undefined || d! === undefined) return null;
    const finalDate = new Date(y!, m!, d!, h, min, sec);
    return isNaN(finalDate.getTime()) ? null : finalDate;
  };

  const conductorNamesMatch = (nameA: string, nameB: string): boolean => {
    if (!nameA || !nameB) return false;
    
    const normalize = (n: string) => n.toLowerCase()
      .replace(/ё/g, 'е')
      .replace(/^(кондуктор|водитель|кассир|контролер)\s+/gi, '')
      .replace(/[^a-zа-я\s]/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    const normA = normalize(nameA);
    const normB = normalize(nameB);
    if (!normA || !normB) return false;
    
    // Strict match
    if (normA === normB || normA.includes(normB) || normB.includes(normA)) return true;

    const partsA = normA.split(/\s+/).filter(p => p.length >= 2);
    const partsB = normB.split(/\s+/).filter(p => p.length >= 2);
    
    if (partsA.length === 0 || partsB.length === 0) {
       // If one side has only initials/one word, do a simple inclusion check
       return normA.includes(normB) || normB.includes(normA);
    }

    // Check if the primary name (usually the longest or first part) exists in both
    // Order-agnostic check: "Farentseva O.V." vs "O.V. Farentseva"
    const hasSharedSignificantPart = partsA.some(pa => partsB.some(pb => pa === pb && pa.length >= 4));
    if (!hasSharedSignificantPart) return false;

    // initials check
    const isInitial = (s: string) => s.length === 1 || (s.length === 2 && s.endsWith('.'));
    const getInitials = (s: string) => s.split(/\s+/).filter(isInitial).map(i => i.replace('.', ''));
    
    const iA = getInitials(normA);
    const iB = getInitials(normB);
    
    if (iA.length > 0 && iB.length > 0) {
      return iA.some(a => iB.includes(a));
    }
    
    return true; 
  };

  const normalizeName = (name: string): string => {
    if (!name) return "";
    return name.toLowerCase()
      .replace(/ё/g, 'е') // Normalize yo/e
      .replace(/[^a-zа-я\s]/gi, '') // Keep only letters and spaces
      .replace(/\s+/g, ' ')
      .trim();
  };

  const normalizeRoute = (r: string): string => {
    if (!r) return '';
    // Remove "маршрут", "№", "route", "номер" and dots/spaces
    return String(r)
      .replace(/Маршрут\s*№?/gi, '')
      .replace(/маршрут/gi, '')
      .replace(/route/gi, '')
      .replace(/№/g, '')
      .replace(/[^0-9a-zа-я]/gi, '')
      .trim()
      .toLowerCase();
  };

  const prilText = await readFileAsText(prilFile);
  const transText = await readFileAsText(transFile);
  
  let krcData: KrcRow[] = [];
  if (krcFile) {
    const isExcel = krcFile.name.toLowerCase().endsWith('.xlsx');
    let krcRowsRaw: any[] = [];

    if (isExcel) {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await krcFile.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.getWorksheet(1);
      
      if (worksheet) {
        let excelHeaders: string[] = [];
        let headerRowIndex = 1;
        
        // Find header row by keywords
        const keywords = ['дата', 'маршрут', 'грз', 'время', 'рейс', 'кондуктор', 'провод', 'марш', 'ффио'];
        for (let i = 1; i <= Math.min(worksheet.rowCount, 20); i++) {
          const row = worksheet.getRow(i);
          let matchCount = 0;
          row.eachCell({ includeEmpty: false }, (cell) => {
            const val = String(cell.value || '').toLowerCase();
            if (keywords.some(kw => val.includes(kw))) matchCount++;
          });
          if (matchCount > 2) {
            headerRowIndex = i;
            break;
          }
        }
        
        const hRow = worksheet.getRow(headerRowIndex);
        hRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          excelHeaders[colNumber] = String(cell.value || '').trim();
        });
        
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber <= headerRowIndex) return;
          const obj: any = {};
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const h = excelHeaders[colNumber];
            if (h) obj[h] = cell.value;
          });
          krcRowsRaw.push(obj);
        });
      }
    } else {
      const krcText = await readFileAsText(krcFile);
      const krcSep = detectSeparator(krcText);
      krcRowsRaw = await parseCsvLocal(krcText, krcSep);
    }
    
    let lastValidDateStr = "";
    for (const row of krcRowsRaw) {
      const kRoute = findKey(row, ['№ марш.', '№ маршрута', 'Route', 'маршрут', 'марш.']);
      const kConductor = findKey(row, ['ФИО кондуктора', 'Conductor', 'фио', 'наименование', 'фио водителя', 'водитель']);
      // Specifically look for the check time, avoiding the work interval
      const kTime = findKey(row, ['Время', 'CR_TIME', 'время', 'чч:мм']);
      // Date can be in 'Дата' or extracted from 'Время работы'
      const kDate = findKey(row, ['Дата', 'Date', 'дата', 'Время работы', 'Период']);

      const route = kRoute ? String(row[kRoute] || '').trim() : '';
      const conductor = kConductor ? String(row[kConductor] || '').trim() : '';
      const timeVal = kTime ? row[kTime] : '';
      const dateVal = kDate ? row[kDate] : '';
      
      const dateStr = String(dateVal || '').trim();
      if (dateStr && dateStr.length > 5) lastValidDateStr = dateStr;
      
      // Use the specific check time for the datetime
      const datetime = parseDate(dateVal || lastValidDateStr, timeVal);
      
      // Format time string for display in the status
      let displayTime = '';
      if (timeVal instanceof Date) {
        displayTime = timeVal.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
      } else if (typeof timeVal === 'number') {
        const totalSec = Math.round(timeVal * 86400);
        const hh = Math.floor(totalSec / 3600);
        const mm = Math.floor((totalSec % 3600) / 60);
        displayTime = `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
      } else {
        const tS = String(timeVal || '').trim();
        const tMatch = tS.match(/(\d{1,2}:\d{1,2})/);
        displayTime = tMatch ? tMatch[1] : tS;
      }

      if (route || conductor) {
        krcData.push({ route, conductor, time: displayTime, datetime });
      }
    }
  }

  const prilSep = detectSeparator(prilText);
  const transSep = detectSeparator(transText);

  const prilRowsRaw = await parseCsvLocal(prilText, prilSep);
  const transRowsRaw = await parseCsvLocal(transText, transSep);

  const prilData: PrilRow[] = [];
  const transData: TransactionRow[] = [];

  let detectedRoute = "";
  let detectedMonth = "";
  let detectedYear = "";

  const monthNames = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
  ];

  for (const row of prilRowsRaw) {
    const kDate = findKey(row, ['Транспортные сутки', 'Дата', 'Date', 'Day']);
    const kRoute = findKey(row, ['№ маршрута', 'Маршрут', 'Route', 'Route Num', 'Маршрут №']);
    const kTime = findKey(row, ['Фактическое время начала рейса', 'Время начала', 'Start Time', 'Время']);
    const kGrz = findKey(row, ['ГРЗ', 'Госномер', 'VREG_NUM', 'Vehicle', 'Гос. номер']);
    const kWork = findKey(row, ['Фактическая транспортная работа', 'Пробег', 'Mileage', 'Работа', 'км']);
    const kDirection = findKey(row, ['Направление', 'Direction', 'Код направления', 'Прям/Обр']);

    const dateStr = (kDate ? row[kDate] : '').trim();
    const route = (kRoute ? row[kRoute] : '').trim();
    const startTimeStr = (kTime ? row[kTime] : '').trim();
    const grzRaw = (kGrz ? row[kGrz] : '').trim();
    const direction = (kDirection ? row[kDirection] : '').trim();
    let actualWorkKm = parseFloat(String((kWork ? row[kWork] : '0')).replace(',', '.'));
    if (isNaN(actualWorkKm)) actualWorkKm = 0;

    if (!detectedRoute && route) detectedRoute = route;

    const dateParts = dateStr.split(/[\/\.\-]/);
    let startDatetime: Date | null = null;
    if (dateParts.length === 3) {
      let year, month, day;
      if (dateParts[0].length === 4) {
        year = parseInt(dateParts[0]);
        month = parseInt(dateParts[1]) - 1;
        day = parseInt(dateParts[2]);
      } else if (dateParts[2].length === 4) {
        year = parseInt(dateParts[2]);
        month = parseInt(dateParts[1]) - 1;
        day = parseInt(dateParts[0]);
      } else {
        // Assume DD.MM.YY or YY.MM.DD
        const p0 = parseInt(dateParts[0]);
        const p2 = parseInt(dateParts[2]);
        if (p0 > 31) {
          year = 2000 + p0;
          month = parseInt(dateParts[1]) - 1;
          day = p2;
        } else {
          year = 2000 + p2;
          month = parseInt(dateParts[1]) - 1;
          day = p0;
        }
      }
      const date = new Date(year, month, day);
      
      if (!isNaN(date.getTime())) {
        if (!detectedMonth) detectedMonth = monthNames[month];
        if (!detectedYear) detectedYear = year.toString();

        if (startTimeStr) {
          // Handle HH:MM:SS or HH:MM
          const timeParts = startTimeStr.split(/[:\-\s]/);
          if (timeParts.length >= 2) {
            const h = parseInt(timeParts[0]);
            const m = parseInt(timeParts[1]);
            const s = parseInt(timeParts[2] || '0');
            if (!isNaN(h) && !isNaN(m)) {
              date.setHours(h, m, s);
              startDatetime = date;
            }
          }
        }
      }
    }

    if (dateStr || route || grzRaw) {
      prilData.push({
        date_str: dateStr,
        route: route,
        start_time_str: startTimeStr,
        grz_raw: grzRaw,
        grz_norm: normalizeGrz(grzRaw),
        actual_work_km: actualWorkKm,
        direction: direction,
        start_datetime: startDatetime
      });
    }
  }

  for (const row of transRowsRaw) {
    const kDate = findKey(row, ['DATE', 'Дата', 'Date', 'Транспортные сутки']);
    const kTime = findKey(row, ['TIME', 'Время', 'Time', 'CR_TIME']);
    const kVreg = findKey(row, ['VREG_NUM', 'ГРЗ', 'Госномер', 'Vehicle', 'Гос. номер']);
    const kRoute = findKey(row, ['ROUTE_NUM', 'Маршрут', 'Route']);
    const kTrip = findKey(row, ['TRIP_NO', 'Рейс', 'Trip']);
    const kCrTime = findKey(row, ['CR_TIME', 'Время закрытия', 'Close Time']);
    const kInName = findKey(row, ['IN_NAME', 'Остановка', 'Stop Name']);
    const kConductor = findKey(row, ['CONDUCTOR', 'Кондуктор', 'ФИО кондуктора']);

    const date = (kDate ? row[kDate] : '').trim();
    const time = (kTime ? row[kTime] : '').trim();
    const vregNum = (kVreg ? row[kVreg] : '').trim();
    const inName = (kInName ? row[kInName] : '').trim();
    const conductor = (kConductor ? row[kConductor] : '').trim();
    
    let tranDatetime: Date | null = null;
    if (date && time) {
      const dateParts = date.split(/[\/\.\-]/);
      if (dateParts.length === 3) {
        let year, month, day;
        if (dateParts[0].length === 4) {
          year = parseInt(dateParts[0]);
          month = parseInt(dateParts[1]) - 1;
          day = parseInt(dateParts[2]);
        } else if (dateParts[2].length === 4) {
          year = parseInt(dateParts[2]);
          month = parseInt(dateParts[1]) - 1;
          day = parseInt(dateParts[0]);
        } else {
          const p0 = parseInt(dateParts[0]);
          const p2 = parseInt(dateParts[2]);
          if (p0 > 31) {
            year = 2000 + p0;
            month = parseInt(dateParts[1]) - 1;
            day = p2;
          } else {
            year = 2000 + p2;
            month = parseInt(dateParts[1]) - 1;
            day = p0;
          }
        }
        const dateObj = new Date(year, month, day);
        
        const timeParts = time.split(/[:\-\s]/);
        if (timeParts.length >= 2) {
          const h = parseInt(timeParts[0]);
          const m = parseInt(timeParts[1]);
          const s = parseInt(timeParts[2] || '0');
          if (!isNaN(h) && !isNaN(m)) {
            dateObj.setHours(h, m, s);
            tranDatetime = dateObj;
          }
        }
      }
    }

    if (date || time || vregNum) {
      transData.push({
        DATE: date,
        TIME: time,
        VREG_NUM: vregNum,
        ROUTE_NUM: (kRoute ? row[kRoute] : '').trim(),
        TRIP_NO: (kTrip ? row[kTrip] : '').trim(),
        CR_TIME: (kCrTime ? row[kCrTime] : '').trim(),
        IN_NAME: inName,
        CONDUCTOR: conductor,
        tran_datetime: tranDatetime,
        vreg_norm: normalizeGrz(vregNum)
      });
    }
  }

  const results: ReconciliationResult[] = [];
  let confirmedCount = 0;
  let unconfirmedCount = 0;
  let krcCheckCount = 0;
  let totalConfirmedMileage = 0;

  for (const flight of prilData) {
    const potentialTrans = transData.filter(t => {
      if (!t.tran_datetime || !flight.start_datetime) return false;
      
      // Match by date (ignoring time)
      const tDate = t.tran_datetime;
      const fDate = flight.start_datetime;
      
      const dateMatch = tDate.getFullYear() === fDate.getFullYear() &&
                        tDate.getMonth() === fDate.getMonth() &&
                        tDate.getDate() === fDate.getDate();
      
      if (!dateMatch) return false;

      // Normalize for comparison
      const tNorm = t.vreg_norm;
      const fNorm = flight.grz_norm;

      // Match by GRZ: exact or one contains another
      const grzMatch = tNorm === fNorm || 
                       (tNorm.length >= 3 && fNorm.includes(tNorm)) ||
                       (fNorm.length >= 3 && tNorm.includes(fNorm));
      
      if (!grzMatch) return false;

      return true;
    });

    let selectedTripNo = "";
    let confirmedTransactions: TransactionRow[] = [];

    if (flight.start_datetime) {
      // Window: from exactly start time to X minutes after
      // User requested that transactions must be later than start time
      const startTimeWindow = flight.start_datetime;
      const endTimeWindow = new Date(flight.start_datetime.getTime() + tripDurationMinutes * 60000);
      
      const inWindow = potentialTrans.filter(t => {
        return t.tran_datetime && t.tran_datetime >= startTimeWindow && t.tran_datetime <= endTimeWindow;
      }).sort((a, b) => {
        // Prefer transactions closer to start time
        const diffA = (a.tran_datetime?.getTime() || 0) - flight.start_datetime!.getTime();
        const diffB = (b.tran_datetime?.getTime() || 0) - flight.start_datetime!.getTime();
        return diffA - diffB;
      });

      if (inWindow.length > 0) {
        selectedTripNo = inWindow[0].TRIP_NO;
        // Only include transactions that are within the trip and strictly within the specified duration window
        confirmedTransactions = potentialTrans.filter(t => 
          t.TRIP_NO === selectedTripNo && 
          t.tran_datetime && 
          t.tran_datetime >= startTimeWindow && 
          t.tran_datetime <= endTimeWindow
        );
      }
    }

    const isConfirmed = selectedTripNo !== "";
    if (isConfirmed) {
      confirmedCount++;
      totalConfirmedMileage += flight.actual_work_km;
    } else unconfirmedCount++;

    let finalDirection = flight.direction;
    if (confirmedTransactions.length > 0) {
      const sampleWithDirection = confirmedTransactions.find(t => t.IN_NAME.includes('_A_') || t.IN_NAME.includes('_B_'));
      if (sampleWithDirection) {
        if (sampleWithDirection.IN_NAME.includes('_A_')) {
          finalDirection = "Прямое";
        } else if (sampleWithDirection.IN_NAME.includes('_B_')) {
          finalDirection = "Обратное";
        }
      } else if (confirmedTransactions[0].IN_NAME) {
         // Fallback check on first transaction even if no _A_ or _B_ tag found via include
         if (confirmedTransactions[0].IN_NAME.includes('_A_')) finalDirection = "Прямое";
         else if (confirmedTransactions[0].IN_NAME.includes('_B_')) finalDirection = "Обратное";
      }
    }
    
    const conductors = confirmedTransactions.length > 0 
      ? Array.from(new Set(confirmedTransactions.map(t => t.CONDUCTOR).filter(Boolean)))
      : [];

    // KRC Matching Logic
    let krcStatus = krcFile ? "Проверка не проводилась" : "";
    if (krcFile && confirmedTransactions.length > 0) {
      const flightRoute = normalizeRoute(flight.route);
      const transactionRoutes = Array.from(new Set(confirmedTransactions.map(t => normalizeRoute(t.ROUTE_NUM)).filter(Boolean)));
      
      const tranStart = confirmedTransactions[0].tran_datetime;
      const tranEnd = confirmedTransactions[confirmedTransactions.length - 1].tran_datetime;
      const scheduledStart = parseDate(flight.date_str, flight.start_time_str);
      
      if (tranStart && tranEnd) {
        const matchingKrc = krcData.find(k => {
          // 1. Route Match
          const kRoute = normalizeRoute(k.route);
          if (kRoute) {
            const kRouteDigits = kRoute.replace(/[^0-9]/g, '');
            const fRouteDigits = flightRoute ? flightRoute.replace(/[^0-9]/g, '') : '';
            
            const isRouteMatch = 
              (flightRoute && (kRoute === flightRoute || (kRouteDigits && kRouteDigits === fRouteDigits))) ||
              transactionRoutes.some(tr => kRoute === tr || (kRouteDigits && kRouteDigits === tr.replace(/[^0-9]/g, '')));
              
            if (!isRouteMatch) return false;
          }

          // 2. Conductor/Driver Match
          const conductorMatch = conductors.some(c => conductorNamesMatch(k.conductor, c));
          if (!conductorMatch) return false;

          // 3. Date & Time matching
          if (k.datetime) {
            const isSameDayStart = k.datetime.getFullYear() === tranStart.getFullYear() &&
                                  k.datetime.getMonth() === tranStart.getMonth() &&
                                  k.datetime.getDate() === tranStart.getDate();
            
            const isSameDayEnd = k.datetime.getFullYear() === tranEnd.getFullYear() &&
                                k.datetime.getMonth() === tranEnd.getMonth() &&
                                k.datetime.getDate() === tranEnd.getDate();

            if (!isSameDayStart && !isSameDayEnd) return false;

            const checkTime = k.datetime.getTime();
            const startT = tranStart.getTime();
            const endT = tranEnd.getTime();
            
            // "попадает в диапазон транзакций"
            // With a small buffer of 15 mins to be safe
            const buffer = 15 * 60000; 

            if (checkTime >= (startT - buffer) && checkTime <= (endT + buffer)) {
              console.log(`Matched KRC: ${k.conductor} on route ${k.route} at ${k.datetime.toLocaleString()}. Range: ${tranStart.toLocaleTimeString()} - ${tranEnd.toLocaleTimeString()}`);
              return true;
            }
          }
          return false;
        });

        if (matchingKrc) {
          krcStatus = `Проверка проводилась, время ${matchingKrc.time}`;
          krcCheckCount++;
        }
      }
    }

    results.push({
      date: flight.date_str,
      route: flight.route,
      startTime: flight.start_time_str,
      grz: flight.grz_raw,
      status: isConfirmed ? 'Подтверждено' : 'Не подтверждено',
      tripNo: selectedTripNo,
      mileage: isConfirmed ? flight.actual_work_km : 0,
      direction: finalDirection,
      transCount: confirmedTransactions.length,
      conductor: conductors.length > 0 ? (() => {
        const full = conductors[0];
        const parts = full.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
          const surname = parts[0].charAt(0).toUpperCase() + parts[0].slice(1).toLowerCase();
          const initials = parts.slice(1).map(p => p.charAt(0).toUpperCase() + '.').join('');
          return `${surname} ${initials}`;
        }
        return full;
      })() : "",
      openTimes: confirmedTransactions.map(t => t.TIME).join('; '),
      closeTimes: confirmedTransactions.map(t => t.CR_TIME).join('; '),
      krcStatus: krcStatus
    });
  }

  return { 
    results, 
    stats: { 
      confirmed: confirmedCount, 
      unconfirmed: unconfirmedCount,
      krcChecks: krcCheckCount,
      totalMileage: totalConfirmedMileage
    },
    metadata: {
      route: detectedRoute || "неизвестно",
      month: detectedMonth || "неизвестно",
      year: detectedYear || ""
    }
  };
}

export async function generateExcel(results: ReconciliationResult[]): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Отчет');

  worksheet.columns = [
    { header: 'Дата', key: 'date', width: 15 },
    { header: 'Маршрут', key: 'route', width: 15 },
    { header: 'Время начала (Отчет)', key: 'startTime', width: 20 },
    { header: 'ГРЗ', key: 'grz', width: 15 },
    { header: 'ФИО водителя', key: 'conductor', width: 25 },
    { header: 'Статус подтверждения', key: 'status', width: 20 },
    { header: 'Номер рейса', key: 'tripNo', width: 15 },
    { header: 'Фактическая транспортная работа (км)', key: 'mileage', width: 25 },
    { header: 'Направление', key: 'direction', width: 20 },
    { header: 'Кол-во транзакций', key: 'transCount', width: 15 },
    { header: 'Проверка КРС', key: 'krcStatus', width: 20 },
    { header: 'Время открытия транзакции', key: 'openTimes', width: 40 },
    { header: 'Время закрытия транзакции', key: 'closeTimes', width: 40 }
  ];

  const filteredResults = results.filter(res => res.date && res.grz && res.startTime);
  filteredResults.forEach(res => worksheet.addRow(res));

  const totalMileage = filteredResults.reduce((sum, r) => sum + r.mileage, 0);
  const totalTrans = filteredResults.reduce((sum, r) => sum + r.transCount, 0);
  
  if (filteredResults.length > 0) {
    worksheet.addRow({
      date: 'Итого',
      mileage: totalMileage,
      transCount: totalTrans
    });
  }

  // Apply styling to all cells: wrap text and vertical alignment
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.alignment = { 
        wrapText: true, 
        vertical: 'middle',
        horizontal: rowNumber === 1 ? 'center' : 'left'
      };
      
      // Add borders for better readability
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  });

  // Auto-fit columns with a maximum width to encourage wrapping
  worksheet.columns.forEach(column => {
    let maxLength = 0;
    column.eachCell!({ includeEmpty: true }, cell => {
      const columnLength = cell.value ? String(cell.value).length : 0;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    
    // Limit max width to 50 characters to force wrapping for very long strings
    const calculatedWidth = maxLength + 2;
    column.width = calculatedWidth > 50 ? 50 : (calculatedWidth < 12 ? 12 : calculatedWidth);
  });

  // Style header specifically
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
  };
  headerRow.height = 35;

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

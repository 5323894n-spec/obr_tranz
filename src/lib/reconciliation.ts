import Papa from 'papaparse';
import ExcelJS from 'exceljs';

/**
 * Normalizes common homoglyphs between Cyrillic and Latin alphabets.
 * This ensures "Фаренцева" with Latin 'A' matches "Фаренцева" with Cyrillic 'А'.
 */
export function normalizeHomoglyphs(text: string): string {
  if (!text) return "";
  const map: Record<string, string> = {
    'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н', 'K': 'К', 
    'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т', 'X': 'Х', 'y': 'у', 'Y': 'У',
    'a': 'а', 'b': 'в', 'c': 'с', 'e': 'е', 'h': 'н', 'k': 'к',
    'm': 'м', 'o': 'о', 'p': 'р', 't': 'т', 'x': 'х'
  };
  return text.split('').map(char => map[char] || char).join('');
}

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
  conductor: string;
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
  plannedMileage?: number;
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
    const keywords = ['дата', 'маршрут', 'грз', 'госномер', 'время', 'рейс', 'date', 'vreg', 'time', 'route', 'сутки', 'кондуктор', 'провод', 'фио', 'номер', 'период', 'автобус', 'водитель', 'наименование'];
    let headerIndex = 0;
    let maxMatches = 0;

    // Check first 30 rows for better coverage in complex files
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
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

export async function parseKrcFile(krcFile: File): Promise<KrcRow[]> {
  const isExcel = krcFile.name.toLowerCase().endsWith('.xlsx');
  let krcRowsRaw: any[] = [];
  const krcData: KrcRow[] = [];
  const keywords = ['дата', 'маршрут', 'грз', 'время', 'рейс', 'кондуктор', 'провод', 'марш', 'ффио', 'период', 'автобус', 'номер', 'водитель', 'проверка', 'контрол', 'тс', 'гос', 'бортовой', 'экспедитор', 'фио', 'fio', 'имя', 'фамилия', 'сотрудник', 'объект', 'время работы'];

  if (isExcel) {
    const workbook = new ExcelJS.Workbook();
    const arrayBuffer = await krcFile.arrayBuffer();
    await workbook.xlsx.load(arrayBuffer);
    
    workbook.eachSheet((worksheet) => {
      let excelHeaders: string[] = [];
      let headerRowIndex = 1;
      
      // Attempt 1: Look for rows with at least 2 keywords
      for (let i = 1; i <= Math.min(worksheet.rowCount, 100); i++) {
        const row = worksheet.getRow(i);
        let matchCount = 0;
        const seenKeywords = new Set<string>();
        row.eachCell({ includeEmpty: false }, (cell) => {
          const val = String(cell.value || '').toLowerCase();
          keywords.forEach(kw => {
            if (val.includes(kw) && !seenKeywords.has(kw)) {
              matchCount++;
              seenKeywords.add(kw);
            }
          });
        });
        if (matchCount >= 2) {
          headerRowIndex = i;
          break;
        }
      }

      // Attempt 2: If failed, take the first row that has at least 1 keyword and multiple columns
      if (headerRowIndex === 1) {
        for (let i = 1; i <= Math.min(worksheet.rowCount, 50); i++) {
          const row = worksheet.getRow(i);
          let cellCount = 0;
          row.eachCell({ includeEmpty: false }, () => cellCount++);
          if (cellCount >= 3 && row.values && Array.isArray(row.values) && row.values.some(v => keywords.some(kw => String(v || '').toLowerCase().includes(kw)))) {
            headerRowIndex = i;
            break;
          }
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
    });
  } else {
    const buffer = await krcFile.arrayBuffer();
    
    // Try multiple encodings for Russian CSVs
    const encodings = ['utf-8', 'windows-1251', 'iso-8859-5', 'utf-16le'];
    let text = '';
    let krcRowsRawCsv: any[] = [];

    for (const encoding of encodings) {
      try {
        const decoder = new TextDecoder(encoding);
        const decodedText = decoder.decode(buffer);
        
        const detectSeparator = (t: string): string => {
          const sample = t.slice(0, 10000);
          const semiCount = (sample.match(/;/g) || []).length;
          const commaCount = (sample.match(/,/g) || []).length;
          const tabCount = (sample.match(/\t/g) || []).length;
          if (tabCount > semiCount && tabCount > commaCount) return '\t';
          return semiCount >= commaCount ? ';' : ',';
        };

        const sep = detectSeparator(decodedText);
        const firstPass = Papa.parse(decodedText, { delimiter: sep, header: false, skipEmptyLines: true });
        const rows = firstPass.data as any[][];
        
        let headerIdx = -1;
        // Attempt 1: Strict match (2+ unique keywords)
        for (let i = 0; i < Math.min(rows.length, 100); i++) {
          let matches = 0;
          const seenKw = new Set<string>();
          rows[i].forEach(cell => {
            const val = String(cell || '').toLowerCase();
            keywords.forEach(kw => {
              if (val.includes(kw) && !seenKw.has(kw)) {
                matches++;
                seenKw.add(kw);
              }
            });
          });
          if (matches >= 2) {
            headerIdx = i;
            break;
          }
        }

        // Attempt 2: Loose match (1 keyword + multiple cells)
        if (headerIdx === -1) {
          for (let i = 0; i < Math.min(rows.length, 50); i++) {
            const row = rows[i];
            if (row.length >= 3 && row.some(cell => keywords.some(kw => String(cell || '').toLowerCase().includes(kw)))) {
              headerIdx = i;
              break;
            }
          }
        }

        if (headerIdx !== -1) {
          const headerRow = rows[headerIdx].map(h => String(h || '').trim());
          const dataRows = rows.slice(headerIdx + 1);
          krcRowsRawCsv = dataRows.map(row => {
            const obj: any = {};
            headerRow.forEach((h, idx) => {
              if (h) obj[h] = row[idx];
            });
            return obj;
          });
          text = decodedText;
          krcRowsRaw = krcRowsRawCsv;
          break; // Found valid data with this encoding
        }
      } catch (e) {
        continue;
      }
    }
  }

  const findKey = (obj: any, target: string[]) => {
    if (!obj) return undefined;
    const keys = Object.keys(obj);
    const targets = target.map(t => t.toLowerCase().trim());
    
    // 1. Exact or strict match
    let found = keys.find(k => {
      const kl = k.toLowerCase().trim();
      return targets.includes(kl);
    });
    if (found) return found;

    // 2. Inclusion with priority to shortest key to avoid over-matching
    const candidates = keys.filter(k => {
      const kl = k.toLowerCase().trim();
      return targets.some(t => kl.includes(t));
    }).sort((a, b) => a.length - b.length);
    
    return candidates[0];
  };

  const parseDate = (dStr: any, tStr: any = ''): Date | null => {
    if (!dStr) return null;
    let y: number, m: number, d: number;
    let h = 0, min = 0, sec = 0;

    if (dStr instanceof Date) {
      y = dStr.getFullYear(); m = dStr.getMonth(); d = dStr.getDate();
    } else if (typeof dStr === 'number') {
      const date = new Date(Math.round((dStr - 25569) * 86400 * 1000));
      y = date.getFullYear(); m = date.getMonth(); d = date.getDate();
    } else {
      let cleanD = String(dStr).trim();
      // Handle "01/03/2026 08:18 - 01/03/2026 09:15" with flexible spacing
      const rangeParts = cleanD.split(/\s*-\s*/);
      const firstPart = rangeParts[0];
      
      const nums = firstPart.split(/[^0-9]+/).filter(Boolean).map(Number);
      if (nums.length >= 3) {
        if (nums[0] > 1000) { y = nums[0]; m = nums[1]-1; d = nums[2]; }
        else if (nums[2] > 1000) { d = nums[0]; m = nums[1]-1; y = nums[2]; }
        else { d = nums[0]; m = nums[1]-1; y = 2000 + nums[2]; }
        // If time is embedded in the date string (like "01.03.2026 14:30")
        if (!tStr && nums.length >= 5) { h = nums[3]; min = nums[4]; sec = nums[5] || 0; }
      } else return null;
    }

    if (tStr) {
      if (tStr instanceof Date) { h = tStr.getHours(); min = tStr.getMinutes(); sec = tStr.getSeconds(); }
      else if (typeof tStr === 'number') {
        const ts = Math.round(tStr * 86400);
        h = Math.floor(ts / 3600); min = Math.floor((ts % 3600) / 60); sec = ts % 60;
      } else {
        const tNums = String(tStr).split(/[^0-9]+/).filter(Boolean).map(Number);
        if (tNums.length >= 2) {
          if (tNums.length >= 5) { h = tNums[3]; min = tNums[4]; sec = tNums[5] || 0; }
          else { h = tNums[0]; min = tNums[1]; sec = tNums[2] || 0; }
        }
      }
    }
    const resValue = new Date(y!, m!, d!, h, min, sec);
    return isNaN(resValue.getTime()) ? null : resValue;
  };

  let lastValidDateSource: any = null;
  let lastValidRoute: string = '';
  let lastValidConductor: string = '';

  for (const row of krcRowsRaw) {
    const kRoute = findKey(row, ['№ марш.', '№ маршрута', 'Route', 'маршрут', 'марш.', 'маршруты', '№марш.', 'маршрут №', 'номер маршрута', 'тс']);
    const kConductor = findKey(row, ['ФИО кондуктора', 'Conductor', 'фио', 'наименование', 'фио водителя', 'водитель', 'DRIVER', 'ф.и.о.', 'кондуктор', 'фио контролера', 'контролер', 'экспедитор', 'сотрудник', 'ФИО водителя']);
    const kTime = findKey(row, ['Время', 'CR_TIME', 'время', 'чч:мм', 'check time', 'время пров.', 'время_пров', 'время проверки', 'время начала проверки', 'время события']);
    const kDate = findKey(row, ['Дата', 'Date', 'дата', 'Время работы', 'Период', 'смена', 'сутки', 'транспортные сутки', 'дата проверки', 'дата события']);

    const routeRaw = kRoute ? String(row[kRoute] || '').trim() : '';
    const conductorRaw = kConductor ? String(row[kConductor] || '').trim() : '';
    const timeVal = kTime ? row[kTime] : '';
    const dateVal = kDate ? row[kDate] : '';
    
    if (routeRaw) lastValidRoute = routeRaw;
    if (conductorRaw) lastValidConductor = conductorRaw;
    
    if (dateVal !== undefined && dateVal !== null && dateVal !== '') {
      lastValidDateSource = dateVal;
    }
    
    const datetime = parseDate(dateVal || lastValidDateSource, timeVal);
    if (!datetime) continue;
    
    let displayTime = '';
    if (timeVal instanceof Date) displayTime = timeVal.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
    else if (typeof timeVal === 'number') {
      const ts = Math.round(timeVal * 86400);
      displayTime = `${String(Math.floor(ts/3600)).padStart(2,'0')}:${String(Math.floor((ts%3600)/60)).padStart(2,'0')}`;
    } else {
      const tS = String(timeVal || '').trim();
      const tMatch = tS.match(/(\d{1,2}:\d{1,2})/);
      displayTime = tMatch ? tMatch[1] : tS;
    }

    if (lastValidRoute || lastValidConductor) {
      krcData.push({ 
        route: lastValidRoute, 
        conductor: lastValidConductor, 
        time: displayTime, 
        datetime 
      });
    }
  }
  return krcData;
}

export async function reconcileFiles(
  prilFile: File, 
  transFile: File, 
  tripDurationMinutes: number = 120,
  krcFile?: File | null,
  forwardMileage: number = 0,
  returnMileage: number = 0
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
    
    const normalize = (n: string) => normalizeHomoglyphs(n.toLowerCase())
      .replace(/ё/g, 'е')
      .replace(/^(кондуктор|водитель|кассир|контролер)\s+/gi, '')
      .replace(/[^a-zа-яё\s]/gi, ' ')
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
       return normA.includes(normB) || normB.includes(normA);
    }

    const hasSharedSignificantPart = partsA.some(pa => partsB.some(pb => pa === pb && pa.length >= 3));
    if (!hasSharedSignificantPart) return false;

    const isInitial = (s: string) => s.length === 1;
    const getInitials = (s: string) => s.split(/\s+/).filter(isInitial);
    
    const iA = getInitials(normA);
    const iB = getInitials(normB);
    
    // If both have initials, they must share at least one
    if (iA.length > 0 && iB.length > 0) {
      const match = iA.some(a => iB.includes(a));
      if (!match) return false;
    }
    
    return true; 
  };

  const normalizeName = (name: string): string => {
    if (!name) return "";
    return normalizeHomoglyphs(name.toLowerCase())
      .replace(/ё/g, 'е')
      .replace(/[^a-zа-яё\s]/gi, '')
      .replace(/\s+/g, ' ')
      .trim();
  };

  const normalizeRoute = (r: string): string => {
    if (!r) return '';
    // Strip everything except alphanumeric, convert homoglyphs, and remove leading zeros
    const str = normalizeHomoglyphs(String(r).toLowerCase())
      .replace(/Маршрут\s*№?/gi, '')
      .replace(/маршрут/gi, '')
      .replace(/route/gi, '')
      .replace(/№/g, '')
      .replace(/[^0-9a-zа-яё]/gi, '')
      .trim();
    // Remove leading zeros for routes like "036" vs "36"
    return str.replace(/^0+/, '');
  };

  const prilText = await readFileAsText(prilFile);
  const transText = await readFileAsText(transFile);
  
  let krcData: KrcRow[] = [];
  if (krcFile) {
    krcData = await parseKrcFile(krcFile);
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
    const kDate = findKey(row, ['Транспортные сутки', 'Дата', 'Date', 'Day', 'Операционные сутки', 'Сутки']);
    const kRoute = findKey(row, ['№ маршрута', 'Маршрут', 'Route', 'Route Num', 'Маршрут №', '№марш.']);
    const kTime = findKey(row, ['Фактическое время начала рейса', 'Время начала', 'Start Time', 'Время', 'Время выезда', 'Начало рейса', 'Начало']);
    const kGrz = findKey(row, ['ГРЗ', 'Госномер', 'VREG_NUM', 'Vehicle', 'Гос. номер', 'Гос-номер']);
    const kWork = findKey(row, ['Фактическая транспортная работа', 'Пробег', 'Mileage', 'Работа', 'км', 'Расстояние']);
    const kDirection = findKey(row, ['Направление', 'Direction', 'Код направления', 'Прям/Обр']);
    const kConductor = findKey(row, ['ФИО кондуктора', 'Водитель', 'Conductor', 'Driver', 'ФИО', 'ФИО водителя', 'Ф.И.О.', 'Кондуктор']);

    const dateStr = (kDate ? row[kDate] : '').trim();
    const route = (kRoute ? row[kRoute] : '').trim();
    const startTimeStr = (kTime ? row[kTime] : '').trim();
    const grzRaw = (kGrz ? row[kGrz] : '').trim();
    const direction = (kDirection ? row[kDirection] : '').trim();
    const conductor = (kConductor ? row[kConductor] : '').trim();
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

    if (dateStr || route || grzRaw || conductor) {
      prilData.push({
        date_str: dateStr,
        route: route,
        start_time_str: startTimeStr,
        grz_raw: grzRaw,
        grz_norm: normalizeGrz(grzRaw),
        actual_work_km: actualWorkKm,
        direction: direction,
        conductor: conductor,
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
    const kConductor = findKey(row, ['CONDUCTOR', 'Кондуктор', 'ФИО кондуктора', 'Водитель', 'DRIVER', 'ФИО']);

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
    
    const conductorsToMatch = Array.from(new Set([
      ...conductors, 
      ...(flight.conductor ? [flight.conductor] : [])
    ])).filter(Boolean);

    // KRC Matching Logic
    let krcStatus = krcFile ? "Проверка не проводилась" : "";
    if (krcFile) {
      const flightRoute = normalizeRoute(flight.route);
      const transactionRoutes = confirmedTransactions.length > 0 
        ? Array.from(new Set(confirmedTransactions.map(t => normalizeRoute(t.ROUTE_NUM)).filter(Boolean)))
        : [];
      
      if (flight.start_datetime || confirmedTransactions.length > 0) {
        const tranDates = confirmedTransactions.map(t => t.tran_datetime).filter((d): d is Date => d !== null).sort((a, b) => a.getTime() - b.getTime());
        const tranStart = tranDates.length > 0 ? tranDates[0] : (flight.start_datetime || new Date());
        const tranEnd = tranDates.length > 0 ? tranDates[tranDates.length - 1] : new Date(tranStart.getTime() + tripDurationMinutes * 60000);
        
        // Find matching KRC check by 4 mandatory criteria: Date, Time, Route, Conductor (FIO)
        const matchingKrc = krcData.find(k => {
          if (!k.datetime) return false;
          
          // 1. CONDUCTOR (FIO) - MUST MATCH
          const conductorMatch = conductorsToMatch.some(c => conductorNamesMatch(k.conductor, c));
          if (!conductorMatch) return false;

          // 2. ROUTE - MUST MATCH
          const kRoute = normalizeRoute(k.route);
          // If KRC row has a route, it must match either the flight route or any route from transactions
          // Special case: if kRoute is empty in KRC row but conductor matches and it's same trip, we might consider it a match
          // but user asked for route matching, so we enforce it if present.
          const isRouteMatch = (!kRoute) || 
                            (flightRoute && (kRoute === flightRoute || flightRoute.includes(kRoute) || kRoute.includes(flightRoute))) ||
                            transactionRoutes.some(tr => kRoute === tr || tr.includes(kRoute) || kRoute.includes(tr));
          if (!isRouteMatch) return false;

          // 3. DATE - MUST MATCH
          const referenceDate = flight.start_datetime || tranStart;
          const isSameDay = k.datetime.getFullYear() === referenceDate.getFullYear() &&
                            k.datetime.getMonth() === referenceDate.getMonth() &&
                            k.datetime.getDate() === referenceDate.getDate();
          
          const kTime = k.datetime.getTime();
          const fTime = referenceDate.getTime();
          
          // Range check for same day: must fall exactly within transaction range
          if (isSameDay) {
            if (tranDates.length > 0) {
              const windowStart = tranStart.getTime(); 
              const windowEnd = tranEnd.getTime();
              return kTime >= windowStart && kTime <= windowEnd;
            }
            return false; // No transactions = no range to fall into
          }

          // Fallback for overnight shifts: if no transaction range yet, match very tight buffer around flight start
          const withinTimeBuffer = Math.abs(kTime - fTime) < 30 * 60000; 

          return withinTimeBuffer;
        });

        if (matchingKrc) {
          krcStatus = `Проверка проводилась, время проверки: ${matchingKrc.time}`;
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
      conductor: conductorsToMatch.length > 0 ? (() => {
        const full = conductorsToMatch[0];
        return full.split(/\s+/)
          .filter(Boolean)
          .map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
          .join(' ');
      })() : "",
      openTimes: confirmedTransactions.map(t => t.TIME).join('; '),
      closeTimes: confirmedTransactions.map(t => t.CR_TIME).join('; '),
      krcStatus: krcStatus,
      plannedMileage: finalDirection === "Прямое" ? forwardMileage : (finalDirection === "Обратное" ? returnMileage : 0)
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

export async function parseExcelReport(file: File): Promise<ReconciliationResult[]> {
  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();
  await workbook.xlsx.load(arrayBuffer);
  const worksheet = workbook.getWorksheet('Отчет') || workbook.getWorksheet(1);
  if (!worksheet) return [];

  const results: ReconciliationResult[] = [];
  const headers: string[] = [];
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber] = String(cell.value || '').trim();
  });

  const findCol = (targets: string[]) => {
    const idx = headers.findIndex(h => h && targets.some(t => h.toLowerCase().includes(t.toLowerCase())));
    return idx !== -1 ? idx : -1;
  };

  const colMap = {
    date: findCol(['Дата']),
    route: findCol(['Маршрут']),
    startTime: findCol(['Время начала']),
    grz: findCol(['ГРЗ']),
    status: findCol(['Статус']),
    tripNo: findCol(['Номер рейса', '№']),
    mileage: findCol(['Пробег', 'км']),
    direction: findCol(['Направление']),
    transCount: findCol(['Кол-во транзакций', 'Транз']),
    conductor: findCol(['Водитель', 'ФИО']),
    openTimes: findCol(['Время открытия']),
    closeTimes: findCol(['Время закрытия']),
    krcStatus: findCol(['Проверка КРС', 'KRC'])
  };

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const dateVal = colMap.date !== -1 ? row.getCell(colMap.date).text : '';
    if (dateVal === 'Итого' || !dateVal) return;

    results.push({
      date: dateVal,
      route: colMap.route !== -1 ? row.getCell(colMap.route).text : '',
      startTime: colMap.startTime !== -1 ? row.getCell(colMap.startTime).text : '',
      grz: colMap.grz !== -1 ? row.getCell(colMap.grz).text : '',
      status: colMap.status !== -1 ? row.getCell(colMap.status).text : '',
      tripNo: colMap.tripNo !== -1 ? row.getCell(colMap.tripNo).text : '',
      mileage: colMap.mileage !== -1 ? parseFloat(row.getCell(colMap.mileage).text.replace(',', '.')) || 0 : 0,
      direction: colMap.direction !== -1 ? row.getCell(colMap.direction).text : '',
      transCount: colMap.transCount !== -1 ? parseInt(row.getCell(colMap.transCount).text) || 0 : 0,
      conductor: colMap.conductor !== -1 ? row.getCell(colMap.conductor).text : '',
      openTimes: colMap.openTimes !== -1 ? row.getCell(colMap.openTimes).text : '',
      closeTimes: colMap.closeTimes !== -1 ? row.getCell(colMap.closeTimes).text : '',
      krcStatus: colMap.krcStatus !== -1 ? row.getCell(colMap.krcStatus).text : ''
    });
  });

  return results;
}

export async function enrichReportWithKrc(reportResults: ReconciliationResult[], krcData: KrcRow[]): Promise<ReconciliationResult[]> {
  const tripDurationMinutes = 120;
  
  const normalizeRoute = (r: string) => {
    if (!r) return "";
    // Remove all non-alphanumeric chars for a common denominator, but keep numbers
    // Actually, just remove leading zeros and prefixes like "м", "а", "№" to compare numeric parts
    return String(r).replace(/^(м|а|№|маршрут|марш|r)\s*/i, '').replace(/^0+/, '').trim().toLowerCase();
  };

  const normalizeHomoglyphs = (text: string): string => {
    if (!text) return "";
    const map: Record<string, string> = {
      'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н', 'K': 'К', 
      'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т', 'X': 'Х', 'y': 'у', 'Y': 'У',
      'a': 'а', 'b': 'в', 'c': 'с', 'e': 'е', 'h': 'н', 'k': 'к',
      'm': 'м', 'o': 'о', 'p': 'р', 't': 'т', 'x': 'х'
    };
    return text.split('').map(char => map[char] || char).join('');
  };

  const conductorNamesMatch = (nameA: string, nameB: string): boolean => {
    if (!nameA || !nameB) return false;
    const normalize = (n: string) => normalizeHomoglyphs(n.toLowerCase())
      .replace(/ё/g, 'е')
      .replace(/^(кондуктор|водитель|кассир|контролер)\s+/gi, '')
      .replace(/[^a-zа-яё\s]/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const normA = normalize(nameA);
    const normB = normalize(nameB);
    if (!normA || !normB) return false;
    
    // Direct match or one contains the other
    if (normA === normB || normA.includes(normB) || normB.includes(normA)) return true;
    
    const partsA = normA.split(/\s+/).filter(p => p.length >= 2);
    const partsB = normB.split(/\s+/).filter(p => p.length >= 2);
    
    // Check if the surname (usually the first part) matches exactly
    if (partsA[0] === partsB[0]) return true;

    const hasSharedSignificantPart = partsA.some(pa => partsB.some(pb => pa === pb && pa.length >= 3));
    if (hasSharedSignificantPart) return true;

    return false;
  };

  return reportResults.map(res => {
    const flightRoute = normalizeRoute(res.route);
    const conductor = res.conductor;
    
    // Parse start time
    const [h, m] = res.startTime.split(':').map(Number);
    // Support multiple date formats (DD.MM.YYYY, DD/MM/YYYY, YYYY-MM-DD, etc.)
    const dateNums = res.date.split(/[^0-9]+/).filter(Boolean).map(Number);
    let d, mon, y;
    if (dateNums.length >= 3) {
      if (dateNums[0] > 1000) { y = dateNums[0]; mon = dateNums[1]; d = dateNums[2]; }
      else if (dateNums[2] > 1000) { d = dateNums[0]; mon = dateNums[1]; y = dateNums[2]; }
      else { d = dateNums[0]; mon = dateNums[1]; y = 2000 + dateNums[2]; }
    } else {
      // Fallback
      return res;
    }
    const startDatetime = new Date(y, mon - 1, d, h, m);
    
    // Parse transaction times for window estimation from openTimes (Время открытия транзакции)
    const openTimesList = res.openTimes.split(';').map(t => t.trim()).filter(Boolean);
    
    const parseClock = (clock: string) => {
      const parts = clock.split(':').map(Number);
      const th = parts[0] || 0;
      const tm = parts[1] || 0;
      const ts = parts[2] || 0;
      return new Date(y, mon - 1, d, th, tm, ts);
    };

    let tranStart: Date | null = null;
    let tranEnd: Date | null = null;

    if (openTimesList.length > 0) {
      const times = openTimesList.map(t => parseClock(t)).sort((a, b) => a.getTime() - b.getTime());
      tranStart = times[0];
      tranEnd = times[times.length - 1];
    } else {
      // Fallback if no transactions are recorded
      tranStart = startDatetime;
      tranEnd = new Date(startDatetime.getTime() + tripDurationMinutes * 60000);
    }

    const matchingKrc = krcData.find(k => {
      if (!k.datetime) return false;
      
      // 1. Date Match
      const isSameDay = k.datetime.getFullYear() === startDatetime.getFullYear() &&
                        k.datetime.getMonth() === startDatetime.getMonth() &&
                        k.datetime.getDate() === startDatetime.getDate();
      if (!isSameDay) return false;

      // 2. Route Match - relaxed logic
      const kRoute = normalizeRoute(k.route);
      const isRouteMatch = (!kRoute) || 
                          (flightRoute && (kRoute === flightRoute || flightRoute.includes(kRoute) || kRoute.includes(flightRoute)));
      if (!isRouteMatch) return false;
      
      // 3. FIO Match (Driver/Conductor Name)
      if (!conductorNamesMatch(k.conductor, conductor)) return false;

      // 4. Time Range Match - Must fall exactly into transaction range
      const kTime = k.datetime.getTime();
      
      if (openTimesList.length > 0 && tranStart && tranEnd) {
        const windowStart = tranStart.getTime();
        const windowEnd = tranEnd.getTime();
        return kTime >= windowStart && kTime <= windowEnd;
      }
      
      return false; // Exactly in range implies range must exist via transactions
    });

    return {
      ...res,
      krcStatus: matchingKrc ? `Проверка проводилась, время проверки: ${matchingKrc.time}` : 'Проверка не проводилась'
    };
  });
}

export async function parseCsvReport(file: File): Promise<ReconciliationResult[]> {
  const text = await readFileAsText(file);
  
  // Custom separator detection for CSV reports
  const detectSeparator = (t: string): string => {
    const sample = t.slice(0, 5000);
    const semiCount = (sample.match(/;/g) || []).length;
    const commaCount = (sample.match(/,/g) || []).length;
    const tabCount = (sample.match(/\t/g) || []).length;
    if (tabCount > semiCount && tabCount > commaCount) return '\t';
    return semiCount >= commaCount ? ';' : ',';
  };

  const sep = detectSeparator(text);
  const parseResult = Papa.parse(text, { 
    delimiter: sep,
    header: true, 
    skipEmptyLines: true 
  });
  const rows = parseResult.data as any[];

  return rows.filter(row => row['Дата'] && row['Дата'] !== 'Итого').map(row => ({
    date: row['Дата'] || '',
    route: row['Маршрут'] || '',
    startTime: row['Время начала (Отчет)'] || row['Время начала'] || '',
    grz: row['ГРЗ'] || '',
    status: row['Статус подтверждения'] || row['Статус'] || '',
    tripNo: row['Номер рейса'] || '',
    mileage: parseFloat(String(row['Фактическая транспортная работа (км)'] || row['Пробег'] || '0').replace(',', '.')) || 0,
    direction: row['Направление'] || '',
    transCount: parseInt(String(row['Кол-во транзакций'] || '0')) || 0,
    conductor: row['ФИО водителя'] || row['Водитель'] || '',
    openTimes: row['Время открытия транзакции'] || row['Время открытия'] || '',
    closeTimes: row['Время закрытия транзакции'] || row['Время закрытия'] || '',
    krcStatus: row['Проверка КРС'] || row['KRC'] || ''
  }));
}

export async function generateCsv(results: ReconciliationResult[]): Promise<Blob> {
  const csvData = results.filter(res => res.date && res.grz && res.startTime).map(res => ({
    'Дата': res.date,
    'Маршрут': res.route,
    'Время начала (Отчет)': res.startTime,
    'ГРЗ': res.grz,
    'ФИО водителя': res.conductor,
    'Статус подтверждения': res.status,
    'Номер рейса': res.tripNo,
    'Фактическая транспортная работа (км)': String(res.mileage.toFixed(2)).replace('.', ','),
    'Плановая транспортная работа (км)': res.plannedMileage ? String(res.plannedMileage.toFixed(2)).replace('.', ',') : '0,00',
    'Направление': res.direction,
    'Кол-во транзакций': res.transCount,
    'Проверка КРС': res.krcStatus,
    'Время открытия транзакции': res.openTimes,
    'Время закрытия транзакции': res.closeTimes
  }));

  const totalMileage = results.reduce((sum, r) => sum + r.mileage, 0);
  const totalTrans = results.reduce((sum, r) => sum + r.transCount, 0);
  
  if (csvData.length > 0) {
    csvData.push({
      'Дата': 'Итого',
      'Маршрут': '',
      'Время начала (Отчет)': '',
      'ГРЗ': '',
      'ФИО водителя': '',
      'Статус подтверждения': '',
      'Номер рейса': '',
      'Фактическая транспортная работа (км)': String(totalMileage.toFixed(2)).replace('.', ','),
      'Плановая транспортная работа (км)': '',
      'Направление': '',
      'Кол-во транзакций': totalTrans as any,
      'Проверка КРС': '',
      'Время открытия транзакции': '',
      'Время закрытия транзакции': ''
    });
  }

  const csv = Papa.unparse(csvData, { delimiter: ';' });
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  return blob;
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
    { header: 'Фактическая транспортная работа (км)', key: 'mileage', width: 25, style: { numFmt: '#,##0.00' } },
    { header: 'Плановая транспортная работа (км)', key: 'plannedMileage', width: 25, style: { numFmt: '#,##0.00' } },
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
    // Determine if row needs highlighting for mismatch
    let isMismatch = false;
    if (rowNumber > 1) {
      const dateVal = row.getCell(1).value;
      if (dateVal !== 'Итого') {
        const actual = parseFloat(row.getCell(8).value as any) || 0;
        const planned = parseFloat(row.getCell(9).value as any) || 0;
        // Highlight only if plan is set and doesn't match actual
        if (planned > 0 && Math.abs(actual - planned) > 0.01) {
          isMismatch = true;
        }
      }
    }

    row.eachCell((cell) => {
      cell.alignment = { 
        wrapText: true, 
        vertical: 'middle',
        horizontal: rowNumber === 1 ? 'center' : 'left'
      };
      
      // Highlight mismatching rows
      if (isMismatch) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFC7CE' } // Light red
        };
        cell.font = { color: { argb: 'FF9C0006' } }; // Dark red text
      }

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

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
  stats: { confirmed: number; unconfirmed: number; krcChecks: number };
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
    const keywords = ['дата', 'маршрут', 'грз', 'госномер', 'время', 'рейс', 'date', 'vreg', 'time', 'route', 'сутки'];
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
  
  // Check for common Russian headers or characters
  const commonHeaders = ['транспортные', 'дата', 'маршрут', 'грз', 'рейс', 'время'];
  const foundUtf8 = commonHeaders.some(h => text.toLowerCase().includes(h));
  
  if (!foundUtf8) {
    try {
      // Try Windows-1251 if UTF-8 doesn't seem to have the headers
      const win1251Decoder = new TextDecoder('windows-1251');
      const winText = win1251Decoder.decode(buffer);
      if (commonHeaders.some(h => winText.toLowerCase().includes(h))) {
        return winText;
      }
    } catch (e) {
      // fallback
    }
  }
  return text;
}

export function detectSeparator(text: string): string {
  const firstLine = text.split('\n')[0];
  const semiCount = (firstLine.match(/;/g) || []).length;
  const commaCount = (firstLine.match(/,/g) || []).length;
  return semiCount >= commaCount ? ';' : ',';
}

export async function reconcileFiles(
  prilFile: File, 
  transFile: File, 
  tripDurationMinutes: number = 120,
  krcFile?: File | null
): Promise<ReconciliationResponse> {
  const findKey = (obj: any, target: string | string[]) => {
    const keys = Object.keys(obj);
    const targets = Array.isArray(target) ? target.map(t => t.toLowerCase()) : [target.toLowerCase()];
    return keys.find(k => targets.includes(k.trim().toLowerCase()));
  };

  const prilText = await readFileAsText(prilFile);
  const transText = await readFileAsText(transFile);
  
  let krcData: KrcRow[] = [];
  if (krcFile) {
    const krcText = await readFileAsText(krcFile);
    const krcSep = detectSeparator(krcText);
    const krcRowsRaw = await parseCsvLocal(krcText, krcSep);
    
    for (const row of krcRowsRaw) {
      const kRoute = findKey(row, ['№ маршрута', 'Маршрут', 'Route', 'Route Num']);
      const kConductor = findKey(row, ['ФИО кондуктора', 'Кондуктор', 'Conductor', 'ФИО']);
      const kTime = findKey(row, ['Время', 'Time', 'Дата и время']);
      const kDate = findKey(row, ['Дата', 'Date']);

      const route = (kRoute ? row[kRoute] : '').trim();
      const conductor = (kConductor ? row[kConductor] : '').trim();
      const timeStr = (kTime ? row[kTime] : '').trim();
      const dateStr = (kDate ? row[kDate] : '').trim();

      let datetime: Date | null = null;
      if (timeStr) {
        // Simple attempt to parse date/time from KRC
        const fullStr = dateStr ? `${dateStr} ${timeStr}` : timeStr;
        const d = new Date(fullStr);
        if (!isNaN(d.getTime())) {
          datetime = d;
        } else {
           // Try specific parsing if standard fails
           const parts = timeStr.split(/[\s,:]/);
           if (parts.length >= 2) {
             const now = new Date();
             const h = parseInt(parts[0]);
             const m = parseInt(parts[1]);
             if (!isNaN(h) && !isNaN(m)) {
               now.setHours(h, m, 0, 0);
               datetime = now;
             }
           }
        }
      }

      if (route || conductor) {
        krcData.push({
          route,
          conductor,
          time: timeStr,
          datetime
        });
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
    if (isConfirmed) confirmedCount++; else unconfirmedCount++;

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

    // KRC Matching Logic
    let krcStatus = krcFile ? "Проверка не проводилась" : "";
    if (krcFile && confirmedTransactions.length > 0) {
      const flightRoute = flight.route;
      const tranStart = confirmedTransactions[0].tran_datetime;
      const tranEnd = confirmedTransactions[confirmedTransactions.length - 1].tran_datetime;
      const conductors = Array.from(new Set(confirmedTransactions.map(t => t.CONDUCTOR).filter(Boolean)));

      if (tranStart && tranEnd) {
        const foundMatch = krcData.some(k => {
          const routeMatch = k.route === flightRoute || flightRoute.includes(k.route) || k.route.includes(flightRoute);
          if (!routeMatch) return false;

          const conductorMatch = conductors.some(c => 
            c.toLowerCase().includes(k.conductor.toLowerCase()) || 
            k.conductor.toLowerCase().includes(c.toLowerCase())
          );
          if (!conductorMatch) return false;

          if (k.datetime) {
            // Buffer of 30 mins around transactions window for KRC check
            const buffer = 30 * 60000;
            return k.datetime >= new Date(tranStart.getTime() - buffer) && 
                   k.datetime <= new Date(tranEnd.getTime() + buffer);
          }
          return true; // If no datetime in KRC, match by route/conductor only
        });

        if (foundMatch) {
          krcStatus = "Проверка проводилась";
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
      krcChecks: krcCheckCount
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

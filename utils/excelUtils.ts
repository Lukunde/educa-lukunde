import * as XLSX from 'xlsx';
import { Sheet, SheetData } from '../types';

// Robust UUID generator that works in all contexts (secure/insecure)
export const generateUUID = () => {
    // Try crypto API first (modern browsers, secure context)
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
        try {
            return crypto.randomUUID();
        } catch (e) {
            // Fallback if randomUUID fails for some reason
        }
    }
    // Fallback timestamp + random generator
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
};

export const parseExcelFile = async (file: File): Promise<Sheet[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        
        const sheets: Sheet[] = workbook.SheetNames.map(name => {
          const ws = workbook.Sheets[name];
          // Convert sheet to array of arrays
          const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as SheetData;
          return {
            id: generateUUID(),
            name,
            data: jsonData
          };
        });

        resolve(sheets);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
};

export const splitSheetByColumn = (sheet: Sheet, columnIndex: number): Sheet[] => {
  const data = sheet.data;
  if (!data || data.length < 2) return []; // No data to split

  const header = data[0];
  const rows = data.slice(1);
  
  const groups: Record<string, any[][]> = {};

  rows.forEach(row => {
    // Skip empty or undefined rows
    if (!row || row.length === 0) return;

    const key = String(row[columnIndex] || "Sem Turma");
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(row);
  });

  const newSheets: Sheet[] = Object.keys(groups).map(key => ({
    id: generateUUID(),
    name: `${sheet.name} - ${key}`,
    data: [header, ...groups[key]]
  }));

  return newSheets;
};
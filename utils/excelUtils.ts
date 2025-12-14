import * as XLSX from 'xlsx';
import { Sheet, SheetData } from '../types';

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
            id: crypto.randomUUID(),
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
  if (data.length < 2) return []; // No data to split

  const header = data[0];
  const rows = data.slice(1);
  
  const groups: Record<string, any[][]> = {};

  rows.forEach(row => {
    const key = String(row[columnIndex] || "Sem Turma");
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(row);
  });

  const newSheets: Sheet[] = Object.keys(groups).map(key => ({
    id: crypto.randomUUID(),
    name: `${sheet.name} - ${key}`,
    data: [header, ...groups[key]]
  }));

  return newSheets;
};
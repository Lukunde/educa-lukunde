import React, { useState, useEffect, useRef } from 'react';
import { SheetData, CellValue, ConditionalRule, ValidationRule } from '../types';

interface SpreadsheetProps {
  data: SheetData;
  rules?: ConditionalRule[];
  validationRules?: ValidationRule[];
  onCellChange: (rowIndex: number, colIndex: number, value: CellValue) => void;
  zoom?: number;
}

const Spreadsheet: React.FC<SpreadsheetProps> = ({ data, rules = [], validationRules = [], onCellChange, zoom = 1 }) => {
  const [editingCell, setEditingCell] = useState<{r: number, c: number} | null>(null);
  const [selectedCell, setSelectedCell] = useState<{r: number, c: number} | null>(null);
  const tableRef = useRef<HTMLDivElement>(null);

  // Determine max columns safely
  const maxCols = data && data.length > 0 ? data.reduce((max, row) => Math.max(max, row && Array.isArray(row) ? row.length : 0), 0) : 0;
  
  // Generate headers safely
  const headers = Array.from({ length: maxCols || 0 }, (_, i) => {
    let label = "";
    let n = i;
    while (n >= 0) {
      label = String.fromCharCode((n % 26) + 65) + label;
      n = Math.floor(n / 26) - 1;
    }
    return label;
  });

  // Keyboard Navigation & Shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // If no data, ignore
      if (!data || data.length === 0) return;

      // If we are editing, let the input handle navigation/typing, 
      // EXCEPT for Enter (Commit) and Escape (Cancel)
      if (editingCell) {
        if (e.key === 'Escape') {
          e.preventDefault();
          setEditingCell(null);
          // Return focus to grid logic if needed
        }
        return;
      }

      if (!selectedCell) return;

      const { r, c } = selectedCell;
      const maxRows = data.length;

      switch (e.key) {
        case 'ArrowUp':
          e.preventDefault();
          setSelectedCell({ r: Math.max(0, r - 1), c });
          break;
        case 'ArrowDown':
          e.preventDefault();
          setSelectedCell({ r: Math.min(maxRows - 1, r + 1), c });
          break;
        case 'ArrowLeft':
          e.preventDefault();
          setSelectedCell({ r, c: Math.max(0, c - 1) });
          break;
        case 'ArrowRight':
          e.preventDefault();
          setSelectedCell({ r, c: Math.min(maxCols - 1, c + 1) });
          break;
        case 'Enter':
        case 'F2':
          e.preventDefault();
          setEditingCell({ r, c });
          break;
        case 'Delete':
        case 'Backspace':
          onCellChange(r, c, "");
          break;
        case 'Tab':
          e.preventDefault();
          if (e.shiftKey) {
             setSelectedCell({ r, c: Math.max(0, c - 1) });
          } else {
             setSelectedCell({ r, c: Math.min(maxCols - 1, c + 1) });
          }
          break;
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selectedCell, editingCell, data, maxCols, onCellChange]); // Added data dependency explicitly

  const getCellStyle = (rowIndex: number, colIndex: number, value: CellValue) => {
    if (!rules || rowIndex === 0) return {}; 

    const columnRules = rules.filter(r => r.columnIndex === colIndex);
    if (columnRules.length === 0) return {};

    // Normalize value for comparison (handle comma as decimal separator)
    const normalizedValue = typeof value === 'string' ? value.replace(',', '.') : value;
    const numValue = Number(normalizedValue);
    const isNumber = !isNaN(numValue) && value !== "" && value !== null;

    for (const rule of columnRules) {
      const ruleValue = Number(rule.value);
      let match = false;

      if (isNumber) {
        switch (rule.condition) {
          case 'gt': match = numValue > ruleValue; break;
          case 'lt': match = numValue < ruleValue; break;
          case 'gte': match = numValue >= ruleValue; break;
          case 'lte': match = numValue <= ruleValue; break;
          case 'eq': match = numValue === ruleValue; break;
        }
      } else {
        const strVal = String(value).toLowerCase();
        const strRule = String(rule.value).toLowerCase();
        if (rule.condition === 'contains') {
          match = strVal.includes(strRule);
        } else if (rule.condition === 'eq') {
          match = strVal === strRule;
        }
      }

      if (match) {
        return {
          backgroundColor: rule.style.backgroundColor,
          color: rule.style.color,
          fontWeight: '500'
        };
      }
    }
    return {};
  };

  const handleCellClick = (rowIndex: number, colIndex: number) => {
    setSelectedCell({ r: rowIndex, c: colIndex });
    // If we click a different cell while editing, stop editing the previous one
    if (editingCell) {
        setEditingCell(null);
    }
  };

  const handleDoubleClick = (rowIndex: number, colIndex: number) => {
    setSelectedCell({ r: rowIndex, c: colIndex });
    setEditingCell({ r: rowIndex, c: colIndex });
  };

  const handleSave = (r: number, c: number, value: string) => {
      onCellChange(r, c, value);
      setEditingCell(null);
      // Focus goes back to selection automatically via state
  };

  const getInputType = (cIdx: number) => {
      const rule = validationRules?.find(r => r.columnIndex === cIdx);
      if (!rule) return 'text';
      if (rule.type === 'number') return 'number';
      if (rule.type === 'date') return 'date';
      return 'text';
  };

  if (!data || data.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full text-gray-400 dark:text-gray-500 bg-gray-50 dark:bg-gray-900">
        <p className="text-lg font-medium">Nenhum dado para exibir</p>
        <p className="text-sm">Carregue um arquivo Excel ou crie uma nova planilha.</p>
      </div>
    );
  }

  return (
    <div className="flex-1 overflow-auto bg-gray-100 dark:bg-gray-900 relative transition-colors duration-200" ref={tableRef}>
      <div 
        className="inline-block min-w-full shadow-sm bg-white dark:bg-gray-800 m-4 rounded-lg border border-gray-200 dark:border-gray-700 transition-colors"
        style={{ zoom: zoom } as any}
      >
        <table className="w-full border-collapse text-sm table-fixed">
          <thead>
            <tr>
              <th className="w-10 border-b border-r border-gray-200 dark:border-gray-700 p-2 text-center bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 sticky top-0 left-0 z-30">
                #
              </th>
              {headers.map((col, idx) => {
                const isColSelected = selectedCell?.c === idx;
                return (
                    <th key={idx} className={`border-b border-r border-gray-200 dark:border-gray-700 px-4 py-2 min-w-[100px] text-left sticky top-0 z-20 transition-colors ${
                        isColSelected 
                        ? 'bg-blue-50/80 dark:bg-blue-900/50 text-blue-800 dark:text-blue-300' 
                        : 'bg-gray-50 dark:bg-gray-700 text-gray-700 dark:text-gray-200'
                    }`}>
                    {col}
                    </th>
                )
              })}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rIdx) => {
              if (!row || !Array.isArray(row)) return null; // Skip undefined or invalid rows
              return (
              <tr key={rIdx} className="hover:bg-blue-50/10 dark:hover:bg-blue-900/10">
                <td className={`border-b border-r border-gray-200 dark:border-gray-700 text-center text-xs font-medium sticky left-0 z-10 transition-colors
                    ${selectedCell?.r === rIdx 
                        ? 'bg-blue-100 dark:bg-blue-900/70 text-blue-800 dark:text-blue-200' 
                        : 'bg-gray-50 dark:bg-gray-700 text-gray-400 dark:text-gray-400'
                    }
                `}>
                  {rIdx + 1}
                </td>
                {headers.map((_, cIdx) => {
                  const cellValue = row[cIdx];
                  const safeValue = cellValue === null || cellValue === undefined ? "" : String(cellValue);
                  
                  const isEditing = editingCell?.r === rIdx && editingCell?.c === cIdx;
                  const isSelected = selectedCell?.r === rIdx && selectedCell?.c === cIdx;
                  const style = !isEditing ? getCellStyle(rIdx, cIdx, cellValue) : {};
                  
                  const validationRule = validationRules?.find(r => r.columnIndex === cIdx);

                  return (
                    <td 
                      key={cIdx} 
                      className={`
                        border-b border-r border-gray-200 dark:border-gray-700 p-0 h-9 min-w-[100px] relative transition-colors duration-75
                        ${isSelected && !isEditing ? 'ring-2 ring-emerald-500 z-10' : ''}
                      `}
                      style={style}
                      onClick={() => handleCellClick(rIdx, cIdx)}
                      onDoubleClick={() => handleDoubleClick(rIdx, cIdx)}
                    >
                      {isEditing ? (
                        validationRule?.type === 'list' && validationRule.options ? (
                             <select
                                autoFocus
                                defaultValue={safeValue}
                                onBlur={(e) => handleSave(rIdx, cIdx, e.target.value)}
                                onChange={(e) => handleSave(rIdx, cIdx, e.target.value)}
                                onKeyDown={(e) => {
                                    if (e.key === 'Enter') {
                                        e.preventDefault();
                                        handleSave(rIdx, cIdx, e.currentTarget.value);
                                        setSelectedCell(prev => prev ? ({ ...prev, r: Math.min(data.length - 1, prev.r + 1) }) : null);
                                    }
                                }}
                                className="w-full h-full px-2 outline-none border-2 border-emerald-500 z-20 absolute top-0 left-0 text-sm shadow-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                             >
                                <option value="" disabled>Selecione...</option>
                                {validationRule.options.map(opt => (
                                    <option key={opt} value={opt}>{opt}</option>
                                ))}
                             </select>
                        ) : (
                            <input
                              autoFocus
                              type={getInputType(cIdx)}
                              defaultValue={safeValue}
                              onBlur={(e) => handleSave(rIdx, cIdx, e.target.value)}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  handleSave(rIdx, cIdx, e.currentTarget.value);
                                  setSelectedCell(prev => prev ? ({ ...prev, r: Math.min(data.length - 1, prev.r + 1) }) : null);
                                }
                              }}
                              className="w-full h-full px-2 outline-none border-2 border-emerald-500 z-20 absolute top-0 left-0 text-sm shadow-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                            />
                        )
                      ) : (
                        <div className="px-2 py-1.5 w-full h-full truncate select-none cursor-cell text-gray-700 dark:text-gray-300">
                          {cellValue}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default Spreadsheet;
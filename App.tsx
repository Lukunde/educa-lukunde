import React, { useState, useEffect, useRef } from 'react';
import { BookOpen, Pencil, Upload, Split, Plus, MessageSquare, Download, Menu, FileSpreadsheet, SaveAll, Palette, X, Trash2, Copy, Edit, ZoomIn, ZoomOut, Share2, Lock, Unlock, Link as LinkIcon, Check, Moon, Sun, ShieldCheck, Calculator, Clock, Calendar, ListChecks, Eye } from 'lucide-react';
import Spreadsheet from './components/Spreadsheet';
import AIAssistant from './components/AIAssistant';
import { Sheet, SheetData, ConditionalRule, ConditionType, ConditionalStyle, ValidationRule, ValidationType } from './types';
import { parseExcelFile, splitSheetByColumn, generateUUID } from './utils/excelUtils';
import { suggestClassColumn } from './services/geminiService';
import * as XLSX from 'xlsx';

// Preset Styles
const PRESET_STYLES: ConditionalStyle[] = [
  { name: 'Vermelho (Reprovado)', backgroundColor: '#FECACA', color: '#991B1B' },
  { name: 'Verde (Aprovado)', backgroundColor: '#BBF7D0', color: '#166534' },
  { name: 'Amarelo (Atenção)', backgroundColor: '#FEF08A', color: '#854D0E' },
  { name: 'Azul (Destaque)', backgroundColor: '#BFDBFE', color: '#1E40AF' },
];

const App: React.FC = () => {
  // Theme State
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    if (typeof window !== 'undefined') {
        try {
            const saved = localStorage.getItem('educa-lukunde-theme');
            if (saved === 'dark' || saved === 'light') return saved;
            return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
        } catch (e) {
            // Fallback if localStorage access is denied
            return 'light';
        }
    }
    return 'light';
  });

  // Initialize sheets from localStorage safely
  const [sheets, setSheets] = useState<Sheet[]>(() => {
    // 1. Environment check
    if (typeof window === 'undefined') return [];

    try {
      // 2. Storage Access check (handles SecurityError/Origin issues)
      const storage = window.localStorage;
      if (!storage) return [];

      const saved = storage.getItem('educa-lukunde-sheets');
      if (!saved) return [];
      
      // 3. Parsing check
      const parsed = JSON.parse(saved);
      
      // 4. Structure check
      if (!Array.isArray(parsed)) return [];
      
      // 5. Deep content validation & MIGRATION of old accessCode
      return parsed.map((s: any) => {
          if (!s || typeof s !== 'object' || typeof s.id !== 'string') return null;
          
          // Legacy migration: If has accessCode but no editCode, map it
          if (s.accessCode && !s.editCode) {
              s.editCode = s.accessCode;
              // We keep accessCode temporarily or delete it. Let's keep s structure valid.
          }
          
          return s;
      }).filter((s: any) => s !== null);

    } catch (e) {
      console.warn("Educa-Lukunde: Could not load data from localStorage (security or corruption issue). Starting fresh.", e);
      return [];
    }
  });
  
  const [activeSheetId, setActiveSheetId] = useState<string | null>(null);
  const [showAI, setShowAI] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  
  // Drag and Drop State using Ref to avoid closure staleness
  const dragItem = useRef<number | null>(null);
  const [draggedSheetIndex, setDraggedSheetIndex] = useState<number | null>(null); // For visual feedback only
  
  // Conditional Formatting State
  const [showFormatModal, setShowFormatModal] = useState(false);
  const [newRule, setNewRule] = useState<{
    colHeader: string;
    condition: ConditionType;
    value: string;
    styleIndex: number;
  }>({
    colHeader: 'A',
    condition: 'lt',
    value: '10',
    styleIndex: 0
  });

  // Data Validation State
  const [showValidationModal, setShowValidationModal] = useState(false);
  const [newValidation, setNewValidation] = useState<{
    colHeader: string;
    type: ValidationType;
    min: string;
    max: string;
    options: string;
    errorMessage: string;
  }>({
    colHeader: 'A',
    type: 'number',
    min: '',
    max: '',
    options: '',
    errorMessage: ''
  });

  // Sharing & Security State
  const [showShareModal, setShowShareModal] = useState(false);
  // unlockedSheets tracks the LEVEL of access: 'edit' or 'view'
  const [unlockedSheets, setUnlockedSheets] = useState<Record<string, 'edit' | 'view'>>({}); 
  const [accessCodeInput, setAccessCodeInput] = useState("");
  const [accessError, setAccessError] = useState<string | null>(null);
  
  const [copyEditSuccess, setCopyEditSuccess] = useState(false);
  const [copyViewSuccess, setCopyViewSuccess] = useState(false);
  const [linkCopySuccess, setLinkCopySuccess] = useState(false);
  
  // Expiration Configuration State
  const [expirationValue, setExpirationValue] = useState(24);
  const [expirationUnit, setExpirationUnit] = useState<'hours' | 'days' | 'weeks' | 'months'>('hours');

  // Context Menu State
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number; sheetId: string } | null>(null);

  // Zoom State
  const [zoomLevel, setZoomLevel] = useState(1);

  const activeSheet = sheets.find(s => s.id === activeSheetId);
  const currentAccess = activeSheetId ? unlockedSheets[activeSheetId] : undefined;
  
  // A sheet is "Locked" (showing the lock screen) if:
  // 1. It has security codes (edit or view)
  // 2. AND the user hasn't unlocked it yet (no entry in unlockedSheets)
  const isSheetLocked = activeSheet && (activeSheet.editCode || activeSheet.viewCode) && !currentAccess;
  
  // Access Levels
  const canEdit = activeSheet && (!activeSheet.editCode && !activeSheet.viewCode) ? true : currentAccess === 'edit';
  const isReadOnly = currentAccess === 'view';

  // Apply Theme
  useEffect(() => {
    if (theme === 'dark') {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    try {
        localStorage.setItem('educa-lukunde-theme', theme);
    } catch (e) {
        // Ignore storage errors
    }
  }, [theme]);

  const toggleTheme = () => setTheme(prev => prev === 'light' ? 'dark' : 'light');

  // Persistence: Save sheets to localStorage whenever they change
  useEffect(() => {
    if (sheets.length > 0) {
      try {
        const dataToSave = JSON.stringify(sheets);
        localStorage.setItem('educa-lukunde-sheets', dataToSave);
      } catch (e) {
        console.error("Erro ao salvar no localStorage (provavelmente cota excedida ou acesso negado):", e);
      }
    }
  }, [sheets]);

  // Sync URL with active sheet ID
  useEffect(() => {
    if (activeSheetId) {
      try {
        const currentUrl = window.location.href;
        // Check for blob protocol to avoid replacing state on blob URLs which causes crashes
        if (currentUrl.startsWith('blob:')) return;

        const url = new URL(currentUrl);
        url.searchParams.set('pauta', activeSheetId);
        window.history.replaceState({}, '', url.toString());
      } catch (e) {
        // Silently fail in restricted environments (sandboxes, iframes, blobs)
        // This is expected in preview environments and should not crash the app
      }
    }
  }, [activeSheetId]);

  // Initial Empty Sheet if nothing loaded
  useEffect(() => {
    if (sheets.length === 0) {
      const initialSheet: Sheet = {
        id: 'init',
        name: 'Pauta 1',
        data: Array(20).fill(Array(10).fill("")),
        conditionalFormats: [],
        validationRules: []
      };
      setSheets([initialSheet]);
      setActiveSheetId('init');
      // Auto unlock 'init' sheet as 'edit'
      setUnlockedSheets(prev => ({ ...prev, 'init': 'edit' }));
    } else if (!activeSheetId) {
       // If loaded from storage, check URL for shared sheet or default to first
       let sharedSheetId: string | null = null;
       try {
           const params = new URLSearchParams(window.location.search);
           sharedSheetId = params.get('pauta');
       } catch (e) {
           // Ignore errors parsing URL
       }
       
       if (sharedSheetId && sheets.find(s => s.id === sharedSheetId)) {
         setActiveSheetId(sharedSheetId);
       } else {
         const firstSheet = sheets[0];
         if (firstSheet) {
            setActiveSheetId(firstSheet.id);
            // Auto unlock if it's the default/init sheet and not shared
            if (!firstSheet.editCode && !firstSheet.viewCode) {
                setUnlockedSheets(prev => ({ ...prev, [firstSheet.id]: 'edit' }));
            }
         }
       }
    }
  }, [sheets.length]); 

  // Drag and Drop Handlers
  const handleDragStart = (e: React.DragEvent, index: number) => {
    if (e.dataTransfer) {
        dragItem.current = index;
        setDraggedSheetIndex(index);
        e.dataTransfer.effectAllowed = "move";
    }
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    
    // Use ref for the current dragged item to ensure we have the latest value without re-render closures
    const currentDragIndex = dragItem.current;
    if (currentDragIndex === null || currentDragIndex === index) return;
    
    // Create new array
    const newSheets = [...sheets];
    const draggedItemContent = newSheets[currentDragIndex];
    
    // Safety check
    if (!draggedItemContent) return;

    // Remove from old index
    newSheets.splice(currentDragIndex, 1);
    // Insert at new index
    newSheets.splice(index, 0, draggedItemContent);
    
    // Update ref to the new position so subsequent dragOvers are correct
    dragItem.current = index;
    
    // Update state
    setSheets(newSheets);
    setDraggedSheetIndex(index);
  };

  const handleDragEnd = () => {
    dragItem.current = null;
    setDraggedSheetIndex(null);
  };

  const handleExport = () => {
    if (!activeSheet) return;
    const ws = XLSX.utils.aoa_to_sheet(activeSheet.data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeSheet.name);
    XLSX.writeFile(wb, `${activeSheet.name}.xlsx`);
  };

  const handleExportAll = () => {
    if (sheets.length === 0) return;
    const wb = XLSX.utils.book_new();
    
    sheets.forEach(sheet => {
      if (!sheet) return;
      const ws = XLSX.utils.aoa_to_sheet(sheet.data);
      let sheetName = (sheet.name || "Sheet").replace(/[\\/?*[\]]/g, " ").trim();
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
      if (!sheetName) sheetName = "Sheet";

      let uniqueName = sheetName;
      let counter = 1;
      while (wb.SheetNames.includes(uniqueName)) {
        uniqueName = `${sheetName.substring(0, 27)}(${counter})`;
        counter++;
      }

      XLSX.utils.book_append_sheet(wb, ws, uniqueName);
    });
    
    XLSX.writeFile(wb, "Educa-Lukunde_Completo.xlsx");
  };

  // Keyboard Shortcuts
  useEffect(() => {
    const handleGlobalShortcuts = (e: KeyboardEvent) => {
        // Ctrl + S: Save/Export
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            handleExport();
        }
        
        // Ctrl + Shift + M: Tab Context Menu
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && (e.key === 'm' || e.key === 'M')) {
            e.preventDefault();
            if (activeSheetId) {
                // Position menu at bottom left near tabs
                setContextMenu({ x: 200, y: window.innerHeight - 150, sheetId: activeSheetId });
            }
        }
    };

    window.addEventListener('keydown', handleGlobalShortcuts);
    return () => window.removeEventListener('keydown', handleGlobalShortcuts);
  }, [activeSheetId, sheets]); // Dependencies to ensure current state is used

  // Close context menu on global click
  useEffect(() => {
    const handleClick = () => setContextMenu(null);
    window.addEventListener('click', handleClick);
    return () => window.removeEventListener('click', handleClick);
  }, []);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    try {
      const parsedSheets = await parseExcelFile(file);
      // Mark uploaded sheets as unlocked ('edit') for the uploader
      const newIds = parsedSheets.map(s => s.id);
      setUnlockedSheets(prev => {
        const next = { ...prev };
        newIds.forEach(id => { next[id] = 'edit'; });
        return next;
      });
      
      setSheets(prev => [...prev, ...parsedSheets]);
      setActiveSheetId(parsedSheets[0].id);
    } catch (error) {
      alert("Erro ao ler arquivo Excel. Verifique se o formato é válido.");
      console.error(error);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleSplitClasses = async () => {
    if (!activeSheet || !activeSheet.data || activeSheet.data.length === 0) return;

    setIsProcessing(true);
    try {
      const headers = activeSheet.data[0];
      if (!headers || headers.length === 0) throw new Error("Headers not found");
      
      const stringHeaders = headers.map(h => String(h || ""));

      // 1. Try AI Suggestion
      let candidateColumn = await suggestClassColumn(stringHeaders);
      
      // 2. Fallback to heuristic if AI returns null
      if (!candidateColumn) {
          candidateColumn = stringHeaders.find(h => 
             h && (
               h.toLowerCase().includes('turma') || 
               h.toLowerCase().includes('classe') || 
               h.toLowerCase().includes('serie')
             )
           ) || null;
      }

      // 3. Ask user to confirm with the suggestion as default
      const userInput = prompt(
          `Qual coluna deve ser usada para separar as turmas?`, 
          candidateColumn || ""
      );

      if (!userInput) {
        setIsProcessing(false);
        return;
      }
      
      const columnIndex = headers.findIndex(h => String(h || "").trim().toLowerCase() === userInput.trim().toLowerCase());

      if (columnIndex === -1) {
        alert("Coluna não encontrada.");
        setIsProcessing(false);
        return;
      }

      const newSheets = splitSheetByColumn(activeSheet, columnIndex);
      
      setUnlockedSheets(prev => {
        const next = { ...prev };
        newSheets.forEach(s => { next[s.id] = 'edit'; });
        return next;
      });

      setSheets(prev => [...prev, ...newSheets]);
      alert(`${newSheets.length} novas turmas separadas com sucesso!`);
      if (newSheets.length > 0) setActiveSheetId(newSheets[0].id);

    } catch (error) {
      console.error(error);
      alert("Erro ao separar turmas.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleSetupClassValidation = async () => {
    if (!activeSheet || !activeSheet.data || activeSheet.data.length === 0) return;

    setIsProcessing(true);
    try {
      const headers = activeSheet.data[0];
      if (!headers || headers.length === 0) {
          alert("Planilha vazia ou sem cabeçalhos.");
          setIsProcessing(false);
          return;
      }

      const stringHeaders = headers.map(h => String(h || ""));
      
      // 1. Try AI Suggestion
      let candidateColumn = await suggestClassColumn(stringHeaders);
      
      // 2. Fallback to heuristic
      if (!candidateColumn) {
          candidateColumn = stringHeaders.find(h => 
             h && (
               h.toLowerCase().includes('turma') || 
               h.toLowerCase().includes('classe') || 
               h.toLowerCase().includes('serie') || 
               h.toLowerCase().includes('ano')
             )
           ) || null;
      }

      // 3. Confirm Column with User
      const userInput = prompt(
          `Qual coluna de Turma você deseja validar?`, 
          candidateColumn || ""
      );

      if (!userInput) {
        setIsProcessing(false);
        return;
      }

      const columnIndex = headers.findIndex(h => String(h || "").trim().toLowerCase() === userInput.trim().toLowerCase());

      if (columnIndex === -1) {
        alert("Coluna não encontrada.");
        setIsProcessing(false);
        return;
      }

      const defaultClasses = "1A, 1B, 2A, 2B, 3A, 3B, 4A, 4B, 5A, 5B, 6A, 6B";
      const userClasses = prompt("Digite as turmas permitidas separadas por vírgula:", defaultClasses);
      
      if (!userClasses) {
          setIsProcessing(false);
          return;
      }

      const options = userClasses.split(',').map(s => s.trim()).filter(s => s.length > 0);

      const rule: ValidationRule = {
          id: generateUUID(),
          columnIndex: columnIndex,
          type: 'list',
          options: options,
          errorMessage: "Turma inválida. Selecione uma da lista."
      };

      // Remove existing rule for this column if any
      const existingRules = activeSheet.validationRules || [];
      const filteredRules = existingRules.filter(r => r.columnIndex !== columnIndex);

      const updatedSheet = {
          ...activeSheet,
          validationRules: [...filteredRules, rule]
      };

      setSheets(prev => prev.map(s => s.id === activeSheet.id ? updatedSheet : s));
      alert(`Validação configurada para a coluna '${String(headers[columnIndex])}'.`);

    } catch (error) {
      console.error(error);
      alert("Erro ao configurar validação de turmas.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCalculateAverages = () => {
    if (!activeSheet || !activeSheet.data || activeSheet.data.length === 0) return;

    const newData = activeSheet.data.map(row => row ? [...row] : []);
    if (newData.length === 0 || !newData[0]) return;

    const headers = newData[0].map(h => String(h || "").toLowerCase().trim());
    
    // Find columns
    const col1Idx = headers.findIndex(h => h.includes('nota 1') || h === 'p1' || h === 'n1');
    const col2Idx = headers.findIndex(h => h.includes('nota 2') || h === 'p2' || h === 'n2');

    if (col1Idx === -1 || col2Idx === -1) {
      alert("Não encontrei as colunas 'Nota 1' e 'Nota 2'.");
      return;
    }

    // Find or Create 'Média' column
    let mediaColIdx = headers.findIndex(h => h === 'média' || h === 'media');
    
    if (mediaColIdx === -1) {
      newData[0].push("Média");
      mediaColIdx = newData[0].length - 1;
    }

    let updatedCount = 0;

    // Calculate for all rows
    for (let i = 1; i < newData.length; i++) {
      const row = newData[i];
      if (!row) continue;
      
      // Ensure row has cell for media
      while (row.length <= mediaColIdx) {
        row.push("");
      }

      const val1Str = String(row[col1Idx] || "").replace(',', '.');
      const val2Str = String(row[col2Idx] || "").replace(',', '.');
      
      const val1 = parseFloat(val1Str);
      const val2 = parseFloat(val2Str);

      if (!isNaN(val1) && !isNaN(val2) && row[col1Idx] !== "" && row[col2Idx] !== "" && row[col1Idx] !== null && row[col2Idx] !== null) {
        const avg = ((val1 + val2) / 2).toFixed(1);
        const useComma = String(row[col1Idx]).includes(',') || String(row[col2Idx]).includes(',');
        row[mediaColIdx] = useComma ? avg.replace('.', ',') : avg;
        updatedCount++;
      }
    }

    // Auto-apply conditional formatting
    const failRule: ConditionalRule = {
        id: generateUUID(),
        columnIndex: mediaColIdx,
        condition: 'lt',
        value: 5,
        style: PRESET_STYLES[0] // Red
    };

    const passRule: ConditionalRule = {
        id: generateUUID(),
        columnIndex: mediaColIdx,
        condition: 'gte',
        value: 5,
        style: PRESET_STYLES[1] // Green
    };

    const existingRules = activeSheet.conditionalFormats || [];
    const otherRules = existingRules.filter(r => r.columnIndex !== mediaColIdx);

    const updatedSheet = { 
        ...activeSheet, 
        data: newData,
        conditionalFormats: [...otherRules, failRule, passRule]
    };
    
    setSheets(prev => prev.map(s => s.id === activeSheet.id ? updatedSheet : s));
    alert(`Média calculada para ${updatedCount} linhas. Formatação aplicada: Verde (>=5), Vermelho (<5).`);
  };

  // Helper: Convert column letter to index
  const getColIndex = (header: string): number => {
      let colIndex = 0;
      const cleanHeader = header.toUpperCase().trim();
      for (let i = 0; i < cleanHeader.length; i++) {
        colIndex = colIndex * 26 + (cleanHeader.charCodeAt(i) - 64);
      }
      return colIndex - 1; // 0-based
  };

  // Validation Logic
  const validateValue = (value: any, rule: ValidationRule): { valid: boolean; msg?: string } => {
    if (value === "" || value === null || value === undefined) return { valid: true };
    const strVal = String(value).trim();

    if (rule.type === 'number') {
       const num = Number(value);
       if (isNaN(num)) return { valid: false, msg: 'O valor deve ser um número.' };
       if (rule.min && num < Number(rule.min)) return { valid: false, msg: `O valor deve ser maior ou igual a ${rule.min}.` };
       if (rule.max && num > Number(rule.max)) return { valid: false, msg: `O valor deve ser menor ou igual a ${rule.max}.` };
    }
    
    if (rule.type === 'date') {
       const date = new Date(strVal);
       if (isNaN(date.getTime())) return { valid: false, msg: 'Data inválida.' };
       // Simple string comparison for dates YYYY-MM-DD
       if (rule.min && strVal < rule.min) return { valid: false, msg: `A data deve ser posterior a ${rule.min}.` };
       if (rule.max && strVal > rule.max) return { valid: false, msg: `A data deve ser anterior a ${rule.max}.` };
    }

    if (rule.type === 'list' && rule.options) {
        if (!rule.options.includes(strVal)) return { valid: false, msg: 'Valor não permitido na lista.' };
    }

    if (rule.type === 'email') {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(strVal)) return { valid: false, msg: 'Endereço de email inválido.' };
    }

    return { valid: true };
  };

  const updateCell = (r: number, c: number, value: any) => {
    if (!activeSheet || !canEdit) return;

    // Check Validation Rules
    const validationRule = activeSheet.validationRules?.find(rule => rule.columnIndex === c);
    if (validationRule) {
        const check = validateValue(value, validationRule);
        if (!check.valid) {
            alert(validationRule.errorMessage || check.msg || "Valor inválido.");
            return; // Cancel update
        }
    }
    
    const newData = [...activeSheet.data];
    if (!newData[r]) newData[r] = [];
    const newRow = [...newData[r]];
    newRow[c] = value;
    newData[r] = newRow;

    // Reactive Average Calculation
    const headers = newData[0] ? newData[0].map(h => String(h || "").toLowerCase().trim()) : [];
    const col1Idx = headers.findIndex(h => h.includes('nota 1') || h === 'p1' || h === 'n1');
    const col2Idx = headers.findIndex(h => h.includes('nota 2') || h === 'p2' || h === 'n2');
    const mediaIdx = headers.findIndex(h => h === 'média' || h === 'media');

    if (mediaIdx !== -1 && col1Idx !== -1 && col2Idx !== -1) {
        if (c === col1Idx || c === col2Idx) {
            const val1Str = String(newRow[col1Idx] || "").replace(',', '.');
            const val2Str = String(newRow[col2Idx] || "").replace(',', '.');
            
            const val1 = parseFloat(val1Str);
            const val2 = parseFloat(val2Str);
            
            if (!isNaN(val1) && !isNaN(val2) && newRow[col1Idx] !== "" && newRow[col2Idx] !== "") {
                const avg = ((val1 + val2) / 2).toFixed(1);
                const useComma = String(newRow[col1Idx]).includes(',') || String(newRow[col2Idx]).includes(',');
                
                while (newRow.length <= mediaIdx) newRow.push("");
                newRow[mediaIdx] = useComma ? avg.replace('.', ',') : avg;
            }
        }
    }

    const updatedSheet = { ...activeSheet, data: newData };
    setSheets(prev => prev.map(s => s.id === activeSheet.id ? updatedSheet : s));
  };

  const handleAddRule = () => {
    if (!activeSheet || !canEdit) return;

    const colIndex = getColIndex(newRule.colHeader);
    if (colIndex < 0) {
      alert("Coluna inválida");
      return;
    }

    const rule: ConditionalRule = {
      id: generateUUID(),
      columnIndex: colIndex,
      condition: newRule.condition,
      value: newRule.value,
      style: PRESET_STYLES[newRule.styleIndex]
    };

    const updatedSheet = {
      ...activeSheet,
      conditionalFormats: [...(activeSheet.conditionalFormats || []), rule]
    };

    setSheets(prev => prev.map(s => s.id === activeSheet.id ? updatedSheet : s));
    setShowFormatModal(false);
  };

  const handleAddValidation = () => {
    if (!activeSheet || !canEdit) return;
    const colIndex = getColIndex(newValidation.colHeader);
    if (colIndex < 0) {
        alert("Coluna inválida");
        return;
    }

    const optionsArray = newValidation.type === 'list' 
        ? newValidation.options.split(',').map(s => s.trim()).filter(s => s.length > 0) 
        : undefined;

    const rule: ValidationRule = {
        id: generateUUID(),
        columnIndex: colIndex,
        type: newValidation.type,
        min: newValidation.min,
        max: newValidation.max,
        options: optionsArray,
        errorMessage: newValidation.errorMessage
    };

    // Remove existing rule for this column if any (simulate overwrite)
    const existingRules = activeSheet.validationRules || [];
    const filteredRules = existingRules.filter(r => r.columnIndex !== colIndex);

    const updatedSheet = {
        ...activeSheet,
        validationRules: [...filteredRules, rule]
    };

    setSheets(prev => prev.map(s => s.id === activeSheet.id ? updatedSheet : s));
    setShowValidationModal(false);
  };

  // Tab Context Menu Handlers
  const handleContextMenu = (e: React.MouseEvent, sheetId: string) => {
    e.preventDefault();
    setContextMenu({ x: e.clientX, y: e.clientY - 120, sheetId }); // Adjust Y to show above cursor/tab
  };

  const handleRenameSheet = () => {
    if (!contextMenu) return;
    const sheet = sheets.find(s => s.id === contextMenu.sheetId);
    if (!sheet) return;

    // Permissions check: user must have edit access to rename
    if (unlockedSheets[sheet.id] !== 'edit' && (sheet.editCode || sheet.viewCode)) {
        alert("Apenas editores podem renomear planilhas.");
        return;
    }

    const newName = prompt("Renomear planilha:", sheet.name);
    if (newName && newName.trim()) {
      setSheets(prev => prev.map(s => s.id === sheet.id ? { ...s, name: newName.trim() } : s));
    }
    setContextMenu(null);
  };

  const handleDuplicateSheet = () => {
    if (!contextMenu) return;
    const sheet = sheets.find(s => s.id === contextMenu.sheetId);
    if (!sheet) return;

    const newSheet: Sheet = {
      ...sheet,
      id: generateUUID(),
      name: `${sheet.name} (Cópia)`,
      data: JSON.parse(JSON.stringify(sheet.data)), // Deep copy data
      conditionalFormats: sheet.conditionalFormats ? [...sheet.conditionalFormats] : [],
      validationRules: sheet.validationRules ? [...sheet.validationRules] : [],
      // Copy security settings
      editCode: sheet.editCode,
      viewCode: sheet.viewCode,
      accessCode: sheet.accessCode, 
      accessCodeExpiration: sheet.accessCodeExpiration,
      isShared: sheet.isShared
    };

    setUnlockedSheets(prev => {
        const next = { ...prev };
        // Inherit access level for the copy
        if(unlockedSheets[sheet.id]) next[newSheet.id] = unlockedSheets[sheet.id];
        return next;
    });

    setSheets(prev => [...prev, ...[newSheet]]);
    setActiveSheetId(newSheet.id);
    setContextMenu(null);
  };

  const handleDeleteSheet = () => {
    if (!contextMenu) return;
    const sheet = sheets.find(s => s.id === contextMenu.sheetId);
    
    // Permission check
    if (sheet && unlockedSheets[sheet.id] !== 'edit' && (sheet.editCode || sheet.viewCode)) {
        alert("Apenas editores podem excluir planilhas.");
        setContextMenu(null);
        return;
    }

    if (sheets.length <= 1) {
      alert("Não é possível excluir a única planilha existente.");
      setContextMenu(null);
      return;
    }

    if (confirm("Tem certeza que deseja excluir esta planilha?")) {
      const newSheets = sheets.filter(s => s.id !== contextMenu.sheetId);
      setSheets(newSheets);
      
      if (activeSheetId === contextMenu.sheetId) {
        setActiveSheetId(newSheets[0].id);
      }
    }
    setContextMenu(null);
  };

  // Zoom Handlers
  const handleZoomIn = () => setZoomLevel(prev => Math.min(prev + 0.1, 2.0));
  const handleZoomOut = () => setZoomLevel(prev => Math.max(prev - 0.1, 0.5));

  // Access Control Logic
  const handleCreateAccessCode = () => {
    if (!activeSheet) return;
    
    // Calculate expiration
    const now = Date.now();
    let multiplier = 1000 * 60 * 60; // 1 hour default
    
    switch(expirationUnit) {
        case 'hours': multiplier = 1000 * 60 * 60; break;
        case 'days': multiplier = 1000 * 60 * 60 * 24; break;
        case 'weeks': multiplier = 1000 * 60 * 60 * 24 * 7; break;
        case 'months': multiplier = 1000 * 60 * 60 * 24 * 30; break;
    }
    
    const expiresAt = now + (expirationValue * multiplier);

    const editCode = Math.random().toString(36).slice(-6).toUpperCase();
    const viewCode = Math.random().toString(36).slice(-6).toUpperCase();

    const updatedSheet: Sheet = { 
        ...activeSheet, 
        editCode: editCode, 
        viewCode: viewCode, 
        accessCodeExpiration: expiresAt, 
        isShared: true 
    };

    setSheets(prev => prev.map(s => s.id === activeSheet.id ? updatedSheet : s));
    
    // Unlock for creator as Editor
    setUnlockedSheets(prev => ({ ...prev, [activeSheet.id]: 'edit' }));
  };

  const handleRemoveAccessCode = () => {
    if (!activeSheet) return;
    const updatedSheet = { 
        ...activeSheet, 
        editCode: undefined, 
        viewCode: undefined,
        accessCode: undefined,
        accessCodeExpiration: undefined, 
        isShared: false 
    };
    setSheets(prev => prev.map(s => s.id === activeSheet.id ? updatedSheet : s));
  };

  const handleUnlockSheet = () => {
     if (!activeSheet) return;
     
     // Check Expiration
     if (activeSheet.accessCodeExpiration && Date.now() > activeSheet.accessCodeExpiration) {
         setAccessError("Os códigos expiraram. O administrador deve gerar novos.");
         return;
     }

     const input = accessCodeInput.trim().toUpperCase();

     if (input === activeSheet.editCode) {
        setUnlockedSheets(prev => ({ ...prev, [activeSheet.id]: 'edit' }));
        setAccessCodeInput("");
        setAccessError(null);
     } else if (input === activeSheet.viewCode) {
        setUnlockedSheets(prev => ({ ...prev, [activeSheet.id]: 'view' }));
        setAccessCodeInput("");
        setAccessError(null);
     } else if (input === activeSheet.accessCode) { 
        // Legacy fallback
        setUnlockedSheets(prev => ({ ...prev, [activeSheet.id]: 'edit' }));
        setAccessCodeInput("");
        setAccessError(null);
     } else {
        setAccessError("Código incorreto.");
     }
  };

  const handleSimulateLock = () => {
      if (!activeSheet) return;
      setUnlockedSheets(prev => {
          const next = { ...prev };
          delete next[activeSheet.id];
          return next;
      });
      setShowShareModal(false);
  };

  // Helper to display remaining time or expiration date
  const formatExpiration = (timestamp: number | undefined) => {
      if (!timestamp) return "";
      return new Date(timestamp).toLocaleString();
  };

  const handleCopyLink = () => {
      if (!activeSheet) return;
      
      let origin = "";
      try {
          origin = window.location.origin;
          if (!origin || origin === 'null') origin = "";
      } catch(e) {
          origin = "";
      }
      
      const link = origin ? `${origin}?pauta=${activeSheet.id}` : `?pauta=${activeSheet.id}`;
      
      if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(link).then(() => {
              setLinkCopySuccess(true);
              setTimeout(() => setLinkCopySuccess(false), 2000);
          }).catch(err => {
              console.error("Clipboard error", err);
              prompt("Copie o link:", link);
          });
      } else {
          prompt("Copie o link:", link);
      }
  };
  
  const handleCopyCode = (code: string | undefined, type: 'edit' | 'view') => {
      if (!code) return;
      if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(code).then(() => {
              if (type === 'edit') {
                  setCopyEditSuccess(true);
                  setTimeout(() => setCopyEditSuccess(false), 2000);
              } else {
                  setCopyViewSuccess(true);
                  setTimeout(() => setCopyViewSuccess(false), 2000);
              }
          }).catch(() => {});
      }
  };

  const isExpired = activeSheet?.accessCodeExpiration ? Date.now() > activeSheet.accessCodeExpiration : false;

  return (
    <div className="flex flex-col h-screen bg-gray-100 dark:bg-gray-900 text-gray-900 dark:text-gray-100 font-sans overflow-hidden transition-colors duration-200">
      {/* Header */}
      <header className="bg-white dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700 h-16 flex items-center justify-between px-4 shadow-sm z-20 transition-colors duration-200">
        <div className="flex items-center gap-3">
          <div className="bg-emerald-600 p-2 rounded-lg text-white shadow-lg shadow-emerald-200 dark:shadow-none">
            <div className="relative">
              <BookOpen size={24} />
              <Pencil size={14} className="absolute -bottom-1 -right-1 bg-white text-emerald-600 rounded-full border-2 border-white dark:border-gray-800" />
            </div>
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-gray-800 dark:text-white">Educa-Lukunde</h1>
            <p className="text-xs text-gray-500 dark:text-gray-400 font-medium">Gestão Inteligente de Pautas</p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {isProcessing && <span className="text-sm text-emerald-600 dark:text-emerald-400 animate-pulse font-medium mr-4">Processando...</span>}
          
          <label className={`flex items-center gap-2 px-3 py-2 bg-gray-50 dark:bg-gray-700 hover:bg-gray-100 dark:hover:bg-gray-600 text-gray-700 dark:text-gray-200 rounded-md cursor-pointer border border-gray-200 dark:border-gray-600 transition-colors text-sm font-medium ${isReadOnly ? 'opacity-50 pointer-events-none' : ''}`}>
            <Upload size={16} />
            <span className="hidden sm:inline">Carregar Excel</span>
            <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" disabled={isReadOnly} />
          </label>

          <button 
            onClick={handleSplitClasses}
            disabled={!activeSheet || isReadOnly}
            className="flex items-center gap-2 px-3 py-2 bg-emerald-50 dark:bg-emerald-900/30 hover:bg-emerald-100 dark:hover:bg-emerald-900/50 text-emerald-700 dark:text-emerald-400 rounded-md border border-emerald-200 dark:border-emerald-800 transition-colors text-sm font-medium disabled:opacity-50 disabled:cursor-not-allowed"
            title="Separar pauta por classes/turmas"
          >
            <Split size={16} />
            <span className="hidden sm:inline">Separar Turmas</span>
          </button>

          <button 
            onClick={handleSetupClassValidation}
            disabled={!activeSheet || isSheetLocked || isReadOnly}
            className="flex items-center gap-2 px-3 py-2 bg-emerald-50 dark:bg-emerald-900/30 hover:bg-emerald-100 dark:hover:bg-emerald-900/50 text-emerald-700 dark:text-emerald-400 rounded-md border border-emerald-200 dark:border-emerald-800 transition-colors text-sm font-medium disabled:opacity-50 disabled:cursor-not-allowed"
            title="Restringir coluna de Turma a valores específicos"
          >
            <ListChecks size={16} />
            <span className="hidden sm:inline">Validar Turmas</span>
          </button>

          <button 
            onClick={handleCalculateAverages}
            disabled={!activeSheet || isSheetLocked || isReadOnly}
            className="flex items-center gap-2 px-3 py-2 bg-emerald-50 dark:bg-emerald-900/30 hover:bg-emerald-100 dark:hover:bg-emerald-900/50 text-emerald-700 dark:text-emerald-400 rounded-md border border-emerald-200 dark:border-emerald-800 transition-colors text-sm font-medium disabled:opacity-50 disabled:cursor-not-allowed"
            title="Calcular Médias (Nota 1 + Nota 2) / 2"
          >
            <Calculator size={16} />
            <span className="hidden sm:inline">Calc. Média</span>
          </button>

          <button 
             onClick={() => setShowValidationModal(true)}
             disabled={!activeSheet || isSheetLocked || isReadOnly}
             className="p-2 text-gray-500 dark:text-gray-400 hover:text-emerald-600 dark:hover:text-emerald-400 hover:bg-gray-50 dark:hover:bg-gray-700 rounded-md transition-colors disabled:opacity-30"
             title="Validação de Dados"
          >
             <ShieldCheck size={20} />
          </button>

          <button 
             onClick={() => setShowFormatModal(true)}
             disabled={!activeSheet || isSheetLocked || isReadOnly}
             className="p-2 text-gray-500 dark:text-gray-400 hover:text-emerald-600 dark:hover:text-emerald-400 hover:bg-gray-50 dark:hover:bg-gray-700 rounded-md transition-colors disabled:opacity-30"
             title="Formatação Condicional"
          >
             <Palette size={20} />
          </button>

          {!isReadOnly && (
              <button 
                onClick={() => setShowShareModal(true)}
                disabled={!activeSheet}
                className={`flex items-center gap-2 px-3 py-2 rounded-md border transition-colors text-sm font-medium disabled:opacity-50
                    ${activeSheet?.isShared 
                        ? 'bg-blue-50 dark:bg-blue-900/30 text-blue-700 dark:text-blue-400 border-blue-200 dark:border-blue-800' 
                        : 'bg-gray-50 dark:bg-gray-700 text-gray-700 dark:text-gray-200 border-gray-200 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-600'
                    }
                `}
                title="Partilhar Pauta com Código"
              >
                <Share2 size={16} />
                <span className="hidden sm:inline">Partilhar Pauta</span>
              </button>
          )}

          <button 
            onClick={handleExport}
            disabled={!activeSheet || isSheetLocked}
            className="p-2 text-gray-500 dark:text-gray-400 hover:text-emerald-600 dark:hover:text-emerald-400 hover:bg-gray-50 dark:hover:bg-gray-700 rounded-md transition-colors disabled:opacity-30"
            title="Exportar atual (Ctrl+S)"
          >
            <Download size={20} />
          </button>

          <button 
            onClick={handleExportAll}
            disabled={sheets.length === 0}
            className="p-2 text-gray-500 dark:text-gray-400 hover:text-emerald-600 dark:hover:text-emerald-400 hover:bg-gray-50 dark:hover:bg-gray-700 rounded-md transition-colors"
            title="Exportar todas as planilhas"
          >
            <SaveAll size={20} />
          </button>
          
          <div className="h-6 w-px bg-gray-200 dark:bg-gray-700 mx-1"></div>

           <button 
            onClick={toggleTheme}
            className="p-2 text-gray-500 dark:text-gray-400 hover:text-amber-500 dark:hover:text-yellow-400 hover:bg-gray-50 dark:hover:bg-gray-700 rounded-md transition-colors"
            title={theme === 'light' ? "Modo Escuro" : "Modo Claro"}
          >
            {theme === 'light' ? <Moon size={20} /> : <Sun size={20} />}
          </button>

          <button 
            onClick={() => setShowAI(!showAI)}
            className={`p-2 rounded-md transition-all ${
                showAI 
                ? 'bg-emerald-100 dark:bg-emerald-900 text-emerald-700 dark:text-emerald-400 shadow-inner' 
                : 'text-gray-500 dark:text-gray-400 hover:bg-gray-50 dark:hover:bg-gray-700'
            }`}
            title="Assistente IA"
          >
            <MessageSquare size={20} />
          </button>
          
          <div className="h-8 w-8 rounded-full bg-gradient-to-tr from-emerald-500 to-teal-400 flex items-center justify-center text-white font-bold text-xs ml-2 cursor-pointer shadow-md border-2 border-white dark:border-gray-800">
            EL
          </div>
        </div>
      </header>

      {/* Main Content */}
      <div className="flex flex-1 overflow-hidden relative">
        <div className="flex-1 flex flex-col min-w-0">
            {/* Toolbar / Formula Bar Placeholder */}
            <div className="h-10 border-b border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-800 flex items-center px-4 gap-4 text-sm text-gray-500 dark:text-gray-400 justify-between transition-colors duration-200">
               <div className="flex items-center gap-4 flex-1">
                  <span className="font-mono bg-gray-100 dark:bg-gray-700 px-2 py-0.5 rounded text-xs text-gray-600 dark:text-gray-300">fx</span>
                  <div className="h-4 w-px bg-gray-300 dark:bg-gray-600"></div>
                  <span className="italic text-gray-400 dark:text-gray-500 text-xs flex items-center gap-2">
                     {isSheetLocked ? 'Pauta Bloqueada - Requer código de acesso' : 
                      isReadOnly ? 
                        <span className="flex items-center gap-1 text-blue-600 dark:text-blue-400 font-medium"><Eye size={12}/> Modo de Visualização (Somente Leitura)</span> : 
                        'Use as setas para navegar, Enter para editar.'}
                  </span>
               </div>
               
               {/* Zoom Controls */}
               <div className="flex items-center gap-2 border-l border-gray-200 dark:border-gray-700 pl-3 ml-2">
                 <button 
                    onClick={handleZoomOut} 
                    className="p-1 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200"
                    title="Diminuir Zoom"
                 >
                    <ZoomOut size={16} />
                 </button>
                 <span className="text-xs font-mono w-10 text-center select-none text-gray-600 dark:text-gray-300">{Math.round(zoomLevel * 100)}%</span>
                 <button 
                    onClick={handleZoomIn} 
                    className="p-1 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200"
                    title="Aumentar Zoom"
                 >
                    <ZoomIn size={16} />
                 </button>
               </div>
            </div>

            {/* Grid or Lock Screen */}
            {isSheetLocked ? (
                <div className="flex-1 flex flex-col items-center justify-center bg-gray-50 dark:bg-gray-900 text-gray-600 dark:text-gray-300">
                    <div className="bg-white dark:bg-gray-800 p-8 rounded-2xl shadow-xl border border-gray-100 dark:border-gray-700 text-center w-96 transition-colors">
                        <div className="w-16 h-16 bg-red-100 dark:bg-red-900/30 text-red-500 dark:text-red-400 rounded-full flex items-center justify-center mx-auto mb-4">
                            <Lock size={32} />
                        </div>
                        <h2 className="text-xl font-bold text-gray-800 dark:text-white mb-2">Acesso Restrito</h2>
                        <p className="text-sm text-gray-500 dark:text-gray-400 mb-6">
                            {isExpired 
                                ? "O código de acesso para esta pauta expirou." 
                                : "Esta pauta está protegida. Digite o código de editor ou visualizador."
                            }
                        </p>
                        
                        {!isExpired && (
                            <div className="mb-4 text-left">
                                <label className="text-xs font-semibold uppercase text-gray-400 dark:text-gray-500 mb-1 block">Código de Acesso</label>
                                <input 
                                    type="text" 
                                    value={accessCodeInput}
                                    onChange={(e) => {
                                        setAccessCodeInput(e.target.value.toUpperCase());
                                        setAccessError(null);
                                    }}
                                    onKeyDown={(e) => e.key === 'Enter' && handleUnlockSheet()}
                                    className={`w-full border-2 rounded-lg p-3 text-center text-lg tracking-widest font-mono uppercase focus:outline-none focus:ring-2 focus:ring-emerald-200 dark:focus:ring-emerald-800 bg-white dark:bg-gray-700 dark:text-white ${accessError ? 'border-red-300 dark:border-red-700 bg-red-50 dark:bg-red-900/20' : 'border-gray-200 dark:border-gray-600'}`}
                                    placeholder="XXXXXX"
                                />
                                {accessError && <p className="text-xs text-red-500 dark:text-red-400 mt-1 text-center font-medium">{accessError}</p>}
                            </div>
                        )}

                        {isExpired ? (
                            <div className="p-3 bg-red-50 dark:bg-red-900/20 border border-red-100 dark:border-red-800 rounded-lg text-red-700 dark:text-red-300 text-sm">
                                Solicite um novo código ao administrador.
                            </div>
                        ) : (
                            <button 
                                onClick={handleUnlockSheet}
                                className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-medium py-3 rounded-lg transition-all shadow-md hover:shadow-lg"
                            >
                                Desbloquear Pauta
                            </button>
                        )}
                    </div>
                </div>
            ) : (
                <Spreadsheet 
                  data={activeSheet?.data || []} 
                  onCellChange={updateCell}
                  rules={activeSheet?.conditionalFormats}
                  validationRules={activeSheet?.validationRules}
                  zoom={zoomLevel}
                  readOnly={isReadOnly}
                />
            )}

            {/* Bottom Tab Bar */}
            <div className="h-10 bg-white dark:bg-gray-800 border-t border-gray-200 dark:border-gray-700 flex items-center px-2 gap-1 overflow-x-auto relative transition-colors duration-200">
              <button 
                className={`p-1.5 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-500 dark:text-gray-400 ${isReadOnly ? 'opacity-50 pointer-events-none' : ''}`}
                onClick={() => {
                   if (isReadOnly) return;
                   const newId = generateUUID();
                   setSheets([...sheets, { id: newId, name: `Nova Planilha ${sheets.length + 1}`, data: [[]], conditionalFormats: [] }]);
                   setActiveSheetId(newId);
                   setUnlockedSheets(prev => ({ ...prev, [newId]: 'edit' }));
                }}
              >
                <Plus size={16} />
              </button>
              
              {sheets.map((sheet, index) => {
                // Safety check for undefined sheets in list
                if (!sheet) return null;
                const sheetAccess = unlockedSheets[sheet.id];
                const sheetLocked = (sheet.editCode || sheet.viewCode) && !sheetAccess;
                
                return (
                  <button
                    key={sheet.id || index}
                    draggable={!isReadOnly}
                    onDragStart={(e) => !isReadOnly && handleDragStart(e, index)}
                    onDragOver={(e) => !isReadOnly && handleDragOver(e, index)}
                    onDragEnd={!isReadOnly ? handleDragEnd : undefined}
                    onClick={() => setActiveSheetId(sheet.id)}
                    onContextMenu={(e) => handleContextMenu(e, sheet.id)}
                    className={`
                      px-4 py-1.5 text-xs font-medium rounded-t-md border-x relative top-[1px] min-w-[100px] truncate transition-colors flex items-center gap-2 select-none cursor-pointer
                      ${activeSheetId === sheet.id 
                        ? 'bg-white dark:bg-gray-800 border-gray-300 dark:border-gray-600 text-emerald-700 dark:text-emerald-400 border-b-white dark:border-b-gray-800 border-t-2 border-t-emerald-500 z-10' 
                        : 'bg-gray-50 dark:bg-gray-700 border-transparent text-gray-500 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-600 hover:text-gray-700 dark:hover:text-gray-200 border-t border-transparent'
                      }
                      ${draggedSheetIndex === index ? 'opacity-50' : ''}
                    `}
                  >
                    {sheet.editCode || sheet.viewCode ? (
                        sheetAccess === 'edit' ? <Unlock size={10} className="text-emerald-500"/> : 
                        sheetAccess === 'view' ? <Eye size={10} className="text-blue-500"/> :
                        <Lock size={10} className="text-red-400"/>
                    ) : (
                        <FileSpreadsheet size={12} className={activeSheetId === sheet.id ? "text-emerald-500" : "text-gray-400"} />
                    )}
                    {sheet.name}
                  </button>
                );
              })}
            </div>
        </div>

        {/* AI Sidebar */}
        {showAI && (
          <AIAssistant 
            activeSheet={activeSheet} 
            onClose={() => setShowAI(false)} 
          />
        )}

        {/* Share Modal */}
        {showShareModal && activeSheet && (
           <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/20 backdrop-blur-sm">
             <div className="bg-white dark:bg-gray-800 rounded-lg shadow-2xl w-[550px] p-0 border border-gray-200 dark:border-gray-700 overflow-hidden" onClick={e => e.stopPropagation()}>
               <div className="bg-emerald-600 dark:bg-emerald-800 p-4 flex justify-between items-center text-white">
                 <h3 className="font-bold flex items-center gap-2 text-lg">
                   <Share2 size={20} />
                   Partilhar Pauta
                 </h3>
                 <button onClick={() => setShowShareModal(false)} className="text-emerald-100 hover:text-white">
                   <X size={20} />
                 </button>
               </div>

               <div className="p-6 space-y-6 text-gray-800 dark:text-gray-200 max-h-[80vh] overflow-y-auto">
                 <div>
                    <h4 className="text-sm font-semibold text-gray-800 dark:text-gray-200 mb-2">Estado da Partilha</h4>
                    {(activeSheet.editCode || activeSheet.viewCode) ? (
                         isExpired ? (
                            <div className="flex items-center gap-2 text-red-700 dark:text-red-400 bg-red-50 dark:bg-red-900/30 p-3 rounded-lg border border-red-100 dark:border-red-800">
                               <Clock size={18} />
                               <span className="text-sm font-medium">Os códigos atuais expiraram em {formatExpiration(activeSheet.accessCodeExpiration)}.</span>
                            </div>
                         ) : (
                            <div className="flex items-center gap-2 text-emerald-700 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-900/30 p-3 rounded-lg border border-emerald-100 dark:border-emerald-800">
                               <Check size={18} />
                               <span className="text-sm font-medium">Ativo até {formatExpiration(activeSheet.accessCodeExpiration)}</span>
                            </div>
                         )
                    ) : (
                        <div className="flex items-center gap-2 text-gray-500 dark:text-gray-400 bg-gray-50 dark:bg-gray-700 p-3 rounded-lg border border-gray-200 dark:border-gray-600">
                           <Lock size={18} />
                           <span className="text-sm">Apenas você tem acesso.</span>
                        </div>
                    )}
                 </div>

                 {activeSheet.editCode && !isExpired ? (
                     <>
                        <div className="grid grid-cols-2 gap-4">
                            {/* Editor Access */}
                            <div className="space-y-2 col-span-2 sm:col-span-1">
                               <label className="text-xs font-semibold uppercase text-red-500 dark:text-red-400 flex items-center gap-1">
                                    <Edit size={12} /> Acesso de Editor (Professor)
                               </label>
                               <div className="flex items-center gap-2">
                                  <div className="flex-1 bg-red-50 dark:bg-red-900/20 border-2 border-dashed border-red-200 dark:border-red-800 rounded-lg p-2 text-center text-xl font-mono tracking-widest text-red-700 dark:text-red-200 select-all">
                                    {activeSheet.editCode}
                                  </div>
                                  <button 
                                    onClick={() => handleCopyCode(activeSheet.editCode, 'edit')}
                                    className={`p-2 border rounded-lg transition-colors ${
                                        copyEditSuccess 
                                        ? 'bg-emerald-100 text-emerald-700 border-emerald-300' 
                                        : 'bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 hover:bg-gray-50'
                                    }`}
                                  >
                                    {copyEditSuccess ? <Check size={16} /> : <Copy size={16} />}
                                  </button>
                               </div>
                               <p className="text-[10px] text-gray-400">Pode editar tudo.</p>
                            </div>

                            {/* Viewer Access */}
                            <div className="space-y-2 col-span-2 sm:col-span-1">
                               <label className="text-xs font-semibold uppercase text-blue-500 dark:text-blue-400 flex items-center gap-1">
                                    <Eye size={12} /> Acesso de Visualizador
                               </label>
                               <div className="flex items-center gap-2">
                                  <div className="flex-1 bg-blue-50 dark:bg-blue-900/20 border-2 border-dashed border-blue-200 dark:border-blue-800 rounded-lg p-2 text-center text-xl font-mono tracking-widest text-blue-700 dark:text-blue-200 select-all">
                                    {activeSheet.viewCode}
                                  </div>
                                  <button 
                                    onClick={() => handleCopyCode(activeSheet.viewCode, 'view')}
                                    className={`p-2 border rounded-lg transition-colors ${
                                        copyViewSuccess 
                                        ? 'bg-emerald-100 text-emerald-700 border-emerald-300' 
                                        : 'bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 hover:bg-gray-50'
                                    }`}
                                  >
                                    {copyViewSuccess ? <Check size={16} /> : <Copy size={16} />}
                                  </button>
                               </div>
                               <p className="text-[10px] text-gray-400">Apenas leitura.</p>
                            </div>
                        </div>

                        <div className="space-y-2">
                           <label className="text-xs font-semibold uppercase text-gray-400 dark:text-gray-500">Link da Pauta</label>
                           <div className="flex items-center gap-2 bg-gray-50 dark:bg-gray-700 p-2 rounded border border-gray-200 dark:border-gray-600">
                              <LinkIcon size={16} className="text-gray-400 shrink-0" />
                              <input 
                                type="text" 
                                readOnly 
                                value={`${window.location.origin}?pauta=${activeSheet.id}`} 
                                className="text-sm text-gray-600 dark:text-gray-300 flex-1 bg-transparent outline-none truncate"
                                onClick={(e) => e.currentTarget.select()}
                              />
                              <button 
                                onClick={handleCopyLink}
                                className={`text-xs font-bold hover:underline transition-colors ${linkCopySuccess ? 'text-emerald-500 dark:text-emerald-400' : 'text-emerald-600 dark:text-emerald-400'}`}
                              >
                                {linkCopySuccess ? "COPIADO!" : "COPIAR"}
                              </button>
                           </div>
                           <p className="text-[10px] text-gray-400">Nota: O link exige um dos códigos acima para abrir.</p>
                        </div>

                        <div className="border-t border-gray-100 dark:border-gray-700 pt-4 flex gap-3">
                           <button 
                              onClick={handleRemoveAccessCode}
                              className="flex-1 py-2 text-sm text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 rounded border border-transparent hover:border-red-100 dark:hover:border-red-900/30 transition-colors"
                           >
                              Revogar Todos os Acessos
                           </button>
                           <button 
                              onClick={handleSimulateLock}
                              className="flex-1 py-2 text-sm bg-gray-800 dark:bg-gray-600 text-white rounded hover:bg-gray-900 dark:hover:bg-gray-500 transition-colors flex items-center justify-center gap-2"
                              title="Bloquear minha visualização para testar"
                           >
                              <Lock size={14} /> Testar Bloqueio
                           </button>
                        </div>
                     </>
                 ) : (
                     <div className="text-center py-4">
                        <p className="text-sm text-gray-600 dark:text-gray-400 mb-4">Gere códigos distintos para editores e visualizadores.</p>
                        
                        <div className="flex gap-2 mb-4 justify-center items-end">
                            <div className="text-left w-24">
                                <label className="text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1 block">Duração</label>
                                <input 
                                    type="number" 
                                    min="1"
                                    value={expirationValue}
                                    onChange={(e) => setExpirationValue(Math.max(1, parseInt(e.target.value) || 1))}
                                    className="w-full border border-gray-300 dark:border-gray-600 rounded p-2 text-sm text-center bg-white dark:bg-gray-700"
                                />
                            </div>
                            <div className="text-left w-32">
                                <label className="text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1 block">Unidade</label>
                                <select 
                                    value={expirationUnit}
                                    onChange={(e) => setExpirationUnit(e.target.value as any)}
                                    className="w-full border border-gray-300 dark:border-gray-600 rounded p-2 text-sm bg-white dark:bg-gray-700"
                                >
                                    <option value="hours">Horas</option>
                                    <option value="days">Dias</option>
                                    <option value="weeks">Semanas</option>
                                    <option value="months">Meses</option>
                                </select>
                            </div>
                        </div>

                        <button 
                            onClick={handleCreateAccessCode}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-2 rounded-lg font-medium transition-colors shadow-lg shadow-emerald-200 dark:shadow-none w-full"
                        >
                            {isExpired ? "Gerar Novos Códigos" : "Gerar Códigos de Acesso"}
                        </button>
                     </div>
                 )}
               </div>
             </div>
           </div>
        )}

        {/* Tab Context Menu */}
        {contextMenu && (
          <div 
            className="absolute z-50 bg-white dark:bg-gray-800 rounded-lg shadow-xl border border-gray-200 dark:border-gray-700 w-48 py-1"
            style={{ top: contextMenu.y, left: contextMenu.x }}
            onClick={(e) => e.stopPropagation()}
          >
             <button onClick={handleRenameSheet} className={`w-full text-left px-4 py-2 text-sm flex items-center gap-2 ${isReadOnly ? 'text-gray-400 cursor-not-allowed' : 'hover:bg-gray-100 dark:hover:bg-gray-700 text-gray-700 dark:text-gray-200'}`} disabled={isReadOnly}>
                <Edit size={14} /> Renomear
             </button>
             <button onClick={handleDuplicateSheet} className={`w-full text-left px-4 py-2 text-sm flex items-center gap-2 ${isReadOnly ? 'text-gray-400 cursor-not-allowed' : 'hover:bg-gray-100 dark:hover:bg-gray-700 text-gray-700 dark:text-gray-200'}`} disabled={isReadOnly}>
                <Copy size={14} /> Duplicar
             </button>
             <div className="border-t border-gray-100 dark:border-gray-700 my-1"></div>
             <button onClick={handleDeleteSheet} className={`w-full text-left px-4 py-2 text-sm flex items-center gap-2 ${isReadOnly ? 'text-gray-400 cursor-not-allowed' : 'hover:bg-red-50 dark:hover:bg-red-900/20 text-red-600 dark:text-red-400'}`} disabled={isReadOnly}>
                <Trash2 size={14} /> Excluir
             </button>
          </div>
        )}
      </div>
    </div>
  );
};

export default App;
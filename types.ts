
export type CellValue = string | number | boolean | null;

export type SheetData = CellValue[][];

export type ConditionType = 'gt' | 'lt' | 'eq' | 'gte' | 'lte' | 'contains';

export interface ConditionalStyle {
  backgroundColor: string;
  color: string;
  name: string; // e.g., "Red", "Green"
}

export interface ConditionalRule {
  id: string;
  columnIndex: number;
  condition: ConditionType;
  value: string | number;
  style: ConditionalStyle;
}

export type ValidationType = 'number' | 'text' | 'date' | 'list' | 'email';

export interface ValidationRule {
  id: string;
  columnIndex: number;
  type: ValidationType;
  min?: string;
  max?: string;
  options?: string[]; // For 'list' type
  errorMessage?: string;
}

export interface Sheet {
  id: string;
  name: string;
  data: SheetData;
  conditionalFormats?: ConditionalRule[];
  validationRules?: ValidationRule[];
  accessCode?: string; // Code required to view/edit
  accessCodeExpiration?: number; // Timestamp when code expires
  isShared?: boolean; // Visual indicator
}

export interface AnalysisResult {
  summary: string;
  insights: string[];
}

export enum MessageRole {
  USER = 'user',
  MODEL = 'model'
}

export interface ChatMessage {
  id: string;
  role: MessageRole;
  text: string;
  isThinking?: boolean;
}
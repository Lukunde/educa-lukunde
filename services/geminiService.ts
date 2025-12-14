import { GoogleGenAI } from "@google/genai";
import { SheetData } from "../types";

// Safely retrieve API key without crashing
const getApiKey = () => {
  try {
    // Check for process.env safely
    if (typeof process !== 'undefined' && process.env && process.env.API_KEY) {
      return process.env.API_KEY;
    }
  } catch (e) {
    // Ignore error
  }
  return '';
};

// Lazy initialization to prevent top-level crashes
let aiInstance: GoogleGenAI | null = null;

const getAI = (): GoogleGenAI | null => {
  if (aiInstance) return aiInstance;
  
  const key = getApiKey();
  if (!key) return null;

  try {
    aiInstance = new GoogleGenAI({ apiKey: key });
    return aiInstance;
  } catch (error) {
    console.error("Failed to initialize GoogleGenAI", error);
    return null;
  }
};

export const analyzeSheetData = async (data: SheetData, query: string): Promise<string> => {
  const ai = getAI();
  if (!ai) {
    return "Erro: Chave de API não configurada ou erro de inicialização.";
  }

  try {
    // Convert a subset of data to CSV for context (limit rows to avoid token limits)
    // We take the first 50 rows as context
    const headers = data[0]?.map(h => String(h || "")).join(",") || "";
    const rows = data.slice(1, 51).map(row => row ? row.map(c => String(c || "")).join(",") : "").join("\n");
    const csvContext = `${headers}\n${rows}`;

    const prompt = `
      Você é um assistente de análise de dados para uma aplicação escolar chamada "Educa-Lukunde".
      Abaixo está uma amostra dos dados da planilha atual (formato CSV).
      
      DADOS:
      ${csvContext}
      
      PERGUNTA DO USUÁRIO:
      "${query}"
      
      Responda de forma concisa, profissional e útil para um professor ou gestor escolar.
      Se a pergunta for sobre separar turmas, explique que eles podem usar o botão "Separar Turmas" na barra de ferramentas.
    `;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
    });

    return response.text || "Não foi possível gerar uma análise.";
  } catch (error) {
    console.error("Gemini Error:", error);
    return "Desculpe, ocorreu um erro ao conectar com a IA.";
  }
};

export const suggestClassColumn = async (headers: string[]): Promise<string | null> => {
   const ai = getAI();
   if (!ai) return null;

   try {
     const prompt = `
       Dada a seguinte lista de cabeçalhos de uma planilha escolar: ${headers.join(", ")}.
       Qual deles é mais provável de representar a "Turma", "Classe" ou "Série"?
       Retorne APENAS o nome exato do cabeçalho. Se nenhum parecer apropriado, retorne "null".
     `;

     const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
     });
     
     const text = response.text?.trim();
     return text === "null" ? null : text || null;
   } catch (e) {
     return null;
   }
};
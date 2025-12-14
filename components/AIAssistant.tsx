import React, { useState, useRef, useEffect } from 'react';
import { Send, Bot, X, Sparkles, Loader2 } from 'lucide-react';
import { analyzeSheetData } from '../services/geminiService';
import { Sheet } from '../types';

interface AIAssistantProps {
  activeSheet: Sheet | undefined;
  onClose: () => void;
}

const AIAssistant: React.FC<AIAssistantProps> = ({ activeSheet, onClose }) => {
  const [query, setQuery] = useState('');
  const [messages, setMessages] = useState<{role: 'user' | 'model', text: string}[]>([
    { role: 'model', text: 'Olá! Sou a IA do Educa-Lukunde. Como posso ajudar com esta pauta?' }
  ]);
  const [isLoading, setIsLoading] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const handleSend = async () => {
    if (!query.trim() || !activeSheet) return;

    const userText = query;
    setQuery('');
    setMessages(prev => [...prev, { role: 'user', text: userText }]);
    setIsLoading(true);

    try {
      const response = await analyzeSheetData(activeSheet.data, userText);
      setMessages(prev => [...prev, { role: 'model', text: response }]);
    } catch (error) {
      setMessages(prev => [...prev, { role: 'model', text: "Desculpe, tive um problema ao analisar." }]);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="w-80 border-l border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-800 flex flex-col shadow-xl z-30 h-full transition-colors duration-200">
      <div className="p-4 border-b border-gray-200 dark:border-gray-700 flex justify-between items-center bg-emerald-50 dark:bg-emerald-900/20">
        <div className="flex items-center gap-2 text-emerald-800 dark:text-emerald-400">
          <Sparkles size={18} />
          <h3 className="font-semibold text-sm">Lukunde IA</h3>
        </div>
        <button onClick={onClose} className="text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200">
          <X size={18} />
        </button>
      </div>

      <div className="flex-1 overflow-y-auto p-4 space-y-4 bg-gray-50/50 dark:bg-gray-900/50">
        {messages.map((msg, idx) => (
          <div key={idx} className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}>
            <div 
              className={`max-w-[85%] rounded-2xl px-4 py-2 text-sm shadow-sm ${
                msg.role === 'user' 
                  ? 'bg-emerald-600 text-white rounded-br-none' 
                  : 'bg-white dark:bg-gray-700 border border-gray-100 dark:border-gray-600 text-gray-800 dark:text-gray-100 rounded-bl-none'
              }`}
            >
              {msg.text}
            </div>
          </div>
        ))}
        {isLoading && (
          <div className="flex justify-start">
             <div className="bg-white dark:bg-gray-700 border border-gray-100 dark:border-gray-600 rounded-2xl rounded-bl-none px-4 py-3 shadow-sm">
               <Loader2 className="animate-spin text-emerald-600 dark:text-emerald-400" size={16} />
             </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      <div className="p-3 border-t border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-800">
        <div className="flex items-center gap-2 bg-gray-100 dark:bg-gray-700 rounded-full px-3 py-2 border border-transparent focus-within:border-emerald-300 dark:focus-within:border-emerald-700 focus-within:bg-white dark:focus-within:bg-gray-900 transition-all">
          <input 
            type="text" 
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && handleSend()}
            placeholder="Pergunte sobre as notas..."
            className="flex-1 bg-transparent text-sm outline-none text-gray-700 dark:text-gray-200 placeholder-gray-400 dark:placeholder-gray-500"
            disabled={!activeSheet}
          />
          <button 
            onClick={handleSend}
            disabled={isLoading || !activeSheet}
            className="text-emerald-600 dark:text-emerald-400 hover:bg-emerald-100 dark:hover:bg-emerald-900/40 p-1.5 rounded-full transition-colors disabled:opacity-50"
          >
            <Send size={16} />
          </button>
        </div>
        {!activeSheet && (
           <p className="text-xs text-center text-amber-500 mt-2">Carregue uma planilha para começar.</p>
        )}
      </div>
    </div>
  );
};

export default AIAssistant;
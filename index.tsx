import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { AlertTriangle, RefreshCw, Trash2 } from 'lucide-react';

interface ErrorBoundaryProps {
  children?: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  public state: ErrorBoundaryState = { hasError: false, error: null };
  public props: ErrorBoundaryProps;

  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.props = props;
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("Uncaught error:", error, errorInfo);
  }

  handleReload = () => {
    window.location.reload();
  };

  handleHardReset = () => {
    if (confirm("Isso apagará todas as planilhas salvas para recuperar o sistema. Deseja continuar?")) {
      try {
        localStorage.clear();
        sessionStorage.clear();
      } catch(e) {
        // Ignore errors clearing storage
      }
      window.location.reload();
    }
  };

  render() {
    if (this.state.hasError) {
      return (
        <div className="flex flex-col items-center justify-center min-h-screen bg-gray-50 text-gray-800 p-4 font-sans">
          <div className="bg-white p-8 rounded-xl shadow-xl max-w-md w-full border border-gray-200 text-center">
            <div className="w-16 h-16 bg-red-100 text-red-500 rounded-full flex items-center justify-center mx-auto mb-4">
              <AlertTriangle size={32} />
            </div>
            <h1 className="text-xl font-bold mb-2 text-gray-900">Ops! Algo deu errado.</h1>
            <p className="text-sm text-gray-500 mb-6">
              Ocorreu um erro inesperado na aplicação. Tente recarregar ou resetar os dados.
            </p>
            
            <div className="bg-gray-100 p-3 rounded text-left text-xs font-mono text-gray-600 mb-6 overflow-auto max-h-32 whitespace-pre-wrap">
              {this.state.error?.message || "Erro desconhecido"}
            </div>

            <div className="space-y-3">
              <button 
                onClick={this.handleReload}
                className="flex items-center justify-center gap-2 w-full bg-emerald-600 hover:bg-emerald-700 text-white font-medium py-2.5 rounded-lg transition-colors"
              >
                <RefreshCw size={16} />
                Tentar Novamente
              </button>
              
              <button 
                onClick={this.handleHardReset}
                className="flex items-center justify-center gap-2 w-full bg-white hover:bg-red-50 text-red-600 border border-red-200 font-medium py-2.5 rounded-lg transition-colors"
              >
                <Trash2 size={16} />
                Resetar Aplicação (Limpar Dados)
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error("Could not find root element to mount to");
}

const root = ReactDOM.createRoot(rootElement);
root.render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>
);
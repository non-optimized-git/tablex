import { useEffect, useState } from 'react';
import Header from './components/Header.jsx';
import FileUploader from './components/FileUploader.jsx';
import GroupSelector from './components/GroupSelector.jsx';
import QuestionSelector from './components/QuestionSelector.jsx';
import ResultPreview from './components/ResultPreview.jsx';
import { useStore } from './store/useStore.js';

function LoadingSpinner() {
  return (
    <div className="flex items-center justify-center p-8">
      <div className="w-8 h-8 border-2 border-muji-beige border-t-transparent rounded-full animate-spin"></div>
    </div>
  );
}

export default function App() {
  const { loadSavedConfig } = useStore();
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    try {
      loadSavedConfig();
    } catch (e) {
      console.error('Failed to load saved config:', e);
    }
  }, [loadSavedConfig]);

  if (error) {
    return (
      <div className="min-h-screen bg-muji-cream flex items-center justify-center">
        <div className="bg-muji-warmWhite border border-muji-lightGray rounded-lg p-6 max-w-md">
          <h2 className="text-lg font-medium text-muji-charcoal mb-2">发生错误</h2>
          <p className="text-sm text-muji-gray mb-4">{error}</p>
          <button
            onClick={() => window.location.reload()}
            className="px-4 py-2 bg-muji-beige text-muji-charcoal rounded hover:bg-muji-lightGray"
          >
            重新加载
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-muji-cream">
      <Header />
      <main className="max-w-6xl mx-auto px-6 py-8">
        {isLoading && <LoadingSpinner />}
        <div className="grid gap-6">
          <FileUploader />
          <GroupSelector />
          <QuestionSelector />
          <ResultPreview />
        </div>
      </main>
      <footer className="text-center py-6 text-xs text-muji-gray">
        基于原始 Python 分析逻辑构建 · MUJI 风格
      </footer>
    </div>
  );
}
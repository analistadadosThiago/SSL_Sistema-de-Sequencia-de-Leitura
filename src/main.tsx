import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';

if (typeof window !== 'undefined') {
  window.addEventListener('unhandledrejection', (e) => {
    const reason = e.reason as { message?: unknown } | undefined;
    const msg = String(reason?.message || e.reason || '');
    if (msg.includes('WebSocket') || msg.includes('vite') || msg.includes('HMR')) {
      e.preventDefault();
      return;
    }
    e.preventDefault();
    console.error('[Unhandled]', e.reason);
  });

  window.addEventListener('error', (e) => {
    const errObj = e.error as { message?: unknown } | undefined;
    const msg = String(errObj?.message || e.message || '');
    if (msg.includes('WebSocket') || msg.includes('vite') || msg.includes('HMR')) {
      e.preventDefault();
      return;
    }
    console.error('[Error]', e.error || e.message);
  });
}

createRoot(document.getElementById('root')!).render(<App />);

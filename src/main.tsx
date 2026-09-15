import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import { ErrorBoundary } from './ErrorBoundary.tsx';
import AdminDataBackupPanel from './features/admin/AdminDataBackupPanel';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <>
        <App />
        <AdminDataBackupPanel />
      </>
    </ErrorBoundary>
  </StrictMode>,
);

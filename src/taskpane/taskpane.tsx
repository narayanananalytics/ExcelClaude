import * as React from 'react';
import { createRoot } from 'react-dom/client';
import App from './components/App';
import './taskpane.css';

let isOfficeInitialized = false;

const render = (Component: React.ComponentType) => {
  const rootElement = document.getElementById('root');
  if (rootElement) {
    const root = createRoot(rootElement);
    root.render(<Component />);
  }
};

Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

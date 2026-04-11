import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'

/* StrictMode disabled: @fortune-sheet/react uses Immer proxies that are revoked on the StrictMode mount/unmount cycle in dev, causing runtime errors after interaction. */
createRoot(document.getElementById('root')!).render(<App />)

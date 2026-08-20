import { createRoot } from 'react-dom/client'
import { App } from './App.js'
import './styles.css'

const root = document.getElementById('root')
if (!root) throw new Error('Missing Office Add-in root')

createRoot(root).render(<App />)

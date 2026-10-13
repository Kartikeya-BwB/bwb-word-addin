import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import './index.css'

const render = () => {
  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  )
}

// Inside Word, wait for Office.js before rendering; in a plain browser, render straight away.
if (typeof Office !== 'undefined') {
  Office.onReady(() => render())
} else {
  render()
}

// src/App.jsx
import React, { useState } from 'react';
import './App.css';
import EnerzhiExcelTransformer from './components/EnerzhiExcelTransformer';
import AmarEnergy from './components/AmarEnergy';

function App() {
  const [activeTab, setActiveTab] = useState('offers'); // 'offers' یا 'amar'

  return (
    <div style={{ minHeight: '100vh', direction: 'rtl', backgroundColor: '#fafafa' }}>
      {/* تب‌بار بالای صفحه */}
      <header
        style={{
          display: 'flex',
          gap: '10px',
          padding: '12px 20px',
          background: '#2c3e50',
          borderBottom: '2px solid #1abc9c',
        }}
      >
        <button
          onClick={() => setActiveTab('offers')}
          style={{
            padding: '10px 22px',
            borderRadius: '6px',
            border: 'none',
            fontWeight: 'bold',
            fontSize: '15px',
            cursor: 'pointer',
            backgroundColor: activeTab === 'offers' ? '#1abc9c' : '#34495e',
            color: '#fff',
            transition: '0.2s',
          }}
        >
          اطلاعیه عرضه‌ها
        </button>

        <button
          onClick={() => setActiveTab('amar')}
          style={{
            padding: '10px 22px',
            borderRadius: '6px',
            border: 'none',
            fontWeight: 'bold',
            fontSize: '15px',
            cursor: 'pointer',
            backgroundColor: activeTab === 'amar' ? '#1abc9c' : '#34495e',
            color: '#fff',
            transition: '0.2s',
          }}
        >
          آمار معاملات
        </button>
      </header>

      {/* محتوای فعال */}
      <main>
        {activeTab === 'offers' ? <EnerzhiExcelTransformer /> : <AmarEnergy />}
      </main>
    </div>
  );
}

export default App;

import React, { useState } from 'react';
import * as XLSX from 'xlsx';

// Panel de control para carga de archivos, filtros y exportación.
// Recibe callbacks para cargar Excel/CSV, cambiar hoja y exportar; además muestra el conteo filtrado.
function ControlPanel({
  onExport,
  onAutocompletarConZoom,
  onLoadZoomCsv,
  isLoading,
  isProcessing,
  displayDataLength,
  displayData,
  availableSheets,
  selectedSheet,
  onSheetChange,
  docenteOptions = [],
  selectedDocente = '',
  onDocenteFilterChange,
  onSaveBackup
}) {

  // Exporta los datos visibles a Excel con formato básico si no se entrega un exportador externo.
  // Cabeceras en negrita, bordes finos y ancho de columna ajustado; detecta fechas/horas como texto legible.
  const handleExport = () => {
    if (onExport) {
      onExport();
      return;
    }

    if (!displayData || displayData.length === 0) {
      alert('No hay datos para exportar.');
      return;
    }

    const headers = displayData[0] ? Object.keys(displayData[0]) : [];

    const wsData = [headers];
    displayData.forEach(row => {
      const rowData = headers.map(header => {
        let value = row[header];
        if (value === null || value === undefined) return '';
        if (typeof value === 'object') return JSON.stringify(value).slice(0, 50) + '...';
        
        const date = new Date(value);
        if (!isNaN(date.getTime()) && typeof value === 'string' && value.includes('-')) {
          return date.toLocaleDateString('es-CL') + ' ' + date.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit' });
        }
        
        return String(value);
      });
      wsData.push(rowData);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    const numCols = headers.length;
    for (let col = 0; col < numCols; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (ws[cellAddress]) {
        ws[cellAddress].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "366092" } },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
          }
        };
      }
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = 0; r <= range.e.r; r++) {
      for (let c = 0; c <= range.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        if (ws[cellAddress] && !ws[cellAddress].s) {
          ws[cellAddress].s = {};
        }
        if (ws[cellAddress].s) {
          ws[cellAddress].s.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
          };
        }
      }
    }

    const colWidths = headers.map((header, i) => {
      const maxLength = Math.max(
        header.length,
        ...wsData.slice(1).map(row => String(row[i] || '').length)
      );
      return { wch: Math.min(maxLength + 2, 30) };
    });
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    const sheetName = 'Monitoreo_USS';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const today = new Date();
    const dateStr = today.toISOString().split('T')[0];
    const fileName = `Monitoreo_USS_${dateStr}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };
  
  // Calcular promedio de un indicador general (por ejemplo: EFICIENCIA / EFICACIA)
  const calculateAverageMetric = (metricLabel) => {
    if (!displayData || displayData.length === 0) return null;
    const normalizedMetric = String(metricLabel || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim()
      .toUpperCase();

    let sum = 0;
    let count = 0;

    displayData.forEach(row => {
      if (!row) return;
      Object.keys(row).forEach(key => {
        const kNorm = String(key || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
        if (kNorm.includes(normalizedMetric)) {
          const raw = String(row[key] || '').trim();
          const numMatch = raw.match(/(\d+(?:[.,]\d+)?)/);
          if (numMatch) {
            const num = parseFloat(numMatch[1].replace(',', '.'));
            if (Number.isFinite(num)) {
              sum += num;
              count++;
            }
          }
        }
      });
    });
    return count > 0 ? (sum / count).toFixed(2) : null;
  };
  
  const averageEfficiency = calculateAverageMetric('EFICIENCIA');
  const averageEfficacy = calculateAverageMetric('EFICACIA');
  
  // Calcular promedios por docente, curso y sección
  const calculateAverageByDocente = () => {
    if (!displayData || displayData.length === 0) return [];
    
    const stats = {};
    displayData.forEach(row => {
      if (!row) return;
      const docente = row['DOCENTE'] || 'Sin Docente';
      const curso = row['CURSO'] || 'Sin Curso';
      
      let seccion = 'Sin Sección';
      for (const key of Object.keys(row)) {
        const kNorm = String(key || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
        if (kNorm === 'SECCION' || kNorm === 'PEAD' || kNorm.includes('SECCION/PEAD') || kNorm.includes('PEAD/SECCION')) {
          if (row[key] && String(row[key]).trim() !== '') {
            seccion = String(row[key]).trim();
            break;
          }
        }
      }

      // Agrupar por docente, curso y sección
      const keyStr = `${docente} - ${curso} - ${seccion}`;
      
      if (!stats[keyStr]) stats[keyStr] = { sum: 0, count: 0, docente, curso, seccion };
      
      Object.keys(row).forEach(key => {
        const kNorm = String(key || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
        if (kNorm.includes('EFICIENCIA')) {
          const raw = String(row[key] || '').trim();
          const numMatch = raw.match(/(\d+(?:[.,]\d+)?)/);
          if (numMatch) {
            const num = parseFloat(numMatch[1].replace(',', '.'));
            if (Number.isFinite(num)) {
              stats[keyStr].sum += num;
              stats[keyStr].count++;
            }
          }
        }
      });
    });

    return Object.keys(stats)
      .map(keyStr => {
        const dStats = stats[keyStr];
        return {
          key: keyStr,
          docente: dStats.docente,
          curso: dStats.curso,
          seccion: dStats.seccion,
          average: dStats.count > 0 ? parseFloat((dStats.sum / dStats.count).toFixed(2)) : null
        };
      })
      .filter(d => d.average !== null)
      .sort((a, b) => b.average - a.average);
  };

  const averagesByCourse = calculateAverageByDocente();
  const [showDocentes, setShowDocentes] = useState(false);
  
  let effConfig = null;
  const indicatorBaseline = averageEfficiency !== null ? averageEfficiency : averageEfficacy;
  if (indicatorBaseline !== null) {
    const avg = parseFloat(indicatorBaseline);
    if (avg === 100) {
      effConfig = { bg: 'bg-[#f4fde9] border-[#63ed12]', text: 'text-[#103b07]', pText: 'text-[#3f7f12]', icon: '🌟', title: 'Excelente desempeño general', desc: 'Los indicadores del archivo muestran un nivel sobresaliente.' };
    } else if (avg >= 86) {
      effConfig = { bg: 'bg-[#e8f6fb] border-[#11acd3]', text: 'text-[#0c5d74]', pText: 'text-[#117a98]', icon: '👍', title: 'Buen nivel general', desc: 'Los resultados son positivos y existe un margen de mejora acotado.' };
    } else if (avg >= 80) {
      effConfig = { bg: 'bg-[#f3ecfb] border-[#5a2290]', text: 'text-[#5a2290]', pText: 'text-[#7444a5]', icon: '⚠️', title: 'Nivel aceptable', desc: 'Hay sesiones que requieren atención para elevar el rendimiento general.' };
    } else {
      effConfig = { bg: 'bg-[#f3ecfb] border-[#5a2290]', text: 'text-[#5a2290]', pText: 'text-[#7444a5]', icon: '🚨', title: 'Atención requerida', desc: 'Los indicadores están por debajo de lo esperado y se recomienda plan de mejora.' };
    }
  }

  return (
    <div className="bg-white rounded-xl shadow-2xl p-6 mb-6">
      <div className="flex items-center justify-between flex-wrap gap-4 mb-6">
        <div className="flex gap-3 flex-wrap">
          {/* Botones existentes */}
          <button
            onClick={handleExport}
            disabled={isLoading || displayDataLength === 0}
            className="bg-[#63ed12] hover:bg-[#54cb0f] text-[#103b07] font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2 text-sm disabled:bg-gray-400 disabled:text-white disabled:cursor-not-allowed"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            Exportar ({displayDataLength})
          </button>
          
          {/* Botón para subir CSV de Zoom */}
          <input
            id="file-input-zoom-csv"
            type="file"
            accept=".csv"
            onChange={onLoadZoomCsv}
            style={{ display: 'none' }}
          />
          <button
            className="bg-[#11acd3] hover:bg-[#0f9bbf] text-white font-bold py-2 px-4 rounded-lg"
            onClick={() => document.getElementById('file-input-zoom-csv').click()}
            disabled={isLoading}
          >
            Subir reporte CSV de Zoom
          </button>

          
          {/* Nuevos botones para guardar y ver historial */}
          <button
            onClick={onSaveBackup}
            disabled={isLoading || displayDataLength === 0}
            className="bg-[#5a2290] hover:bg-[#4b1c78] text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 flex items-center gap-2 text-sm disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4" />
            </svg>
            Guardar copia
          </button>
          
        </div>
      </div>

      {/* Resto de tu componente */}
      {availableSheets && availableSheets.length > 1 && (
        <div className="mb-4 bg-gray-100 rounded-lg p-2">
          <div className="flex gap-1 overflow-x-auto">
            {availableSheets.map((sheet) => (
              <button
                key={sheet.index}
                onClick={() => onSheetChange(sheet.index)}
                className={`px-4 py-2 rounded-lg font-semibold transition-all whitespace-nowrap text-sm ${
                  selectedSheet === sheet.index
                    ? 'bg-[#11acd3] text-white shadow-lg'
                    : 'bg-white text-gray-700 hover:bg-[#e8f6fb] border border-[#d7f3fa]'
                }`}
              >
                {sheet.name}
              </button>
            ))}
          </div>
        </div>
      )}
      


      <div className="bg-[#e8f6fb] border-2 border-[#11acd3] rounded-lg p-4">
        <div className="flex flex-col lg:flex-row lg:items-end gap-3">
          <div className="flex flex-wrap items-center gap-2">
            {[
              { key: 'regular', label: 'FILAS REGULAR', bg: '#11acd3', hover: '#0f9bbf', text: '#ffffff' },
              { key: 'intensivo', label: 'FILAS INTENSIVO', bg: '#5a2290', hover: '#4b1c78', text: '#ffffff' },
              { key: 'superintensivo', label: 'FILAS SUPERINTENSIVO', bg: '#63ed12', hover: '#54cb0f', text: '#103b07' }
            ].map((mode) => (
              <button
                key={mode.key}
                onClick={() => onAutocompletarConZoom && onAutocompletarConZoom(mode.key)}
                disabled={isLoading || isProcessing}
                className="font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 disabled:bg-gray-400 disabled:text-white disabled:cursor-not-allowed flex items-center justify-center gap-2 text-sm"
                style={{ backgroundColor: mode.bg, color: mode.text }}
                onMouseEnter={(e) => {
                  if (!isLoading && !isProcessing) e.currentTarget.style.backgroundColor = mode.hover;
                }}
                onMouseLeave={(e) => {
                  if (!isLoading && !isProcessing) e.currentTarget.style.backgroundColor = mode.bg;
                }}
              >
                {isProcessing ? (
                  <>
                    <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                    <span className="text-sm">Procesando...</span>
                  </>
                ) : (
                  <>
                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
                    </svg>
                    <span className="text-sm">{mode.label}</span>
                  </>
                )}
              </button>
            ))}
          </div>

          {/* Combobox para filtrar por docente */}
          <div className="min-w-[300px] lg:min-w-[360px]">
            <label htmlFor="docente-filter" className="block text-[11px] font-bold uppercase tracking-wide text-[#5a2290] mb-1">
              Filtro de docente
            </label>
            <div className="flex items-center gap-2">
              <select
                id="docente-filter"
                value={selectedDocente}
                onChange={(e) => onDocenteFilterChange && onDocenteFilterChange(e.target.value)}
                disabled={isLoading}
                className="h-[42px] flex-1 bg-white border-2 border-[#11acd3] rounded-lg px-4 text-sm text-gray-800 focus:outline-none focus:ring-2 focus:ring-[#11acd3] focus:border-[#11acd3] transition-shadow shadow-sm"
              >
                <option value="">Todos los docentes</option>
                {docenteOptions.map((docente) => (
                  <option key={docente} value={docente}>{docente}</option>
                ))}
              </select>
              {selectedDocente && (
                <button
                  onClick={() => onDocenteFilterChange && onDocenteFilterChange('')}
                  className="h-[42px] bg-white hover:bg-red-50 text-red-600 border border-red-200 hover:border-red-300 rounded-lg px-3 text-xs font-bold transition-colors shrink-0"
                  title="Quitar filtro"
                >
                  Limpiar
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

    </div>
  );
}

export default ControlPanel;

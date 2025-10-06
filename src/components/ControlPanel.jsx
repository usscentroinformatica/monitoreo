import React from "react";
import * as XLSX from 'xlsx';

function ControlPanel({
  selectedDocente,
  setSelectedDocente,
  numFilas,
  setNumFilas,
  uniqueDocentes,
  onCreateRows,
  onCreateRowsForAll,
  onExport,
  onLoadExcel,
  onLoadZoomCsv,
  onAutocompletarConZoom, // Nueva prop
  isLoading,
  displayDataLength,
  displayData,
  availableSheets,
  selectedSheet,
  onSheetChange
}) {
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
    const sheetName = selectedDocente ? `Sesiones_${selectedDocente.replace(/\s+/g, '_')}` : 'Monitoreo_USS';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const today = new Date();
    const dateStr = today.toISOString().split('T')[0];
    const fileName = `Monitoreo_USS_${dateStr}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  return (
    <div className="bg-white rounded-xl shadow-2xl p-6 mb-6">
      <div className="flex items-center justify-between flex-wrap gap-4 mb-6">
        <div>
          <h1 className="text-3xl font-bold text-blue-900 mb-1">
            Sistema de Monitoreo USS
          </h1>
        </div>
        
        <div className="flex gap-3">
          <label className="bg-orange-600 hover:bg-orange-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2 cursor-pointer text-sm">
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            {isLoading ? "Cargando..." : "Cargar Excel"}
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={onLoadExcel}
              className="hidden"
              disabled={isLoading}
            />
          </label>

          <label className="bg-purple-600 hover:bg-purple-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2 cursor-pointer text-sm">
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            {isLoading ? "Procesando..." : "Cargar CSV Zoom"}
            <input
              type="file"
              accept=".csv"
              onChange={onLoadZoomCsv}
              className="hidden"
              disabled={isLoading}
            />
          </label>
          
          <button
            onClick={handleExport}
            disabled={isLoading || displayDataLength === 0}
            className="bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2 text-sm disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            Exportar ({displayDataLength})
          </button>
        </div>
      </div>

      {availableSheets && availableSheets.length > 1 && (
        <div className="mb-4 bg-gray-100 rounded-lg p-2">
          <div className="flex gap-1 overflow-x-auto">
            {availableSheets.map((sheet) => (
              <button
                key={sheet.index}
                onClick={() => {
                  onSheetChange(sheet.index);
                  setSelectedDocente("");
                }}
                className={`px-4 py-2 rounded-lg font-semibold transition-all whitespace-nowrap text-sm ${
                  selectedSheet === sheet.index
                    ? 'bg-blue-600 text-white shadow-lg'
                    : 'bg-white text-gray-700 hover:bg-gray-200 border border-gray-300'
                }`}
              >
                {sheet.name}
              </button>
            ))}
          </div>
        </div>
      )}

      <div className="bg-blue-50 border-2 border-blue-200 rounded-lg p-4">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
          <div className="flex flex-col md:flex-row gap-2 items-end">
            <div className="flex-1 w-full md:w-auto">
              <label className="block text-sm font-bold text-blue-900 mb-2">
                Filtrar por Docente:
              </label>
              <select
                value={selectedDocente}
                onChange={(e) => setSelectedDocente(e.target.value)}
                className="w-full px-4 py-2 border-2 border-blue-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
              >
                <option value="">Todos los docentes</option>
                {uniqueDocentes
                  .filter(docente => docente && docente.trim() !== '')
                  .sort((a, b) => a.localeCompare(b))
                  .map((docente) => (
                    <option key={docente} value={docente}>
                      {docente}
                    </option>
                ))}
              </select>
            </div>
            <button
              onClick={onAutocompletarConZoom}
              disabled={isLoading}
              className="bg-gradient-to-r from-blue-600 to-slate-600 hover:from-blue-700 hover:to-slate-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 disabled:bg-gray-400 disabled:cursor-not-allowed flex items-center justify-center gap-2 text-sm w-full md:w-auto"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
              </svg>
              <span className="text-sm">Autocompletar</span>
            </button>
          </div>
        </div>
        
        <div className="mt-3">
          <p className="text-xs text-gray-600">
            Autocompletar MONITOREO (16 sesiones/curso): Crea automáticamente 16 filas por curso/docente y las autocompleta con datos de Zoom
          </p>
        </div>
        
        {selectedDocente && (
          <div className="mt-3 text-sm text-blue-800 font-semibold">
            Mostrando {displayDataLength} sesiones de {selectedDocente}
          </div>
        )}
      </div>
    </div>
  );
}

export default ControlPanel;

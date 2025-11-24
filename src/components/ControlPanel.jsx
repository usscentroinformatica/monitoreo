import * as XLSX from 'xlsx';

// Panel de control para carga de archivos, filtros y exportación.
// Recibe callbacks para cargar Excel/CSV, cambiar hoja y exportar; además muestra el conteo filtrado.
function ControlPanel({
  onExport,
  onAutocompletarConZoom,
  onLoadZoomCsv,
  isLoading,
  displayDataLength,
  displayData,
  availableSheets,
  selectedSheet,
  onSheetChange,
  onSelectRandomDocente,  // Nueva prop
  randomDocente,          // Nueva prop
  onClearRandomDocente,    // Nueva prop
  onSaveBackup,        // Nueva prop
  onOpenBackupModal    // Nueva prop
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
  
  return (
    <div className="bg-white rounded-xl shadow-2xl p-6 mb-6">
      <div className="flex items-center justify-between flex-wrap gap-4 mb-6">
        <div>
          <h1 className="text-3xl font-bold text-blue-900 mb-1">
            Sistema de Monitoreo USS
          </h1>
        </div>
        
        <div className="flex gap-3 flex-wrap">
          {/* Botones existentes */}
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
          
          {/* Botón para subir CSV de Zoom */}
          <input
            id="file-input-zoom-csv"
            type="file"
            accept=".csv"
            onChange={onLoadZoomCsv}
            style={{ display: 'none' }}
          />
          <button
            className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-lg"
            onClick={() => document.getElementById('file-input-zoom-csv').click()}
            disabled={isLoading}
          >
            Subir reporte CSV de Zoom
          </button>

          
          {/* Botón para elegir docente aleatorio */}
          <button
            onClick={onSelectRandomDocente}
            disabled={isLoading}
            className="bg-orange-500 hover:bg-orange-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 flex items-center gap-2 text-sm"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
            </svg>
            Elegir docente aleatorio
          </button>
          
          {/* Nuevos botones para guardar y ver historial */}
          <button
            onClick={onSaveBackup}
            disabled={isLoading || displayDataLength === 0}
            className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 flex items-center gap-2 text-sm disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4" />
            </svg>
            Guardar copia
          </button>
          
          <button
            onClick={onOpenBackupModal}
            className="bg-teal-600 hover:bg-teal-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 flex items-center gap-2 text-sm"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
            </svg>
            Ver archivos guardados
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
      
      {/* Mostrar información del docente aleatorio seleccionado */}
      {randomDocente && (
        <div className="mb-4 bg-orange-100 border-2 border-orange-300 rounded-lg p-4 flex justify-between items-center">
          <div>
            <p className="text-orange-800 font-bold">
              Mostrando solo registros del docente: <span className="text-orange-600 text-lg">{randomDocente}</span>
            </p>
            <p className="text-sm text-orange-700">
              Filas mostradas: {displayDataLength} 
            </p>
          </div>
          <button
            onClick={onClearRandomDocente}
            className="bg-red-500 hover:bg-red-700 text-white font-bold py-2 px-4 rounded-lg"
          >
            Mostrar todos
          </button>
        </div>
      )}

      <div className="bg-blue-50 border-2 border-blue-200 rounded-lg p-4">
        <button
          onClick={onAutocompletarConZoom}
          disabled={isLoading}
          className="bg-gradient-to-r from-blue-600 to-slate-600 hover:from-blue-700 hover:to-slate-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 disabled:bg-gray-400 disabled:cursor-not-allowed flex items-center justify-center gap-2 text-sm"
        >
          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
          </svg>
          <span className="text-sm">Autocompletar</span>
        </button>
      </div>

    </div>
  );
}

export default ControlPanel;

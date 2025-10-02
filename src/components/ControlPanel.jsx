import React from "react";

function ControlPanel({
  selectedDocente,
  setSelectedDocente,
  numFilas,
  setNumFilas,
  uniqueDocentes,
  onCreateRows,
  onAddRow,
  onExport,
  onLoadExcel,
  onLoadZoomCsv,
  isLoading,
  displayDataLength
}) {
  return (
    <div className="bg-white rounded-xl shadow-2xl p-6 mb-6">
      <div className="flex items-center justify-between flex-wrap gap-4 mb-6">
        <div>
          <h1 className="text-3xl font-bold text-blue-900 mb-1">
            Sistema de Monitoreo USS - Tabla Editable
          </h1>
          <p className="text-gray-600">
            Filtra por docente y crea filas automáticamente
          </p>
        </div>
        
        <div className="flex gap-3">
          <label className="bg-orange-600 hover:bg-orange-700 text-white font-bold py-3 px-6 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2 cursor-pointer">
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
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

          <label className="bg-purple-600 hover:bg-purple-700 text-white font-bold py-3 px-6 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2 cursor-pointer">
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
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
            onClick={onAddRow}
            className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-6 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
            </svg>
            Agregar Fila
          </button>
          
          <button
            onClick={onExport}
            className="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 flex items-center gap-2"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            Exportar Excel
          </button>
        </div>
      </div>

      <div className="bg-blue-50 border-2 border-blue-200 rounded-lg p-4">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-end">
          <div>
            <label className="block text-sm font-bold text-blue-900 mb-2">
              Filtrar por Docente:
            </label>
            <select
              value={selectedDocente}
              onChange={(e) => setSelectedDocente(e.target.value)}
              className="w-full px-4 py-2 border-2 border-blue-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            >
              <option value="">Todos los docentes</option>
              {uniqueDocentes.map((docente) => (
                <option key={docente} value={docente}>
                  {docente}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-sm font-bold text-blue-900 mb-2">
              Número de Filas a Crear:
            </label>
            <input
              type="number"
              min="1"
              value={numFilas}
              onChange={(e) => setNumFilas(e.target.value)}
              placeholder="Ej: 16"
              className="w-full px-4 py-2 border-2 border-blue-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              disabled={!selectedDocente}
            />
          </div>

          <div>
            <button
              onClick={onCreateRows}
              disabled={!selectedDocente || !numFilas}
              className="w-full bg-purple-600 hover:bg-purple-700 text-white font-bold py-2 px-6 rounded-lg shadow-lg transition-all duration-200 transform hover:scale-105 disabled:bg-gray-400 disabled:cursor-not-allowed flex items-center justify-center gap-2"
            >
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
              Crear Filas
            </button>
          </div>
        </div>
        
        {selectedDocente && (
          <div className="mt-3 text-sm text-blue-800 font-semibold">
            📊 Mostrando {displayDataLength} sesiones de {selectedDocente}
          </div>
        )}
      </div>
    </div>
  );
}

export default ControlPanel;
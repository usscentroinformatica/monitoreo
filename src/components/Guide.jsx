import React from "react";

const Guide = () => {
  const handlePrint = () => {
    window.print();
  };

  return (
    <>
      {/* ESTILOS DE IMPRESIÓN: solo se aplican al imprimir */}
      <style media="print">{`
        @page { size: A4 portrait; margin: 12mm; }
        body * { visibility: hidden; }
        #print-guide, #print-guide * { visibility: visible; }
        #print-guide { position: absolute; top: 0; left: 0; width: 100%; font-size: 14px; line-height: 1.5; color: #000; background: #fff; }
        .no-print { display: none !important; }
      `}</style>

      {/* GUIA VISUAL */}
      <div id="print-guide" className="bg-gradient-to-br from-blue-50 to-indigo-100 rounded-2xl shadow-xl p-6 md:p-8">
        {/* ENCABEZADO */}
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-3">
            <div className="bg-blue-900 text-white rounded-full p-2">
              <svg className="w-7 h-7" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
            </div>
            <div>
              <h1 className="text-2xl md:text-3xl font-extrabold text-blue-900">Guía Paso a Paso</h1>
              <p className="text-sm text-blue-700">Tiempo estimado: 10 minutos</p>
            </div>
          </div>
          <button onClick={handlePrint} className="no-print bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded-lg shadow flex items-center gap-2 transition">
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z" />
            </svg>
            Generar PDF
          </button>
        </div>

        {/* PASOS */}
        <div className="grid gap-4">
          {steps.map((step, idx) => (
            <div key={idx} className="bg-white rounded-xl shadow p-4 flex items-start gap-4">
              <div className={`flex-shrink-0 w-10 h-10 rounded-full flex items-center justify-center text-white font-bold ${step.color}`}>
                {idx + 1}
              </div>
              <div className="flex-1">
                <h3 className="font-bold text-gray-800 mb-1">{step.title}</h3>
                <p className="text-sm text-gray-600 mb-2">{step.desc}</p>
                {step.details?.length > 0 && (
                  <ul className="list-disc list-inside space-y-1 text-xs text-gray-600">
                    {step.details.map((d, i) => (
                      <li key={i}><span className="font-semibold">{d.label}:</span> {d.text}</li>
                    ))}
                  </ul>
                )}
                {step.extra && (<div className="mt-2 text-xs text-gray-500">{step.extra}</div>)}
              </div>
            </div>
          ))}
        </div>

        {/* TIPS FINALES */}
        <div className="mt-6 bg-yellow-50 border-l-4 border-yellow-400 rounded p-4">
          <h4 className="font-bold text-yellow-800 mb-2">Tips rápidos</h4>
          <ul className="list-disc list-inside text-sm text-yellow-700 space-y-1">
            <li>Si la carga falla, prueba con la plantilla disponible.</li>
            <li>Las alertas en color resaltado indican puntos a revisar.</li>
            <li>Guarda copias frecuentes con “Guardar copia”.</li>
            <li>Para imprimir la guía, usa el botón “Generar PDF”.</li>
          </ul>
        </div>
      </div>
    </>
  );
};

export default Guide;

// Datos de los pasos
const steps = [
  {
    title: "Carga tu archivo Excel",
    desc: "Importa tu hoja de programación para comenzar.",
    details: [
      { label: "Acción", text: "Clic en “+ Nueva Pestaña”." },
      { label: "Selección", text: "Elige tu archivo .xlsx desde tu equipo." },
      { label: "Resultado", text: "Verás una tabla editable organizada automáticamente." },
    ],
    color: "bg-blue-600",
  },
  {
    title: "Sube el reporte de Zoom",
    desc: "Integra datos de uso para autocompletar tiempos.",
    details: [
      { label: "Acción", text: "Clic en “Subir reporte CSV de Zoom”." },
      { label: "Origen", text: "Descarga el CSV desde la sección de reportes en Zoom." },
      { label: "Confirmación", text: "Aparecerá una notificación de integración exitosa." },
    ],
    color: "bg-green-600",
  },
  {
    title: "Autocompleta sesiones y tiempos",
    desc: "Completa fechas, horas y duraciones automáticamente.",
    details: [
      { label: "Acción", text: "Clic en “Autocompletar” (botón azul)." },
      { label: "Proceso", text: "Se llenan campos y se generan sesiones faltantes cuando corresponde." },
      { label: "Resumen", text: "Verás una alerta final con estadísticas del proceso." },
    ],
    color: "bg-indigo-600",
  },
  {
    title: "Revisa y edita tus datos",
    desc: "Realiza ajustes rápidos y precisos.",
    details: [
      { label: "Filtro", text: "Usa “Elegir docente aleatorio” (naranja) para enfocarte; “Mostrar todos” para volver." },
      { label: "Edición", text: "Haz clic en una celda para cambiar su valor." },
      { label: "Eliminación", text: "Usa el botón “Eliminar” en la fila que quieras quitar." },
    ],
    color: "bg-orange-600",
  },
  {
    title: "Exporta y guarda respaldos",
    desc: "Genera tu archivo final y conserva copias de seguridad.",
    details: [
      { label: "Exportar", text: "Clic en “Exportar (N)” (verde) para descargar .xlsx." },
      { label: "Backup", text: "Clic en “Guardar copia” (índigo) para almacenar un respaldo." },
      { label: "Historial", text: "“Ver archivos guardados” (turquesa) para gestionar descargas y restauraciones." },
    ],
    color: "bg-purple-600",
  },
];
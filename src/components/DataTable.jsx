import React from "react";

function DataTable({ 
  data, 
  headers, 
  dropdownOptions, 
  onCellChange, 
  onDeleteRow 
}) {
  // Helper function to extract information from course titleE
  const extractCourseInfo = (value) => {
    if (!value || typeof value !== 'string') return { course: '', section: '' };
    
    const courseMatch = value.match(/^(.+?)–\s*(PEAD-[a-zA-Z]+)/i);
    if (courseMatch) {
      return {
        course: courseMatch[1].trim(),
        section: courseMatch[2].trim()
      };
    }
    
    const altMatch = value.match(/^(.+?)\s+(PEAD-[a-zA-Z]+)/i);
    if (altMatch) {
      return {
        course: altMatch[1].trim(),
        section: altMatch[2].trim()
      };
    }
    
    return { course: value, section: '' };
  };

  const suggestCourse = (value, rowData) => {
    if (!value || typeof value !== 'string') return '';
    return extractCourseInfo(value).course;
  };

  const suggestSection = (value, rowData) => {
    if (!value || typeof value !== 'string') return '';
    return extractCourseInfo(value).section;
  };

  // FUNCIÓN MEJORADA para manejar correctamente los objetos
  const ensureString = (value) => {
    if (value === null || value === undefined) return '';
    
    // Si ya es string, devolverlo directamente
    if (typeof value === 'string') return value;
    
    // Si es número o booleano
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    
    // Si es una fecha
    if (value instanceof Date) {
      return value.toLocaleDateString('es-ES', {
        day: '2-digit',
        month: '2-digit', 
        year: 'numeric'
      });
    }
    
    // Si es un array
    if (Array.isArray(value)) {
      return value.map(v => ensureString(v)).join(', ');
    }
    
    // Si es un objeto
    if (typeof value === 'object') {
      // Intentar obtener propiedades comunes
      if (value.text !== undefined) return String(value.text);
      if (value.value !== undefined) return String(value.value);
      if (value.name !== undefined) return String(value.name);
      if (value.richText !== undefined && Array.isArray(value.richText)) {
        return value.richText.map(rt => rt.text || '').join('');
      }
      
      // Si tiene un método toString personalizado
      if (value.toString && value.toString !== Object.prototype.toString) {
        const strValue = value.toString();
        if (strValue !== '[object Object]') return strValue;
      }
      
      // Último recurso: intentar JSON
      try {
        const jsonStr = JSON.stringify(value);
        if (jsonStr && jsonStr !== '{}' && jsonStr !== '[]') {
          return jsonStr.length > 50 ? jsonStr.substring(0, 50) + '...' : jsonStr;
        }
      } catch (e) {
        // Si falla JSON, devolver vacío
      }
      
      return '';
    }
    
    // Conversión final por defecto
    return String(value);
  };

  const displayHeaders = headers;

  if (!displayHeaders || displayHeaders.length === 0) {
    return (
      <div className="bg-white rounded-xl shadow-2xl overflow-hidden p-8 text-center">
        <p className="text-gray-500 text-lg">No hay datos cargados. Por favor, carga un archivo Excel para comenzar.</p>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-xl shadow-2xl overflow-hidden">
      <div className="overflow-x-auto">
        <table className="min-w-full text-xs border-collapse">
          <thead>
            <tr className="bg-blue-900">
              <th className="px-3 py-3 text-center font-bold border border-blue-800 text-white uppercase sticky left-0 z-10" style={{ backgroundColor: '#203864', minWidth: '80px' }}>
                Acciones
              </th>
              {displayHeaders.map((header) => (
                <th
                  key={header}
                  className="px-3 py-3 text-center font-bold border border-blue-800 text-white uppercase tracking-wide"
                  style={{ 
                    backgroundColor: '#203864', 
                    minWidth: header.length > 30 ? '200px' : header.length > 20 ? '150px' : '100px',
                    fontSize: '10px'
                  }}
                >
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rowIndex) => (
              <tr 
                key={rowIndex}
                className="hover:bg-blue-100 transition-colors"
                style={{ backgroundColor: rowIndex % 2 === 0 ? '#E8F4F8' : '#FFFFFF' }}
              >
                <td className="px-2 py-2 border border-gray-300 text-center sticky left-0 z-10" style={{ backgroundColor: rowIndex % 2 === 0 ? '#E8F4F8' : '#FFFFFF' }}>
                  <button
                    onClick={() => onDeleteRow(rowIndex)}
                    className="bg-red-500 hover:bg-red-600 text-white px-3 py-1 rounded text-xs font-bold"
                  >
                    Eliminar
                  </button>
                </td>
                {displayHeaders.map((header) => (
                  <td 
                    key={header}
                    className="px-1 py-1 border border-gray-300"
                  >
                    {dropdownOptions[header] ? (
                      <select
                        value={ensureString(row[header]) || ""}
                        onChange={(e) => onCellChange(rowIndex, header, e.target.value)}
                        className="w-full px-2 py-1 text-center bg-transparent focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-400 rounded appearance-none cursor-pointer hover:bg-blue-50 transition-colors"
                        style={{ 
                          minWidth: '100px',
                          backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 10 10'%3E%3Cpath fill='%23888' d='M5 7L1 3h8z'/%3E%3C/svg%3E")`,
                          backgroundRepeat: 'no-repeat',
                          backgroundPosition: 'right 6px center',
                          paddingRight: '24px'
                        }}
                      >
                        <option value="">--</option>
                        {dropdownOptions[header].map((option) => (
                          <option key={option} value={option}>
                            {option}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <input
                        type="text"
                        value={ensureString(
                          header === 'CURSO' ? suggestCourse(row[header], row) :
                          header === 'SECCION' ? (row[header] || suggestSection(row['TEMA'] || row['CURSO'], row)) :
                          (row[header] || "")
                        )}
                        onChange={(e) => {
                          const newValue = ensureString(e.target.value);
                          if (header === 'CURSO' || header === 'TEMA') {
                            const courseInfo = extractCourseInfo(newValue);
                            onCellChange(rowIndex, header, ensureString(courseInfo.course));
                            if (courseInfo.section && headers.includes('SECCION')) {
                              onCellChange(rowIndex, 'SECCION', ensureString(courseInfo.section));
                            }
                          } else {
                            onCellChange(rowIndex, header, newValue);
                          }
                        }}
                        className="w-full px-2 py-1 text-center bg-transparent focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-400 rounded"
                        style={{ minWidth: '100px' }}
                      />
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export default DataTable;

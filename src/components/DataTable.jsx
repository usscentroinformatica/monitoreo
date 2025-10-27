import React from "react";

function DataTable({ 
  data, 
  headers, 
  dropdownOptions, 
  onCellChange, 
  onDeleteRow 
}) {
  // Helper function to extract information from course title
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

  const ensureString = (value) => {
    if (value === null || value === undefined) {
      return "";
    }
    
    // CASO ESPECIAL: Si es un string que parece ser toString() de un Date (contiene GMT)
    if (typeof value === "string" && value.includes("GMT")) {
      try {
        const dateObj = new Date(value);
        if (!isNaN(dateObj.getTime())) {
          const year = dateObj.getUTCFullYear();
          const hours = dateObj.getUTCHours();
          const minutes = dateObj.getUTCMinutes();
          const seconds = dateObj.getUTCSeconds();
          
          if (year === 1899 || year === 1900) {
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
          }
          
          if (hours !== 0 || minutes !== 0 || seconds !== 0) {
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
          }
        }
      } catch (e) {
        // Si falla, continuar con el resto del código
      }
    }
    
    // INTERCEPTAR OBJETOS DATE PRIMERO
    if (value instanceof Date) {
      const year = value.getUTCFullYear();
      const hours = value.getUTCHours();
      const minutes = value.getUTCMinutes();
      const seconds = value.getUTCSeconds();
      
      if (year === 1899 || year === 1900) {
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      }
      
      if (hours !== 0 || minutes !== 0 || seconds !== 0) {
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      }
      
      return value.toLocaleDateString('es-ES', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });
    }
    
    // Si ya es un string
    if (typeof value === "string") {
      const trimmed = value.trim();
      
      // Si parece ser una hora (HH:MM:SS), devolverla sin modificar
      if (/^\d{1,2}:\d{2}:\d{2}$/.test(trimmed)) {
        return trimmed;
      }
      
      // Si es muy largo, truncar
      if (trimmed.length > 100) {
        return trimmed.substring(0, 97) + '...';
      }
      
      return trimmed;
    }
    
    if (typeof value === "number") {
      // Caso especial: si es 0, devolverlo como "0"
      if (value === 0) {
        return "0";
      }
      
      // Verificar si es una fracción que representa tiempo (0-1)
      if (value > 0 && value < 1) {
        const totalSeconds = Math.round(value * 24 * 60 * 60);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      }
      
      // Verificar si es fecha de Excel (número grande)
      if (value > 1 && value < 100000) {
        try {
          const excelDate = new Date((value - 25569) * 86400 * 1000);
          if (!isNaN(excelDate.getTime()) && excelDate.getFullYear() > 1900) {
            return excelDate.toLocaleDateString('es-ES', {
              day: '2-digit',
              month: '2-digit',
              year: 'numeric'
            });
          }
        } catch (e) {
          // Si falla, devolver como número normal
        }
      }
      
      return String(value);
    }
    
    if (typeof value === "boolean") {
      return String(value);
    }
    
    if (Array.isArray(value)) {
      return value.join(", ");
    }
    
    if (typeof value === "object") {
      // Hipervínculos
      if (value.hyperlink !== undefined) {
        const text = String(value.text || value.hyperlink || '').trim();
        if (text.length > 50) {
          return text.substring(0, 47) + '...';
        }
        return text;
      }
      
      // Texto enriquecido
      if (value.richText !== undefined) {
        return value.richText.map(rt => rt.text || '').join('').trim();
      }
      
      // Texto directo
      if (value.text !== undefined) {
        return String(value.text).trim();
      }
      
      // Fórmulas
      if (value.formula !== undefined) {
        return String(value.result !== undefined ? value.result : value.formula).trim();
      }
      
      // Valores anidados
      if (value.value !== undefined) {
        if (value.value instanceof Date) {
          const hours = value.value.getUTCHours();
          const minutes = value.value.getUTCMinutes();
          const seconds = value.value.getUTCSeconds();
          
          if (hours !== 0 || minutes !== 0 || seconds !== 0) {
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
          } else {
            return value.value.toLocaleDateString('es-ES', {
              day: '2-digit',
              month: '2-digit',
              year: 'numeric'
            });
          }
        }
        return String(value.value).trim();
      }
      
      // Nombres
      if (value.name !== undefined) {
        return String(value.name).trim();
      }
      
      // Otros objetos
      try {
        const jsonStr = JSON.stringify(value);
        if (jsonStr === '{}' || jsonStr === 'null') return "";
        return jsonStr.length > 50 ? "[Datos complejos]" : jsonStr;
      } catch (e) {
        return "[Error en datos]";
      }
    }
    
    return String(value).trim();
  };

  // Filtrar headers vacíos y asegurar que tenemos headers válidos
  const displayHeaders = headers.filter(header => header && header.trim() !== "");

  if (!displayHeaders || displayHeaders.length === 0) {
    // Si no hay headers pero sí hay datos, generar headers automáticamente
    if (data && data.length > 0) {
      const firstRow = data[0];
      const autoHeaders = Object.keys(firstRow);
      if (autoHeaders.length > 0) {
        return (
          <DataTable 
            data={data} 
            headers={autoHeaders} 
            dropdownOptions={dropdownOptions} 
            onCellChange={onCellChange} 
            onDeleteRow={onDeleteRow} 
          />
        );
      }
    }
    
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
                        value={(() => {
                          const val = row[header];
                          // Si es null, undefined, o string vacío, usar ""
                          if (val === null || val === undefined || val === '') return "";
                          // Si es un string, usarlo directamente (sin ensureString que puede modificarlo)
                          if (typeof val === 'string') return val.trim();
                          // Para otros tipos, convertir a string
                          return String(val);
                        })()}
                        onChange={(e) => {
                          console.log(`📝 Dropdown cambió - Fila: ${rowIndex}, Columna: ${header}, Nuevo valor: "${e.target.value}"`);
                          onCellChange(rowIndex, header, e.target.value);
                        }}
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
                    ) : (() => {
                      const cellValue = ensureString(
                        header === 'CURSO' ? suggestCourse(row[header], row) :
                        header === 'SECCION' || header === 'SECCIÓN' ? (row[header] || suggestSection(row['TEMA'] || row['CURSO'], row)) :
                        (row[header] !== undefined ? row[header] : "")
                      );
                      
                      // Detectar si es fecha, hora o texto largo
                      const isTime = /^\d{1,2}:\d{2}(:\d{2})?$/.test(cellValue);
                      const isDate = /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cellValue) || /^\d{1,2}-\d{1,2}-\d{4}$/.test(cellValue);
                      const isLongText = cellValue.length > 50;
                      
                      // Determinar clases CSS y estilos
                      let cellClass = "w-full px-2 py-1 text-center bg-transparent focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-400 rounded";
                      let cellStyle = {
                        minWidth: header.includes('AULA') || header.length > 20 ? '200px' : '100px',
                        fontSize: header.includes('AULA') || header.length > 20 ? '10px' : '12px'
                      };
                      
                      if (isLongText) {
                        cellClass += " cell-truncated";
                        cellStyle.maxWidth = "200px";
                        cellStyle.whiteSpace = "nowrap";
                        cellStyle.overflow = "hidden";
                        cellStyle.textOverflow = "ellipsis";
                      }
                      
                      return (
                        <input
                          type="text"
                          value={cellValue}
                          onChange={(e) => {
                            const newValue = ensureString(e.target.value);
                            if (header === 'CURSO' || header === 'TEMA') {
                              const courseInfo = extractCourseInfo(newValue);
                              onCellChange(rowIndex, header, ensureString(courseInfo.course));
                              if (courseInfo.section && (headers.includes('SECCION') || headers.includes('SECCIÓN'))) {
                                onCellChange(rowIndex, headers.includes('SECCION') ? 'SECCION' : 'SECCIÓN', ensureString(courseInfo.section));
                              }
                            } else {
                              onCellChange(rowIndex, header, newValue);
                            }
                          }}
                          className={cellClass}
                          style={cellStyle}
                          title={isLongText ? cellValue : undefined} // Tooltip para texto largo
                        />
                      );
                    })()}
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
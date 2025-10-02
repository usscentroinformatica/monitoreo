import React from "react";

function DataTable({ 
  data, 
  headers, 
  dropdownOptions, 
  onCellChange, 
  onDeleteRow 
}) {
  return (
    <div className="bg-white rounded-xl shadow-2xl overflow-hidden">
      <div className="overflow-x-auto">
        <table className="min-w-full text-xs border-collapse">
          <thead>
            <tr className="bg-blue-900">
              <th className="px-3 py-3 text-center font-bold border border-blue-800 text-white uppercase sticky left-0 z-10" style={{ backgroundColor: '#203864', minWidth: '80px' }}>
                Acciones
              </th>
              {headers.map((header) => (
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
                {headers.map((header) => (
                  <td 
                    key={header}
                    className="px-1 py-1 border border-gray-300"
                  >
                    {dropdownOptions[header] ? (
                      <select
                        value={row[header] || ""}
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
                        value={row[header] || ""}
                        onChange={(e) => onCellChange(rowIndex, header, e.target.value)}
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
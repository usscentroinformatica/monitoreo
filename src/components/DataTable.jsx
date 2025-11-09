import React, { memo, useCallback, useRef, useEffect, useState } from "react";

const DataTable = memo(({ data, headers, dropdownOptions = {}, onCellChange, onDeleteRow }) => {
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

  // Handler optimizado para cambios en celdas
  const handleCellChange = useCallback((rowIndex, header, value) => {
    onCellChange(rowIndex, header, value);
  }, [onCellChange]);

  // Handler optimizado para eliminar filas
  const handleDeleteRow = useCallback((rowIndex) => {
    onDeleteRow(rowIndex);
  }, [onDeleteRow]);

  // Refs y estado para barra horizontal fija y sincronizada
  const scrollContainerRef = useRef(null);
  const stickyScrollbarRef = useRef(null);
  const [dummyWidth, setDummyWidth] = useState(0);

  // Sincronizar desplazamiento horizontal entre la tabla y la barra fija
  useEffect(() => {
    const main = scrollContainerRef.current;
    const fake = stickyScrollbarRef.current;
    if (!main || !fake) return;

    const onMainScroll = () => {
      if (fake.scrollLeft !== main.scrollLeft) fake.scrollLeft = main.scrollLeft;
    };
    const onFakeScroll = () => {
      if (main.scrollLeft !== fake.scrollLeft) main.scrollLeft = fake.scrollLeft;
    };

    main.addEventListener('scroll', onMainScroll, { passive: true });
    fake.addEventListener('scroll', onFakeScroll, { passive: true });

    return () => {
      main.removeEventListener('scroll', onMainScroll);
      fake.removeEventListener('scroll', onFakeScroll);
    };
  }, []);

  // Actualizar ancho del dummy para que la barra aparezca siempre
  useEffect(() => {
    const main = scrollContainerRef.current;
    if (!main) return;
    const updateWidth = () => setDummyWidth(main.scrollWidth);
    updateWidth();
    window.addEventListener('resize', updateWidth);
    return () => window.removeEventListener('resize', updateWidth);
  }, [data, displayHeaders]);

  return (
    <div className="bg-white rounded-xl shadow-2xl overflow-hidden">
      <div className="relative max-h-[70vh] overflow-y-auto">
        <div ref={scrollContainerRef} className="overflow-x-auto">
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
                    <div className="flex items-center justify-center">
                      <span 
                        className="truncate max-w-full"
                        title={header}
                        style={{
                          display: 'inline-block',
                          maxWidth: header.length > 30 ? '180px' : header.length > 20 ? '130px' : '80px',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'nowrap'
                        }}
                      >
                        {header}
                      </span>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.map((row, rowIndex) => (
                <TableRow 
                  key={rowIndex}
                  row={row}
                  rowIndex={rowIndex}
                  displayHeaders={displayHeaders}
                  dropdownOptions={dropdownOptions}
                  onCellChange={handleCellChange}
                  onDeleteRow={handleDeleteRow}
                />
              ))}
            </tbody>
          </table>
        </div>
        {/* Barra horizontal fija y sincronizada */}
        <div
          ref={stickyScrollbarRef}
          className="sticky bottom-0 left-0 right-0 h-4 overflow-x-auto bg-gray-100 border-t border-gray-300"
          style={{ zIndex: 20 }}
        >
          <div style={{ width: dummyWidth, height: 1 }} />
        </div>
      </div>
    </div>
  );
});

// Componente memoizado para las filas
const TableRow = memo(({ row, rowIndex, displayHeaders, dropdownOptions = {}, onCellChange, onDeleteRow }) => {
  return (
    <tr 
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
        <TableCell 
          key={header}
          header={header}
          value={row[header]}
          rowIndex={rowIndex}
          dropdownOptions={dropdownOptions}
          onCellChange={onCellChange}
        />
      ))}
    </tr>
  );
});

// Componente memoizado para las celdas
const TableCell = memo(({ header, value, rowIndex, dropdownOptions = {}, onCellChange }) => {
  const handleChange = useCallback((e) => {
    onCellChange(rowIndex, header, e.target.value);
  }, [rowIndex, header, onCellChange]);

  const hasDropdown = dropdownOptions?.[header];
  const hasValue = value && String(value).trim() !== "";

  return (
    <td className="px-1 py-1 border border-gray-300">
      {hasDropdown ? (
        hasValue ? (
          <select
            value={value || ""}
            onChange={handleChange}
            className="w-full px-2 py-1 text-center bg-transparent focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-400 rounded appearance-none cursor-pointer hover:bg-blue-50 transition-colors"
            style={{ 
              minWidth: '100px',
              backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 10 10'%3E%3Cpath fill='%23888' d='M5 7L1 3h8z'/%3E%3C/svg%3E")`,
              backgroundRepeat: 'no-repeat',
              backgroundPosition: 'right 6px center',
              paddingRight: '24px'
            }}
          >
            <option value=""></option>
            {(dropdownOptions[header] || []).map((option) => (
              <option key={option} value={option}>{option}</option>
            ))}
          </select>
        ) : (
          <>
            <input
              list={`list-${header}`}
              value={value || ""}
              onChange={handleChange}
              className="w-full px-2 py-1 text-center bg-transparent focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-400 rounded appearance-none hover:bg-blue-50 transition-colors"
              style={{ minWidth: '100px' }}
            />
            <datalist id={`list-${header}`}>
              {(dropdownOptions[header] || []).map((option) => (
                <option key={option} value={option} />
              ))}
            </datalist>
          </>
        )
      ) : (
        <input
          type="text"
          value={value || ""}
          onChange={handleChange}
          className="w-full px-2 py-1 text-center bg-transparent focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-400 rounded appearance-none hover:bg-blue-50 transition-colors"
          style={{ minWidth: '100px' }}
        />
      )}
    </td>
  );
});

export default DataTable;
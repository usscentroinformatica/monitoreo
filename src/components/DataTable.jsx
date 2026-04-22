import React, { memo, useCallback, useRef, useEffect, useState, useMemo } from "react";

// Función para convertir "HH:MM:SS" a segundos totales
const timeToSeconds = (timeStr) => {
  if (!timeStr || typeof timeStr !== 'string') return 0;
  const match = timeStr.match(/(\d+):(\d+):(\d+)/);
  if (!match) return 0;
  return parseInt(match[1]) * 3600 + parseInt(match[2]) * 60 + parseInt(match[3]);
};

// Función para convertir hora con AM/PM a segundos desde medianoche
const hourToSeconds = (hourStr) => {
  if (!hourStr || typeof hourStr !== 'string') return 0;
  const match = hourStr.match(/^(\d{1,2}):(\d{2}):(\d{2})\s+(AM|PM)$/i);
  if (!match) return 0;

  let hours = parseInt(match[1]);
  const minutes = parseInt(match[2]);
  const seconds = parseInt(match[3]);
  const ampm = match[4].toUpperCase();

  if (ampm === 'PM' && hours !== 12) hours += 12;
  if (ampm === 'AM' && hours === 12) hours = 0;

  return hours * 3600 + minutes * 60 + seconds;
};

// Función para convertir segundos a formato "HH:MM:SS AM/PM"
const secondsToHour = (totalSeconds) => {
  let hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12 || 12;
  return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')} ${ampm}`;
};

// Función para truncar hora al minuto exacto
const truncateToMinute = (value) => {
  if (!value || typeof value !== 'string') return value;
  const match = value.match(/^(\d{1,2}):(\d{2}):(\d{2})\s+(AM|PM)$/i);
  if (!match) return value;
  const hours = match[1].padStart(2, '0');
  const minutes = match[2];
  const ampm = match[4].toUpperCase();
  return `${hours}:${minutes}:00 ${ampm}`;
};

// Función para calcular INICIO REAL CLASE automáticamente


const getByAliases = (row, aliases) => {
  for (const k of aliases) {
    if (row && Object.prototype.hasOwnProperty.call(row, k)) return row[k];
  }
  return '';
};

const firstExistingAlias = (aliases, row) => aliases.find((k) => row && Object.prototype.hasOwnProperty.call(row, k));

// Función para verificar si una fila está vacía (es un separador)
const isRowEmpty = (row) => {
  if (!row) return true;
  return Object.values(row).every(value => 
    value === null || 
    value === undefined || 
    value === '' || 
    (typeof value === 'string' && value.trim() === '')
  );
};

// Función para obtener el valor numérico de la sesión
const getSessionNumber = (sessionValue) => {
  if (!sessionValue && sessionValue !== 0) return null;
  const num = parseInt(sessionValue);
  return isNaN(num) ? null : num;
};

// Función para insertar separadores cada 16 sesiones
const insertSeparatorsBySession = (data) => {
  if (!data || data.length === 0) return data;
  
  const result = [];
  let lastSessionNum = null;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const sessionValue = row?.["SESION"];
    const currentSessionNum = getSessionNumber(sessionValue);
    
    if (isRowEmpty(row)) {
      result.push(row);
      continue;
    }
    
    if (currentSessionNum === 1 && lastSessionNum === 16) {
      const emptyRow = {};
      if (row) {
        Object.keys(row).forEach(key => {
          emptyRow[key] = '';
        });
      }
      result.push(emptyRow);
    }
    
    result.push(row);
    lastSessionNum = currentSessionNum;
  }
  
  return result;
};

// Función para normalizar datos
const normalizeData = (data) => {
  if (!data || data.length === 0) return data;
  
  const cleanedData = [];
  let lastWasEmpty = false;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const isEmpty = isRowEmpty(row);
    
    if (isEmpty && lastWasEmpty) {
      continue;
    }
    
    cleanedData.push(row);
    lastWasEmpty = isEmpty;
  }
  
  return insertSeparatorsBySession(cleanedData);
};

const DataTable = memo(({ data, headers, dropdownOptions = {}, onCellChange, onDeleteRow }) => {
  const displayHeaders = useMemo(() => {
    const baseHeaders = headers.filter(header => header && header.trim() !== "");
    const norm = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
    const existing = new Map(baseHeaders.map((h) => [norm(h), h]));

    const choose = (preferredAliases) => {
      for (const alias of preferredAliases) {
        const found = existing.get(norm(alias));
        if (found) return found;
      }
      return null;
    };

    const aliasGroups = [
      ['HORA INICIO', 'Hora Inicio', 'INICIO', 'inicio'],
      ['HORA FIN', 'Hora Fin', 'FIN', 'fin'],
      ['HORAS PROGRAMADAS', 'Horas Programadas', 'TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACION PROGRAMADA', 'DURACIÓN PROGRAMADA', 'Duración Programada'],
      ['DURACION TOTAL CLASE', 'DURACIÓN TOTAL CLASE', 'Duración total clase', 'FINALIZA LA CLASE (ZOOM)', 'Hora Finalización Zoom'],
      ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'],
      ['H.I', 'H.I.', 'HI', 'H I', 'INICIO A LA HORA', 'Inicio a la Hora'],
      ['H.F', 'H.F.', 'HF', 'H F', 'FIN A LA HORA', 'Fin a la Hora'],
      ['INICIO REAL CLASE', 'Inicio Real Clase', 'INICIO SESION 10 A 5 MINUTOS ANTES DE INICIAR CLASE', 'INICIO SESION 10 A 5 MINUTOS ANTES DE INICIAR LA CLASE', 'Inicio sesion 10 a 5 minutos antes de iniciar clase']
    ];

    const hidden = new Set();
    for (const group of aliasGroups) {
      const keep = choose(group);
      if (!keep) continue;
      const keepNorm = norm(keep);
      group.forEach((alias) => {
        const found = existing.get(norm(alias));
        if (found && norm(found) !== keepNorm) {
          hidden.add(norm(found));
        }
      });
    }

    return baseHeaders.filter((h) => !hidden.has(norm(h)));
  }, [headers]);
  const [filters, setFilters] = useState({});
  const [openFilter, setOpenFilter] = useState(null);

  const handleFilterChange = useCallback((header, value) => {
    setFilters(prev => ({
      ...prev,
      [header]: value
    }));
    setOpenFilter(null); // Cerrar el dropdown después de seleccionar
  }, []);

  const toggleFilter = useCallback((header) => {
    setOpenFilter(openFilter === header ? null : header);
  }, [openFilter]);

  // Cerrar dropdown cuando se hace clic fuera
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (!event.target.closest('.filter-dropdown')) {
        setOpenFilter(null);
      }
    };

    if (openFilter) {
      document.addEventListener('mousedown', handleClickOutside);
      return () => document.removeEventListener('mousedown', handleClickOutside);
    }
  }, [openFilter]);

  if (!displayHeaders || displayHeaders.length === 0) {
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

  const indexedData = useMemo(() => {
    return (data || []).map((row, index) => {
      if (!row || typeof row !== 'object') return row;
      return { ...row, __sourceIndex: index };
    });
  }, [data]);

  const processedData = useMemo(() => {
    return normalizeData(indexedData);
  }, [indexedData]);

  const filteredData = useMemo(() => {
    if (Object.keys(filters).length === 0) {
      return processedData.map((row, index) => ({ row, originalIndex: index }));
    }
    
    return processedData
      .map((row, index) => ({ row, originalIndex: index }))
      .filter(({ row }) => {
        return Object.entries(filters).every(([header, filterValue]) => {
          if (!filterValue || filterValue === '') return true;
          const cellValue = row[header];
          if (!cellValue) return false;
          
          // Para SESION, hacer coincidencia exacta (no parcial)
          if (header === 'SESION' || header === 'Sesión') {
            return String(cellValue).trim() === String(filterValue).trim();
          }
          
          // Para otros campos, usar coincidencia parcial (includes)
          
          // Para SESION, hacer coincidencia exacta (no parcial)
          if (header === 'SESION' || header === 'Sesión') {
            return String(cellValue).trim() === String(filterValue).trim();
          }
          
          // Para otros campos, usar coincidencia parcial (includes)
          return String(cellValue).toLowerCase().includes(String(filterValue).toLowerCase());
        });
      });
  }, [processedData, filters]);

  const metricHeaders = useMemo(() => {
    const normalize = (text) =>
      String(text || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, '');

    const findHeader = (aliases) => {
      const aliasSet = new Set(aliases.map((a) => normalize(a)));
      return displayHeaders.find((h) => aliasSet.has(normalize(h))) || null;
    };

    return {
      eficienciaHeader: findHeader(['EFICIENCIA', 'INDICE EFICIENCIA', 'ÍNDICE DE EFICIENCIA']),
      hiHeader: findHeader(['H.I', 'HI', 'INICIO A LA HORA']),
      hfHeader: findHeader(['H.F', 'HF', 'FIN A LA HORA'])
    };
  }, [displayHeaders]);

  const calculateMetricsForRows = useCallback((rows) => {
    const { eficienciaHeader, hiHeader, hfHeader } = metricHeaders;
    const validRows = (rows || []).filter((row) => !isRowEmpty(row));
    const totalRows = validRows.length;

    let eficienciaSum = 0;
    let eficienciaCount = 0;
    let hiSiCount = 0;
    let hfSiCount = 0;

    validRows.forEach((row) => {
      if (eficienciaHeader) {
        const raw = String(row?.[eficienciaHeader] || '').trim();
        const numeric = raw.match(/(\d+(?:[.,]\d+)?)/);
        if (numeric) {
          const parsed = parseFloat(numeric[1].replace(',', '.'));
          if (Number.isFinite(parsed)) {
            eficienciaSum += parsed;
            eficienciaCount += 1;
          }
        }
      }

      if (hiHeader) {
        const hiValue = String(row?.[hiHeader] || '').trim().toUpperCase();
        if (hiValue === 'SI' || hiValue === 'SÍ') hiSiCount += 1;
      }

      if (hfHeader) {
        const hfValue = String(row?.[hfHeader] || '').trim().toUpperCase();
        if (hfValue === 'SI' || hfValue === 'SÍ') hfSiCount += 1;
      }
    });

    const eficienciaPromedio = eficienciaCount > 0 ? (eficienciaSum / eficienciaCount) : NaN;
    const hiPct = totalRows > 0 ? (hiSiCount / totalRows) * 100 : NaN;
    const hfPct = totalRows > 0 ? (hfSiCount / totalRows) * 100 : NaN;

    return {
      totalRows,
      eficienciaPromedio,
      hiPct,
      hfPct,
      hiSiCount,
      hfSiCount
    };
  }, [metricHeaders]);

  const summaryMetrics = useMemo(() => {
    const rows = filteredData.map(({ row }) => row).filter((row) => !isRowEmpty(row));
    const metrics = calculateMetricsForRows(rows);

    return {
      ...metricHeaders,
      ...metrics
    };
  }, [filteredData, metricHeaders, calculateMetricsForRows]);

  const cycleSeparatorSummaryByIndex = useMemo(() => {
    const summaries = {};
    let cycleRows = [];

    for (let idx = 0; idx < filteredData.length; idx += 1) {
      const currentRow = filteredData[idx]?.row;
      const currentIsEmpty = isRowEmpty(currentRow);
      const isSessionSeparator = currentIsEmpty && idx > 0 && idx < filteredData.length - 1 &&
        getSessionNumber(filteredData[idx - 1]?.row?.['SESION']) === 16 &&
        getSessionNumber(filteredData[idx + 1]?.row?.['SESION']) === 1;

      if (isSessionSeparator) {
        const metrics = calculateMetricsForRows(cycleRows);
        summaries[idx] = {
          eficienciaPct: Number.isFinite(metrics.eficienciaPromedio) ? `${metrics.eficienciaPromedio.toFixed(2)}%` : '',
          hiPct: metrics.totalRows > 0 && Number.isFinite(metrics.hiPct) ? `${metrics.hiPct.toFixed(2)}%` : '',
          hfPct: metrics.totalRows > 0 && Number.isFinite(metrics.hfPct) ? `${metrics.hfPct.toFixed(2)}%` : ''
        };
        cycleRows = [];
        continue;
      }

      if (!currentIsEmpty) {
        cycleRows.push(currentRow);
      }
    }

    return summaries;
  }, [filteredData, calculateMetricsForRows]);

  const handleCellChange = useCallback((sourceIndex, header, value) => {
    if (!Number.isInteger(sourceIndex) || sourceIndex < 0) return;

    const currentRow = data[sourceIndex];
    if (!currentRow) return;
    
    if (isRowEmpty(currentRow)) return;
    
    onCellChange(sourceIndex, header, value);
  }, [data, onCellChange]);

  const handleDeleteRow = useCallback((sourceIndex) => {
    if (!Number.isInteger(sourceIndex) || sourceIndex < 0) return;

    const currentRow = data[sourceIndex];
    if (isRowEmpty(currentRow)) return;
    onDeleteRow(sourceIndex);
  }, [data, onDeleteRow]);

  const scrollContainerRef = useRef(null);
  const stickyScrollbarRef = useRef(null);
  const [dummyWidth, setDummyWidth] = useState(0);

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

  useEffect(() => {
    const main = scrollContainerRef.current;
    if (!main) return;
    const updateWidth = () => setDummyWidth(main.scrollWidth);
    updateWidth();
    const resizeObserver = new ResizeObserver(() => updateWidth());
    resizeObserver.observe(main);
    return () => resizeObserver.disconnect();
  }, []);

  return (
    <div className="bg-white rounded-xl shadow-2xl overflow-hidden">
      <div className="relative max-h-[70vh] overflow-y-auto">
        <div ref={scrollContainerRef} className="overflow-x-auto">
          <table className="min-w-full text-xs border-collapse">
            <thead>
              <tr className="bg-blue-900">
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
                    <div className="flex flex-col items-center justify-center relative">
                      <span
                        className="block w-full text-center mb-1"
                        title={header}
                        style={{
                          maxWidth: header.length > 30 ? '180px' : header.length > 20 ? '130px' : '80px',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'normal',
                          wordBreak: 'break-word',
                          lineHeight: '1.1',
                          maxHeight: '2.2em',
                          display: '-webkit-box',
                          WebkitLineClamp: 2,
                          WebkitBoxOrient: 'vertical'
                        }}
                      >
                        {header}
                      </span>
                      <div className="flex items-center justify-center">
                        <button
                          onClick={() => toggleFilter(header)}
                          className={`filter-dropdown p-1 rounded transition-colors duration-200 relative ${
                            filters[header] ? 'bg-blue-700 text-white shadow-md' : 'bg-blue-500 hover:bg-blue-600 text-white'
                          }`}
                          title={`Filtrar por ${header}${filters[header] ? ` (filtrando: ${filters[header]})` : ''}`}
                        >
                          <svg
                            className={`w-3 h-3 transition-transform duration-200 ${openFilter === header ? 'rotate-180' : ''}`}
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                          >
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                          </svg>
                          {filters[header] && (
                            <div className="absolute -top-1 -right-1 w-2 h-2 bg-green-400 rounded-full border border-white"></div>
                          )}
                        </button>
                      </div>
                      
                      {openFilter === header && (
                        <div className="filter-dropdown absolute top-full mt-1 bg-white border border-gray-300 rounded-md shadow-lg z-20 max-h-48 overflow-y-auto min-w-32">
                          <div
                            className="px-3 py-2 hover:bg-gray-100 cursor-pointer text-sm text-gray-700 border-b border-gray-200 font-medium"
                            onClick={() => handleFilterChange(header, '')}
                          >
                            📂 Todos
                          </div>
                          {Array.from(new Set(processedData.map(row => row[header]).filter(val => val !== null && val !== undefined && val !== '')))
                            .sort((a, b) => {
                              const numA = Number(a);
                              const numB = Number(b);
                              if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
                              return String(a).localeCompare(String(b), 'es', { sensitivity: 'base' });
                            })
                            .map(value => (
                            <div
                              key={value}
                              className={`px-3 py-2 hover:bg-gray-100 cursor-pointer text-sm ${
                                filters[header] === value ? 'bg-blue-50 text-blue-700 font-medium' : 'text-gray-700'
                              }`}
                              onClick={() => handleFilterChange(header, value)}
                            >
                              {String(value).length > 25 ? String(value).substring(0, 25) + '...' : value}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </th>
                ))}
               </tr>
            </thead>
            <tbody>
              {filteredData.map(({ row, originalIndex }, rowIndex) => {
                const isEmpty = isRowEmpty(row);
                const isSessionSeparator = isEmpty && rowIndex > 0 && rowIndex < filteredData.length - 1 &&
                  getSessionNumber(filteredData[rowIndex - 1]?.row?.["SESION"]) === 16 && 
                  getSessionNumber(filteredData[rowIndex + 1]?.row?.["SESION"]) === 1;
                const separatorSummaryMetrics = isSessionSeparator
                  ? cycleSeparatorSummaryByIndex[rowIndex]
                  : null;
                const sourceIndex = Number.isInteger(row?.__sourceIndex) ? row.__sourceIndex : -1;
                
                return (
                  <TableRow
                    key={originalIndex}
                    row={row}
                    rowIndex={sourceIndex}
                    displayHeaders={displayHeaders}
                    dropdownOptions={dropdownOptions}
                    onCellChange={handleCellChange}
                    onDeleteRow={handleDeleteRow}
                    isEmpty={isEmpty}
                    isSessionSeparator={isSessionSeparator}
                    separatorSummaryMetrics={separatorSummaryMetrics}
                  />
                );
              })}
            </tbody>
            <tfoot>
              <tr className="bg-[#dbe7fb] border-t-2 border-blue-300">
                {displayHeaders.map((header) => {
                  const isEficiencia = summaryMetrics.eficienciaHeader === header;
                  const isHI = summaryMetrics.hiHeader === header;
                  const isHF = summaryMetrics.hfHeader === header;

                  let text = '';
                  if (isEficiencia) {
                    text = Number.isFinite(summaryMetrics.eficienciaPromedio)
                      ? `${summaryMetrics.eficienciaPromedio.toFixed(2)}%`
                      : '';
                  } else if (isHI) {
                    text = summaryMetrics.totalRows > 0
                      ? `${summaryMetrics.hiPct.toFixed(2)}%`
                      : '';
                  } else if (isHF) {
                    text = summaryMetrics.totalRows > 0
                      ? `${summaryMetrics.hfPct.toFixed(2)}%`
                      : '';
                  }

                  return (
                    <td
                      key={`summary-${header}`}
                      className="px-2 py-2 border border-blue-200 text-center font-bold text-blue-900"
                      style={{ minWidth: header.length > 30 ? '200px' : header.length > 20 ? '150px' : '100px' }}
                      title={undefined}
                    >
                      {text}
                    </td>
                  );
                })}
              </tr>
            </tfoot>
          </table>
        </div>
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

const TableRow = memo(({ row, rowIndex, displayHeaders, dropdownOptions = {}, onCellChange, onDeleteRow, isEmpty = false, isSessionSeparator = false, separatorSummaryMetrics = null }) => {
  if (isEmpty) {
    if (isSessionSeparator) {
      const normalize = (text) =>
        String(text || '')
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .toUpperCase()
          .replace(/[^A-Z0-9]/g, '');

      return (
        <tr className="bg-blue-100 border-t-2 border-b-2 border-blue-300">
          {displayHeaders.map((header) => {
            const headerNorm = normalize(header);
            const isEficiencia = headerNorm.includes('EFICIENCIA');
            const isHI = headerNorm === 'HI' || headerNorm === 'INICIOALAHORA';
            const isHF = headerNorm === 'HF' || headerNorm === 'FINALAHORA';

            let value = '';
            if (isEficiencia) value = separatorSummaryMetrics?.eficienciaPct || '';
            if (isHI) value = separatorSummaryMetrics?.hiPct || '';
            if (isHF) value = separatorSummaryMetrics?.hfPct || '';

            return (
              <td
                key={`cycle-separator-${header}`}
                className="px-2 py-2 border border-blue-200 text-center text-sm font-bold text-blue-900"
              >
                {value}
              </td>
            );
          })}
        </tr>
      );
    }

    return (
      <tr className="bg-gray-100 border-t-2 border-b-2 border-gray-400">
        <td 
          colSpan={displayHeaders.length} 
          className="px-2 py-3 text-center text-gray-500 italic text-sm font-semibold"
          style={{ backgroundColor: '#F3F4F6' }}
        >
          {'──────────────────────────────────────────────────'}
        </td>
      </tr>
    );
  }

  return (
    <tr
      className="hover:bg-blue-100 transition-colors"
      style={{ backgroundColor: rowIndex % 2 === 0 ? '#E8F4F8' : '#FFFFFF' }}
    >
      {displayHeaders.map((header) => (
        <TableCell
          key={header}
          header={header}
          value={row[header]}
          rowIndex={rowIndex}
          rowData={row}
          dropdownOptions={dropdownOptions}
          onCellChange={onCellChange}
        />
      ))}
    </tr>
  );
});

const TableCell = memo(({ header, value, rowIndex, rowData, dropdownOptions = {}, onCellChange }) => {
  const isRowEmpty = !rowData || Object.values(rowData).every(v => !v || (typeof v === 'string' && v.trim() === ''));

  // Estado local SOLO para el spinner de tiempo de espera (display instantáneo)
  const [localWait, setLocalWait] = useState(value);
  const pendingWaitRef = useRef(null); // guarda el valor más reciente pendiente de commit
  const waitTimerRef = useRef(null);

  // Sincronizar cuando el valor externo cambia (ej: autocomplete masivo)
  useEffect(() => {
    setLocalWait(value);
  }, [value]);

  // Limpiar timer al desmontar
  useEffect(() => () => { if (waitTimerRef.current) clearTimeout(waitTimerRef.current); }, []);

  const handleChange = useCallback((e) => {
    if (isRowEmpty) return;
    const newValue = e.target.value;
    onCellChange(rowIndex, header, newValue);
  }, [rowIndex, header, onCellChange, isRowEmpty, isRowEmpty]);

  const hasDropdown = dropdownOptions?.[header];
  const hasValue = value && String(value).trim() !== "";

  const hNorm = String(header || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
  const hCompact = hNorm.replace(/[^A-Z0-9]/g, '');
  const isEfficiency = hNorm.includes('EFICIENCIA');
  const isHI = hCompact === 'HI' || hCompact === 'INICIOALAHORA';
  const isHF = hCompact === 'HF' || hCompact === 'FINALAHORA';
  const isWaitBeforeStart = hNorm.includes('TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE') || hNorm === 'TIEMPO DE ESPERA' || hNorm.includes('ESPERA ANTES DE INICIAR');
  let cellBg = undefined;
  let titleMsg = undefined;

  const parseWaitDurationParts = useCallback((rawValue) => {
    const empty = { hours: '', minutes: '', seconds: '' };
    if (rawValue === null || rawValue === undefined) return empty;

    const trimmed = String(rawValue).trim();
    if (trimmed === '') return empty;

    const hhmmss = trimmed.match(/^(\d{1,3}):(\d{1,2})(?::(\d{1,2}))?$/);
    if (hhmmss) {
      const h = Math.max(0, parseInt(hhmmss[1] || '0', 10) || 0);
      const m = Math.min(59, Math.max(0, parseInt(hhmmss[2] || '0', 10) || 0));
      const s = Math.min(59, Math.max(0, parseInt(hhmmss[3] || '0', 10) || 0));
      return { hours: String(h), minutes: String(m), seconds: String(s) };
    }

    const onlyMinutesMatch = trimmed.match(/\d+/);
    if (onlyMinutesMatch) {
      const totalMinutes = Math.max(0, parseInt(onlyMinutesMatch[0], 10) || 0);
      const h = Math.floor(totalMinutes / 60);
      const m = totalMinutes % 60;
      return { hours: String(h), minutes: String(m), seconds: '0' };
    }

    return empty;
  }, []);

  const toHHMMSS = (parts) => {
    const hasAny = parts.hours !== '' || parts.minutes !== '' || parts.seconds !== '';
    if (!hasAny) return '';

    const h = Math.max(0, parseInt(parts.hours || '0', 10) || 0);
    const m = Math.min(59, Math.max(0, parseInt(parts.minutes || '0', 10) || 0));
    const s = Math.min(59, Math.max(0, parseInt(parts.seconds || '0', 10) || 0));
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  };

  const handleWaitDurationPartChange = useCallback((part, rawValue) => {
    if (isRowEmpty) return;

    // Calcular nuevo valor usando el estado local más reciente
    const currentParts = parseWaitDurationParts(localWait);
    const sanitized = String(rawValue || '').replace(/[^\d]/g, '');

    const nextParts = { ...currentParts };
    if (sanitized === '') {
      nextParts[part] = '';
    } else {
      const parsed = parseInt(sanitized, 10);
      if (Number.isNaN(parsed)) return;

      if (part === 'hours') {
        nextParts.hours = String(Math.max(0, parsed));
      } else {
        nextParts[part] = String(Math.min(59, Math.max(0, parsed)));
      }
    }

    const formatted = toHHMMSS(nextParts);

    // Actualizar display local INMEDIATAMENTE (sin lag)
    setLocalWait(formatted);
    pendingWaitRef.current = formatted;

    // Disparar lógica pesada solo tras 600ms sin más clicks
    if (waitTimerRef.current) clearTimeout(waitTimerRef.current);
    waitTimerRef.current = setTimeout(() => {
      onCellChange(rowIndex, header, pendingWaitRef.current);
    }, 600);
  }, [rowIndex, header, onCellChange, isRowEmpty, parseWaitDurationParts, localWait]);

  if (isEfficiency && !isRowEmpty) {
    const raw = String(value ?? '').trim();
    const numMatch = raw.match(/(\d+(?:[.,]\d+)?)/);
    const num = numMatch ? parseFloat(numMatch[1].replace(',', '.')) : NaN;
    if (Number.isFinite(num)) {
      if (num === 100) {
        cellBg = '#dcfce7';
        titleMsg = '🌟 ¡Excelente! Eficiencia perfecta.';
      } else if (num >= 86 && num < 100) {
        cellBg = '#fef08a';
        titleMsg = '👍 ¡Buen trabajo! Aunque hay un pequeño margen para mejorar y llegar al 100%.';
      } else if (num >= 80 && num < 86) {
        cellBg = '#fed7aa';
        titleMsg = '⚠️ Aceptable, pero pon atención: ¡la eficiencia debe mejorar pronto!';
      } else if (num < 80) {
        cellBg = '#fee2e2';
        titleMsg = '🚨 Atención: La eficiencia está por debajo de lo esperado. ¡Es necesario mejorar!';
      }
    }
  }

  if ((isHI || isHF) && !isRowEmpty) {
    const vNorm = String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
    if (vNorm === 'SI') {
      cellBg = '#dcfce7';
      titleMsg = isHI
        ? 'Inicia dentro del rango permitido.'
        : 'Finaliza en la hora esperada o despues.';
    } else if (vNorm === 'NO') {
      cellBg = '#fee2e2';
      titleMsg = isHI
        ? 'No inicia dentro del rango de tolerancia.'
        : 'Finaliza antes de la hora esperada.';
    }
  }

  const displayValue = useMemo(() => {
    if (isRowEmpty) return '';
    return value;
  }, [value, isRowEmpty]);

  const waitDurationParts = useMemo(() => parseWaitDurationParts(localWait), [localWait, parseWaitDurationParts]);

  if (isRowEmpty) {
    return (
      <td className="px-1 py-2 border border-gray-300 text-center bg-gray-100">
        <span className="text-gray-300">─</span>
      </td>
    );
  }

  return (
    <td className="px-1 py-1 border border-gray-300 relative group" style={{ backgroundColor: cellBg }}>
      {titleMsg && (
        <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 invisible opacity-0 group-hover:visible group-hover:opacity-100 z-[99] w-48 p-3 bg-gray-900 text-white rounded-lg shadow-2xl transition-all duration-300 transform scale-95 group-hover:scale-100 pointer-events-none">
          <p className="text-xs font-bold leading-relaxed text-center drop-shadow-sm">
            {titleMsg}
          </p>
          <div className="absolute top-full left-1/2 -translate-x-1/2 border-8 border-transparent border-t-gray-900"></div>
        </div>
      )}

      {hasDropdown ? (
        hasValue ? (
          <select
            value={displayValue || ""}
            onChange={handleChange}
            className={`w-full px-2 py-1 text-center ${(isEfficiency || isHI || isHF) ? '' : 'bg-transparent focus:bg-white hover:bg-blue-50'} focus:outline-none focus:ring-2 focus:ring-blue-400 rounded appearance-none cursor-pointer transition-colors`}
            style={{
              minWidth: '100px',
              backgroundColor: cellBg,
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
              value={displayValue || ""}
              onChange={handleChange}
              className={`w-full px-2 py-1 text-center ${(isEfficiency || isHI || isHF) ? '' : 'bg-transparent focus:bg-white hover:bg-blue-50'} focus:outline-none focus:ring-2 focus:ring-blue-400 rounded appearance-none transition-colors`}
              style={{ minWidth: '100px', backgroundColor: cellBg }}
            />
            <datalist id={`list-${header}`}>
              {(dropdownOptions[header] || []).map((option) => (
                <option key={option} value={option} />
              ))}
            </datalist>
          </>
        )
      ) : (
        isWaitBeforeStart ? (
          <div
            className={`w-full flex items-center justify-center gap-1 ${(isEfficiency || isHI || isHF) ? '' : 'bg-transparent hover:bg-blue-50'} rounded px-1 py-1 transition-colors`}
            style={{ minWidth: '130px', backgroundColor: cellBg }}
            title="Tiempo de espera (HH:MM:SS)"
          >
            <input
              type="number"
              min="0"
              step="1"
              value={waitDurationParts.hours}
              onChange={(e) => handleWaitDurationPartChange('hours', e.target.value)}
              className="w-12 px-1 py-1 text-center bg-white focus:outline-none focus:ring-1 focus:ring-blue-400 rounded border border-gray-200"
              placeholder="HH"
            />
            <span className="text-gray-500 font-semibold">:</span>
            <input
              type="number"
              min="0"
              max="59"
              step="1"
              value={waitDurationParts.minutes}
              onChange={(e) => handleWaitDurationPartChange('minutes', e.target.value)}
              className="w-12 px-1 py-1 text-center bg-white focus:outline-none focus:ring-1 focus:ring-blue-400 rounded border border-gray-200"
              placeholder="MM"
            />
            <span className="text-gray-500 font-semibold">:</span>
            <input
              type="number"
              min="0"
              max="59"
              step="1"
              value={waitDurationParts.seconds}
              onChange={(e) => handleWaitDurationPartChange('seconds', e.target.value)}
              className="w-12 px-1 py-1 text-center bg-white focus:outline-none focus:ring-1 focus:ring-blue-400 rounded border border-gray-200"
              placeholder="SS"
            />
          </div>
        ) : (
          <input
            type="text"
            value={displayValue || ""}
            onChange={handleChange}
            className={`w-full px-2 py-1 text-center ${(isEfficiency || isHI || isHF) ? '' : 'bg-transparent focus:bg-white hover:bg-blue-50'} focus:outline-none focus:ring-2 focus:ring-blue-400 rounded appearance-none transition-colors`}
            style={{ minWidth: '100px', backgroundColor: cellBg }}
            readOnly={false}
          />
        )
      )}
    </td>
  );
});

export default DataTable;

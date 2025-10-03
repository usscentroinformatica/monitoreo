import React, { useState, useMemo } from "react";
import ExcelJS from "exceljs";
import ControlPanel from "./components/ControlPanel";
import DataTable from "./components/DataTable";

function App() {
  const [data, setData] = useState([]);
  const [zoomData, setZoomData] = useState([]);
  const [selectedDocente, setSelectedDocente] = useState("");
  const [numFilas, setNumFilas] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [availableSheets, setAvailableSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState(0);
  const [workbookData, setWorkbookData] = useState(null);
  const [currentHeaders, setCurrentHeaders] = useState([]);

  // ===== FUNCIONES DE UTILIDAD =====
  const normalizeDocenteName = (name) => {
    if (!name) return "";
    return name.toUpperCase().trim().split(/\s+/).sort().join(" ");
  };

  const normalizeCursoName = (name) => {
  if (!name) return "";
  
  // Primero convertir números romanos a arábigos
  let normalized = convertRomanToArabic(name);
  
  return normalized
    .toUpperCase()
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .split(/\s+/)
    .filter(word => word.length > 1)
    .join(" ");
};

  const convertRomanToArabic = (text) => {
  if (!text) return text;
  
  const romanToArabic = {
    'II': '2',
    'III': '3',
    'IV': '4',
    'V': '5'
  };
  
  let result = text;
  Object.keys(romanToArabic).forEach(roman => {
    // Reemplazar números romanos que estén como palabra completa
    const regex = new RegExp(`\\b${roman}\\b`, 'gi');
    result = result.replace(regex, romanToArabic[roman]);
  });
  
  return result;
};

  const matchDocente = (docenteExcel, docenteZoom) => {
    const normalizedExcel = normalizeDocenteName(docenteExcel);
    const normalizedZoom = normalizeDocenteName(docenteZoom);
    
    if (normalizedExcel === normalizedZoom) return true;
    
    const wordsExcel = normalizedExcel.split(" ");
    const wordsZoom = normalizedZoom.split(" ");
    const commonWords = wordsExcel.filter(word => wordsZoom.includes(word));
    
    return commonWords.length >= 2;
  };

  const extractDate = (dateTimeStr) => {
    if (!dateTimeStr) return "";
    const match = dateTimeStr.match(/^([A-Za-z]+\s+\d{1,2},\s+\d{4})/);
    return match ? match[1] : "";
  };

  const extractTime = (dateTimeStr) => {
    if (!dateTimeStr) return "";
    
    let match = dateTimeStr.match(/(\d{1,2}:\d{2}:\d{2}\s*[AP]M)/i);
    if (match) return match[1];
    
    match = dateTimeStr.match(/(\d{1,2}:\d{2}:\d{2}\s*[ap]\.\s*m\.)/i);
    if (match) return match[1];
    
    match = dateTimeStr.match(/^(\d{1,2}:\d{2}:\d{2})/);
    if (match) return match[1];
    
    return dateTimeStr;
  };

  const detectTurno = (horaStr) => {
  if (!horaStr) return "";
  
  // Extraer la hora en formato 24h
  let hour = 0;
  
  // Intentar parsear diferentes formatos
  // Formato: "06:50:01 PM" o "10:02:10 AM"
  const match12h = horaStr.match(/(\d{1,2}):(\d{2}):(\d{2})\s*([AP]M)/i);
  if (match12h) {
    hour = parseInt(match12h[1]);
    const period = match12h[4].toUpperCase();
    
    if (period === 'PM' && hour !== 12) {
      hour += 12;
    } else if (period === 'AM' && hour === 12) {
      hour = 0;
    }
  } else {
    // Formato: "10:02:10 p. m." o "10:02:10 a. m."
    const matchPeriod = horaStr.match(/(\d{1,2}):(\d{2}):(\d{2})\s*([ap])\.\s*m\./i);
    if (matchPeriod) {
      hour = parseInt(matchPeriod[1]);
      const period = matchPeriod[4].toLowerCase();
      
      if (period === 'p' && hour !== 12) {
        hour += 12;
      } else if (period === 'a' && hour === 12) {
        hour = 0;
      }
    } else {
      // Formato 24h: "18:50:01"
      const match24h = horaStr.match(/(\d{1,2}):/);
      if (match24h) {
        hour = parseInt(match24h[1]);
      }
    }
  }
  
  // Determinar turno basado en la hora
  if (hour >= 6 && hour < 12) {
    return "MAÑANA";
  } else if (hour >= 12 && hour < 18) {
    return "TARDE";
  } else if (hour >= 18 && hour <= 23) {
    return "NOCHE";
  } else {
    return "NOCHE"; // 0-5 AM también es noche
  }
};

  const extractCursoFromTema = (tema) => {
    if (!tema) return "";
    const match = tema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
    return match ? match[1].trim() : tema;
  };

  // ===== HANDLERS =====
  const handleZoomCsvUpload = async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  setIsLoading(true);
  try {
    const text = await file.text();
    
    const lines = text.split('\n').filter(line => line.trim());
    let delimiter = ';';
    
    if (!lines[0].includes(';')) {
      delimiter = lines[0].includes('\t') ? '\t' : ',';
    }
    
    const headers = lines[0].split(delimiter).map(h => h.trim().replace(/^"|"$/g, ''));
    
    console.log("Delimitador detectado:", delimiter);
    console.log("Headers detectados:", headers);
    
    const parsedZoomData = [];
    for (let i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      
      const values = lines[i].split(delimiter).map(v => v.trim().replace(/^"|"$/g, ''));
      const row = {};
      headers.forEach((header, index) => {
        row[header] = values[index] || "";
      });
      parsedZoomData.push(row);
    }

    console.log("Total de registros Zoom:", parsedZoomData.length);
    setZoomData(parsedZoomData);

    // Determinar docentes a procesar
    const docentesToProcess = selectedDocente 
      ? [selectedDocente] 
      : [...new Set(data.map(row => row.DOCENTE).filter(d => d && d.trim() !== ''))];

    if (docentesToProcess.length === 0) {
      alert("No hay docentes registrados en el Excel para autocompletar");
      return;
    }

    console.log(`Procesando ${docentesToProcess.length} docente(s):`, docentesToProcess);

    let updatedCount = 0;
    let createdCount = 0;
    let deletedCount = 0;
    let newData = [...data];

    // Procesar cada docente
    docentesToProcess.forEach(docenteActual => {
      console.log(`\n--- Procesando docente: ${docenteActual} ---`);

      // PASO 1: Identificar todas las sesiones del docente en Zoom
      const sesionesZoomPorCurso = new Map(); // Key: CURSO|||SECCION, Value: Set de números de sesión

      parsedZoomData.forEach(zoomRow => {
        const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
        const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";

        if (!matchDocente(docenteActual, zoomDocente) || !zoomTema) return;

        let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
        
        if (temaMatch) {
          const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
          const cursoZoom = cursoParte.trim();
          const sesionNumero = sesionNumeroStr || "";
          const key = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}`;
          
          if (!sesionesZoomPorCurso.has(key)) {
            sesionesZoomPorCurso.set(key, new Set());
          }
          if (sesionNumero) {
            sesionesZoomPorCurso.get(key).add(parseInt(sesionNumero));
          }
        }
      });

      console.log(`Sesiones encontradas en Zoom para ${docenteActual}:`, 
        Array.from(sesionesZoomPorCurso.entries()).map(([key, sessions]) => 
          `${key}: [${Array.from(sessions).sort((a,b) => a-b).join(', ')}]`
        )
      );

      // PASO 2: Eliminar filas del Excel que NO están en Zoom
      const filasOriginales = newData.length;
      newData = newData.filter(row => {
        if (row.DOCENTE !== docenteActual) return true; // Mantener otras filas

        // Si la fila tiene datos de Zoom completos, mantenerla
        if (row["Columna 13"] && row.inicio && row.fin) return true;

        // Si la fila está vacía o incompleta, verificar si existe en Zoom
        const keyCurso = `${normalizeCursoName(row.CURSO || '')}|||${(row.SECCION || '').toUpperCase()}`;
        const sesionesZoom = sesionesZoomPorCurso.get(keyCurso);
        
        // Si no hay curso/sección o no existe en Zoom, eliminar
        if (!row.CURSO || !row.SECCION || !sesionesZoom) {
          console.log(`Eliminando fila vacía sin coincidencia: ${row.CURSO} - ${row.SECCION} - Sesión ${row.SESION}`);
          deletedCount++;
          return false;
        }

        // Si la sesión no existe en Zoom, eliminar
        if (row.SESION && !sesionesZoom.has(parseInt(row.SESION))) {
          console.log(`Eliminando fila: Sesión ${row.SESION} no existe en Zoom para ${row.CURSO} - ${row.SECCION}`);
          deletedCount++;
          return false;
        }

        return true; // Mantener la fila para intentar autocompletar
      });

      // Buscar template del docente
      let templateRow = newData.find(row => row.DOCENTE === docenteActual);
      
      if (!templateRow) {
        const firstZoomForDocente = parsedZoomData.find(zoomRow => {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
          return matchDocente(docenteActual, zoomDocente);
        });

        if (firstZoomForDocente) {
          templateRow = {
            PERIODO: "",
            MODELO: "PROTECH XP",
            MODALIDAD: "VIRTUAL",
            CURSO: "",
            SECCION: "",
            "AULA USS": "",
            DOCENTE: docenteActual,
            TURNO: "",
            DIAS: "",
            "HORA INICIO": "",
            "HORA FIN": "",
            SESION: "",
            "Columna 13": "",
            inicio: "",
            fin: "",
            "Columna 16": "",
            "Columna 17": "",
            TOTAL: ""
          };
        } else {
          console.log(`No se encontraron registros de Zoom para ${docenteActual}`);
          return;
        }
      }

      // PASO 3: Autocompletar filas existentes vacías
      parsedZoomData.forEach(zoomRow => {
        const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
        const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
        const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
        const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";

        if (!matchDocente(docenteActual, zoomDocente) || !zoomTema) return;

        let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
        
        if (!temaMatch) return;

        const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
        const cursoZoom = cursoParte.trim();
        const sesionNumero = parseInt(sesionNumeroStr || "0");

        let matched = false;

        // Buscar fila existente para autocompletar
        newData.forEach((row, index) => {
          if (row.DOCENTE !== docenteActual) return;
          if (matched) return; // Ya se encontró match para este registro de Zoom

          const isEmptyRow = !row["Columna 13"] || !row.inicio || !row.fin;
          if (!isEmptyRow) return; // Solo autocompletar filas vacías

          const cursoExcelNorm = normalizeCursoName(row.CURSO || '');
          const cursoZoomNorm = normalizeCursoName(cursoZoom);
          const seccionMatch = (row.SECCION || '').toUpperCase() === seccionZoom.toUpperCase();
          const sesionMatch = sesionNumero && parseInt(String(row.SESION)) === sesionNumero;

          let cursoMatch = cursoExcelNorm === cursoZoomNorm;
          if (!cursoMatch && cursoExcelNorm && cursoZoomNorm) {
            const wordsExcel = cursoExcelNorm.split(" ");
            const wordsZoom = cursoZoomNorm.split(" ");
            const commonWords = wordsExcel.filter(word => wordsZoom.includes(word));
            cursoMatch = commonWords.length >= 2;
          }

          if (seccionMatch && sesionMatch && cursoMatch) {
            const fechaExtraida = extractDate(fechaInicio);
            const horaInicioExtraida = extractTime(fechaInicio);
            const horaFinExtraida = extractTime(fechaFin);
            const turnoDetectado = detectTurno(fechaInicio);
            
            newData[index] = {
              ...newData[index],
              CURSO: cursoZoom,
              TURNO: turnoDetectado,
              "Columna 13": fechaExtraida,
              inicio: horaInicioExtraida,
              fin: horaFinExtraida
            };
            updatedCount++;
            matched = true;
            console.log(`Autocompletada: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionNumero}`);
          }
        });

        // PASO 4: Si no hay fila existente, crear una nueva
        if (!matched) {
          const fechaExtraida = extractDate(fechaInicio);
          const horaInicioExtraida = extractTime(fechaInicio);
          const horaFinExtraida = extractTime(fechaFin);
          const turnoDetectado = detectTurno(fechaInicio);

          const newRow = {
            PERIODO: templateRow.PERIODO,
            MODELO: templateRow.MODELO,
            MODALIDAD: templateRow.MODALIDAD,
            CURSO: cursoZoom,
            SECCION: seccionZoom,
            "AULA USS": templateRow["AULA USS"],
            DOCENTE: docenteActual,
            TURNO: turnoDetectado,
            DIAS: templateRow.DIAS,
            "HORA INICIO": templateRow["HORA INICIO"],
            "HORA FIN": templateRow["HORA FIN"],
            SESION: sesionNumero,
            "Columna 13": fechaExtraida,
            inicio: horaInicioExtraida,
            fin: horaFinExtraida,
            "Columna 16": "",
            "Columna 17": "",
            TOTAL: ""
          };

          const alreadyExists = newData.some(row => 
            row.DOCENTE === docenteActual &&
            row.SECCION.toUpperCase() === seccionZoom.toUpperCase() &&
            parseInt(String(row.SESION)) === sesionNumero &&
            row["Columna 13"] === fechaExtraida
          );

          if (!alreadyExists) {
            newData.push(newRow);
            createdCount++;
            console.log(`Nueva fila: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionNumero}`);
          }
        }
      });
    });

    setData(newData);

    const mensaje = selectedDocente 
      ? `Procesado para ${selectedDocente}:\n${deletedCount} filas eliminadas\n${updatedCount} filas autocompletadas\n${createdCount} filas nuevas creadas`
      : `Procesados ${docentesToProcess.length} docentes:\n${deletedCount} filas eliminadas\n${updatedCount} filas autocompletadas\n${createdCount} filas nuevas creadas`;
    
    alert(mensaje);

  } catch (error) {
    alert("Error al procesar el archivo CSV: " + error.message);
    console.error(error);
  } finally {
    setIsLoading(false);
    event.target.value = "";
  }
};

  const handleFileUpload = async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  setIsLoading(true);
  try {
    const workbook = new ExcelJS.Workbook();
    const arrayBuffer = await file.arrayBuffer();
    await workbook.xlsx.load(arrayBuffer);
    setWorkbookData(workbook);
    
    const sheetNames = workbook.worksheets.map((sheet, index) => ({
      index,
      name: sheet.name
    }));
    setAvailableSheets(sheetNames);
    
    const worksheet = workbook.worksheets[0];
    const { data: loadedData, headers: sheetHeaders } = loadSheetData(worksheet);
    
    setData(loadedData);
    setCurrentHeaders(sheetHeaders);  // Guarda los headers dinámicos
    
  } catch (error) {
    alert("❌ Error al cargar el archivo: " + error.message);
    console.error(error);
  } finally {
    setIsLoading(false);
    event.target.value = "";
  }
};

const loadSheetData = (worksheet) => {
  const loadedData = [];
  let sheetHeaders = [];
  let headerRowIndex = 1;
  let maxColumns = 0;

  // Buscar la primera fila con datos
  for (let i = 1; i <= 10; i++) {
    const row = worksheet.getRow(i);
    let hasData = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.value && String(cell.value).trim() !== "") {
        hasData = true;
      }
    });
    if (hasData) {
      headerRowIndex = i;
      break;
    }
  }

  const headerRow = worksheet.getRow(headerRowIndex);
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    maxColumns = Math.max(maxColumns, colNumber);
  });

  if (maxColumns === 0 && worksheet.actualColumnCount) {
    maxColumns = worksheet.actualColumnCount;
  }

  // Leer encabezados
  for (let col = 1; col <= maxColumns; col++) {
    const cell = headerRow.getCell(col);
    const value = cell.value;
    if (value && String(value).trim() !== "") {
      const headerText = typeof value === 'object' && value.text ? value.text : String(value).trim();
      sheetHeaders.push(headerText);
    } else {
      sheetHeaders.push(`Columna ${col}`);
    }
  }

  // NUEVO: Detectar si la primera fila de datos son realmente encabezados
  const firstDataRow = worksheet.getRow(headerRowIndex + 1);
  let firstRowValues = [];
  for (let col = 1; col <= Math.min(5, sheetHeaders.length); col++) {
    const cell = firstDataRow.getCell(col);
    if (cell.value) {
      firstRowValues.push(String(cell.value).trim().toUpperCase());
    }
  }
  
  // Si la primera fila de datos contiene palabras como "PERIODO", "MODELO", etc., usarla como encabezado
  const commonHeaders = ['PERIODO', 'CICLO', 'MODELO', 'MODALIDAD', 'CURSO', 'SECCION', 'DOCENTE'];
  const matchCount = firstRowValues.filter(val => commonHeaders.includes(val)).length;
  
  if (matchCount >= 3) {
    console.log("⚠️ Los encabezados reales están en la fila de datos. Ajustando...");
    sheetHeaders = [];
    for (let col = 1; col <= maxColumns; col++) {
      const cell = firstDataRow.getCell(col);
      const value = cell.value;
      if (value && String(value).trim() !== "") {
        sheetHeaders.push(String(value).trim());
      } else {
        sheetHeaders.push(`Columna ${col}`);
      }
    }
    headerRowIndex++; // Saltar esta fila ahora que son encabezados
  }

  console.log(`Hoja: ${worksheet.name}`);
  console.log(`Fila de encabezados: ${headerRowIndex}`);
  console.log(`Encabezados detectados (${sheetHeaders.length}):`, sheetHeaders);

  // Leer datos
  worksheet.eachRow((row, rowIndex) => {
    if (rowIndex <= headerRowIndex) return;
    
    const getCellValue = (cell) => {
      if (!cell || !cell.value) return "";
      if (cell.value.hyperlink) return cell.value.hyperlink;
      if (typeof cell.value === 'object' && cell.value.text) return cell.value.text;
      if (cell.value instanceof Date) return cell.value.toLocaleDateString();
      return String(cell.value).trim();
    };
    
    const rowData = {};
    let hasAnyData = false;
    
    for (let col = 1; col <= sheetHeaders.length; col++) {
      const header = sheetHeaders[col - 1];
      const cellValue = getCellValue(row.getCell(col));
      rowData[header] = cellValue;
      if (cellValue !== "") {
        hasAnyData = true;
      }
    }

    if (hasAnyData) {
      loadedData.push(rowData);
    }
  });

  console.log(`Total de registros cargados: ${loadedData.length}`);
  if (loadedData.length > 0) {
    console.log("Primera fila:", loadedData[0]);
  }

  return { data: loadedData, headers: sheetHeaders };
};


const handleSheetChange = (sheetIndex) => {
  if (!workbookData) {
    alert("Por favor, carga primero un archivo Excel");
    return;
  }

  setSelectedSheet(sheetIndex);
  
  const worksheet = workbookData.worksheets[sheetIndex];
  const { data: loadedData, headers: sheetHeaders } = loadSheetData(worksheet);
  
  setData(loadedData);
  setCurrentHeaders(sheetHeaders);
  setSelectedDocente("");  
  
};

  const getUniqueCursosFromZoom = (zoomData, docente) => {
  const cursos = new Set();
  
  zoomData.forEach(zoomRow => {
    const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
    const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
    
    if (!matchDocente(docente, zoomDocente) || !zoomTema) return;
    
    const match = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)/i);
    if (match) {
      const curso = match[1].trim();
      cursos.add(curso);
    }
  });
  
  return Array.from(cursos);
};


const createRowsForDocente = () => {
  if (!selectedDocente || !numFilas || numFilas <= 0) {
    alert("Por favor selecciona un docente y especifica el número de filas");
    return;
  }

  const docenteRow = data.find(row => row.DOCENTE === selectedDocente);
  
  if (!docenteRow) {
    alert("No se encontró información del docente seleccionado");
    return;
  }

  const teacherZoom = zoomData.filter(zoomRow => {
    const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
    return matchDocente(selectedDocente, zoomDocente);
  });

  // Agrupar por CURSO y SECCION - esto detecta TODOS los cursos del docente
  const cursosSeccionesMap = new Map();
  
  teacherZoom.forEach(zoomRow => {
    const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
    let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
    
    if (temaMatch) {
      const [, cursoParte, seccionZoom] = temaMatch;
      const curso = cursoParte.trim();
      const key = `${curso}|||${seccionZoom}`; // Clave única por curso+sección
      
      if (!cursosSeccionesMap.has(key)) {
        cursosSeccionesMap.set(key, { curso, seccion: seccionZoom, zoomRows: [] });
      }
      cursosSeccionesMap.get(key).zoomRows.push(zoomRow);
    }
  });

  // Si no hay datos de Zoom, usar sección del Excel
  if (cursosSeccionesMap.size === 0 && docenteRow.SECCION && docenteRow.CURSO) {
    const key = `${docenteRow.CURSO}|||${docenteRow.SECCION}`;
    cursosSeccionesMap.set(key, { 
      curso: docenteRow.CURSO, 
      seccion: docenteRow.SECCION, 
      zoomRows: [] 
    });
  }

  if (cursosSeccionesMap.size === 0) {
    alert("No se encontró información de cursos/secciones para crear filas");
    return;
  }

  console.log(`\n📚 Docente: ${selectedDocente}`);
  console.log(`📖 Total de cursos encontrados: ${cursosSeccionesMap.size}`);
  cursosSeccionesMap.forEach(({ curso, seccion }, key) => {
    console.log(`   - ${curso} (${seccion})`);
  });

  // Contar filas existentes por CURSO+SECCION
  const existingRowsByCursoSeccion = {};
  data.forEach(row => {
    if (row.DOCENTE === selectedDocente) {
      const key = `${row.CURSO}|||${row.SECCION}`;
      if (!existingRowsByCursoSeccion[key]) {
        existingRowsByCursoSeccion[key] = 0;
      }
      existingRowsByCursoSeccion[key]++;
    }
  });

  const totalCursosSecciones = cursosSeccionesMap.size;
  const rowsPerCursoSeccion = Math.ceil(parseInt(numFilas) / totalCursosSecciones);
  const allNewRows = [];
  let totalCreated = 0;
  let totalAutoCompleted = 0;

  console.log(`\n🎯 Distribución: ${numFilas} filas entre ${totalCursosSecciones} cursos = ${rowsPerCursoSeccion} filas por curso\n`);

  cursosSeccionesMap.forEach(({ curso, seccion, zoomRows }, key) => {
    const existingCount = existingRowsByCursoSeccion[key] || 0;
    const rowsToCreate = rowsPerCursoSeccion - existingCount;
    
    console.log(`📖 ${curso} - ${seccion}:`);
    console.log(`   Filas existentes: ${existingCount}`);
    console.log(`   Filas a crear: ${rowsToCreate}`);
    console.log(`   Total final: ${rowsPerCursoSeccion}`);
    
    if (rowsToCreate <= 0) {
      console.log(`   ⚠️ Ya tiene suficientes filas, omitiendo...`);
      return;
    }

    const startSession = existingCount + 1;

    for (let i = 0; i < rowsToCreate; i++) {
      const sesionActual = startSession + i;
      
      const matchingZoom = zoomRows.find(zoomRow => {
        const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
        let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
        if (!temaMatch) return false;

        const [, , , sesionNumeroStr] = temaMatch;
        const sesionNum = parseInt(sesionNumeroStr || "");

        return sesionNum === sesionActual;
      });

      const thisRow = {
        PERIODO: docenteRow.PERIODO,
        MODELO: docenteRow.MODELO,
        MODALIDAD: docenteRow.MODALIDAD,
        CURSO: curso,
        SECCION: seccion,
        "AULA USS": docenteRow["AULA USS"],
        DOCENTE: docenteRow.DOCENTE,
        TURNO: docenteRow.TURNO,
        DIAS: docenteRow.DIAS,
        "HORA INICIO": docenteRow["HORA INICIO"],
        "HORA FIN": docenteRow["HORA FIN"],
        SESION: sesionActual,
        "Columna 13": "",
        inicio: "",
        fin: "",
        "Columna 16": "",
        "Columna 17": "",
        TOTAL: ""
      };

      if (matchingZoom) {
        const fechaInicio = matchingZoom['Hora de inicio'] || matchingZoom['Start Time'] || "";
        const fechaFin = matchingZoom['Hora de finalización'] || matchingZoom['End Time'] || "";
        const fechaExtraida = extractDate(fechaInicio);
        const horaInicioExtraida = extractTime(fechaInicio);
        const horaFinExtraida = extractTime(fechaFin);
        const turnoDetectado = detectTurno(fechaInicio);
        
        thisRow["Columna 13"] = fechaExtraida;
        thisRow.inicio = horaInicioExtraida;
        thisRow.fin = horaFinExtraida;
        thisRow.TURNO = turnoDetectado;
        
        console.log(`   ✓ Sesión ${sesionActual} autocompletada`);
        totalAutoCompleted++;
      }

      allNewRows.push(thisRow);
      totalCreated++;
    }
  });

  setData([...data, ...allNewRows]);
  setNumFilas("");

  console.log(`\n✅ Resumen:`);
  console.log(`   Total de filas creadas: ${totalCreated}`);
  console.log(`   Filas autocompletadas: ${totalAutoCompleted}`);
  
  alert(`✅ Se crearon ${totalCreated} filas para ${selectedDocente}\n${totalAutoCompleted} filas fueron autocompletadas con datos de Zoom`);
};


const createRowsForAllDocentes = () => {
  if (!numFilas || numFilas <= 0) {
    alert("Por favor especifica el número de filas a crear por docente");
    return;
  }

  if (zoomData.length === 0) {
    alert("⚠️ Primero carga el CSV de Zoom para poder autocompletar los datos");
    return;
  }

  let totalCreated = 0;
  let totalAutoCompleted = 0;
  const processedDocentes = [];

  // Obtener todos los docentes únicos que tienen datos en Zoom
  const docentesEnZoom = new Set();
  zoomData.forEach(zoomRow => {
    const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
    if (zoomDocente) {
      // Buscar el docente normalizado en la lista
      const docenteMatch = uniqueDocentes.find(d => matchDocente(d, zoomDocente));
      if (docenteMatch) {
        docentesEnZoom.add(docenteMatch);
      }
    }
  });

  console.log(`📋 Docentes encontrados en Zoom:`, Array.from(docentesEnZoom));

  const allNewRows = [];

  docentesEnZoom.forEach(docente => {
    const docenteRow = data.find(row => row.DOCENTE === docente);
    
    // Si no hay template, crear uno básico
    const template = docenteRow || {
      PERIODO: "",
      MODELO: "PROTECH XP",
      MODALIDAD: "VIRTUAL",
      CURSO: "",
      SECCION: "",
      "AULA USS": "",
      DOCENTE: docente,
      TURNO: "",
      DIAS: "",
      "HORA INICIO": "",
      "HORA FIN": "",
      SESION: "",
      "Columna 13": "",
      inicio: "",
      fin: "",
      "Columna 16": "",
      "Columna 17": "",
      TOTAL: ""
    };

    const teacherZoom = zoomData.filter(zoomRow => {
      const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
      return matchDocente(docente, zoomDocente);
    });

    // Agrupar por CURSO y SECCION
    const cursosSeccionesMap = new Map();
    
    teacherZoom.forEach(zoomRow => {
      const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
      let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
      
      if (temaMatch) {
        const [, cursoParte, seccionZoom] = temaMatch;
        const curso = cursoParte.trim();
        const key = `${curso}|||${seccionZoom}`;
        
        if (!cursosSeccionesMap.has(key)) {
          cursosSeccionesMap.set(key, { curso, seccion: seccionZoom, zoomRows: [] });
        }
        cursosSeccionesMap.get(key).zoomRows.push(zoomRow);
      }
    });

    if (cursosSeccionesMap.size === 0) return;

    // Contar filas existentes por CURSO+SECCION
    const existingRowsByCursoSeccion = {};
    data.forEach(row => {
      if (row.DOCENTE === docente) {
        const key = `${row.CURSO}|||${row.SECCION}`;
        if (!existingRowsByCursoSeccion[key]) {
          existingRowsByCursoSeccion[key] = 0;
        }
        existingRowsByCursoSeccion[key]++;
      }
    });

    const totalCursosSecciones = cursosSeccionesMap.size;
    const rowsPerCursoSeccion = Math.ceil(parseInt(numFilas) / totalCursosSecciones);

    cursosSeccionesMap.forEach(({ curso, seccion, zoomRows }, key) => {
      const existingCount = existingRowsByCursoSeccion[key] || 0;
      const rowsToCreate = rowsPerCursoSeccion - existingCount;
      
      if (rowsToCreate <= 0) return;

      const startSession = existingCount + 1;

      for (let i = 0; i < rowsToCreate; i++) {
        const sesionActual = startSession + i;
        
        const matchingZoom = zoomRows.find(zoomRow => {
          const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
          let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
          if (!temaMatch) return false;

          const [, , , sesionNumeroStr] = temaMatch;
          const sesionNum = parseInt(sesionNumeroStr || "");

          return sesionNum === sesionActual;
        });

        const thisRow = {
          PERIODO: template.PERIODO,
          MODELO: template.MODELO,
          MODALIDAD: template.MODALIDAD,
          CURSO: curso,
          SECCION: seccion,
          "AULA USS": template["AULA USS"],
          DOCENTE: template.DOCENTE,
          TURNO: template.TURNO,
          DIAS: template.DIAS,
          "HORA INICIO": template["HORA INICIO"],
          "HORA FIN": template["HORA FIN"],
          SESION: sesionActual,
          "Columna 13": "",
          inicio: "",
          fin: "",
          "Columna 16": "",
          "Columna 17": "",
          TOTAL: ""
        };

        if (matchingZoom) {
          const fechaInicio = matchingZoom['Hora de inicio'] || matchingZoom['Start Time'] || "";
          const fechaFin = matchingZoom['Hora de finalización'] || matchingZoom['End Time'] || "";
          const fechaExtraida = extractDate(fechaInicio);
          const horaInicioExtraida = extractTime(fechaInicio);
          const horaFinExtraida = extractTime(fechaFin);
          const turnoDetectado = detectTurno(fechaInicio);
          
          thisRow["Columna 13"] = fechaExtraida;
          thisRow.inicio = horaInicioExtraida;
          thisRow.fin = horaFinExtraida;
          thisRow.TURNO = turnoDetectado;
          
          totalAutoCompleted++;
        }

        allNewRows.push(thisRow);
        totalCreated++;
      }
    });

    processedDocentes.push(docente);
  });

  setData([...data, ...allNewRows]);
  setNumFilas("");

  alert(`✅ Se crearon ${totalCreated} filas para ${processedDocentes.length} docentes\n${totalAutoCompleted} filas fueron autocompletadas con datos de Zoom`);
};

  const exportToExcel = async () => {
    if (data.length === 0) {
      alert('No hay datos para exportar.');
      return;
    }

    // CAMBIO CLAVE: Usa currentHeaders dinámicos (de la hoja cargada) o headers fijos si no hay
    const exportHeaders = currentHeaders.length > 0 ? currentHeaders : headers;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Monitoreo");

    // Agrega headers dinámicos
    worksheet.addRow(exportHeaders);

    const headerRow = worksheet.getRow(1);
    headerRow.height = 40;
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF203864" }
      };
      cell.font = {
        color: { argb: "FFFFFFFF" },
        bold: true,
        size: 11,
        name: "Arial"
      };
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFFFFFFF" } },
        left: { style: "thin", color: { argb: "FFFFFFFF" } },
        bottom: { style: "thin", color: { argb: "FFFFFFFF" } },
        right: { style: "thin", color: { argb: "FFFFFFFF" } }
      };
    });

    // Exporta solo los datos filtrados actuales (displayData)
    displayData.forEach((row, index) => {
      const rowData = exportHeaders.map(h => row[h] !== undefined ? row[h] : "");
      const excelRow = worksheet.addRow(rowData);
      excelRow.height = 25;

      const isEven = (index + 2) % 2 === 0;
      excelRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: isEven ? "FFE8F4F8" : "FFFFFFFF" }
        };
        cell.font = {
          size: 10,
          name: "Arial",
          color: { argb: "FF000000" }
        };
        cell.alignment = {
          vertical: "middle",
          horizontal: "center"
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE0E0E0" } },
          left: { style: "thin", color: { argb: "FFE0E0E0" } },
          bottom: { style: "thin", color: { argb: "FFE0E0E0" } },
          right: { style: "thin", color: { argb: "FFE0E0E0" } }
        };
      });
    });

    // CAMBIO: Anchos de columnas dinámicos basados en headers (mínimo 10, máximo 50)
    const dynamicWidths = exportHeaders.map(() => Math.min(50, Math.max(10, 15))); // Ajusta según necesidad
    worksheet.columns = exportHeaders.map((header, index) => ({
      key: header,
      width: dynamicWidths[index]
    }));

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Monitoreo_USS_${new Date().toISOString().split('T')[0]}.xlsx`; // Con fecha actual
    link.click();
    window.URL.revokeObjectURL(url);
  };

  const addRow = () => {
    const newRow = {};
    const useHeaders = currentHeaders.length > 0 ? currentHeaders : headers;
    useHeaders.forEach(header => {
      newRow[header] = "";
    });
    setData([...data, newRow]);
  };

  const deleteRow = (index) => {
    if (selectedDocente) {
      const realIndex = data.findIndex((row) => row === filteredData[index]);
      const newData = data.filter((_, i) => i !== realIndex);
      setData(newData);
    } else {
      const newData = data.filter((_, i) => i !== index);
      setData(newData);
    }
  };

  const handleCellChange = (rowIndex, columnName, value) => {
    if (selectedDocente) {
      const realIndex = data.findIndex((row, idx) => 
        row === filteredData[rowIndex]
      );
      const newData = [...data];
      newData[realIndex][columnName] = value;
      setData(newData);
    } else {
      const newData = [...data];
      newData[rowIndex][columnName] = value;
      setData(newData);
    }
  };

  const uniqueDocentes = [
    "CARRASCO CHEVEZ HENRY", "LARA PERLECHE LOURDES", "CALLACNA SENCIO IVAN",
    "MEJIA ZELADA CARLOS", "MOGOLLON GALECIO POLO", "DIAZ ESPINO MIGUEL",
    "GARCIA CABRERA MARTIN", "DIAZ MUSAYON CESAR", "CASTAÑEDA BALCAZAR EDWARD",
    "QUINTANA DAVILA TONY", "SANCHEZ JAEGER CRISTHIAM", "TULLUME DIAZ JOSE",
    "GUERRERO AGURTO GINO", "CRIOLLO VALVERDE STEPHANY", "TICONA TAPIA ESTRELLA",
    "DELGADO MARIELLA", "SANCHEZ PEREZ DIANA", "QUESADA QUIROZ JENNIE",
    "BURGOS SUERO LUCIANA", "SALAZAR LLUEN IVONNE", "SANCHEZ GUEVARA OMAR",
    "BRUNO SARMIENTO JOSE", "GONZALES ÑIQUE PERCY", "OLIVOS KATHERIN",
    "SANDOVAL HORNA CARMEN", "SALAZAR LLUEN JAIRO", "MECHAN FRANCISCO",
    "CISNEROS LEANDRO", "MURO EFRAIN", "QUESADA JENNIE", "NIETO NELSON",
    "SALAZAR DANIEL", "DIAZ CESAR"
  ];

  const dropdownOptions = {
    MODELO: ["PROTECH XP", "TRADICIONAL"],
    MODALIDAD: ["PRESENCIAL", "VIRTUAL"],
    CURSO: [
      "AUTOCAD 2D", "AUTOCAD 3D", "COMPUTACION 2", "COMPUTACION 2 ARQ",
      "COMPUTACION 2 CIV", "COMPUTACIÓN 2 MEC", "COMPUTACIÓN 2 SIST",
      "COMPUTACION 3", "COMPUTACION 3 ARQ", "COMPUTACION 3 CIV",
      "COMPUTACIÓN 3 MEC", "COMPUTACIÓN 3 SIST", "WORD 365", "WORD ASOCIADO",
      "EXCEL 365", "EXCEL ASOCIADO", "DISEÑO CON CANVA", "DISEÑO WEB",
      "POWER BI", "TALLER WORD 2019/365", "TALLER EXCEL 2019/365"
    ],
    SECCION: [
      "A", "PEAD-a", "PEAD-b", "PEAD-C", "PEAD-d", "PEAD-e", "PEAD-f",
      "PEAD-g", "PEAD-h", "PEAD-i", "PEAD-j", "PEAD-k", "PEAD-l", "PEAD-m",
      "PEAD-n", "PEAD-o", "PEAD-p", "PEAD-q", "PEAD-A"
    ],
    TURNO: ["MAÑANA", "TARDE", "NOCHE"],
    DIAS: [
      "LUN", "MAR", "MIE", "JUE", "VIE", "SAB", "LUN y MIE", "LUN-MIE",
      "LUN-VIE", "MAR Y JUE", "MAR y VIE", "MAR-JUE", "MAR-VIE", "MIE y JUE",
      "MIE y VIE", "MIE-VIE", "JUE-SAB", "VIE y SAB", "VIE-SAB"
    ],
    DOCENTE: uniqueDocentes
  };

  // ===== DATOS COMPUTADOS =====
  const filteredData = useMemo(() => {
    if (!selectedDocente) return data;
    return data.filter(row => row.DOCENTE === selectedDocente);
  }, [data, selectedDocente]);

  const displayData = useMemo(() => {
    if (!selectedDocente) return data;
    
    return filteredData.map((row, index) => ({
      ...row,
      SESION: index + 1
    }));
  }, [filteredData, selectedDocente, data]);

  // ===== RENDER =====
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 p-4">
      <div className="max-w-full mx-auto">
        <ControlPanel
          selectedDocente={selectedDocente}
          setSelectedDocente={setSelectedDocente}
          numFilas={numFilas}
          setNumFilas={setNumFilas}
          uniqueDocentes={uniqueDocentes}
          onCreateRows={createRowsForDocente}
          onCreateRowsForAll={createRowsForAllDocentes}
          onAddRow={addRow}
          onExport={exportToExcel}
          onLoadExcel={handleFileUpload}
          onLoadZoomCsv={handleZoomCsvUpload}
          isLoading={isLoading}
          displayDataLength={displayData.length}
          displayData={displayData} // NUEVO: Pasa displayData para export local si se necesita
          availableSheets={availableSheets}
          selectedSheet={selectedSheet}
          onSheetChange={handleSheetChange}
        />

        <DataTable
  data={displayData}
  headers={currentHeaders.length > 0 ? currentHeaders : []}  // Array vacío si no hay headers cargados
  dropdownOptions={dropdownOptions}
  onCellChange={handleCellChange}
  onDeleteRow={deleteRow}
/>
      </div>
    </div>
  );
}

export default App;
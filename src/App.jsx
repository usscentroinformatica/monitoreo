import React, { useState, useMemo } from "react";
import ExcelJS from "exceljs";
import ControlPanel from "./components/ControlPanel";
import DataTable from "./components/DataTable";

function App() {
  // Sistema de pestañas
const [tabs, setTabs] = useState([]);
const [activeTabId, setActiveTabId] = useState(null);
const [nextTabId, setNextTabId] = useState(1);

// Obtener la pestaña activa
const activeTab = tabs.find(tab => tab.id === activeTabId);

// Estados de la pestaña activa (si existe)
const data = activeTab?.data || [];
const zoomData = activeTab?.zoomData || [];
const selectedDocente = activeTab?.selectedDocente || "";
const numFilas = activeTab?.numFilas || "";
const isLoading = activeTab?.isLoading || false;
const availableSheets = activeTab?.availableSheets || [];
const selectedSheet = activeTab?.selectedSheet || 0;
const workbookData = activeTab?.workbookData || null;
const currentHeaders = activeTab?.currentHeaders || [];

// Función para actualizar la pestaña activa
const updateActiveTab = (updates) => {
  setTabs(prevTabs => 
    prevTabs.map(tab => 
      tab.id === activeTabId 
        ? { ...tab, ...updates }
        : tab
    )
  );
};

// Función para crear nueva pestaña
const createNewTab = (fileName, initialData = {}) => {
  const newTab = {
    id: nextTabId,
    name: fileName || `Archivo ${nextTabId}`,
    data: initialData.data || [],
    zoomData: [],
    selectedDocente: "",
    numFilas: "",
    isLoading: false,
    availableSheets: initialData.availableSheets || [],
    selectedSheet: 0,
    workbookData: initialData.workbookData || null,
    currentHeaders: initialData.currentHeaders || []
  };
  
  setTabs(prev => [...prev, newTab]);
  setActiveTabId(nextTabId);
  setNextTabId(prev => prev + 1);
};

// Función para cerrar pestaña
const closeTab = (tabId) => {
  const confirmClose = window.confirm("¿Estás seguro de cerrar esta pestaña? Los cambios no guardados se perderán.");
  if (!confirmClose) return;

  const newTabs = tabs.filter(tab => tab.id !== tabId);
  setTabs(newTabs);
  
  if (activeTabId === tabId) {
    setActiveTabId(newTabs.length > 0 ? newTabs[0].id : null);
  }
};

// Wrappers para los setters
const setData = (newData) => updateActiveTab({ data: newData });
const setZoomData = (newZoomData) => updateActiveTab({ zoomData: newZoomData });
const setSelectedDocente = (docente) => updateActiveTab({ selectedDocente: docente });
const setNumFilas = (num) => updateActiveTab({ numFilas: num });
const setIsLoading = (loading) => updateActiveTab({ isLoading: loading });
const setAvailableSheets = (sheets) => updateActiveTab({ availableSheets: sheets });
const setSelectedSheet = (sheet) => updateActiveTab({ selectedSheet: sheet });
const setWorkbookData = (wb) => updateActiveTab({ workbookData: wb });
const setCurrentHeaders = (headers) => updateActiveTab({ currentHeaders: headers });

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

    console.log("=== INICIANDO PROCESAMIENTO CSV ZOOM ===");
    console.log("Total registros Zoom:", parsedZoomData.length);
    
    // IMPORTANTE: Mostrar los headers del Excel actual
    console.log("Headers del Excel actual:", currentHeaders);
    
    setZoomData(parsedZoomData);

    const docentesToProcess = selectedDocente 
      ? [selectedDocente] 
      : [...new Set(data.map(row => row.DOCENTE).filter(d => d && d.trim() !== ''))];

    if (docentesToProcess.length === 0) {
      alert("No hay docentes registrados en el Excel para autocompletar");
      return;
    }

    console.log(`📋 Modo: ${selectedDocente ? 'Docente específico' : 'TODOS los docentes'}`);
    console.log(`📋 Docentes a procesar (${docentesToProcess.length}):`, docentesToProcess);

    let updatedCount = 0;
    let createdCount = 0;
    const newData = [...data];

    // IMPORTANTE: Variable global para todas las sesiones usadas (no por docente)
    const sesionesUsadasGlobal = new Set();

    // Función auxiliar para detectar si una fila está vacía
    const isRowEmpty = (row) => {
      // Buscar en todas las posibles columnas de fecha/hora
      const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
      const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
      const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
      
      let hasDate = false;
      let hasStart = false;
      
      // Verificar fecha
      for (const col of possibleDateCols) {
        const value = row[col];
        if (value && value.toString().trim() !== '') {
          hasDate = true;
          break;
        }
      }
      
      // Verificar hora inicio
      for (const col of possibleStartCols) {
        const value = row[col];
        if (value && value.toString().trim() !== '') {
          hasStart = true;
          break;
        }
      }
      
      // Una fila está vacía si NO tiene fecha O NO tiene hora inicio
      return !hasDate || !hasStart;
    };

    // Función para actualizar una fila con datos de Zoom
    const updateRowWithZoom = (row, zoomInfo) => {
      const updatedRow = { ...row };
      
      // Buscar columnas y actualizar
      const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
      const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
      const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
      
      // Actualizar fecha
      for (const col of possibleDateCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomInfo.fecha;
          break;
        }
      }
      
      // Actualizar inicio
      for (const col of possibleStartCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomInfo.horaInicio;
          break;
        }
      }
      
      // Actualizar fin
      for (const col of possibleEndCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomInfo.horaFin;
          break;
        }
      }
      
      updatedRow.CURSO = zoomInfo.curso;
      updatedRow.TURNO = zoomInfo.turno;
      
      return updatedRow;
    };

    docentesToProcess.forEach(docenteActual => {
      console.log(`\n--- Procesando: ${docenteActual} ---`);
      
      // Debug: Mostrar sesiones de Zoom disponibles para este docente
      const sesionesZoomDocente = parsedZoomData.filter(zoomRow => {
        const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
        return matchDocente(docenteActual, zoomDocente);
      });
      console.log(`📊 Sesiones Zoom encontradas para ${docenteActual}:`, sesionesZoomDocente.length);

      // PASO 1: Autocompletar filas existentes (prioridad máxima)
      console.log("Buscando filas para autocompletar...");
      
      // Primera pasada: Autocompletar filas que coinciden exactamente
      newData.forEach((row, index) => {
        if (row.DOCENTE !== docenteActual) return;

        // Buscar coincidencia en Zoom para esta fila específica
        for (const zoomRow of parsedZoomData) {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
          const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
          
          if (!matchDocente(docenteActual, zoomDocente) || !zoomTema) continue;

          const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
          
          if (!temaMatch) continue;

          const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
          const cursoZoom = cursoParte.trim();
          const sesionZoom = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;

          const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;
          
          if (sesionesUsadasGlobal.has(claveZoom)) continue;

          // Verificar si esta fila coincide con los datos de Zoom
          const cursoMatch = row.CURSO && normalizeCursoName(row.CURSO) === normalizeCursoName(cursoZoom);
          const seccionMatch = row.SECCION && row.SECCION.toUpperCase() === seccionZoom.toUpperCase();
          const sesionMatch = row.SESION && parseInt(String(row.SESION)) === sesionZoom;

          // Si coincide exactamente, autocompletar (sin importar si tiene o no fechas/horas)
          if (cursoMatch && seccionMatch && sesionMatch) {
            const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
            const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
            
            // Autocompletar solo los campos de fecha/hora que estén vacíos
            const updatedRow = { ...row };
            
            // Buscar y actualizar fecha solo si está vacía
            const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
            for (const col of possibleDateCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractDate(fechaInicio);
                break;
              }
            }
            
            // Actualizar hora inicio solo si está vacía
            const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
            for (const col of possibleStartCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractTime(fechaInicio);
                break;
              }
            }
            
            // Actualizar hora fin solo si está vacía
            const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
            for (const col of possibleEndCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractTime(fechaFin);
                break;
              }
            }
            
            // Actualizar turno solo si está vacío
            if (!updatedRow.TURNO || updatedRow.TURNO.toString().trim() === '') {
              updatedRow.TURNO = detectTurno(fechaInicio);
            }
            
            newData[index] = updatedRow;
            sesionesUsadasGlobal.add(claveZoom);
            updatedCount++;
            
            console.log(`✓ Fila ${index} AUTOCOMPLETADA: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
            break;
          }
        }
      });

      // Segunda pasada: Autocompletar filas completamente vacías del docente
      console.log("Buscando filas vacías para autocompletar...");
      
      newData.forEach((row, index) => {
        if (row.DOCENTE !== docenteActual) return;

        // Verificar si esta fila está realmente vacía (sin curso, sección, sesión definidos)
        const hasEmptySession = !row.CURSO || row.CURSO.toString().trim() === '' ||
                               !row.SECCION || row.SECCION.toString().trim() === '' ||
                               !row.SESION || row.SESION.toString().trim() === '';

        if (!hasEmptySession) return;

        console.log(`Fila ${index} tiene datos incompletos:`, {
          CURSO: row.CURSO,
          SECCION: row.SECCION,
          SESION: row.SESION
        });

        // Buscar cualquier sesión de Zoom no usada para este docente
        for (const zoomRow of parsedZoomData) {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
          const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
          
          if (!matchDocente(docenteActual, zoomDocente) || !zoomTema) continue;

          const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
          
          if (!temaMatch) continue;

          const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
          const cursoZoom = cursoParte.trim();
          const sesionZoom = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;

          const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;
          
          if (sesionesUsadasGlobal.has(claveZoom)) continue;

          // Autocompletar esta fila vacía con los datos de Zoom
          const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
          const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
          
          newData[index] = updateRowWithZoom(row, {
            curso: cursoZoom,
            fecha: extractDate(fechaInicio),
            horaInicio: extractTime(fechaInicio),
            horaFin: extractTime(fechaFin),
            turno: detectTurno(fechaInicio)
          });
          
          // Actualizar también los campos básicos
          newData[index].CURSO = cursoZoom;
          newData[index].SECCION = seccionZoom;
          newData[index].SESION = sesionZoom;
          
          sesionesUsadasGlobal.add(claveZoom);
          updatedCount++;
          
          console.log(`✓ Fila vacía ${index} COMPLETADA con: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
          break;
        }
      });

      // PASO 2: Crear solo lo que realmente falta
      console.log("\nVerificando si hay sesiones realmente faltantes...");
      
      parsedZoomData.forEach((zoomRow) => {
        const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
        const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
        
        if (!matchDocente(docenteActual, zoomDocente) || !zoomTema) return;

        const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
        
        if (!temaMatch) return;

        const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
        const cursoZoom = cursoParte.trim();
        const sesionZoom = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;

        const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;

        if (sesionesUsadasGlobal.has(claveZoom)) return;

        // VERIFICAR SI YA EXISTE UNA FILA CON ESTA COMBINACIÓN (incluso si tiene datos)
        const existingRow = newData.find(row => 
          row.DOCENTE === docenteActual &&
          normalizeCursoName(row.CURSO || "") === normalizeCursoName(cursoZoom) &&
          (row.SECCION || "").toUpperCase() === seccionZoom.toUpperCase() &&
          parseInt(String(row.SESION || 0)) === sesionZoom
        );

        if (existingRow) {
          console.log(`⚠️ Ya existe fila para ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}. NO se crea duplicado.`);
          sesionesUsadasGlobal.add(claveZoom);
          return;
        }

        // Solo crear nueva fila si realmente no existe
        let templateRow = newData.find(row => row.DOCENTE === docenteActual) || {};

        const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
        const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";

        // Crear objeto con TODAS las columnas del Excel actual
        const newRow = {};
        currentHeaders.forEach(header => {
          newRow[header] = templateRow[header] || "";
        });

        // Actualizar con datos específicos
        newRow.DOCENTE = docenteActual;
        newRow.CURSO = cursoZoom;
        newRow.SECCION = seccionZoom;
        newRow.SESION = sesionZoom;
        newRow.TURNO = detectTurno(fechaInicio);
        newRow.MODELO = templateRow.MODELO || "PROTECH XP";
        newRow.MODALIDAD = templateRow.MODALIDAD || "VIRTUAL";

        // Asignar fecha y horas según columnas disponibles
        const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA'];
        const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio'];
        const possibleEndCols = ['fin', 'FIN', 'Hora Fin'];
        
        for (const col of possibleDateCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = extractDate(fechaInicio);
            break;
          }
        }
        
        for (const col of possibleStartCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = extractTime(fechaInicio);
            break;
          }
        }
        
        for (const col of possibleEndCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = extractTime(fechaFin);
            break;
          }
        }

        newData.push(newRow);
        sesionesUsadasGlobal.add(claveZoom);
        createdCount++;
        
        console.log(`✓ Nueva fila realmente necesaria: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
      });
    });

    setData(newData);

    alert(`✅ Completado:\n\n${updatedCount} filas autocompletadas\n${createdCount} filas nuevas creadas`);

  } catch (error) {
    alert("❌ Error: " + error.message);
    console.error(error);
  } finally {
    setIsLoading(false);
    event.target.value = "";
  }
};

  const handleFileUpload = async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const tempLoading = { isLoading: true };
  if (activeTab) updateActiveTab(tempLoading);

  try {
    const workbook = new ExcelJS.Workbook();
    const arrayBuffer = await file.arrayBuffer();
    await workbook.xlsx.load(arrayBuffer);
    
    const sheetNames = workbook.worksheets.map((sheet, index) => ({
      index,
      name: sheet.name
    }));
    
    const worksheet = workbook.worksheets[0];
    const { data: loadedData, headers: sheetHeaders } = loadSheetData(worksheet);
    
    // Crear nueva pestaña con el archivo
    createNewTab(file.name, {
      data: loadedData,
      availableSheets: sheetNames,
      workbookData: workbook,
      currentHeaders: sheetHeaders
    });
    
  } catch (error) {
    alert("Error al cargar el archivo: " + error.message);
    console.error(error);
  } finally {
    if (activeTab) updateActiveTab({ isLoading: false });
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
      {/* SISTEMA DE PESTAÑAS */}
      <div className="bg-white rounded-t-xl shadow-lg mb-0">
        <div className="flex items-center bg-gray-100 border-b-2 border-gray-300 overflow-x-auto">
          {tabs.map((tab) => (
            <div
              key={tab.id}
              className={`flex items-center px-4 py-3 cursor-pointer border-r border-gray-300 transition-all whitespace-nowrap ${
                activeTabId === tab.id
                  ? 'bg-white border-b-4 border-blue-600 font-bold'
                  : 'bg-gray-200 hover:bg-gray-300'
              }`}
              onClick={() => setActiveTabId(tab.id)}
            >
              <span className="mr-2 text-sm">{tab.name}</span>
              <button
                onClick={(e) => {
                  e.stopPropagation();
                  closeTab(tab.id);
                }}
                className="text-red-500 hover:text-red-700 font-bold text-xl ml-2"
              >
                ×
              </button>
            </div>
          ))}
          <button
            onClick={() => document.getElementById('file-input-new-tab').click()}
            className="px-6 py-3 bg-blue-600 text-white hover:bg-blue-700 font-bold whitespace-nowrap text-sm"
          >
            + Nueva Pestaña
          </button>
          <input
            id="file-input-new-tab"
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="hidden"
          />
        </div>
      </div>

      {/* CONTENIDO DE LA PESTAÑA ACTIVA */}
      {activeTab ? (
        <>
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
            displayData={displayData}
            availableSheets={availableSheets}
            selectedSheet={selectedSheet}
            onSheetChange={handleSheetChange}
          />

          <DataTable
            data={displayData}
            headers={currentHeaders.length > 0 ? currentHeaders : []}
            dropdownOptions={dropdownOptions}
            onCellChange={handleCellChange}
            onDeleteRow={deleteRow}
          />
        </>
      ) : (
        <div className="bg-white rounded-b-xl shadow-2xl p-12 text-center">
          <h2 className="text-2xl font-bold text-gray-700 mb-4">
            No hay archivos abiertos
          </h2>
          <p className="text-gray-500 mb-6">
            Haz clic en "+ Nueva Pestaña" para cargar un archivo Excel
          </p>
        </div>
      )}
    </div>
  </div>
);
}

export default App;
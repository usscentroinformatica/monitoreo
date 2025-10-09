import React, { useState, useMemo } from "react";
import ExcelJS from "exceljs";
import ControlPanel from "./components/ControlPanel";
import DataTable from "./components/DataTable";

// Rutas a los archivos de plantilla (resueltas por Vite)
import templateExcelUrl from "./archivos/EJEMPLO.xlsx?url";
import templateCsvUrl from "./archivos/meetings_Docentes_CIS_2025_09_08_2025_09_21.csv?url";

// Función para descargar archivo Excel de plantilla
const downloadExcelTemplate = () => {
  try {
    const a = document.createElement('a');
    a.href = templateExcelUrl; // URL generada por Vite
    a.download = 'EJEMPLO.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (error) {
    console.error('Error al descargar la plantilla Excel:', error);
  }
};

// Función para descargar archivo CSV de plantilla
const downloadCSVTemplate = () => {
  try {
    const a = document.createElement('a');
    a.href = templateCsvUrl; // URL generada por Vite
    a.download = 'meetings_Docentes_CIS_template.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (error) {
    console.error('Error al descargar la plantilla CSV:', error);
  }
};

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
      currentHeaders: initialData.currentHeaders || [],
      // Caché por hoja
      sheetData: initialData.sheetData || { 0: { data: initialData.data || [], headers: initialData.currentHeaders || [] } }
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
  const setData = (newData) => {
    const currentIndex = activeTab?.selectedSheet ?? 0;
    const prevSheetData = activeTab?.sheetData || {};
    const updatedSheetData = { ...prevSheetData, [currentIndex]: { data: newData, headers: currentHeaders } };
    updateActiveTab({ data: newData, sheetData: updatedSheetData });
  };
  const setZoomData = (newZoomData) => updateActiveTab({ zoomData: newZoomData });
  const setSelectedDocente = (docente) => updateActiveTab({ selectedDocente: docente });
  const setNumFilas = (num) => updateActiveTab({ numFilas: num });
  const setIsLoading = (loading) => updateActiveTab({ isLoading: loading });
  const setAvailableSheets = (sheets) => updateActiveTab({ availableSheets: sheets });
  const setSelectedSheet = (sheet) => updateActiveTab({ selectedSheet: sheet });
  const setWorkbookData = (wb) => updateActiveTab({ workbookData: wb });
  const setCurrentHeaders = (headers) => {
    const currentIndex = activeTab?.selectedSheet ?? 0;
    const prevSheetData = activeTab?.sheetData || {};
    const prevData = prevSheetData[currentIndex]?.data || data;
    const updatedSheetData = { ...prevSheetData, [currentIndex]: { data: prevData, headers } };
    updateActiveTab({ currentHeaders: headers, sheetData: updatedSheetData });
  };

  // ===== FUNCIONES DE UTILIDAD =====
  const normalizeDocenteName = (name) => {
    if (!name) return "";
    // Ignorar iniciales de una sola letra y ordenar palabras para empatar distintos órdenes
    return name
      .toUpperCase()
      .trim()
      .split(/\s+/)
      .filter(w => w.length > 1)
      .sort()
      .join(" ");
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

  // Normalización robusta de SECCION (PEAD-a, PEAD a, A, a1, etc.)
  const normalizeSeccion = (value) => {
    if (!value) return "";
    return String(value)
      .toUpperCase()
      .trim()
      .replace(/^PEAD[-_ ]?/, "")
      .replace(/[^A-Z0-9]/g, "");
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
    const s = String(dateTimeStr).trim();

    // Mes con nombre (es/en): "September 22, 2024" / "Septiembre 22, 2024"
    const m1 = s.match(/([A-Za-zÁÉÍÓÚáéíóúñÑ]+)\s+(\d{1,2}),\s*(\d{4})/);
    if (m1) {
      const monthMap = {
        JANUARY:1,FEBRUARY:2,MARCH:3,APRIL:4,MAY:5,JUNE:6,JULY:7,AUGUST:8,SEPTEMBER:9,OCTOBER:10,NOVEMBER:11,DECEMBER:12,
        ENERO:1,FEBRERO:2,MARZO:3,ABRIL:4,MAYO:5,JUNIO:6,JULIO:7,AGOSTO:8,SEPTIEMBRE:9,OCTUBRE:10,NOVIEMBRE:11,DICIEMBRE:12
      };
      const mon = (m1[1] || "").toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      const d = String(m1[2]).padStart(2,'0');
      const y = m1[3];
      const mmNum = monthMap[mon];
      if (mmNum) {
        const mm = String(mmNum).padStart(2,'0');
        return `${d}/${mm}/${y}`;
      }
    }

    // yyyy-mm-dd
    const m2 = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m2) {
      const y = m2[1], m = String(m2[2]).padStart(2,'0'), d = String(m2[3]).padStart(2,'0');
      return `${d}/${m}/${y}`;
    }

    // dd/mm/yyyy o mm/dd/yyyy
    const m3 = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m3) {
      let a = parseInt(m3[1],10), b = parseInt(m3[2],10); let y = m3[3];
      if (a > 12 && b <= 12) { // dd/mm
        const dd = String(a).padStart(2,'0'); const mm = String(b).padStart(2,'0');
        y = y.length === 2 ? `20${y}`: y; return `${dd}/${mm}/${y}`;
      }
      if (b > 12 && a <= 12) { // mm/dd
        const dd = String(b).padStart(2,'0'); const mm = String(a).padStart(2,'0');
        y = y.length === 2 ? `20${y}`: y; return `${dd}/${mm}/${y}`;
      }
      const dd = String(a).padStart(2,'0'); const mm = String(b).padStart(2,'0');
      y = y.length === 2 ? `20${y}`: y; return `${dd}/${mm}/${y}`;
    }

    // Fallback con Date()
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      const dd = String(d.getDate()).padStart(2,'0');
      const mm = String(d.getMonth()+1).padStart(2,'0');
      const yy = String(d.getFullYear());
      return `${dd}/${mm}/${yy}`;
    }
    return s;
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
    const match = tema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
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
    console.log("Headers del Excel actual:", currentHeaders);
    
    // Merge semanal: acumula CSVs sin perder semanas anteriores y elimina duplicados básicos
    const mergedZoom = [...zoomData, ...parsedZoomData];
    const seen = new Set();
    const uniqueMerged = mergedZoom.filter(z => {
      const host = z['Anfitrión'] || z['Host'] || '';
      const topic = z['Tema'] || z['Topic'] || '';
      const start = z['Hora de inicio'] || z['Start Time'] || '';
      const end = z['Hora de finalización'] || z['End Time'] || '';
      const key = `${host}|||${topic}|||${start}|||${end}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    setZoomData(uniqueMerged);

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
    const sesionesUsadasGlobal = new Set();

    const updateRowWithZoom = (row, zoomInfo) => {
      const updatedRow = { ...row };
      
      const possibleDateCols = ['DIA', 'Dia', 'Fecha', 'FECHA', 'Columna 13', 'COLUMNA 13'];
      const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
      const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
      const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom'];
      
      for (const col of possibleDateCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomInfo.fecha;
          break;
        }
      }
      
      for (const col of possibleStartCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomInfo.horaInicio;
          break;
        }
      }
      
      for (const col of possibleEndCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomInfo.horaFin;
          break;
        }
      }
      
      // NUEVO: Guardar hora de finalización del Zoom
      for (const col of possibleFinalizaCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomInfo.horaFinalizacion;
          break;
        }
      }
      
      updatedRow.CURSO = zoomInfo.curso;
      updatedRow.TURNO = zoomInfo.turno;
      
      return updatedRow;
    };

    docentesToProcess.forEach(docenteActual => {
      console.log(`\n--- Procesando: ${docenteActual} ---`);
      
      const sesionesZoomDocente = parsedZoomData.filter(zoomRow => {
        const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
        return matchDocente(docenteActual, zoomDocente);
      });
      console.log(`📊 Sesiones Zoom encontradas para ${docenteActual}:`, sesionesZoomDocente.length);

      console.log("Buscando filas para autocompletar...");
      
      // Primera pasada: Autocompletar filas que coinciden exactamente
      newData.forEach((row, index) => {
        if (row.DOCENTE !== docenteActual) return;

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

          const cursoMatch = row.CURSO && normalizeCursoName(row.CURSO) === normalizeCursoName(cursoZoom);
          const seccionMatch = row.SECCION && normalizeSeccion(row.SECCION) === normalizeSeccion(seccionZoom);
          const sesionMatch = row.SESION && parseInt(String(row.SESION)) === sesionZoom;

          if (cursoMatch && seccionMatch && sesionMatch) {
            const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
            const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
            
            const updatedRow = { ...row };
            
            const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
            for (const col of possibleDateCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractDate(fechaInicio);
                break;
              }
            }
            
            const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
            for (const col of possibleStartCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractTime(fechaInicio);
                break;
              }
            }
            
            const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
            for (const col of possibleEndCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractTime(fechaFin);
                break;
              }
            }
            
            // NUEVO: Guardar hora de finalización del Zoom
            const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom'];
            for (const col of possibleFinalizaCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractTime(fechaFin);
                break;
              }
            }
            
            if (!updatedRow.TURNO || updatedRow.TURNO.toString().trim() === '') {
              updatedRow.TURNO = detectTurno(fechaInicio);
            }
            
            const horaInicioZoom = extractTime(fechaInicio);
            const horaProgramada = updatedRow['HORA INICIO'] || row['HORA INICIO'];
            updatedRow['INICIO SESION 10 MINUTOS ANTES'] = verificarInicio10MinutesAntes(horaInicioZoom, horaProgramada);
            
            newData[index] = updatedRow;
            sesionesUsadasGlobal.add(claveZoom);
            updatedCount++;
            
            console.log(`✓ Fila ${index} AUTOCOMPLETADA: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
            break;
          }
        }
      });

      // Segunda pasada: Autocompletar filas vacías
      console.log("Buscando filas vacías para autocompletar...");
      
      newData.forEach((row, index) => {
        if (row.DOCENTE !== docenteActual) return;

        const hasEmptySession = !row.CURSO || row.CURSO.toString().trim() === '' ||
                               !row.SECCION || row.SECCION.toString().trim() === '' ||
                               !row.SESION || row.SESION.toString().trim() === '';

        if (!hasEmptySession) return;

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

          const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
          const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
          
          newData[index] = updateRowWithZoom(row, {
            curso: cursoZoom,
            fecha: extractDate(fechaInicio),
            horaInicio: extractTime(fechaInicio),
            horaFin: extractTime(fechaFin),
            horaFinalizacion: extractTime(fechaFin),
            turno: detectTurno(fechaInicio)
          });
          
          newData[index].CURSO = cursoZoom;
          newData[index].SECCION = seccionZoom;
          newData[index].SESION = sesionZoom;
          
          const horaInicioZoom = extractTime(fechaInicio);
          const horaProgramada = newData[index]['HORA INICIO'];
          newData[index]['INICIO SESION 10 MINUTOS ANTES'] = verificarInicio10MinutesAntes(horaInicioZoom, horaProgramada);
          
          sesionesUsadasGlobal.add(claveZoom);
          updatedCount++;
          
          console.log(`✓ Fila vacía ${index} COMPLETADA con: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
          break;
        }
      });

      // Fallback adicional: emparejar por horario cuando el Tema no contiene PEAD-
      // Para filas del docente con fecha/hora vacías y con hora programada, buscar la sesión de Zoom más cercana en minutos
      {
        const usedZoomByStart = new Set();
        const timeToMinutes = (timeStr) => {
          if (!timeStr || typeof timeStr !== 'string') return 0;
          let s = timeStr.trim();
          s = s.replace(/a\.\s*m\.|p\.\s*m\./gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
          const m12 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i);
          if (m12) {
            let h = parseInt(m12[1]) || 0; const min = parseInt(m12[2]) || 0; const sec = parseInt(m12[3]||'0')||0; const p = m12[4].toUpperCase();
            if (p === 'PM' && h !== 12) h += 12; if (p === 'AM' && h === 12) h = 0;
            return h * 60 + min + sec / 60;
          }
          const m24 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
          if (m24) {
            const h = parseInt(m24[1]) || 0; const min = parseInt(m24[2]) || 0; const sec = parseInt(m24[3]||'0')||0;
            return h * 60 + min + sec / 60;
          }
          return 0;
        };

        newData.forEach((row, index) => {
          if (row.DOCENTE !== docenteActual) return;

          const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
          const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
          const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];

          const horaProg = row['HORA INICIO'] || row['Hora Inicio'] || row['INICIO'] || row['inicio'] || '';
          const hasFecha = possibleDateCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
          const hasHoraInicio = possibleStartCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
          const hasHoraFin = possibleEndCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');

          if (!horaProg || (hasFecha && hasHoraInicio && hasHoraFin)) return;

          const tProg = timeToMinutes(String(horaProg));
          if (tProg === 0) return;

          let bestZoom = null; let bestStartStr = null; let bestEndStr = null; let bestDiff = Infinity;

          parsedZoomData.forEach((zoomRow) => {
            const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
            if (!matchDocente(docenteActual, zoomDocente)) return;

            const startStr = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
            const endStr = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
            if (!startStr || usedZoomByStart.has(startStr)) return;

            const tz = timeToMinutes(extractTime(startStr));
            if (tz === 0) return;

            const diff = Math.abs(tz - tProg);
            // Umbral de 120 minutos para asociar sesiones por horario
            if (diff < bestDiff && diff <= 120) {
              bestDiff = diff;
              bestZoom = zoomRow;
              bestStartStr = startStr;
              bestEndStr = endStr;
            }
          });

          if (bestZoom) {
            const fechaInicio = bestStartStr || "";
            const fechaFin = bestEndStr || "";

            const updatedRow = updateRowWithZoom(row, {
              curso: row.CURSO || extractCursoFromTema(bestZoom['Tema'] || bestZoom['Topic'] || ""),
              fecha: extractDate(fechaInicio),
              horaInicio: extractTime(fechaInicio),
              horaFin: extractTime(fechaFin),
              horaFinalizacion: extractTime(fechaFin),
              turno: detectTurno(fechaInicio)
            });

            newData[index] = updatedRow;

            const horaInicioZoom = extractTime(fechaInicio);
            const horaProgramada = updatedRow['HORA INICIO'] || row['HORA INICIO'];
            newData[index]['INICIO SESION 10 MINUTOS ANTES'] = verificarInicio10MinutesAntes(horaInicioZoom, horaProgramada);

            usedZoomByStart.add(bestStartStr);
            updatedCount++;
            console.log(`✓ Fallback por horario aplicado en fila ${index} (dif ${Math.round(bestDiff)} min)`);
          }
        });
      }

      // Tercera pasada: Crear nuevas filas necesarias
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

        const existingRow = newData.find(row => 
          row.DOCENTE === docenteActual &&
          normalizeCursoName(row.CURSO || "") === normalizeCursoName(cursoZoom) &&
          normalizeSeccion(row.SECCION || "") === normalizeSeccion(seccionZoom) &&
          parseInt(String(row.SESION || 0)) === sesionZoom
        );

        if (existingRow) {
          console.log(`⚠️ Ya existe fila para ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}. NO se crea duplicado.`);
          sesionesUsadasGlobal.add(claveZoom);
          return;
        }

        let templateRow = newData.find(row => row.DOCENTE === docenteActual) || {};

        const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
        const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";

        const newRow = {};
        currentHeaders.forEach(header => {
          newRow[header] = templateRow[header] || "";
        });

        newRow.DOCENTE = docenteActual;
        newRow.CURSO = cursoZoom;
        newRow.SECCION = seccionZoom;
        newRow.SESION = sesionZoom;
        newRow.TURNO = detectTurno(fechaInicio);
        newRow.MODELO = templateRow.MODELO || "PROTECH XP";
        newRow.MODALIDAD = templateRow.MODALIDAD || "VIRTUAL";

        const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA'];
        const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio'];
        const possibleEndCols = ['fin', 'FIN', 'Hora Fin'];
        const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom'];
        
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

        // Guardar hora de finalización del Zoom
        for (const col of possibleFinalizaCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = extractTime(fechaFin);
            break;
          }
        }

        const horaInicioZoom = extractTime(fechaInicio);
        const horaProgramada = newRow['HORA INICIO'];
        newRow['INICIO SESION 10 MINUTOS ANTES'] = verificarInicio10MinutesAntes(horaInicioZoom, horaProgramada);

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

  // ===== FUNCIÓN PRINCIPAL: AUTOCOMPLETAR CON ZOOM =====
  const handleAutocompletarConZoom = async () => {
  if (data.length === 0) {
    alert("⚠️ Primero carga el archivo Excel TRADICIONAL");
    return;
  }

  setIsLoading(true);
  
  try {
    console.log("=== INICIANDO PROCESO COMPLETO ===");
    
    // PASO 1: Autocompletar filas existentes con datos de Zoom (si hay CSV cargado)
    let dataProcesada = [...data];
    
    if (zoomData.length > 0) {
      console.log("\n📋 PASO 1: Autocompletando filas existentes con datos de Zoom");
      
      dataProcesada.forEach((row, index) => {
        const docente = row.DOCENTE;
        const curso = row.CURSO;
        const seccion = row.SECCION;
        const sesion = row.SESION;

        if (!docente || !curso || !seccion || !sesion) return;

        const sesionZoom = zoomData.find(zoomRow => {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
          const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
          
          if (!matchDocente(docente, zoomDocente)) return false;
          
          const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
          if (!temaMatch) return false;
          
          const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
          const cursoZoom = cursoParte.trim();
          const sesionZoom = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;
          
          return normalizeCursoName(cursoZoom) === normalizeCursoName(curso) &&
                 normalizeSeccion(seccionZoom) === normalizeSeccion(seccion) &&
                 sesionZoom === parseInt(sesion);
        });

        if (sesionZoom) {
          const fechaInicio = sesionZoom['Hora de inicio'] || sesionZoom['Start Time'] || "";
          const fechaFin = sesionZoom['Hora de finalización'] || sesionZoom['End Time'] || "";
          
          const fechaExtraida = extractDate(fechaInicio);
          const horaInicioExtraida = extractTime(fechaInicio);
          const horaFinExtraida = extractTime(fechaFin);
          
          // Autocompletar campos
          const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
          for (const col of possibleDateCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = fechaExtraida;
              break;
            }
          }
          
          const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
          for (const col of possibleStartCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = horaInicioExtraida;
              break;
            }
          }
          
          const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
          for (const col of possibleEndCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = horaFinExtraida;
              break;
            }
          }
          
          dataProcesada[index].TURNO = detectTurno(fechaInicio);
          
          if (currentHeaders.includes('INICIO SESION 10 MINUTOS ANTES')) {
            dataProcesada[index]['INICIO SESION 10 MINUTOS ANTES'] = verificarInicio10MinutesAntes(
              horaInicioExtraida,
              row['HORA INICIO'] || ""
            );
          }
          
          const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)'];
          for (const col of possibleFinalizaCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = horaFinExtraida;
              break;
            }
          }
          
          console.log(`  ✓ Autocompletado: ${docente} - ${curso} - Sesión ${sesion}`);
        }
      });
    }
    
    // Fallback adicional (Autocompletar): emparejar por horario usando zoomData
    if (zoomData.length > 0) {
      const usedZoomByStart = new Set();
      const timeToMinutes = (timeStr) => {
        if (!timeStr || typeof timeStr !== 'string') return 0;
        let s = timeStr.trim();
        s = s.replace(/a\.\s*m\.|p\.\s*m\./gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
        const m12 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i);
        if (m12) {
          let h = parseInt(m12[1]) || 0; const min = parseInt(m12[2]) || 0; const sec = parseInt(m12[3]||'0')||0; const p = m12[4].toUpperCase();
          if (p === 'PM' && h !== 12) h += 12; if (p === 'AM' && h === 12) h = 0;
          return h * 60 + min + sec / 60;
        }
        const m24 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        if (m24) {
          const h = parseInt(m24[1]) || 0; const min = parseInt(m24[2]) || 0; const sec = parseInt(m24[3]||'0')||0;
          return h * 60 + min + sec / 60;
        }
        return 0;
      };

      dataProcesada.forEach((row, index) => {
        const docente = row.DOCENTE || '';
        const horaProg = row['HORA INICIO'] || row['Hora Inicio'] || row['INICIO'] || row['inicio'] || '';
        if (!docente || !horaProg) return;

        // Detectar si faltan fecha/hora reales
        const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
        const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
        const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
        const hasFecha = possibleDateCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
        const hasHoraInicio = possibleStartCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
        const hasHoraFin = possibleEndCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
        if (hasFecha && hasHoraInicio && hasHoraFin) return;

        const tProg = timeToMinutes(String(horaProg));
        if (tProg === 0) return;

        let bestZoom = null; let bestStartStr = null; let bestEndStr = null; let bestDiff = Infinity;
        zoomData.forEach(zoomRow => {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || '';
          if (!matchDocente(docente, zoomDocente)) return;
          const startStr = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || '';
          const endStr = zoomRow['Hora de finalización'] || zoomRow['End Time'] || '';
          if (!startStr || usedZoomByStart.has(startStr)) return;
          const tz = timeToMinutes(extractTime(startStr));
          if (tz === 0) return;
          const diff = Math.abs(tz - tProg);
          if (diff < bestDiff && diff <= 120) {
            bestDiff = diff; bestZoom = zoomRow; bestStartStr = startStr; bestEndStr = endStr;
          }
        });

        if (bestZoom) {
          const fechaInicio = bestStartStr || '';
          const fechaFin = bestEndStr || '';

          // Completar campos reales
          const setIfHasHeader = (obj, colNames, val) => {
            for (const c of colNames) { if (currentHeaders.includes(c)) { obj[c] = val; break; } }
          };

          setIfHasHeader(row, possibleDateCols, extractDate(fechaInicio));
          setIfHasHeader(row, possibleStartCols, extractTime(fechaInicio));
          setIfHasHeader(row, possibleEndCols, extractTime(fechaFin));

          // Turno y 10 minutos antes
          row.TURNO = row.TURNO && String(row.TURNO).trim() !== '' ? row.TURNO : detectTurno(fechaInicio);
          row['INICIO SESION 10 MINUTOS ANTES'] = verificarInicio10MinutesAntes(extractTime(fechaInicio), row['HORA INICIO'] || '');

          usedZoomByStart.add(bestStartStr);
        }
      });
    }

    // PASO 2: Detectar docentes/cursos únicos
    console.log("\n📋 PASO 2: Detectando docentes y cursos");
    
    const docentesCursosMap = new Map();
    
    dataProcesada.forEach(row => {
      const docente = row.DOCENTE;
      const curso = row.CURSO;
      const seccion = row.SECCION;
      
      if (!docente || !curso || !seccion) return;
      
      const key = `${docente}|||${curso}|||${seccion}`;
      
      if (!docentesCursosMap.has(key)) {
        docentesCursosMap.set(key, {
          docente,
          curso,
          seccion,
          periodo: row.PERIODO || "",
          sesionesExistentes: []
        });
      }
      
      // Guardar las sesiones que ya tienen datos autocompletados
      if (row.SESION) {
        docentesCursosMap.get(key).sesionesExistentes.push({
          sesion: parseInt(row.SESION),
          datos: row
        });
      }
    });

    // Extender con docentes del combo para asegurar generación aunque no exista Zoom
    if (Array.isArray(uniqueDocentes) && uniqueDocentes.length > 0) {
      uniqueDocentes.forEach(docName => {
        const rowsDoc = dataProcesada.filter(r => r.DOCENTE === docName);
        if (rowsDoc.length === 0) return; // sin base en Excel, no se generan combos
        rowsDoc.forEach(r => {
          const curso = r.CURSO;
          const seccion = r.SECCION || r['SECCIÓN'];
          if (!curso || String(curso).trim() === '' || !seccion || String(seccion).trim() === '') return;
          const key = `${docName}|||${curso}|||${seccion}`;
          if (!docentesCursosMap.has(key)) {
            docentesCursosMap.set(key, {
              docente: docName,
              curso,
              seccion,
              periodo: r.PERIODO || "",
              sesionesExistentes: []
            });
          }
        });
      });
    }

    console.log(`Total combinaciones encontradas: ${docentesCursosMap.size}`);
    
    // PASO 3: Crear 16 sesiones por cada curso
    console.log("\n📋 PASO 3: Creando 16 sesiones por cada curso/docente");
    
    const nuevasFilas = [];
    let totalCreadas = 0;
    let totalConDatos = 0;

    docentesCursosMap.forEach(({ docente, curso, seccion, periodo, sesionesExistentes }) => {
      console.log(`\n--- ${docente} - ${curso} - ${seccion} ---`);
      
      // Crear un mapa de sesiones existentes para búsqueda rápida
      const mapaSesiones = new Map();
      sesionesExistentes.forEach(item => {
        mapaSesiones.set(item.sesion, item.datos);
      });
      
      // Fila de referencia para horario programado (prioriza sesiones existentes del mismo curso/sección)
      const baseRefRow = (sesionesExistentes[0] && sesionesExistentes[0].datos) ||
                         data.find(r => r.DOCENTE === docente && r.CURSO === curso && r.SECCION === seccion) ||
                         data.find(r => r.DOCENTE === docente) || {};
      
      // Crear 16 sesiones (del 1 al 16)
      for (let sesion = 1; sesion <= 16; sesion++) {
        const nuevaFila = {};
        
        // Inicializar todas las columnas como vacías
        currentHeaders.forEach(header => {
          nuevaFila[header] = "";
        });
        
        // Establecer valores básicos
        nuevaFila.SESION = sesion;
        nuevaFila.DOCENTE = docente;
        nuevaFila.PERIODO = periodo;
        nuevaFila.CURSO = curso;
        nuevaFila.SECCION = seccion;
        
        // Prellenar horario programado y otros campos base desde fila de referencia
        if (baseRefRow) {
          if (baseRefRow['HORA INICIO']) nuevaFila['HORA INICIO'] = baseRefRow['HORA INICIO'];
          if (baseRefRow['HORA FIN']) nuevaFila['HORA FIN'] = baseRefRow['HORA FIN'];
          if (baseRefRow['DIAS']) nuevaFila['DIAS'] = baseRefRow['DIAS'];
        }
        
        // Si esta sesión ya tiene datos autocompletados, copiarlos
        // Si esta sesión ya tiene datos autocompletados, copiarlos
if (mapaSesiones.has(sesion)) {
  const datosExistentes = mapaSesiones.get(sesion);
  
  // Copiar solo los campos que fueron autocompletados desde Zoom
  const camposAutocompletables = [
    'Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia',
    'inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO',
    'fin', 'FIN', 'Hora Fin', 'HORA FIN',
    'TURNO', // ⭐ Solo copiar si viene de Zoom
    'INICIO SESION 10 MINUTOS ANTES',
    'FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)'
  ];
  
  camposAutocompletables.forEach(campo => {
    if (currentHeaders.includes(campo) && datosExistentes[campo]) {
      nuevaFila[campo] = datosExistentes[campo];
    }
  });
  
  totalConDatos++;
  console.log(`  ✓ Sesión ${sesion}: CON DATOS`);
} else {
  console.log(`  ○ Sesión ${sesion}: VACÍA`);
}
        
        nuevasFilas.push(nuevaFila);
        totalCreadas++;
      }
    });

    // PASO 4: Fusionar 16 sesiones con datos existentes (no borrar)
    const mergedDataFinal = [...data];
    const normalizeSeccionLocal = (v) => (v ? String(v).toUpperCase().trim().replace(/^PEAD[-_ ]?/i, '').replace(/[^A-Z0-9]/g, '') : '');
    nuevasFilas.forEach(nfila => {
      const idx = mergedDataFinal.findIndex(r =>
        (r.DOCENTE || '') === (nfila.DOCENTE || '') &&
        (normalizeCursoName(r.CURSO || '') === normalizeCursoName(nfila.CURSO || '')) &&
        (normalizeSeccionLocal(r.SECCION || r['SECCIÓN'] || '') === normalizeSeccionLocal(nfila.SECCION || nfila['SECCIÓN'] || '')) &&
        (parseInt(r.SESION || 0) === parseInt(nfila.SESION || 0))
      );
      if (idx >= 0) {
        // Completar solo campos vacíos
        Object.keys(nfila).forEach(k => {
          const v = mergedDataFinal[idx][k];
          if (v === undefined || String(v).trim() === '') {
            mergedDataFinal[idx][k] = nfila[k];
          }
        });
      } else {
        mergedDataFinal.push(nfila);
      }
    });
    setData(mergedDataFinal);
    
    alert(`✅ Proceso completado:\n\n${totalCreadas} sesiones generadas (16 por curso)\n${totalConDatos} sesiones con datos de Zoom\nFusionado sin borrar registros existentes`);
    
    console.log("=== PROCESO FINALIZADO ===");
    
  } catch (error) {
    console.error("Error en proceso:", error);
    alert("❌ Error: " + error.message);
  } finally {
    setIsLoading(false);
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
        currentHeaders: sheetHeaders,
        sheetData: { 0: { data: loadedData, headers: sheetHeaders } }
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
  const isMonitoreo = (worksheet.name || '').toLowerCase().includes('monitoreo');
  const MAX_SCAN_ROWS = isMonitoreo ? 10 : 3;
  const MAX_COLS = isMonitoreo ? 60 : 25;

  // Función auxiliar MUY ROBUSTA para extraer texto de celdas
  const extractCellText = (cell) => {
    if (!cell) return "";
    
    // Primero intentar obtener el valor crudo
    let rawValue = cell.value;
    if (rawValue === null || rawValue === undefined) return "";
    
    // Si es un string simple, devolverlo directamente
    if (typeof rawValue === 'string') {
      return rawValue.trim();
    }
    
    // Si es un número, convertirlo a string
    if (typeof rawValue === 'number') {
      return String(rawValue);
    }
    
    // Si es un objeto complejo de Excel
    if (typeof rawValue === 'object') {
      // Texto de hipervínculo
      if (rawValue.hyperlink && rawValue.text) {
        return String(rawValue.text).trim();
      }
      
      // Texto enriquecido
      if (Array.isArray(rawValue.richText)) {
        return rawValue.richText.map(rt => rt.text || '').join('').trim();
      }
      
      // Texto directo
      if (rawValue.text !== undefined) {
        return String(rawValue.text).trim();
      }
      
      // Resultado de fórmula
      if (rawValue.result !== undefined) {
        const res = rawValue.result;
        if (res instanceof Date) {
          return res.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).trim();
        }
        if (typeof res === 'number' || typeof res === 'string') {
          return String(res).trim();
        }
        return '';
      }
      // Si es un objeto no reconocible, no devolver [object Object]
      return '';
    }
    
    // Fallback: convertir a string
    return String(rawValue).trim();
  };

  // Función SÚPER SIMPLE: leer TODO directamente sin filtros
  const readAllCellsInRow = (row, maxCols = 30) => {
    const cells = [];
    for (let col = 1; col <= maxCols; col++) {
      const cell = row.getCell(col);
      const text = extractCellText(cell);
      cells.push(text);
    }
    return cells;
  };

  console.log('🔍 DETECTANDO ENCABEZADOS - MODO SIMPLE Y DIRECTO');
  
  let bestRowNum = -1; let bestCount = 0; let bestKeywordRowNum = -1; let bestKeywordCount = 0; const headerKeywords = ['DOCENTE','CURSO','SECCION','SECCIÓN','SESION','HORA INICIO','HORA FIN','FECHA','DIA','COLUMNA 13'];
  // MÉTODO 1: Leer directamente las primeras filas SIN filtros
  for (let rowNum = 1; rowNum <= MAX_SCAN_ROWS; rowNum++) {
    console.log(`\n📋 === FILA ${rowNum} ===`);
    
    const row = worksheet.getRow(rowNum);
    const allCells = readAllCellsInRow(row, MAX_COLS);
    
    // Mostrar TODO lo que encuentra
    allCells.forEach((cellText, index) => {
      if (cellText && cellText.trim() !== '') {
        console.log(`   Col ${index + 1}: "${cellText}"`);
      }
    });
    
    // Contar celdas con contenido real
    const nonEmptyCells = allCells.filter(cell => cell && cell.trim() !== '').length;
    console.log(`   📊 Total celdas con contenido: ${nonEmptyCells}`);
    if (isMonitoreo && nonEmptyCells > bestCount) { bestCount = nonEmptyCells; bestRowNum = rowNum; }
    // Conteo de palabras clave típicas de MONITOREO
    const upperCells = allCells.map(c => (c || '').toString().toUpperCase().trim());
    const matchCount = headerKeywords.reduce((acc, k) => acc + (upperCells.includes(k) ? 1 : 0), 0);
    if (isMonitoreo && matchCount > bestKeywordCount) { bestKeywordCount = matchCount; bestKeywordRowNum = rowNum; }
    
    // Si tiene un número razonable de celdas con contenido, usar esta fila
    if (nonEmptyCells >= 5) {
      headerRowIndex = rowNum;
      console.log(`✅ SELECCIONANDO FILA ${rowNum} COMO ENCABEZADOS`);
      
      // Filtrar solo las celdas vacías del final
      let validHeaders = [...allCells];
      while (validHeaders.length > 0 && (!validHeaders[validHeaders.length - 1] || validHeaders[validHeaders.length - 1].trim() === '')) {
        validHeaders.pop();
      }
      
      // Usar los textos reales, sin cambiarlos
      sheetHeaders = validHeaders.map((header, index) => {
        if (header && header.trim() !== '') {
          return header.trim();
        } else {
          return `COLUMNA_${index + 1}`;
        }
      });
      
      console.log(`📋 ENCABEZADOS FINALES:`, sheetHeaders);
      break;
    }
  }

  // Ajuste adicional (MONITOREO): si detectamos una fila con palabras clave típicas, úsala como encabezados
  if (isMonitoreo && bestKeywordRowNum > 0) {
    const headersUpper = sheetHeaders.map(h => (h || '').toString().toUpperCase());
    const hasAnyKeyword = headerKeywords.some(k => headersUpper.includes(k));
    if (sheetHeaders.length === 0 || !hasAnyKeyword) {
      const kwRow = worksheet.getRow(bestKeywordRowNum);
      const kwCells = readAllCellsInRow(kwRow, MAX_COLS);
      let validHeaders = [...kwCells];
      while (validHeaders.length > 0 && (!validHeaders[validHeaders.length - 1] || validHeaders[validHeaders.length - 1].trim() === '')) {
        validHeaders.pop();
      }
      sheetHeaders = validHeaders.map((header, index) => header && header.trim() !== '' ? header.trim() : `COLUMNA_${index + 1}`);
      headerRowIndex = bestKeywordRowNum;
    }
  }
  // FALLBACK: Si no encontró nada en el método 1, intentar con la mejor fila detectada para MONITOREO
  if (sheetHeaders.length === 0 && isMonitoreo && bestRowNum > 0) {
    const bestRow = worksheet.getRow(bestRowNum);
    const bestCells = readAllCellsInRow(bestRow, MAX_COLS);
    let validHeaders = [...bestCells];
    while (validHeaders.length > 0 && (!validHeaders[validHeaders.length - 1] || validHeaders[validHeaders.length - 1].trim() === '')) {
      validHeaders.pop();
    }
    sheetHeaders = validHeaders.map((header, index) => header && header.trim() !== '' ? header.trim() : `COLUMNA_${index + 1}`);
    headerRowIndex = bestRowNum;
  }
  // FALLBACK: Si no encontró nada, usar la primera fila que tenga cualquier dato
  if (sheetHeaders.length === 0) {
    console.log('⚠️ FALLBACK: Buscando cualquier fila con datos...');
    
    for (let i = 1; i <= (isMonitoreo ? 10 : 5); i++) {
      console.log(`   Probando fila ${i}...`);
      const row = worksheet.getRow(i);
      const cells = readAllCellsInRow(row, MAX_COLS);
      const nonEmpty = cells.filter(c => c && c.trim() !== '');
      
      if (nonEmpty.length > 0) {
        console.log(`   ✅ Encontré ${nonEmpty.length} celdas en fila ${i}`);
        headerRowIndex = i;
        sheetHeaders = cells.map((cell, idx) => cell || `COLUMNA_${idx + 1}`);
        
        // Limpiar headers vacíos del final
        while (sheetHeaders.length > 0 && sheetHeaders[sheetHeaders.length - 1].startsWith('COLUMNA_')) {
          sheetHeaders.pop();
        }
        break;
      }
    }
    
    // Último recurso extremo
    if (sheetHeaders.length === 0) {
      console.log('💥 ÚLTIMO RECURSO: Creando headers genéricos');
      sheetHeaders = Array.from({ length: 20 }, (_, i) => `COLUMNA_${i + 1}`);
      headerRowIndex = 1;
    }
  }

  console.log(`🎯 RESULTADO FINAL:`);
  console.log(`📄 Hoja: ${worksheet.name}`);
  console.log(`📏 Fila de encabezados: ${headerRowIndex}`);
  console.log(`📋 Encabezados detectados (${sheetHeaders.length}):`, sheetHeaders);
  
  // Verificar si tenemos headers reales o genéricos
  const genericHeaders = sheetHeaders.filter(h => h && h.startsWith('COLUMNA_')).length;
  if (genericHeaders > sheetHeaders.length / 2) {
    console.warn('⚠️ ADVERTENCIA: La mayoría de headers son genéricos. Es posible que la detección no haya funcionado correctamente.');
  } else {
    console.log('✅ Headers reales detectados correctamente');
  }

  console.log(`\n🗂️ LEYENDO DATOS desde fila ${headerRowIndex + 1}...`);
  const maxRowsToCheck = Math.min(worksheet.rowCount || 1000, 1000);

  // ⭐ FUNCIÓN getCellValue - CORREGIDA
  const getCellValue = (cell, columnIndex) => {
    if (!cell || cell.value === null || cell.value === undefined) return "";

    const rawValue = cell.value;
    const headerName = sheetHeaders[columnIndex - 1];

    // Números específicos solo para SESION (no tocar tiempos ni duraciones)
    if (headerName && ['SESION'].includes(headerName.toUpperCase())) {
      if (typeof rawValue === 'number') {
        if (rawValue === 0) return "0";
        return Math.round(rawValue);
      }
      if (typeof rawValue === 'string') {
        const num = parseInt(rawValue);
        if (!isNaN(num)) {
          if (num === 0) return "0";
          return num;
        }
      }
      if (typeof rawValue === 'object') {
        const res = rawValue.result ?? rawValue.value;
        if (res !== undefined) {
          const num = parseInt(res);
          if (!isNaN(num)) return num;
          return String(res);
        }
        return "";
      }
    }

    // EFICIENCIA
    if (headerName && headerName.toUpperCase() === 'EFICIENCIA') {
      if (typeof rawValue === 'object') {
        if (rawValue.result !== undefined) {
          const result = rawValue.result;
          if (typeof result === 'number') {
            return `${Math.round(result * 100)}%`;
          }
          return String(result);
        }
        if (rawValue.formula) return rawValue.formula;
        if (rawValue.value !== undefined) {
          const value = rawValue.value;
          if (typeof value === 'number') {
            return `${Math.round(value * 100)}%`;
          }
          return String(value);
        }
        return "";
      }

      if (typeof rawValue === 'number') {
        if (rawValue > 0 && rawValue <= 1) {
          return `${Math.round(rawValue * 100)}%`;
        }
        return `${Math.round(rawValue)}%`;
      }

      if (typeof rawValue === 'string') {
        if (rawValue.includes('%')) return rawValue;
        const num = parseFloat(rawValue);
        if (!isNaN(num)) {
          if (num > 0 && num <= 1) {
            return `${Math.round(num * 100)}%`;
          }
          return `${Math.round(num)}%`;
        }
      }
      return rawValue || '';
    }

    // TIEMPO EFECTIVO DICTADO
    if (headerName && headerName === 'TIEMPO EFECTIVO DICTADO') {
      const toHMS = (num) => {
        const totalSeconds = Math.round(num * 24 * 60 * 60);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      };
      if (typeof rawValue === 'number') {
        if (rawValue >= 0 && rawValue < 1) return toHMS(rawValue);
        return String(rawValue);
      }
      if (rawValue instanceof Date) {
        const hours = rawValue.getUTCHours();
        const minutes = rawValue.getUTCMinutes();
        const seconds = rawValue.getUTCSeconds();
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      }
      if (typeof rawValue === 'object') {
        const res = rawValue.result ?? rawValue.value;
        if (res instanceof Date) {
          const hours = res.getUTCHours();
          const minutes = res.getUTCMinutes();
          const seconds = res.getUTCSeconds();
          return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
        }
        if (typeof res === 'number') {
          if (res >= 0 && res < 1) return toHMS(res);
          return String(res);
        }
        if (typeof res === 'string') return res;
        return "";
      }
    }

    // Fechas
    if (rawValue instanceof Date) {
      const year = rawValue.getUTCFullYear();
      const hours = rawValue.getUTCHours();
      const minutes = rawValue.getUTCMinutes();
      const seconds = rawValue.getUTCSeconds();

      if (year === 1899 || year === 1900) {
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      }

      if (hours !== 0 || minutes !== 0 || seconds !== 0) {
        return `${rawValue.toLocaleDateString('es-ES', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        })} ${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      } else {
        return rawValue.toLocaleDateString('es-ES', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        });
      }
    }

    // Números
    if (typeof rawValue === 'number') {
      if (cell.numFmt && cell.numFmt.includes('%')) {
        return Math.round(rawValue * 100) + '%';
      }

      const cellFormat = (cell.numFmt || '').toLowerCase();
      const isTimeFormat = cellFormat.includes('h:mm') || cellFormat.includes('hh:mm') || cellFormat.includes('[h]') || cellFormat.includes('h:mm:ss') || cellFormat.includes('hh:mm:ss') || cellFormat.includes('am/pm') || cellFormat.includes('a/p');
      const isDateFormat = cellFormat.includes('d/m') || cellFormat.includes('dd/mm') || cellFormat.includes('m/d') || cellFormat.includes('yyyy') || cellFormat.includes('dd-mm') || cellFormat.includes('mm-dd');

      if (rawValue >= 0 && rawValue < 1 && (!isDateFormat || isTimeFormat)) {
        const totalSeconds = Math.round(rawValue * 24 * 60 * 60);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      }

      if (rawValue >= 1 && rawValue < 100000) {
        try {
          const excelDate = new Date((rawValue - 25569) * 86400 * 1000);
          if (!isNaN(excelDate.getTime()) && excelDate.getFullYear() > 1900 && excelDate.getFullYear() < 2100) {
            if (isTimeFormat && !isDateFormat) {
              const hours = excelDate.getUTCHours();
              const minutes = excelDate.getUTCMinutes();
              const seconds = excelDate.getUTCSeconds();
              return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
            }
            return excelDate.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
          }
        } catch (e) {
          console.error('Error convirtiendo fecha:', e);
        }
      }
      return String(rawValue);
    }

    // Objetos complejos (manejar correctamente resultados de fórmulas y tiempos)
    if (typeof rawValue === 'object') {
      if (rawValue.hyperlink) {
        const text = rawValue.text || rawValue.hyperlink || '';
        return text.length > 50 ? text.substring(0, 47) + '...' : String(text).trim();
      }
      if (Array.isArray(rawValue.richText)) {
        return rawValue.richText.map(rt => rt.text || '').join('').trim();
      }
      if (rawValue.text !== undefined) return String(rawValue.text).trim();
      const handleNumericLike = (num) => {
        const cellFormat = (cell.numFmt || '').toLowerCase();
        const isTimeFormat = cellFormat.includes('h:mm') || cellFormat.includes('hh:mm') || cellFormat.includes('[h]') || cellFormat.includes('h:mm:ss') || cellFormat.includes('hh:mm:ss') || cellFormat.includes('am/pm') || cellFormat.includes('a/p');
        const isDateFormat = cellFormat.includes('d/m') || cellFormat.includes('dd/mm') || cellFormat.includes('m/d') || cellFormat.includes('yyyy') || cellFormat.includes('dd-mm') || cellFormat.includes('mm-dd');
        if (cell.numFmt && cell.numFmt.includes('%')) {
          return Math.round(num * 100) + '%';
        }
        if (num >= 0 && num < 1 && (!isDateFormat || isTimeFormat)) {
          const totalSeconds = Math.round(num * 24 * 60 * 60);
          const hours = Math.floor(totalSeconds / 3600);
          const minutes = Math.floor((totalSeconds % 3600) / 60);
          const seconds = totalSeconds % 60;
          return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
        }
        if (num >= 1 && num < 100000) {
          try {
            const excelDate = new Date((num - 25569) * 86400 * 1000);
            if (!isNaN(excelDate.getTime()) && excelDate.getFullYear() > 1900 && excelDate.getFullYear() < 2100) {
              if (isTimeFormat && !isDateFormat) {
                const hours = excelDate.getUTCHours();
                const minutes = excelDate.getUTCMinutes();
                const seconds = excelDate.getUTCSeconds();
                return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
              }
              return excelDate.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
            }
          } catch (e) {}
        }
        return String(num);
      };
      if (rawValue.result !== undefined) {
        const res = rawValue.result;
        if (res instanceof Date) {
          const hours = res.getUTCHours();
          const minutes = res.getUTCMinutes();
          const seconds = res.getUTCSeconds();
          return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
        }
        if (typeof res === 'number') {
          return handleNumericLike(res);
        }
        if (typeof res === 'string') {
          return res.trim();
        }
        return "";
      }
      if (rawValue.formula) return `=${rawValue.formula}`;
      if (rawValue.value !== undefined) {
        const val = rawValue.value;
        if (val instanceof Date) {
          const hours = val.getUTCHours();
          const minutes = val.getUTCMinutes();
          const seconds = val.getUTCSeconds();
          if (hours !== 0 || minutes !== 0 || seconds !== 0) {
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
          }
          return val.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
        }
        if (typeof val === 'number') {
          return handleNumericLike(val);
        }
        if (typeof val === 'object' && val !== null) return "";
        return String(val).trim();
      }
      console.warn(`⚠️ Objeto no procesado en columna "${headerName}":`, rawValue);
      return "";
    }

    // Strings
    const stringValue = String(rawValue).trim();
    const datePatterns = [/(\d{1,2}\/\d{1,2}\/\d{4})/, /(\d{1,2}-\d{1,2}-\d{4})/, /(\d{4}-\d{1,2}-\d{1,2})/];
    for (const pattern of datePatterns) {
      const match = stringValue.match(pattern);
      if (match) return match[1];
    }
    const timePattern = /(\d{1,2}:\d{2}:\d{2})/;
    const timeMatch = stringValue.match(timePattern);
    if (timeMatch) return timeMatch[1];
    if (stringValue.length > 100) return stringValue.substring(0, 97) + '...';
    return stringValue;
  }; // ⭐ CIERRE DE getCellValue

  for (let rowIndex = headerRowIndex + 1; rowIndex <= maxRowsToCheck; rowIndex++) {
    const row = worksheet.getRow(rowIndex);
    const rowData = {};
    let hasData = false;

    sheetHeaders.forEach((header, colIndex) => {
      const cell = row.getCell(colIndex + 1);
      const cellValue = getCellValue(cell, colIndex + 1);
      rowData[header] = cellValue;
      const stringValue = String(cellValue || '');
      if (stringValue.trim() !== '') {
        hasData = true;
      }
    });

    if (hasData) {
      loadedData.push(rowData);
    }
  }

  console.log(`Total de registros cargados: ${loadedData.length}`);
  if (loadedData.length > 0) {
    console.log("Primera fila cargada:", loadedData[0]);
    console.log("Encabezados finales:", sheetHeaders);
  }

  // Post-procesado SOLO para MONITOREO: calcular campos derivados si vinieron vacíos o sin resultado
  if (isMonitoreo && loadedData.length > 0) {
    const toSeconds = (timeStr) => {
      if (!timeStr) return 0;
      if (typeof timeStr === 'number') {
        // Fracción de día a segundos
        if (timeStr >= 0 && timeStr < 1) return Math.round(timeStr * 24 * 60 * 60);
        return Math.round(timeStr * 60 * 60); // Interpretar como horas
      }
      if (timeStr instanceof Date) {
        return timeStr.getUTCHours() * 3600 + timeStr.getUTCMinutes() * 60 + timeStr.getUTCSeconds();
      }
      const s = String(timeStr).trim();
      // HH:MM:SS o H:MM:SS
      const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*[AP]M)?/i);
      if (m) {
        const h = parseInt(m[1] || '0', 10);
        const mi = parseInt(m[2] || '0', 10);
        const se = parseInt(m[3] || '0', 10);
        return h * 3600 + mi * 60 + se;
      }
      return 0;
    };
    const secondsToHMS = (totalSeconds) => {
      if (!totalSeconds || totalSeconds <= 0) return '00:00:00';
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    };
    const minutesFrom = (val) => Math.round(toSeconds(val) / 60);

    loadedData.forEach((row) => {
      // Calcular TIEMPO EFECTIVO DICTADO si está vacío
      const finClase = row['FINALIZA LA CLASE (ZOOM)'];
      const tiempoEspera = row['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE'];
      const tieneTiempoEfectivo = row['TIEMPO EFECTIVO DICTADO'] && String(row['TIEMPO EFECTIVO DICTADO']).trim() !== '';
      if (!tieneTiempoEfectivo && finClase && tiempoEspera) {
        const finSegs = toSeconds(finClase);
        const esperaSegs = toSeconds(tiempoEspera);
        const toleranciaSegs = 10 * 60; // 10 minutos
        const ajuste = Math.max(esperaSegs - toleranciaSegs, 0);
        const efectivo = Math.max(finSegs - ajuste, 0);
        if (efectivo > 0) row['TIEMPO EFECTIVO DICTADO'] = secondsToHMS(efectivo);
      }

      // Calcular EFICIENCIA si está vacía y hay datos suficientes
      const tieneEficiencia = row['EFICIENCIA'] && String(row['EFICIENCIA']).trim() !== '';
      const horasProgramadas = row['HORAS PROGRAMADAS'];
      const tiempoEfectivo = row['TIEMPO EFECTIVO DICTADO'];
      if (!tieneEficiencia && horasProgramadas && tiempoEfectivo) {
        const hpMin = minutesFrom(horasProgramadas);
        const teMin = minutesFrom(tiempoEfectivo);
        if (hpMin > 0) {
          const pct = Math.round((teMin / hpMin) * 100);
          row['EFICIENCIA'] = `${pct}%`;
        }
      }
    });
  }

  return { data: loadedData, headers: sheetHeaders };
} // ⭐ CIERRE DE loadSheetData

  const handleSheetChange = (sheetIndex) => {
    if (!workbookData) {
      alert("Por favor, carga primero un archivo Excel");
      return;
    }

    const cache = activeTab?.sheetData || {};
    const prevSheetData = activeTab?.sheetData || {};

    // Usar caché si existe para esta hoja
    if (cache[sheetIndex]) {
      const headers = cache[sheetIndex].headers || [];
      const dataForSheet = cache[sheetIndex].data || [];
      const updatedSheetData = { ...prevSheetData, [sheetIndex]: { data: dataForSheet, headers } };
      updateActiveTab({
        selectedSheet: sheetIndex,
        currentHeaders: headers,
        data: dataForSheet,
        sheetData: updatedSheetData,
        selectedDocente: ""
      });
      return;
    }

    // Si no hay caché, leer del workbook
    const worksheet = workbookData.worksheets[sheetIndex];
    const { data: loadedData, headers: sheetHeaders } = loadSheetData(worksheet);

    // Guardar en caché y actualizar estado en un solo paso
    const updatedSheetData = { ...prevSheetData, [sheetIndex]: { data: loadedData, headers: sheetHeaders } };
    updateActiveTab({
      selectedSheet: sheetIndex,
      currentHeaders: sheetHeaders,
      data: loadedData,
      sheetData: updatedSheetData,
      selectedDocente: ""
    });
  };

  const getUniqueCursosFromZoom = (zoomData, docente) => {
    const cursos = new Set();
    
    zoomData.forEach(zoomRow => {
      const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
      const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
      
      if (!matchDocente(docente, zoomDocente) || !zoomTema) return;
      
      const match = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)/i);
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
        
        // NUEVO: Verificar inicio 10 minutos antes
        thisRow['INICIO SESION 10 MINUTOS ANTES'] = verificarInicio10MinutesAntes(
          horaInicioExtraida, 
          thisRow['HORA INICIO']
        );
        
        totalAutoCompleted++;
      }

      allNewRows.push(thisRow);
      totalCreated++;
    }
  });

  setData([...data, ...allNewRows]);
  setNumFilas("");

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
    const exportHeaders = currentHeaders.length > 0 ? currentHeaders : [];  // ← CORREGIDO: [] en lugar de 'headers'

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
    const useHeaders = currentHeaders.length > 0 ? currentHeaders : [];  // ← CORREGIDO: [] en lugar de 'headers'
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
    // Mapear índice visible a índice real cuando MONITOREO reordena filas por DOCENTE
    let realIndex = index;
    if (isMonitoreoView) {
    const groups = new Map();
    data.forEach((r, idx) => {
    const key = (r.DOCENTE ?? '').toString();
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ r, idx });
    });
    const flat = [];
    groups.forEach(list => list.forEach(item => flat.push(item)));
    realIndex = flat[index]?.idx ?? index;
    }
    
    const newData = data.filter((_, i) => i !== realIndex);
    setData(newData);
    }
    };

  // Función para calcular TIEMPO EFECTIVO DICTADO
// Fórmula: =MAX(L3-MAX(K3-$N$1;0);0)
// Donde L3 = FINALIZA LA CLASE (ZOOM)
//       K3 = TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE
//       N1 = TOLERANCIA (00:10:00)
const calcularTiempoEfectivo = (row) => {
  const finClase = row['FINALIZA LA CLASE (ZOOM)'];
  const tiempoEspera = row['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE'];
  const tolerancia = '00:10:00'; // Valor de N1
  
  // Convertir tiempo en formato HH:MM:SS a segundos
  const timeToSeconds = (timeStr) => {
    if (!timeStr || typeof timeStr !== 'string') return 0;
    const parts = timeStr.split(':');
    if (parts.length !== 3) return 0;
    const hours = parseInt(parts[0]) || 0;
    const minutes = parseInt(parts[1]) || 0;
    const seconds = parseInt(parts[2]) || 0;
    return hours * 3600 + minutes * 60 + seconds;
  };
  
  // Convertir segundos a formato HH:MM:SS
  const secondsToTime = (totalSeconds) => {
    if (totalSeconds <= 0) return '00:00:00';
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
  };
  
  try {
    const finClaseSegs = timeToSeconds(finClase);
    const tiempoEsperaSegs = timeToSeconds(tiempoEspera);
    const toleranciaSegs = timeToSeconds(tolerancia);
    
    // Aplicar fórmula: MAX(finClase - MAX(tiempoEspera - tolerancia, 0), 0)
    const diferencia = tiempoEsperaSegs - toleranciaSegs;
    const ajuste = Math.max(diferencia, 0);
    const resultado = Math.max(finClaseSegs - ajuste, 0);
    
    return secondsToTime(resultado);
  } catch (error) {
    console.error('Error calculando tiempo efectivo:', error);
    return null;
  }
};

// Función para calcular EFICIENCIA
// Fórmula: =(MAX(L3-MAX(K3-$N$1;0);0))/M3
// Donde L3 = FINALIZA LA CLASE (ZOOM)
//       K3 = TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE
//       N1 = 10 minutos (tolerancia en minutos)
//       M3 = HORAS PROGRAMADAS
const calcularEficiencia = (row) => {
  // Obtener los valores directamente de las celdas
  const finClase = row['FINALIZA LA CLASE (ZOOM)'];
  const tiempoEspera = row['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE'];
  const horasProgramadas = row['HORAS PROGRAMADAS'];
  const toleranciaMinutos = 10; // N1 = 10 minutos

  // Si es un valor que ya tiene porcentaje, devolverlo tal cual
  if (row['EFICIENCIA']) {
    const existingValue = row['EFICIENCIA'];
    if (typeof existingValue === 'string' && existingValue.includes('%')) {
      return existingValue;
    }
    if (typeof existingValue === 'number') {
      return `${Math.round(existingValue)}%`;
    }
  }
  
  // Función mejorada para convertir tiempos a minutos
  const timeToMinutes = (timeStr) => {
    if (!timeStr) return 0;
    
    // Si es un objeto de Excel con valor
    if (typeof timeStr === 'object' && timeStr.value !== undefined) {
      timeStr = timeStr.value;
    }
    
    // Si es número
    if (typeof timeStr === 'number') {
      if (timeStr === 0) return 0;
      // Si es fracción de día (formato Excel)
      if (timeStr < 1) return timeStr * 24 * 60;
      // Si es número de horas
      return timeStr * 60;
    }
    
    // Si es string, procesar como HH:MM:SS
    if (typeof timeStr === 'string') {
      // Limpiar el string
      timeStr = timeStr.trim();
      
      // Intentar primero como HH:MM:SS
      const parts = timeStr.split(':');
      if (parts.length >= 2) {
        const hours = parseInt(parts[0]) || 0;
        const minutes = parseInt(parts[1]) || 0;
        const seconds = parts.length > 2 ? (parseInt(parts[2]) || 0) : 0;
        return (hours * 60) + minutes + (seconds / 60);
      }
      
      // Intentar como número
      const num = parseFloat(timeStr);
      if (!isNaN(num)) {
        if (num < 1) return num * 24 * 60; // Fracción de día
        return num * 60; // Horas
      }
    }
    
    return 0;
  };
  
  try {
    // Convertir todos los tiempos a minutos
    const finClaseMin = timeToMinutes(finClase);
    const tiempoEsperaMin = timeToMinutes(tiempoEspera);
    const horasProgramadasMin = timeToMinutes(horasProgramadas);
    
    // Si no hay horas programadas, mantener el valor original
    if (horasProgramadasMin === 0) {
      return row['EFICIENCIA'] || '0%';
    }
    
    // Aplicar la fórmula de Excel
    const diferencia = tiempoEsperaMin - toleranciaMinutos;
    const ajuste = Math.max(diferencia, 0);
    const tiempoEfectivoMin = Math.max(finClaseMin - ajuste, 0);
    
    // Calcular eficiencia como porcentaje
    const eficiencia = (tiempoEfectivoMin / horasProgramadasMin) * 100;
    
    // Si el resultado es válido, devolverlo formateado
    if (!isNaN(eficiencia) && eficiencia !== Infinity) {
      return `${Math.round(eficiencia)}%`;
    }
    
    // Si algo falló, mantener el valor original
    return row['EFICIENCIA'] || '0%';
  } catch (error) {
    console.error('Error calculando eficiencia:', error);
    // En caso de error, mantener el valor original
    return row['EFICIENCIA'] || '0%';
  }
};

  const handleCellChange = (rowIndex, columnName, value) => {
  if (selectedDocente) {
    const realIndex = data.findIndex((row) => row === filteredData[rowIndex]);
    const newData = [...data];
    newData[realIndex][columnName] = value;
    
    // Calcular TIEMPO EFECTIVO DICTADO automáticamente
    if (columnName === 'FINALIZA LA CLASE (ZOOM)' || 
        columnName === 'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE') {
      const row = newData[realIndex];
      const tiempoEfectivo = calcularTiempoEfectivo(row);
      if (tiempoEfectivo !== null) {
        newData[realIndex]['TIEMPO EFECTIVO DICTADO'] = tiempoEfectivo;
      }
    }
    
    // Calcular EFICIENCIA automáticamente
    if (columnName === 'FINALIZA LA CLASE (ZOOM)' || 
        columnName === 'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE' ||
        columnName === 'HORAS PROGRAMADAS') {
      const row = newData[realIndex];
      const eficiencia = calcularEficiencia(row);
      if (eficiencia !== null) {
        newData[realIndex]['EFICIENCIA'] = eficiencia;
      }
    }
    
    setData(newData);
  } else {
    // Mapear índice visible a índice real cuando MONITOREO reordena filas por DOCENTE
    let realIndex = rowIndex;
    if (isMonitoreoView) {
      const groups = new Map();
      data.forEach((r, idx) => {
        const key = (r.DOCENTE ?? '').toString();
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push({ r, idx });
      });
      const flat = [];
      groups.forEach(list => list.forEach(item => flat.push(item)));
      realIndex = flat[rowIndex]?.idx ?? rowIndex;
    }

    const newData = [...data];
    newData[realIndex][columnName] = value;
    
    // Calcular TIEMPO EFECTIVO DICTADO automáticamente
    if (columnName === 'FINALIZA LA CLASE (ZOOM)' || 
        columnName === 'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE') {
      const tiempoEfectivo = calcularTiempoEfectivo(newData[realIndex]);
      if (tiempoEfectivo !== null) {
        newData[realIndex]['TIEMPO EFECTIVO DICTADO'] = tiempoEfectivo;
      }
    }
    
    // Calcular EFICIENCIA automáticamente
    if (columnName === 'FINALIZA LA CLASE (ZOOM)' || 
        columnName === 'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE' ||
        columnName === 'HORAS PROGRAMADAS') {
      const eficiencia = calcularEficiencia(newData[realIndex]);
      if (eficiencia !== null) {
        newData[realIndex]['EFICIENCIA'] = eficiencia;
      }
    }
    
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

  // ===== FUNCIONES AUXILIARES PARA AUTOCOMPLETADO CON ZOOM =====
  
  // Función para calcular tiempo de espera antes de iniciar la clase
  const calcularTiempoEspera = (horaInicioZoom, horaProgramada) => {
    if (!horaInicioZoom || !horaProgramada) return "";
    
    // Convertir tiempo HH:MM:SS a minutos totales desde medianoche
    const timeToMinutes = (timeStr) => {
      if (!timeStr || typeof timeStr !== 'string') return 0;
      
      let timeOnly = timeStr;
      if (timeStr.includes('AM') || timeStr.includes('PM') || timeStr.includes('a. m.') || timeStr.includes('p. m.')) {
        const match = timeStr.match(/(\d{1,2}):(\d{2}):(\d{2})\s*([AP]M|[ap]\.\s*m\.)/i);
        if (match) {
          let hours = parseInt(match[1]);
          const minutes = parseInt(match[2]);
          const seconds = parseInt(match[3]);
          const period = match[4].toUpperCase().replace(/\s|\./g, '');
          
          if (period.includes('P') && hours !== 12) hours += 12;
          if (period.includes('A') && hours === 12) hours = 0;
          
          return hours * 60 + minutes + seconds / 60;
        }
      }
      
      const parts = timeOnly.split(':');
      if (parts.length >= 2) {
        const hours = parseInt(parts[0]) || 0;
        const minutes = parseInt(parts[1]) || 0;
        const seconds = parts.length >= 3 ? (parseInt(parts[2]) || 0) : 0;
        return hours * 60 + minutes + seconds / 60;
      }
      
      return 0;
    };
    
    const minutosInicioZoom = timeToMinutes(horaInicioZoom);
    const minutosProgramado = timeToMinutes(horaProgramada);
    
    // Calcular diferencia en minutos (programado - real)
    const diferencia = minutosProgramado - minutosInicioZoom;
    
    return Math.max(0, Math.round(diferencia)); // No permitir valores negativos
  };

  // Función para calcular tiempo efectivo dictado
  const calcularTiempoEfectivoDictado = (horaInicio, horaFin) => {
    if (!horaInicio || !horaFin) return "";
    
    const timeToMinutes = (timeStr) => {
      if (!timeStr || typeof timeStr !== 'string') return 0;
      
      let timeOnly = timeStr;
      if (timeStr.includes('AM') || timeStr.includes('PM') || timeStr.includes('a. m.') || timeStr.includes('p. m.')) {
        const match = timeStr.match(/(\d{1,2}):(\d{2}):(\d{2})\s*([AP]M|[ap]\.\s*m\.)/i);
        if (match) {
          let hours = parseInt(match[1]);
          const minutes = parseInt(match[2]);
          const seconds = parseInt(match[3]);
          const period = match[4].toUpperCase().replace(/\s|\./g, '');
          
          if (period.includes('P') && hours !== 12) hours += 12;
          if (period.includes('A') && hours === 12) hours = 0;
          
          return hours * 60 + minutes + seconds / 60;
        }
      }
      
      const parts = timeOnly.split(':');
      if (parts.length >= 2) {
        const hours = parseInt(parts[0]) || 0;
        const minutes = parseInt(parts[1]) || 0;
        const seconds = parts.length >= 3 ? (parseInt(parts[2]) || 0) : 0;
        return hours * 60 + minutes + seconds / 60;
      }
      
      return 0;
    };
    
    const minutosInicio = timeToMinutes(horaInicio);
    const minutosFin = timeToMinutes(horaFin);
    
    let duracionMinutos = minutosFin - minutosInicio;
    
    // Manejar el caso de sesiones que cruzan medianoche
    if (duracionMinutos < 0) {
      duracionMinutos += 24 * 60;
    }
    
    // Convertir a formato HH:MM:SS
    const horas = Math.floor(duracionMinutos / 60);
    const minutos = Math.floor(duracionMinutos % 60);
    const segundos = Math.round((duracionMinutos % 1) * 60);
    
    return `${horas.toString().padStart(2, '0')}:${minutos.toString().padStart(2, '0')}:${segundos.toString().padStart(2, '0')}`;
  };

  // Función para calcular eficiencia
  const calcularEficienciaZoom = (tiempoEfectivo, horasProgramadas) => {
    if (!tiempoEfectivo || !horasProgramadas) return "";
    
    // Convertir tiempo efectivo a horas decimales
    const timeToHours = (timeStr) => {
      if (!timeStr) return 0;
      const parts = timeStr.split(':');
      if (parts.length >= 2) {
        const hours = parseInt(parts[0]) || 0;
        const minutes = parseInt(parts[1]) || 0;
        const seconds = parts.length >= 3 ? (parseInt(parts[2]) || 0) : 0;
        return hours + (minutes / 60) + (seconds / 3600);
      }
      return 0;
    };
    
    const horasEfectivas = timeToHours(tiempoEfectivo);
    const horasProgram = parseFloat(horasProgramadas) || 0;
    
    if (horasProgram === 0) return "0%";
    
    const eficiencia = (horasEfectivas / horasProgram) * 100;
    return `${Math.round(eficiencia)}%`;
  };

  // Función para determinar observación según duración
  const determinarObservacion = (tiempoEfectivo) => {
    if (!tiempoEfectivo) return "";
    
    const timeToMinutes = (timeStr) => {
      if (!timeStr) return 0;
      const parts = timeStr.split(':');
      if (parts.length >= 2) {
        const hours = parseInt(parts[0]) || 0;
        const minutes = parseInt(parts[1]) || 0;
        return hours * 60 + minutes;
      }
      return 0;
    };
    
    const minutos = timeToMinutes(tiempoEfectivo);
    
    if (minutos >= 170) { // 2h50m o más
      return "✅";
    } else if (minutos >= 120) { // Entre 2h y 2h50m
      return "⚠️";
    } else { // Menos de 2h
      return "❌";
    }
  };

  // Función para verificar si inició 10 minutos antes
const verificarInicio10MinutesAntes = (horaInicioZoom, horaProgramada) => {
  if (!horaInicioZoom || !horaProgramada) return "NO";
  
  // Convertir tiempo HH:MM:SS a minutos totales desde medianoche
  const timeToMinutes = (timeStr) => {
    if (!timeStr || typeof timeStr !== 'string') return 0;
    
    // Extraer solo la parte de tiempo si viene con fecha
    let timeOnly = timeStr;
    if (timeStr.includes('AM') || timeStr.includes('PM') || timeStr.includes('a. m.') || timeStr.includes('p. m.')) {
      const match = timeStr.match(/(\d{1,2}):(\d{2}):(\d{2})\s*([AP]M|[ap]\.\s*m\.)/i);
      if (match) {
        let hours = parseInt(match[1]);
        const minutes = parseInt(match[2]);
        const seconds = parseInt(match[3]);
        const period = match[4].toUpperCase().replace(/\s|\./g, '');
        
        // Convertir a formato 24h
        if (period.includes('P') && hours !== 12) hours += 12;
        if (period.includes('A') && hours === 12) hours = 0;
        
        return hours * 60 + minutes + seconds / 60;
      }
    }
    
    // Formato HH:MM:SS estándar
    const parts = timeOnly.split(':');
    if (parts.length >= 2) {
      const hours = parseInt(parts[0]) || 0;
      const minutes = parseInt(parts[1]) || 0;
      const seconds = parts.length >= 3 ? (parseInt(parts[2]) || 0) : 0;
      return hours * 60 + minutes + seconds / 60;
    }
    
    return 0;
  };
  
  const minutosInicioZoom = timeToMinutes(horaInicioZoom);
  const minutosProgramado = timeToMinutes(horaProgramada);
  
  // Calcular diferencia (programado - zoom)
  const diferencia = minutosProgramado - minutosInicioZoom;
  
  // Si inició 10 minutos o más antes, retornar "SI"
  return diferencia >= 10 ? "SI" : "NO";
};


  // ===== DATOS COMPUTADOS =====
const filteredData = useMemo(() => {
  if (!selectedDocente) return data;
  return data.filter(row => row.DOCENTE === selectedDocente);
}, [data, selectedDocente]);

// Generar opciones dinámicas desde los datos
const uniqueCursos = useMemo(() => {
  const cursos = new Set();
  data.forEach(row => {
    if (row.CURSO && row.CURSO.trim() !== '') {
      cursos.add(row.CURSO.trim());
    }
  });
  return Array.from(cursos).sort();
}, [data]);

const uniqueSecciones = useMemo(() => {
  const secciones = new Set();
  data.forEach(row => {
    if (row.SECCION && row.SECCION.trim() !== '') {
      secciones.add(row.SECCION.trim());
    }
  });
  return Array.from(secciones).sort();
}, [data]);

const uniqueTurnos = useMemo(() => {
  const turnos = new Set();
  data.forEach(row => {
    if (row.TURNO && row.TURNO.trim() !== '') {
      turnos.add(row.TURNO.trim());
    }
  });
  return Array.from(turnos).sort();
}, [data]);

const uniqueDias = useMemo(() => {
  const dias = new Set();
  data.forEach(row => {
    if (row.DIAS && row.DIAS.trim() !== '') {
      dias.add(row.DIAS.trim());
    }
  });
  return Array.from(dias).sort();
}, [data]);

const uniquePeriodos = useMemo(() => {
  const periodos = new Set();
  data.forEach(row => {
    if (row.PERIODO && row.PERIODO.trim() !== '') {
      periodos.add(row.PERIODO.trim());
    }
  });
  return Array.from(periodos).sort();
}, [data]);

// Opciones dinámicas para DOCENTE SOLO en la hoja MONITOREO
const selectedSheetName = (availableSheets[selectedSheet]?.name || '').toString();
const isMonitoreoView = selectedSheetName.toLowerCase().includes('monitoreo');
const docentesFromData = useMemo(() => {
  const set = new Set();
  data.forEach(row => {
    const v = row && row.DOCENTE != null ? String(row.DOCENTE).trim() : '';
    if (v) set.add(v);
  });
  return Array.from(set).sort();
}, [data]);
const docenteOptions = isMonitoreoView
  ? Array.from(new Set([...docentesFromData, ...uniqueDocentes]))
  : uniqueDocentes;

const dropdownOptions = {
  MODELO: ["PROTECH XP", "TRADICIONAL"],
  MODALIDAD: ["PRESENCIAL", "VIRTUAL"],
  CURSO: uniqueCursos.length > 0 ? uniqueCursos : ["COMPUTACION 2", "COMPUTACION 3"],
  SECCION: uniqueSecciones.length > 0 ? uniqueSecciones : ["A", "PEAD-a", "PEAD-b"],
  TURNO: uniqueTurnos.length > 0 ? uniqueTurnos : ["MAÑANA", "TARDE", "NOCHE"],
  DIAS: uniqueDias.length > 0 ? uniqueDias : ["LUN", "MAR", "MIE", "JUE", "VIE", "SAB"],
  DOCENTE: docenteOptions,
  "INICIO SESION 10 MINUTOS ANTES": ["SI", "NO"],
  CICLO: ["SUPER INTENSIVO", "INTENSIVO", "REGULAR"],
  PERIODO: uniquePeriodos.length > 0 ? uniquePeriodos : ["2025 2: AGO", "2025 1: ENE", "2024 2: JUL"]
};

const displayData = useMemo(() => {
  // MONITOREO sin docente seleccionado: agrupar por DOCENTE respetando el orden original del Excel
  if (!selectedDocente && isMonitoreoView) {
    const groups = new Map();
    data.forEach((row) => {
      const docenteKey = (row.DOCENTE ?? '').toString();
      if (!groups.has(docenteKey)) groups.set(docenteKey, []);
      groups.get(docenteKey).push(row);
    });
    const ordered = [];
    groups.forEach((rows) => {
      ordered.push(...rows);
    });
    return ordered;
  }

  if (!selectedDocente) return data;

  // Agrupar por CURSO/SECCIÓN, ordenar dentro del grupo y renumerar 1..16.
  const normalizeSeccionLocal = (v) => (v ? String(v).toUpperCase().trim().replace(/^PEAD[-_ ]?/i, '').replace(/[^A-Z0-9]/g, '') : '');
  const parseIntSafe = (v) => {
    const n = parseInt(String(v));
    return Number.isNaN(n) ? null : n;
  };

  const grouped = new Map();
  filteredData.forEach((row, idx) => {
    const key = `${normalizeCursoName(row.CURSO || '')}|||${normalizeSeccionLocal(row.SECCION || row['SECCIÓN'] || '')}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push({ row, idx, originalIndex: idx });
  });

  const out = [];

  const keys = Array.from(grouped.keys()).sort(); // ordenar grupos por curso/sección normalizados

  keys.forEach((key) => {
    const entries = grouped.get(key).slice();

    // Orden dentro del grupo: por SESION numérica si existe, si no por índice original
    entries.sort((a, b) => {
      const aS = parseIntSafe(a.row.SESION);
      const bS = parseIntSafe(b.row.SESION);
      if (aS !== null && bS !== null && aS !== bS) return aS - bS;
      return a.originalIndex - b.originalIndex;
    });

    // Renumerar 1..16 y limitar a 16 visibles por grupo
    entries.slice(0, 16).forEach((entry, i) => {
      out.push({ ...entry.row, SESION: i + 1 });
    });
  });

  return out;
}, [filteredData, selectedDocente, data, isMonitoreoView]);

  // ===== RENDER =====
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
              onAutocompletarConZoom={handleAutocompletarConZoom}
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
            <h2 className="text-2xl font-bold text-gray-700 mb-6">
              No hay archivos abiertos
            </h2>
            <p className="text-gray-500 mb-8">
              Haz clic en "+ Nueva Pestaña" para cargar un archivo Excel
            </p>
            
            {/* Botones de descarga de plantillas - Pantalla principal */}
            <div className="flex flex-col items-center gap-4 mt-8">
              <p className="text-gray-600 font-semibold text-lg mb-2">
                📥 Descargar plantillas de ejemplo:
              </p>
              <div className="flex gap-4">
                <button
                  onClick={downloadExcelTemplate}
                  className="px-6 py-3 bg-green-600 text-white hover:bg-green-700 rounded-lg font-semibold text-sm shadow-md transition-all hover:shadow-lg"
                >
                  📄 Descargar Plantilla Excel
                </button>
                <button
                  onClick={downloadCSVTemplate}
                  className="px-6 py-3 bg-blue-600 text-white hover:bg-blue-700 rounded-lg font-semibold text-sm shadow-md transition-all hover:shadow-lg"
                >
                  📊 Descargar Plantilla CSV (Zoom)
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;

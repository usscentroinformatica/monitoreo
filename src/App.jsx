import React, { useState, useMemo } from "react";
import ExcelJS from "exceljs";
import ControlPanel from "./components/ControlPanel";
import DataTable from "./components/DataTable";
import { BackupHistoryModal, useBackupManager } from './components/BackupManager';
import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';
import { getStorage } from 'firebase/storage';
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
  const isLoading = activeTab?.isLoading || false;
  const availableSheets = activeTab?.availableSheets || [];
  const selectedSheet = activeTab?.selectedSheet || 0;
  const workbookData = activeTab?.workbookData || null;
  const currentHeaders = activeTab?.currentHeaders || [];
  const [randomDocente, setRandomDocente] = useState(null);

  const mostrarToast = (mensaje, tipo = 'info') => {
  // Crear elemento toast
  const toast = document.createElement('div');
  toast.className = `toast toast-${tipo}`;
  
  // Posición superior derecha y estilos mejorados
  toast.style.position = 'fixed';
  toast.style.top = '20px';
  toast.style.right = '20px';
  toast.style.backgroundColor = tipo === 'error' ? '#f44336' : tipo === 'warning' ? '#ff9800' : '#4CAF50';
  toast.style.color = 'white';
  toast.style.padding = '12px 15px';
  toast.style.borderRadius = '5px';
  toast.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
  toast.style.zIndex = '10000';
  toast.style.maxWidth = '350px';
  toast.style.fontSize = '14px';
  toast.style.fontWeight = 'bold';
  
  // Agregar al body
  document.body.appendChild(toast);
  
  // Contenido HTML para mejor formateo
  toast.innerHTML = mensaje;
  
  // Remover después de 10 segundos
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.5s';
    
    setTimeout(() => {
      if (document.body.contains(toast)) {
        document.body.removeChild(toast);
      }
    }, 500);
  }, 10000);  // 10 segundos
  
  // Sistema de apilamiento para múltiples notificaciones
  const toasts = document.querySelectorAll('.toast');
  if (toasts.length > 1) {
    const offset = Array.from(toasts).slice(0, -1).reduce((total, t) => {
      return total + t.offsetHeight + 10;
    }, 0);
    toast.style.top = `${20 + offset}px`;
  }
};

  // Usa el hook personalizado para manejar backups
const {
  backupHistory,
  isBackupModalOpen,
  setIsBackupModalOpen,
  saveBackup: saveBackupToStorage,
  downloadBackup,
  deleteBackup
} = useBackupManager(ExcelJS, mostrarToast);

// Función wrapper para saveBackup
const saveBackup = () => {
  saveBackupToStorage(data, currentHeaders, activeTab, setIsLoading);
};


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


const selectRandomDocente = () => {
  // Obtener lista de docentes únicos
  const uniqueDocentes = [...new Set(data
    .map(row => row.DOCENTE)
    .filter(docente => docente && docente.trim() !== '')
  )];
  
  if (uniqueDocentes.length === 0) {
    mostrarToast('❌ No hay docentes para seleccionar', 'error');
    return;
  }
  
  // Seleccionar un docente aleatorio
  const randomIndex = Math.floor(Math.random() * uniqueDocentes.length);
  const selectedDocente = uniqueDocentes[randomIndex];
  
  // Establecer el docente aleatorio
  setRandomDocente(selectedDocente);
  
  mostrarToast(`🎲 Docente seleccionado: <br><b>${selectedDocente}</b>`, 'success');
};

  // Función para crear nueva pestaña
  const createNewTab = (fileName, initialData = {}) => {
    const newTab = {
      id: nextTabId,
      name: fileName || `Archivo ${nextTabId}`,
      data: initialData.data || [],
      zoomData: [],
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
  const setIsLoading = (loading) => updateActiveTab({ isLoading: loading });
  const setSelectedSheet = (sheet) => updateActiveTab({ selectedSheet: sheet });
  const setWorkbookData = (wb) => updateActiveTab({ workbookData: wb });
  const setCurrentHeaders = (headers) => {
    const currentIndex = activeTab?.selectedSheet ?? 0;
    const prevSheetData = activeTab?.sheetData || {};
    const prevData = prevSheetData[currentIndex]?.data || data;
    const updatedSheetData = { ...prevSheetData, [currentIndex]: { data: prevData, headers } };
    updateActiveTab({ currentHeaders: headers, sheetData: updatedSheetData });
  };
  // Función para setear TIEMPO EFECTIVO DICTADO de forma condicional
  // Función para setear TIEMPO EFECTIVO DICTADO y EFICIENCIA de forma condicional
const setEffectiveTimeConditionally = (row) => {
  // Calcular TIEMPO EFECTIVO DICTADO si está vacío
  if (!row['TIEMPO EFECTIVO DICTADO'] || row['TIEMPO EFECTIVO DICTADO'] === null || row['TIEMPO EFECTIVO DICTADO'] === '') {
    row['TIEMPO EFECTIVO DICTADO'] = calculateEffectiveTime(row);
  }
  
  // 🆕 Calcular EFICIENCIA siempre que se tenga TIEMPO EFECTIVO DICTADO
  if (row['TIEMPO EFECTIVO DICTADO']) {
    row['EFICIENCIA'] = calculateEfficiency(row);
  }
};
  // ===== FUNCIONES DE UTILIDAD =====
  const normalizeDocenteName = (name) => {
    if (!name) return "";
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
      const regex = new RegExp(`\\b${roman}\\b`, 'gi');
      result = result.replace(regex, romanToArabic[roman]);
    });
   
    return result;
  };
  const normalizeSeccion = (value) => {
  if (!value) return "";
  
  let normalized = String(value)
    .toUpperCase()
    .trim()
    .replace(/^PEAD[-_ ]?/, ""); // Elimina "PEAD-" al inicio
  
  // Elimina caracteres que no sean alfanuméricos
  normalized = normalized.replace(/[^A-Z0-9]/g, "");
  
  return normalized;
};

const matchSecciones = (seccionExcel, seccionZoom) => {
  const normalizedExcel = normalizeSeccion(seccionExcel);
  const normalizedZoom = normalizeSeccion(seccionZoom);
  
  // Coincidencia exacta
  if (normalizedExcel === normalizedZoom) return true;
  
  // Comprobar si uno está contenido en el otro (para casos como "A" vs "AA")
  if (normalizedExcel.includes(normalizedZoom) || normalizedZoom.includes(normalizedExcel)) {
    // Mostrar mensaje más claro sobre la discrepancia
    mostrarToast(`⚠️ Diferencia en secciones:<br>
      <b>Docente:</b> ${seccionExcel.includes('PEAD') ? seccionExcel.split('-')[0] : 'DOCENTE'}<br>
      <b>Secciones en Excel:</b> ${seccionExcel}<br>
      <b>Sección en Zoom:</b> ${seccionZoom}`, 'warning');
    console.log(`⚠️ Discrepancia detectada: Excel "${seccionExcel}" vs Zoom "${seccionZoom}"`);
    return true;
  }
  
  return false;
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
    if (!dateTimeStr) return null;

    const s = String(dateTimeStr).trim();
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
    const m2 = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m2) {
      const y = m2[1], m = String(m2[2]).padStart(2,'0'), d = String(m2[3]).padStart(2,'0');
      return `${d}/${m}/${y}`;
    }
    const m3 = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m3) {
      let a = parseInt(m3[1],10), b = parseInt(m3[2],10); let y = m3[3];
      if (a > 12 && b <= 12) {
        const dd = String(a).padStart(2,'0'); const mm = String(b).padStart(2,'0');
        y = y.length === 2 ? `20${y}`: y; return `${dd}/${mm}/${y}`;
      }
      if (b > 12 && a <= 12) {
        const dd = String(b).padStart(2,'0'); const mm = String(a).padStart(2,'0');
        y = y.length === 2 ? `20${y}`: y; return `${dd}/${mm}/${y}`;
      }
      const dd = String(a).padStart(2,'0'); const mm = String(b).padStart(2,'0');
      y = y.length === 2 ? `20${y}`: y; return `${dd}/${mm}/${y}`;
    }
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      const dd = String(d.getDate()).padStart(2,'0');
      const mm = String(d.getMonth()+1).padStart(2,'0');
      const yy = String(d.getFullYear());
      return `${dd}/${mm}/${yy}`;
    }
    return null;
  };
  const extractTime = (dateTimeStr) => {
    if (!dateTimeStr) return null;
    let s = String(dateTimeStr).trim();

    // Match HH:MM:SS [AP]M
    let match = s.match(/(\d{1,2}):(\d{2}):(\d{2})\s*([AP]M)/i);
    if (match) {
      return `${match[1].padStart(2, '0')}:${match[2]}:${match[3]} ${match[4].toUpperCase()}`;
    }

    // Match HH:MM [AP]M
    match = s.match(/(\d{1,2}):(\d{2})\s*([AP]M)/i);
    if (match) {
      return `${match[1].padStart(2, '0')}:${match[2]}:00 ${match[3].toUpperCase()}`;
    }

    // Match HH:MM:SS a.m./p.m.
    match = s.match(/(\d{1,2}):(\d{2}):(\d{2})\s*([ap])\.\s*m\./i);
    if (match) {
      return `${match[1].padStart(2, '0')}:${match[2]}:${match[3]} ${match[4].toUpperCase()}M`;
    }

    // Match HH:MM a.m./p.m.
    match = s.match(/(\d{1,2}):(\d{2})\s*([ap])\.\s*m\./i);
    if (match) {
      return `${match[1].padStart(2, '0')}:${match[2]}:00 ${match[4].toUpperCase()}M`;
    }

    // Match HH:MM:SS (24h)
    match = s.match(/(\d{1,2}):(\d{2}):(\d{2})/);
    if (match) {
      return `${match[1].padStart(2, '0')}:${match[2]}:${match[3]}`;
    }

    // Match HH:MM (24h)
    match = s.match(/(\d{1,2}):(\d{2})/);
    if (match) {
      return `${match[1].padStart(2, '0')}:${match[2]}:00`;
    }

    return null;
  };
  const extractDuration = (zoomRow) => {
    if (!zoomRow) return "";
    let durStr = zoomRow['Duración'] || zoomRow['Duration'] || zoomRow['Duration (Minutes)'] || "";
    if (!durStr) return "";
    let minutes = parseInt(durStr);
    if (isNaN(minutes)) {
      // Si ya es un string de tiempo como "01:00:00", retornar tal cual
      return String(durStr).trim();
    }
    // Convertir minutos a HH:MM:00
    let hours = Math.floor(minutes / 60);
    let mins = minutes % 60;
    return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}:00`;
  };
  const detectTurno = (horaStr) => {
    if (!horaStr) return null;

    
    let hour = 0;
   
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
        const match24h = horaStr.match(/(\d{1,2}):/);
        if (match24h) {
          hour = parseInt(match24h[1]);
        }
      }
    }
   
    if (hour >= 6 && hour < 12) {
      return "MAÑANA";
    } else if (hour >= 12 && hour < 18) {
      return "TARDE";
    } else if (hour >= 18 && hour <= 23) {
      return "NOCHE";
    } else {
      return "NOCHE";
    }
  };
  const extractCursoFromTema = (tema) => {
    if (!tema) return null;
    const match = tema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
    return match ? match[1].trim() : tema;
  };

  // NUEVA FUNCIÓN: Detectar MODELO basado en CURSO
  const detectModeloFromCurso = (cursoStr) => {
    if (!cursoStr) return null;
    const upperCurso = cursoStr.toUpperCase().trim();
    if (upperCurso.includes('COMPUTACION') && 
        (upperCurso.includes('I') || upperCurso.includes('II') || upperCurso.includes('III'))) {
      return "TRADICIONAL";
    }
    return null; // Si no coincide, no setear nada (dejar que el usuario lo defina manualmente)
  };

  // NUEVA FUNCIÓN: Convertir tiempo a minutos
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

  // NUEVA FUNCIÓN: Calcular TIEMPO EFECTIVO DICTADO
  const calculateEffectiveTime = (row) => {
    const endCol = row['FINALIZA LA CLASE (ZOOM)'] || row['Finaliza la Clase (Zoom)'] || row['Hora Finalización Zoom'] || null;
    const waitTimeCol = row['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE'] || null;
    const tolerancia = '00:10:00'; // Tolerancia fija (N1)

    if (!endCol) return null;

    const endMin = timeToMinutes(endCol);
    const waitMin = waitTimeCol ? timeToMinutes(waitTimeCol) : 0;
    const toleranciaMin = timeToMinutes(tolerancia);

    const waitAdjusted = Math.max(waitMin - toleranciaMin, 0);
    let diffMin = endMin - waitAdjusted;
    if (diffMin < 0) diffMin = 0;

    const hours = Math.floor(diffMin / 60);
    const mins = Math.floor(diffMin % 60);
    const secs = 0; // Asumiendo sin segundos

    return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
  };

  // NUEVA FUNCIÓN: Calcular EFICIENCIA
const calculateEfficiency = (row) => {
  const tiempoEfectivo = row['TIEMPO EFECTIVO DICTADO'] || calculateEffectiveTime(row);
  const horasProgramadas = row['HORAS PROGRAMADAS'] || row['Horas Programadas'] || row['horas programadas'] || '03:00:00';

  if (!tiempoEfectivo || !horasProgramadas) return null;

  const efectivoMin = timeToMinutes(tiempoEfectivo);
  const programadasMin = timeToMinutes(horasProgramadas);

  if (programadasMin === 0) return null;

  const eficiencia = (efectivoMin / programadasMin) * 100;
  
  // Redondear a 0 decimales y agregar símbolo %
  return `${Math.round(eficiencia)}%`;
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
        row[header] = values[index] || null;
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

    const docentesToProcess = [...new Set(data.map(row => row.DOCENTE).filter(d => d && d.trim() !== ''))];

    if (docentesToProcess.length === 0) {
      alert("No hay docentes registrados en el Excel para autocompletar");
      return;
    }

    console.log(`📋 Modo: TODOS los docentes`);
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
      
      updatedRow.CURSO = zoomInfo.curso;
      updatedRow.TURNO = zoomInfo.turno;
      
      // Detectar y setear MODELO si aplica
      const modeloDetectado = detectModeloFromCurso(zoomInfo.curso);
      if (modeloDetectado && (!updatedRow.MODELO || updatedRow.MODELO === null || updatedRow.MODELO.toString().trim() === '')) {
        updatedRow.MODELO = modeloDetectado;
      }

      // Calcular TIEMPO EFECTIVO DICTADO de forma condicional
      setEffectiveTimeConditionally(updatedRow);
      
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
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col] === null || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractDate(fechaInicio);
                break;
              }
            }
            
            const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
            for (const col of possibleStartCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col] === null || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractTime(fechaInicio);
                break;
              }
            }
            
            const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
            for (const col of possibleEndCols) {
              if (currentHeaders.includes(col) && (!updatedRow[col] || updatedRow[col] === null || updatedRow[col].toString().trim() === '')) {
                updatedRow[col] = extractTime(fechaFin);
                break;
              }
            }
            
            if (!updatedRow.TURNO || updatedRow.TURNO === null || updatedRow.TURNO.toString().trim() === '') {
              updatedRow.TURNO = detectTurno(fechaInicio);
            }

            // Detectar y setear MODELO si aplica
            const modeloDetectado = detectModeloFromCurso(cursoZoom);
            if (modeloDetectado && (!updatedRow.MODELO || updatedRow.MODELO === null || updatedRow.MODELO.toString().trim() === '')) {
              updatedRow.MODELO = modeloDetectado;
            }

            // Calcular TIEMPO EFECTIVO DICTADO de forma condicional
            setEffectiveTimeConditionally(updatedRow);
            
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

        const hasEmptySession = !row.CURSO || row.CURSO === null || row.CURSO.toString().trim() === '' ||
                               !row.SECCION || row.SECCION === null || row.SECCION.toString().trim() === '' ||
                               !row.SESION || row.SESION === null || row.SESION.toString().trim() === '';

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
            turno: detectTurno(fechaInicio)
          });
          
          newData[index].CURSO = cursoZoom;
          newData[index].SECCION = seccionZoom;
          newData[index].SESION = sesionZoom;
          
          sesionesUsadasGlobal.add(claveZoom);
          updatedCount++;
          
          console.log(`✓ Fila vacía ${index} COMPLETADA con: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
          break;
        }
      });

      // Fallback adicional: emparejar por horario cuando el Tema no contiene PEAD-
      {
        const usedZoomByStart = new Set();

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
              turno: detectTurno(fechaInicio)
            });

            newData[index] = updatedRow;

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

        const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
        const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";

        const newRow = {};
        currentHeaders.forEach(header => {
          newRow[header] = null;
        });

        newRow.DOCENTE = docenteActual;
        newRow.CURSO = cursoZoom;
        newRow.SECCION = seccionZoom;
        newRow.SESION = sesionZoom;
        newRow.TURNO = detectTurno(fechaInicio);

        // Detectar y setear MODELO si aplica
        const modeloDetectado = detectModeloFromCurso(cursoZoom);
        if (modeloDetectado) {
          newRow.MODELO = modeloDetectado;
        }

        // 🆕 HORAS PROGRAMADAS siempre 3 horas
        const possibleHorasProgramadasCols = ['HORAS PROGRAMADAS', 'Horas Programadas', 'horas programadas'];
        for (const col of possibleHorasProgramadasCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = '03:00:00';
            break;
          }
        }

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

        // Calcular TIEMPO EFECTIVO DICTADO para nueva fila (siempre, ya que es nueva)
        newRow['TIEMPO EFECTIVO DICTADO'] = calculateEffectiveTime(newRow);

        newData.push(newRow);
        sesionesUsadasGlobal.add(claveZoom);
        createdCount++;
        
        console.log(`✓ Nueva fila realmente necesaria: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
      });
    });

    // Asegurar que todas las filas tengan el cálculo de forma condicional
    newData.forEach(row => {
      setEffectiveTimeConditionally(row);
    });

    // Si no existe la columna en headers, agregarla
if (!currentHeaders.includes('TIEMPO EFECTIVO DICTADO')) {
  const newHeaders = [...currentHeaders, 'TIEMPO EFECTIVO DICTADO'];
  if (!newHeaders.includes('EFICIENCIA')) {
    newHeaders.push('EFICIENCIA');
  }
  setCurrentHeaders(newHeaders);
} else if (!currentHeaders.includes('EFICIENCIA')) {
  setCurrentHeaders([...currentHeaders, 'EFICIENCIA']);
}

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
    alert("⚠️ Primero carga el archivo Excel");
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
          
          const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)'];
          for (const col of possibleFinalizaCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = horaFinExtraida;
              break;
            }
          }

          // NUEVO: Detectar y setear MODELO si aplica
          const modeloDetectado = detectModeloFromCurso(curso);
          if (modeloDetectado && (!dataProcesada[index].MODELO || dataProcesada[index].MODELO === null || dataProcesada[index].MODELO.toString().trim() === '')) {
            dataProcesada[index].MODELO = modeloDetectado;
          }

          // NUEVO: Calcular TIEMPO EFECTIVO DICTADO de forma condicional
          setEffectiveTimeConditionally(dataProcesada[index]);

          // 🆕 Completar HORAS PROGRAMADAS si está vacío
          const possibleHorasProgramadasCols = ['HORAS PROGRAMADAS', 'Horas Programadas', 'horas programadas'];
          for (const col of possibleHorasProgramadasCols) {
            if (currentHeaders.includes(col) && (!dataProcesada[index][col] || dataProcesada[index][col] === null || String(dataProcesada[index][col]).trim() === '')) {
              dataProcesada[index][col] = '03:00:00';
              break;
            }
          }
          
          console.log(`  ✓ Autocompletado: ${docente} - ${curso} - Sesión ${sesion}`);
        }
      });

      // Fallback por horario
      const usedZoomByStart = new Set();
      const timeToMinutes = (timeStr) => {
        if (!timeStr || typeof timeStr !== 'string') return 0;
        let s = timeStr.trim();
        s = s.replace(/a\.\s*m\.|p\.\s*m\./gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
        const m12 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i);
        if (m12) {
          let h = parseInt(m12[1]) || 0;
          const min = parseInt(m12[2]) || 0;
          const sec = parseInt(m12[3]||'0')||0;
          const p = m12[4].toUpperCase();
          if (p === 'PM' && h !== 12) h += 12;
          if (p === 'AM' && h === 12) h = 0;
          return h * 60 + min + sec / 60;
        }
        const m24 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        if (m24) {
          const h = parseInt(m24[1]) || 0;
          const min = parseInt(m24[2]) || 0;
          const sec = parseInt(m24[3]||'0')||0;
          return h * 60 + min + sec / 60;
        }
        return 0;
      };

      dataProcesada.forEach((row, index) => {
        const docente = row.DOCENTE || '';
        const horaProg = row['HORA INICIO'] || row['Hora Inicio'] || row['INICIO'] || row['inicio'] || '';
        if (!docente || !horaProg) return;

        const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
        const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
        const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
        const hasFecha = possibleDateCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
        const hasHoraInicio = possibleStartCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
        const hasHoraFin = possibleEndCols.some(col => currentHeaders.includes(col) && row[col] && String(row[col]).trim() !== '');
        if (hasFecha && hasHoraInicio && hasHoraFin) return;

        const tProg = timeToMinutes(String(horaProg));
        if (tProg === 0) return;

        let bestZoom = null;
        let bestStartStr = null;
        let bestEndStr = null;
        let bestDiff = Infinity;
        
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
            bestDiff = diff;
            bestZoom = zoomRow;
            bestStartStr = startStr;
            bestEndStr = endStr;
          }
          return;
        });

        const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
        const cursoZoom = cursoParte.trim();
        const sesionZoom = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;
        const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;
        
        if (sesionesUsadasGlobal.has(claveZoom)) return;
        
        // Buscar usando matchSecciones en lugar de igualdad exacta
        const existingRow = newData.find(row =>
          row.DOCENTE === docenteActual &&
          normalizeCursoName(row.CURSO || "") === normalizeCursoName(cursoZoom) &&
          matchSecciones(row.SECCION || "", seccionZoom) &&
          parseInt(String(row.SESION || 0)) === sesionZoom
        );
        
        if (existingRow) {
          console.log(`⚠️ Ya existe fila para ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}. NO se crea duplicado.`);
          mostrarToast(`✅ Coincidencia encontrada:<br>
            <b>Docente:</b> ${docenteActual}<br>
            <b>Excel:</b> ${existingRow.SECCION}<br>
            <b>Zoom:</b> ${seccionZoom}`, 'info');
          sesionesUsadasGlobal.add(claveZoom);
          return;
        }
        
        // Si se encontró el docente pero no coincide la sección, mostrar notificación
        const filasMismoDocente = newData.filter(row => 
          row.DOCENTE === docenteActual && 
          normalizeCursoName(row.CURSO || "") === normalizeCursoName(cursoZoom)
        );
        
        // En la parte de crear nuevas filas, cuando detectas un docente con sección diferente
if (filasMismoDocente.length > 0) {
  const seccionesExistentes = [...new Set(
    filasMismoDocente
      .map(row => row.SECCION)
      .filter(Boolean)
  )].join(", ");
  
  mostrarToast(`⚠️ Diferencia en secciones:<br>
    <b>Docente:</b> ${docenteActual}<br>
    <b>Secciones en Excel:</b> ${seccionesExistentes}<br>
    <b>Sección en Zoom:</b> ${seccionZoom}`, 'warning');
}
        
        const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
        const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
        const newRow = {};
        currentHeaders.forEach(header => {
          newRow[header] = "";
        });

        if (bestZoom) {
          const fechaInicio = bestStartStr || '';
          const fechaFin = bestEndStr || '';

          const setIfHasHeader = (obj, colNames, val) => {
            for (const c of colNames) {
              if (currentHeaders.includes(c)) {
                obj[c] = val;
                break;
              }
            }
          };

          setIfHasHeader(row, possibleDateCols, extractDate(fechaInicio));
          setIfHasHeader(row, possibleStartCols, extractTime(fechaInicio));
          setIfHasHeader(row, possibleEndCols, extractTime(fechaFin));

          row.TURNO = row.TURNO && String(row.TURNO).trim() !== '' ? row.TURNO : detectTurno(fechaInicio);

          // NUEVO: Detectar y setear MODELO si aplica en fallback
          const cursoFallback = row.CURSO || extractCursoFromTema(bestZoom['Tema'] || bestZoom['Topic'] || "");
          const modeloDetectado = detectModeloFromCurso(cursoFallback);
          if (modeloDetectado && (!row.MODELO || row.MODELO === null || row.MODELO.toString().trim() === '')) {
            row.MODELO = modeloDetectado;
          }

          // NUEVO: Calcular TIEMPO EFECTIVO DICTADO de forma condicional en fallback
          setEffectiveTimeConditionally(row);

          usedZoomByStart.add(bestStartStr);
        }
      });
    }

    // PASO 2: Detectar grupos únicos por DOCENTE+CURSO+SECCIÓN
    console.log("\n📋 PASO 2: Detectando grupos únicos por DOCENTE+CURSO+SECCIÓN");
    
    const gruposPorSeccion = new Map();
    
    dataProcesada.forEach((row, originalIndex) => {
      const docente = row.DOCENTE || '';
      const curso = row.CURSO || '';
      const seccion = normalizeSeccion(row.SECCION || row['SECCIÓN'] || '');
      
      if (!docente || !curso || !seccion) return;
      
      const key = `${docente}|||${normalizeCursoName(curso)}|||${seccion}`;
      
      if (!gruposPorSeccion.has(key)) {
        gruposPorSeccion.set(key, {
          docente,
          curso,
          seccion: row.SECCION || row['SECCIÓN'],
          primeraFila: row,
          indices: [],
          filas: [],
          sesionesExistentes: new Set()
        });
      }
      
      const grupo = gruposPorSeccion.get(key);
      grupo.indices.push(originalIndex);
      grupo.filas.push(row);
      
      if (row.SESION) {
        grupo.sesionesExistentes.add(parseInt(row.SESION));
      }
    });

    console.log(`Total grupos detectados: ${gruposPorSeccion.size}`);

    // PASO 3: Crear exactamente 16 sesiones por cada grupo
    console.log("\n📋 PASO 3: Creando exactamente 16 sesiones por cada grupo");

    const resultadoFinal = [];
    const gruposOrdenados = Array.from(gruposPorSeccion.entries())
      .sort((a, b) => Math.min(...a[1].indices) - Math.min(...b[1].indices));

    gruposOrdenados.forEach(([key, grupo]) => {
      const { docente, curso, seccion, primeraFila, filas, sesionesExistentes } = grupo;
      
      console.log(`\n--- ${docente} - ${curso} - ${seccion} ---`);
      console.log(`   Sesiones existentes: ${Array.from(sesionesExistentes).sort((a,b) => a-b).join(', ')}`);
      console.log(`   Total filas existentes: ${filas.length}`);
      
      const sesionesCompletas = [];
      
      // Crear Map con las filas existentes del rango 1-16
      const existingInRange = new Map();
      filas.forEach(f => {
        const s = parseInt(String(f.SESION || 0));
        if (s >= 1 && s <= 16 && !existingInRange.has(s)) {
          existingInRange.set(s, f);
        }
      });
      
      // Si hay filas existentes pero ninguna tiene SESION en rango 1-16, 
      // asignar la primera fila como SESION 1
      if (existingInRange.size === 0 && filas.length > 0) {
        const primeraFilaConDatos = filas[0];
        primeraFilaConDatos.SESION = 1;
        existingInRange.set(1, primeraFilaConDatos);
        console.log(`   📌 Primera fila asignada como SESION 1`);
      }
      
      // Crear exactamente 16 sesiones (1-16)
      for (let sesion = 1; sesion <= 16; sesion++) {
        if (existingInRange.has(sesion)) {
          // Usar la fila ORIGINAL completa SIN MODIFICAR
          const filaExistente = existingInRange.get(sesion);
          
          // Asegurarse de que SESION sea el número correcto
          filaExistente.SESION = sesion;
          
          // NUEVO: Recalcular TIEMPO EFECTIVO DICTADO de forma condicional para filas existentes
          setEffectiveTimeConditionally(filaExistente);

          // 🆕 Completar HORAS PROGRAMADAS si está vacío
          const possibleHorasProgramadasCols = ['HORAS PROGRAMADAS', 'Horas Programadas', 'horas programadas'];
          for (const col of possibleHorasProgramadasCols) {
            if (currentHeaders.includes(col) && (!filaExistente[col] || filaExistente[col] === null || String(filaExistente[col]).trim() === '')) {
              filaExistente[col] = '03:00:00';
              break;
            }
          }
          
          sesionesCompletas.push(filaExistente);
          console.log(`  ○ Sesión ${sesion}: YA EXISTE (mantenida con todos sus datos)`);
        } else {
          // Crear nueva fila SOLO con METADATOS básicos copiados de la primera fila
          const nuevaFila = {};
          currentHeaders.forEach(header => {
            nuevaFila[header] = null;
          });
          
          // ===== METADATOS que SÍ se copian de la primera fila =====
          nuevaFila.DOCENTE = primeraFila.DOCENTE || null;
          nuevaFila.CURSO = primeraFila.CURSO || null;
          nuevaFila.SECCION = primeraFila.SECCION || null;
          nuevaFila.MODELO = primeraFila.MODELO || null;
          nuevaFila.MODALIDAD = primeraFila.MODALIDAD || null;
          nuevaFila.CICLO = primeraFila.CICLO || null;
          nuevaFila.PERIODO = primeraFila.PERIODO || null;
          
          // Aula USS copiada TAL CUAL de la primera fila
          nuevaFila['Aula USS'] = primeraFila['Aula USS'] || primeraFila['AULA USS'] || null;
          nuevaFila['AULA USS'] = primeraFila['Aula USS'] || primeraFila['AULA USS'] || null;
          
          // Otros campos de programación que deben copiarse
          nuevaFila.DIAS = primeraFila.DIAS || null;
          nuevaFila['HORA INICIO'] = primeraFila['HORA INICIO'] || null;
          nuevaFila['HORA FIN'] = primeraFila['HORA FIN'] || null;
          
          // TURNO solo si existe en la primera fila (puede sobrescribirse con Zoom)
          nuevaFila.TURNO = primeraFila.TURNO || null;
          
          // Campo único de esta fila
          nuevaFila.SESION = sesion;

          // 🆕 HORAS PROGRAMADAS siempre 3 horas
          const possibleHorasProgramadasCols = ['HORAS PROGRAMADAS', 'Horas Programadas', 'horas programadas'];
          for (const col of possibleHorasProgramadasCols) {
            if (currentHeaders.includes(col)) {
              nuevaFila[col] = '03:00:00';
              break;
            }
          }

          // Buscar datos de Zoom para esta sesión específica
          if (zoomData.length > 0) {
            const sesionZoom = zoomData.find(zoomRow => {
              const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
              const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
              
              if (!matchDocente(docente, zoomDocente)) return false;
              
              const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
              if (!temaMatch) return false;
              
              const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
              const cursoZoom = cursoParte.trim();
              const sesionZoomNum = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;
              
              return normalizeCursoName(cursoZoom) === normalizeCursoName(curso) &&
                     normalizeSeccion(seccionZoom) === normalizeSeccion(seccion) &&
                     sesionZoomNum === sesion;
            });
            
            if (sesionZoom) {
              const fechaInicio = sesionZoom['Hora de inicio'] || sesionZoom['Start Time'] || "";
              const fechaFin = sesionZoom['Hora de finalización'] || sesionZoom['End Time'] || "";
              
              // Completar SOLO los campos que vienen de Zoom
              const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
              for (const col of possibleDateCols) {
                if (currentHeaders.includes(col)) {
                  nuevaFila[col] = extractDate(fechaInicio);
                  break;
                }
              }
              
              const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
              for (const col of possibleStartCols) {
                if (currentHeaders.includes(col)) {
                  nuevaFila[col] = extractTime(fechaInicio);
                  break;
                }
              }
              
              const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
              for (const col of possibleEndCols) {
                if (currentHeaders.includes(col)) {
                  nuevaFila[col] = extractTime(fechaFin);
                  break;
                }
              }
              
              const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)'];
              for (const col of possibleFinalizaCols) {
                if (currentHeaders.includes(col)) {
                  nuevaFila[col] = extractTime(fechaFin);
                  break;
                }
              }
              
              // Solo actualizar TURNO si estaba vacío Y viene de Zoom
              const turnoDetectado = detectTurno(fechaInicio);
              if (turnoDetectado && (!nuevaFila.TURNO || nuevaFila.TURNO === null || String(nuevaFila.TURNO).trim() === '')) {
                nuevaFila.TURNO = turnoDetectado;
              }

              // NUEVO: Detectar y setear MODELO si aplica
              const modeloDetectado = detectModeloFromCurso(curso);
              if (modeloDetectado && (!nuevaFila.MODELO || nuevaFila.MODELO === null || String(nuevaFila.MODELO).trim() === '')) {
                nuevaFila.MODELO = modeloDetectado;
              }

              // NUEVO: Calcular TIEMPO EFECTIVO DICTADO (siempre para nueva)
              nuevaFila['TIEMPO EFECTIVO DICTADO'] = calculateEffectiveTime(nuevaFila);
              
              console.log(`  ✓ Sesión ${sesion}: CREADA CON DATOS ZOOM`);
            } else {
              // Calcular con datos disponibles (puede ser null) - siempre para nueva
              nuevaFila['TIEMPO EFECTIVO DICTADO'] = calculateEffectiveTime(nuevaFila);
              console.log(`  + Sesión ${sesion}: CREADA (solo metadatos copiados)`);
            }
          } else {
            // Calcular con datos disponibles (puede ser null) - siempre para nueva
            nuevaFila['TIEMPO EFECTIVO DICTADO'] = calculateEffectiveTime(nuevaFila);
            console.log(`  + Sesión ${sesion}: CREADA (solo metadatos copiados)`);
          }
          
          sesionesCompletas.push(nuevaFila);
        }
      }
      
      // Agregar el bloque completo al resultado (SOLO las 16 sesiones)
      resultadoFinal.push(...sesionesCompletas);
      
      const nuevasCreadas = 16 - existingInRange.size;
      console.log(`  📊 Total final para grupo: 16 sesiones exactas`);
      console.log(`  📊 Sesiones existentes mantenidas: ${existingInRange.size}`);
      console.log(`  📊 Sesiones nuevas creadas: ${nuevasCreadas}`);
    });

    // Agregar filas que no tienen grupo definido al final
    dataProcesada.forEach((row) => {
      const docente = row.DOCENTE || '';
      const curso = row.CURSO || '';
      const seccion = normalizeSeccion(row.SECCION || row['SECCIÓN'] || '');
      
      if (!docente || !curso || !seccion) {
        // NUEVO: Calcular para filas sin grupo de forma condicional
        setEffectiveTimeConditionally(row);
        resultadoFinal.push(row);
      }
    });

    // Si no existe la columna en headers, agregarla
    if (!currentHeaders.includes('TIEMPO EFECTIVO DICTADO')) {
      setCurrentHeaders([...currentHeaders, 'TIEMPO EFECTIVO DICTADO']);
    }

    setData(resultadoFinal);

    const totalGrupos = gruposPorSeccion.size;
    const totalSesionesCreadas = Array.from(gruposPorSeccion.values())
      .reduce((sum, grupo) => sum + (16 - grupo.sesionesExistentes.size), 0);
    const totalConDatos = Array.from(gruposPorSeccion.values())
      .reduce((sum, grupo) => {
        let conDatos = 0;
        for (let s = 1; s <= 16; s++) {
          if (!grupo.sesionesExistentes.has(s)) {
            const existeZoom = zoomData.some(z => {
              const zd = z['Anfitrión'] || z['Host'] || "";
              const zt = z['Tema'] || z['Topic'] || "";
              if (!matchDocente(grupo.docente, zd)) return false;
              const tm = zt.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
              if (!tm) return false;
              const [, cp, sz, sn] = tm;
              return normalizeCursoName(cp.trim()) === normalizeCursoName(grupo.curso) &&
                     normalizeSeccion(sz) === normalizeSeccion(grupo.seccion) &&
                     (sn ? parseInt(sn) : 0) === s;
            });
            if (existeZoom) conDatos++;
          }
        }
        return sum + conDatos;
      }, 0);

    alert(`✅ Proceso completado:\n\n${totalGrupos} grupos procesados\n${totalSesionesCreadas} sesiones nuevas creadas\n${totalConDatos} sesiones con datos de Zoom\n\n✅ Registros existentes mantenidos SIN modificar\n✅ Cada grupo ahora tiene EXACTAMENTE 16 sesiones`);

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
  
  console.log('🔍 INICIANDO DETECCIÓN DE ENCABEZADOS');
  console.log('📄 Nombre de hoja:', worksheet.name);
  console.log('📏 Total de filas:', worksheet.rowCount);
  
  // Función mejorada para extraer texto de celdas
  const extractCellText = (cell) => {
    if (!cell || cell.value === null || cell.value === undefined) return null;
    
    const rawValue = cell.value;
    
    // String directo
    if (typeof rawValue === 'string') {
      const trimmed = rawValue.trim();
      return trimmed || null;
    }
    
    // Número
    if (typeof rawValue === 'number') {
      return String(rawValue);
    }
    
    // Fecha
    if (rawValue instanceof Date) {
      return rawValue.toLocaleDateString('es-ES');
    }
    
    // Objetos complejos
    if (typeof rawValue === 'object') {
      // Hiperlink
      if (rawValue.hyperlink && rawValue.text) {
        return String(rawValue.text).trim() || null;
      }
      
      // RichText
      if (Array.isArray(rawValue.richText)) {
        return rawValue.richText.map(rt => rt.text || '').join('').trim() || null;
      }
      
      // Formula con resultado
      if (rawValue.result !== undefined) {
        if (rawValue.result instanceof Date) {
          return rawValue.result.toLocaleDateString('es-ES');
        }
        if (typeof rawValue.result === 'number' || typeof rawValue.result === 'string') {
          return String(rawValue.result).trim() || null;
        }
      }
      
      // Texto dentro del objeto
      if (rawValue.text !== undefined) {
        return String(rawValue.text).trim() || null;
      }
    }
    
    return String(rawValue).trim() || null;
  };

  // PASO 1: Buscar fila de encabezados (scan las primeras 10 filas)
  let bestHeaderRow = 0;
  let maxNonEmptyCells = 0;
  
  for (let rowNum = 1; rowNum <= Math.min(10, worksheet.rowCount); rowNum++) {
    const row = worksheet.getRow(rowNum);
    const cellValues = [];
    let nonEmptyCount = 0;
    
    // Leer hasta 50 columnas
    for (let col = 1; col <= 50; col++) {
      const cell = row.getCell(col);
      const text = extractCellText(cell);
      cellValues.push(text);
      if (text && text.trim() !== '') {
        nonEmptyCount++;
      }
    }
    
    console.log(`📋 Fila ${rowNum}: ${nonEmptyCount} celdas con contenido`);
    
    // La fila con más celdas no vacías probablemente es el encabezado
    if (nonEmptyCount > maxNonEmptyCells && nonEmptyCount >= 3) {
      maxNonEmptyCells = nonEmptyCount;
      bestHeaderRow = rowNum;
      sheetHeaders = cellValues;
    }
  }
  
  headerRowIndex = bestHeaderRow || 1;
  console.log(`✅ Fila de encabezados detectada: ${headerRowIndex}`);
  console.log(`📊 Total de columnas detectadas: ${maxNonEmptyCells}`);
  
  // Limpiar headers: eliminar nulls del final y asignar nombres genéricos
  while (sheetHeaders.length > 0 && !sheetHeaders[sheetHeaders.length - 1]) {
    sheetHeaders.pop();
  }
  
  sheetHeaders = sheetHeaders.map((h, idx) => {
    if (h && h.trim() !== '') {
      return h.trim();
    }
    return `Columna ${idx + 1}`;
  });
  
  console.log('📋 Encabezados finales:', sheetHeaders);
  
  // PASO 2: Leer datos desde la fila siguiente al encabezado
  const getCellValue = (cell, colIndex) => {
  if (!cell || cell.value === null || cell.value === undefined) return null;
  
  const rawValue = cell.value;
  const headerName = sheetHeaders[colIndex - 1];
  
  // Fechas directas
  if (rawValue instanceof Date) {
    const year = rawValue.getUTCFullYear();
    const hours = rawValue.getUTCHours();
    const minutes = rawValue.getUTCMinutes();
    const seconds = rawValue.getUTCSeconds();
    
    // Si es solo hora (año 1899/1900 = fecha base de Excel)
    if (year === 1899 || year === 1900) {
      return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }
    
    // Fecha completa con hora
    if (hours !== 0 || minutes !== 0 || seconds !== 0) {
      const dateStr = rawValue.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
      const timeStr = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
      return `${dateStr} ${timeStr}`;
    }
    
    // Solo fecha
    return rawValue.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }
  
  // Números
  if (typeof rawValue === 'number') {
    // Porcentajes
    if (cell.numFmt && cell.numFmt.includes('%')) {
      return `${Math.round(rawValue * 100)}%`;
    }
    
    // Verificar si es formato de tiempo
    const cellFormat = (cell.numFmt || '').toLowerCase();
    const isTimeFormat = cellFormat.includes('h:mm') || 
                        cellFormat.includes('[h]') || 
                        cellFormat.includes('hh:mm') ||
                        cellFormat.includes('h:mm:ss') ||
                        cellFormat.includes('hh:mm:ss');
    
    const isDateFormat = cellFormat.includes('d/m') || 
                        cellFormat.includes('dd/mm') || 
                        cellFormat.includes('yyyy') ||
                        cellFormat.includes('dd-mm');
    
    // Si es menor a 1 y NO tiene formato de fecha, es un tiempo
    if (rawValue >= 0 && rawValue < 1 && (!isDateFormat || isTimeFormat)) {
      const totalSeconds = Math.round(rawValue * 24 * 60 * 60);
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }
    
    // Si es mayor o igual a 1, podría ser una fecha de Excel
    if (rawValue >= 1 && rawValue < 100000) {
      try {
        const excelDate = new Date((rawValue - 25569) * 86400 * 1000);
        if (!isNaN(excelDate.getTime()) && excelDate.getFullYear() > 1900 && excelDate.getFullYear() < 2100) {
          // Si tiene formato de tiempo explícito, mostrar solo la hora
          if (isTimeFormat && !isDateFormat) {
            const hours = excelDate.getUTCHours();
            const minutes = excelDate.getUTCMinutes();
            const seconds = excelDate.getUTCSeconds();
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
          }
          return excelDate.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
        }
      } catch (e) {
        console.error('❌ Error convirtiendo fecha:', e);
      }
    }
    
    return String(rawValue);
  }
  
  // Objetos complejos (AQUÍ ESTÁ LA CORRECCIÓN PRINCIPAL)
  if (typeof rawValue === 'object') {
    // Hiperlinks
    if (rawValue.hyperlink) {
      return rawValue.text || rawValue.hyperlink || null;
    }
    
    // RichText
    if (Array.isArray(rawValue.richText)) {
      return rawValue.richText.map(rt => rt.text || '').join('').trim() || null;
    }
    
    // ⭐ RESULTADO DE FÓRMULA (AQUÍ ESTÁ EL FIX)
    if (rawValue.result !== undefined) {
      const res = rawValue.result;
      
      // ✅ SI EL RESULTADO ES UN DATE OBJECT
      if (res instanceof Date) {
        const year = res.getUTCFullYear();
        const hours = res.getUTCHours();
        const minutes = res.getUTCMinutes();
        const seconds = res.getUTCSeconds();
        
        // Si es solo hora (año 1899/1900 = fecha base de Excel)
        if (year === 1899 || year === 1900) {
          return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
        }
        
        // Fecha completa con hora
        if (hours !== 0 || minutes !== 0 || seconds !== 0) {
          const dateStr = res.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
          const timeStr = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
          return `${dateStr} ${timeStr}`;
        }
        
        // Solo fecha
        return res.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
      }
      
      // Si el resultado es un número
      if (typeof res === 'number') {
        const cellFormat = (cell.numFmt || '').toLowerCase();
        const isTimeFormat = cellFormat.includes('h:mm') || cellFormat.includes('[h]');
        const isDateFormat = cellFormat.includes('d/m') || cellFormat.includes('dd/mm');
        
        // Porcentaje
        if (cell.numFmt && cell.numFmt.includes('%')) {
          return `${Math.round(res * 100)}%`;
        }
        
        // Tiempo (número decimal < 1)
        if (res >= 0 && res < 1 && (!isDateFormat || isTimeFormat)) {
          const totalSeconds = Math.round(res * 24 * 60 * 60);
          const hours = Math.floor(totalSeconds / 3600);
          const minutes = Math.floor((totalSeconds % 3600) / 60);
          const seconds = totalSeconds % 60;
          return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
        }
        
        return String(res);
      }
      
      // Si el resultado es string
      if (typeof res === 'string') {
        return res.trim() || null;
      }
    }
    
    // Texto directo
    if (rawValue.text !== undefined) {
      return String(rawValue.text).trim() || null;
    }
    
    return null;
  }
  
  // String
  const stringValue = String(rawValue).trim();
  return stringValue || null;
};
  
  console.log(`\n📖 Leyendo datos desde fila ${headerRowIndex + 1}...`);
  
  for (let rowIndex = headerRowIndex + 1; rowIndex <= worksheet.rowCount; rowIndex++) {
    const row = worksheet.getRow(rowIndex);
    const rowData = {};
    let hasAnyData = false;
    
    sheetHeaders.forEach((header, colIndex) => {
      const cell = row.getCell(colIndex + 1);
      const cellValue = getCellValue(cell, colIndex + 1);
      rowData[header] = cellValue;
      
      if (cellValue !== null && String(cellValue).trim() !== '') {
        hasAnyData = true;
      }
    });
    
    // Solo agregar filas que tengan al menos un dato
    if (hasAnyData) {
      loadedData.push(rowData);
    }
  }

  // NUEVO: Calcular TIEMPO EFECTIVO DICTADO para todas las filas cargadas de forma condicional
  loadedData.forEach(row => {
    setEffectiveTimeConditionally(row);
  });

  // NUEVO: Si no existe la columna en headers, agregarla al final
if (!sheetHeaders.includes('TIEMPO EFECTIVO DICTADO')) {
  sheetHeaders.push('TIEMPO EFECTIVO DICTADO');
}

// 🆕 Agregar columna EFICIENCIA si no existe
if (!sheetHeaders.includes('EFICIENCIA')) {
  sheetHeaders.push('EFICIENCIA');
}
  
  console.log(`✅ Total de registros cargados: ${loadedData.length}`);
  if (loadedData.length > 0) {
    console.log('📄 Primera fila de ejemplo:', loadedData[0]);
  }
  
  return { data: loadedData, headers: sheetHeaders };
};

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
        sheetData: updatedSheetData
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
      sheetData: updatedSheetData
    });
  };
  const exportToExcel = async () => {
    if (data.length === 0) {
      alert('No hay datos para exportar.');
      return;
    }
    // Crear un workbook NUEVO (sin clonar, ya que ExcelJS no tiene .clone() nativo)
    const workbook = new ExcelJS.Workbook();
    const sheetDataCache = activeTab?.sheetData || {};
    const availableSheetsList = availableSheets || [{ index: 0, name: 'Monitoreo' }]; // Fallback para single sheet
    availableSheetsList.forEach(({ index: sheetIndex, name: sheetName }) => {
      // Crear worksheet nuevo con nombre original
      const worksheet = workbook.addWorksheet(sheetName || `Hoja${sheetIndex + 1}`);

      // Obtener datos cacheados para esta hoja (incluye cambios como filas agregadas)
      const cached = sheetDataCache[sheetIndex] || { data: data, headers: currentHeaders };
      const sheetHeaders = cached.headers || [];
      const sheetDataToExport = cached.data || [];

      // Agregar headers si existen
      if (sheetHeaders.length > 0) {
        const headerRow = worksheet.addRow(sheetHeaders);
        headerRow.height = 40;
        headerRow.eachCell((cell, colNumber) => {
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
      }
      // Agregar filas de datos con cambios aplicados
      sheetDataToExport.forEach((row, rowIdx) => {
        const rowData = sheetHeaders.map(h => row[h] !== undefined ? row[h] : null);
        const excelRow = worksheet.addRow(rowData);
        excelRow.height = 25;
        const isEven = (rowIdx + 2) % 2 === 0; // +2 porque fila 1 es header
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
      // Ajustar anchos de columnas dinámicamente
      if (sheetHeaders.length > 0) {
        const dynamicWidths = sheetHeaders.map(() => Math.min(50, Math.max(10, 15)));
        worksheet.columns = sheetHeaders.map((header, idx) => ({
          key: header,
          width: dynamicWidths[idx] || 15
        }));
      }
    });
    // Generar y descargar el archivo
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Monitoreo_USS_${activeTab?.name ? activeTab.name.replace(/\.[^/.]+$/, "") : "datos"}_${new Date().toISOString().split('T')[0]}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);

    console.log(`✅ Exportado ${availableSheetsList.length} hojas con todos los cambios (filas agregadas incluidas).`);
    alert(`✅ ¡Archivo exportado exitosamente con ${availableSheetsList.length} hoja(s) y todos los cambios guardados!`);
  };
  const deleteRow = (index) => {
    // Mapear índice visible a índice real cuando MONITOREO reordena filas por DOCENTE
    let realIndex = index;
    const selectedSheetName = (availableSheets[selectedSheet]?.name || '').toString();
    const isMonitoreoView = selectedSheetName.toLowerCase().includes('monitoreo');
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
  };
  const handleCellChange = (rowIndex, columnName, value) => {
  console.log(`🔧 handleCellChange llamado:`);
  console.log(`   Fila visual: ${rowIndex}`);
  console.log(`   Columna: ${columnName}`);
  console.log(`   Nuevo valor: "${value}"`);
  console.log(`   Total filas: ${data.length}`);
  
  // Validar que el índice sea válido
  if (rowIndex < 0 || rowIndex >= data.length) {
    console.error(`❌ ERROR: Índice ${rowIndex} fuera de rango (0-${data.length - 1})`);
    return;
  }

  // Crear copia del array de datos
  const newData = [...data];
  
  // Verificar que la fila existe
  if (!newData[rowIndex]) {
    console.error(`❌ ERROR: Fila no encontrada en índice ${rowIndex}`);
    return;
  }
  
  console.log(`   Valor anterior en fila ${rowIndex}[${columnName}]:`, newData[rowIndex][columnName]);
  
  // Actualizar el valor (incluso si es vacío)
  newData[rowIndex] = {
    ...newData[rowIndex],
    [columnName]: value === "" ? null : value // Guardar null en vez de string vacío
  };
  
  console.log(`   Valor nuevo guardado:`, newData[rowIndex][columnName]);

  // NUEVO: Si el cambio afecta el cálculo de TIEMPO EFECTIVO DICTADO, recalcularlo (sobrescribir siempre en este caso, ya que es cambio manual)
  // NUEVO: Si el cambio afecta el cálculo de TIEMPO EFECTIVO DICTADO, recalcularlo
const relevantColumns = [
  'FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom',
  'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE',
  'HORAS PROGRAMADAS', 'Horas Programadas', 'horas programadas'
];
if (relevantColumns.some(col => col === columnName)) {
  newData[rowIndex]['TIEMPO EFECTIVO DICTADO'] = calculateEffectiveTime(newData[rowIndex]);
  newData[rowIndex]['EFICIENCIA'] = calculateEfficiency(newData[rowIndex]);
  console.log(`   📊 Recalculado TIEMPO EFECTIVO DICTADO: ${newData[rowIndex]['TIEMPO EFECTIVO DICTADO']}`);
  console.log(`   📊 Recalculado EFICIENCIA: ${newData[rowIndex]['EFICIENCIA']}`);
}
  
  console.log(`✅ Cambio aplicado correctamente`);
  
  // Actualizar estado
  setData(newData);
};

  // ===== DATOS COMPUTADOS =====
  // Generar opciones dinámicas desde los datos
  const uniqueCursos = useMemo(() => {
    const cursos = new Set();
    data.forEach(row => {
      if (row.CURSO && row.CURSO.toString().trim() !== '') {
        cursos.add(row.CURSO.toString().trim());
      }
    });
    return Array.from(cursos).sort();
  }, [data]);
  const uniqueSecciones = useMemo(() => {
    const secciones = new Set();
    data.forEach(row => {
      if (row.SECCION && row.SECCION.toString().trim() !== '') {
        secciones.add(row.SECCION.toString().trim());
      }
    });
    return Array.from(secciones).sort();
  }, [data]);
  const uniqueTurnos = useMemo(() => {
    const turnos = new Set();
    data.forEach(row => {
      if (row.TURNO && row.TURNO.toString().trim() !== '') {
        turnos.add(row.TURNO.toString().trim());
      }
    });
    return Array.from(turnos).sort();
  }, [data]);
  const uniqueDias = useMemo(() => {
    const dias = new Set();
    data.forEach(row => {
      if (row.DIAS && row.DIAS.toString().trim() !== '') {
        dias.add(row.DIAS.toString().trim());
      }
    });
    return Array.from(dias).sort();
  }, [data]);
  const uniqueModelos = useMemo(() => {
    const modelos = new Set();
    data.forEach(row => {
      if (row.MODELO && row.MODELO.toString().trim() !== '') {
        modelos.add(row.MODELO.toString().trim());
      }
    });
    return Array.from(modelos).sort();
  }, [data]);
  const uniqueModalidades = useMemo(() => {
    const modalidades = new Set();
    data.forEach(row => {
      if (row.MODALIDAD && row.MODALIDAD.toString().trim() !== '') {
        modalidades.add(row.MODALIDAD.toString().trim());
      }
    });
    return Array.from(modalidades).sort();
  }, [data]);
  const uniqueCiclos = useMemo(() => {
    const ciclos = new Set();
    data.forEach(row => {
      if (row.CICLO && row.CICLO.toString().trim() !== '') {
        ciclos.add(row.CICLO.toString().trim());
      }
    });
    return Array.from(ciclos).sort();
  }, [data]);
  const uniquePeriodos = useMemo(() => {
    const periodos = new Set();
    data.forEach(row => {
      if (row.PERIODO && row.PERIODO.toString().trim() !== '') {
        periodos.add(row.PERIODO.toString().trim());
      }
    });
    return Array.from(periodos).sort();
  }, [data]);
  const selectedSheetName = (availableSheets[selectedSheet]?.name || '').toString();
  const isMonitoreoView = selectedSheetName.toLowerCase().includes('monitoreo');
  const dropdownOptions = {
    MODELO: uniqueModelos.length > 0 ? uniqueModelos : [],
    MODALIDAD: uniqueModalidades.length > 0 ? uniqueModalidades : [],
    CURSO: uniqueCursos.length > 0 ? uniqueCursos : [],
    SECCION: uniqueSecciones.length > 0 ? uniqueSecciones : [],
    TURNO: uniqueTurnos.length > 0 ? uniqueTurnos : [],
    DIAS: uniqueDias.length > 0 ? uniqueDias : [],
    CICLO: uniqueCiclos.length > 0 ? uniqueCiclos : [],
    PERIODO: uniquePeriodos.length > 0 ? uniquePeriodos : []
  };
  const displayData = useMemo(() => {
    // Simplemente retornar los datos en su orden original, SIN REORDENAR
    // NUEVO: Asegurar cálculo en display si es necesario (aunque ya se hace en setData)
    return data.map(row => ({
      ...row,
      'TIEMPO EFECTIVO DICTADO': row['TIEMPO EFECTIVO DICTADO'] || calculateEffectiveTime(row)
    }));
  }, [data]);

  // ===== RENDER =====
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 p-4">
      <div className="max-w-full mx-auto">
        {/* DESCARGA DE ARCHIVOS PRINCIPALES */}
        <div className="mb-6 flex gap-4 items-center">
          <span className="font-bold text-blue-900 text-lg">Descargar archivos principales:</span>
          <a
            href="/EJEMPLO.xlsx"
            download
            className="bg-blue-900 text-white px-3 py-1 rounded-lg shadow hover:bg-blue-700 text-xs font-bold"
            style={{ textDecoration: 'none' }}
          >
            Descargar Excel
          </a>
          <a
            href="/meetings_Docentes_CIS_2025_09_08_2025_09_21.csv"
            download
            className="bg-blue-900 text-white px-3 py-1 rounded-lg shadow hover:bg-blue-700 text-xs font-bold"
            style={{ textDecoration: 'none' }}
          >
            Descargar CSV
          </a>
        </div>
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
  onSelectRandomDocente={selectRandomDocente}
  randomDocente={randomDocente}
  onClearRandomDocente={() => setRandomDocente(null)}
  onSaveBackup={saveBackup}
  onOpenBackupModal={() => setIsBackupModalOpen(true)}
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
          </div>
        )}

        {/* Modal para historial de backups - MOVIDO AQUÍ DENTRO */}
      <BackupHistoryModal
        isOpen={isBackupModalOpen}
        onClose={() => setIsBackupModalOpen(false)}
        backups={backupHistory}
        onDownload={downloadBackup}
        onDelete={deleteBackup}
      />


      </div>
    </div>    
  );
}
export default App;

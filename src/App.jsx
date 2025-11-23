import React, { useState, useMemo } from "react";
import ExcelJS from "exceljs";
import ControlPanel from "./components/ControlPanel";
import DataTable from "./components/DataTable";
import { BackupHistoryModal, useBackupManager } from './components/BackupManager';
import Guide from './components/Guide';
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
  
  // Botón de cierre manual
  const closeBtn = document.createElement('button');
  closeBtn.textContent = '×';
  closeBtn.style.marginLeft = '10px';
  closeBtn.style.background = 'transparent';
  closeBtn.style.border = 'none';
  closeBtn.style.color = 'white';
  closeBtn.style.fontSize = '16px';
  closeBtn.style.cursor = 'pointer';
  closeBtn.style.float = 'right';
  closeBtn.setAttribute('aria-label', 'Cerrar');
  closeBtn.addEventListener('click', () => {
    if (document.body.contains(toast)) {
      document.body.removeChild(toast);
    }
  });
  toast.appendChild(closeBtn);
  
  // Permitir cerrar haciendo clic en el toast
  toast.style.cursor = 'pointer';
  toast.addEventListener('click', () => {
    if (document.body.contains(toast)) {
      document.body.removeChild(toast);
    }
  });
  
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
  deleteBackup,
  restoreBackup  // ✅ AGREGAR ESTA LÍNEA
} = useBackupManager(ExcelJS, mostrarToast);

// Función wrapper para saveBackup
const saveBackup = () => {
  saveBackupToStorage(data, currentHeaders, activeTab, setIsLoading);
};

const handleRestoreBackup = (backup) => {
  restoreBackup(backup, createNewTab, ExcelJS);
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

// Agregador de diferencias de sección (se establece durante el procesamiento)
let collectSeccionDiff = null;

const matchSecciones = (seccionExcel, seccionZoom) => {
  const normalizedExcel = normalizeSeccion(seccionExcel);
  const normalizedZoom = normalizeSeccion(seccionZoom);
  
  // Coincidencia exacta
  if (normalizedExcel === normalizedZoom) return true;
  
  // Si hay similitud por contención ("A" vs "AA"), registrar discrepancia pero NO autocompletar
  if (normalizedExcel.includes(normalizedZoom) || normalizedZoom.includes(normalizedExcel)) {
    if (typeof collectSeccionDiff === 'function') {
      collectSeccionDiff(seccionExcel, seccionZoom);
    }
    console.log(`⚠️ Discrepancia detectada (secciones distintas): Excel "${seccionExcel}" vs Zoom "${seccionZoom}"`);
    return false;
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
    if (!dateTimeStr) return "";
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
  const extractDuration = (zoomRow) => {
    if (!zoomRow) return "";
    // Intentar múltiples variantes de encabezado de Zoom
    const candidates = [
      'Duración (hh:mm:ss)',
      'Duration (hh:mm:ss)',
      'Duración',
      'Duration',
      'Duración (minutos)',
      'Duración (Minutos)',
      'Duration (Minutes)'
    ];
    let durStr = candidates.map(k => zoomRow[k]).find(v => v && String(v).trim() !== '');
    if (!durStr) {
      // Fallback: calcular a partir de hora de inicio y fin
      const startStr = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || '';
      const endStr = zoomRow['Hora de finalización'] || zoomRow['End Time'] || '';
      const startTime = extractTime(startStr);
      const endTime = extractTime(endStr);
      const sMin = timeToMinutes(startTime);
      const eMin = timeToMinutes(endTime);
      if (isFinite(sMin) && isFinite(eMin) && eMin >= sMin) {
        const diffSec = (eMin - sMin) * 60;
        return secondsToHHMMSS(diffSec);
      }
      return "";
    }
    const trimmed = String(durStr).trim();
    // Si el valor trae ":", interpretarlo como HH:MM:SS
    if (trimmed.includes(':')) {
      const secs = durationToSeconds(trimmed);
      if (isFinite(secs)) return secondsToHHMMSS(secs);
      return trimmed;
    }
    // Si no trae ":", asumir minutos enteros
    const minutes = parseFloat(trimmed);
    if (!isNaN(minutes)) {
      const hours = Math.floor(minutes / 60);
      const mins = Math.round(minutes % 60);
      return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}:00`;
    }
    return trimmed;
  };
  // Utilidades para manejar duraciones tipo HH:MM:SS
  const durationToSeconds = (str) => {
    if (!str) return NaN;
    const s = String(str).trim();
    const m = s.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
    if (m) {
      const h = parseInt(m[1] || '0');
      const min = parseInt(m[2] || '0');
      const sec = parseInt(m[3] || '0');
      return h * 3600 + min * 60 + sec;
    }
    const onlyMin = parseInt(s);
    if (!isNaN(onlyMin)) return onlyMin * 60; // minutos a segundos
    return NaN;
  };
  const secondsToHHMMSS = (sec) => {
    if (!isFinite(sec) || sec < 0) sec = 0;
    const h = Math.floor(sec / 3600);
    const rem = sec % 3600;
    const m = Math.floor(rem / 60);
    const s = Math.floor(rem % 60);
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };
  const detectTurno = (horaStr) => {
    if (!horaStr) return "";
   
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
    if (!tema) return "";
    const match = tema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
    return match ? match[1].trim() : tema;
  };
  // ===== HANDLERS =====

  const handleAutocompletarConZoom = async () => {
  if (data.length === 0) {
    mostrarToast(`⚠️ Primero carga el archivo Excel`, "warning");
    alert("⚠️ Primero carga el archivo Excel");
    return;
  }
  setIsLoading(true);
 
  try {
    console.log("=== INICIANDO PROCESO COMPLETO ===");
    // Acumuladores para notificaciones agregadas
    const docentesSinPEAD = new Set();
   
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
         
          // Usar matchSecciones en lugar de igualdad exacta
          return normalizeCursoName(cursoZoom) === normalizeCursoName(curso) &&
                 matchSecciones(seccion, seccionZoom) &&
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
         
          // ACTUALIZADO: Guardar duración de la grabación en FINALIZA LA CLASE (ZOOM)
          const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom'];
          for (const col of possibleFinalizaCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = extractDuration(sesionZoom);
              break;
            }
          }
          // Calcular TIEMPO EFECTIVO DICTADO y EFICIENCIA en esta pasada
          const possibleWaitCols = ['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE', 'Tiempo de espera antes de iniciar la clase', 'TIEMPO DE ESPERA', 'Espera antes de iniciar'];
          const possibleEffectiveCols = ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'];
          const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada'];
          const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
          const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
          const zoomDurStr = extractDuration(sesionZoom);
          for (const col of possibleTotalCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = zoomDurStr;
              break;
            }
          }
          const waitStr = possibleWaitCols.map(c => dataProcesada[index][c]).find(v => v && String(v).trim() !== '');
          const durationSec = durationToSeconds(zoomDurStr);
          const waitSec = durationToSeconds(waitStr);
          let effectiveSec = NaN;
          if (isFinite(durationSec) && isFinite(waitSec)) {
            effectiveSec = Math.max(durationSec - waitSec, 0);
            for (const col of possibleEffectiveCols) {
              if (currentHeaders.includes(col)) {
                dataProcesada[index][col] = secondsToHHMMSS(effectiveSec);
                break;
              }
            }
          }
          const progStr = possibleProgramadoCols.map(c => dataProcesada[index][c]).find(v => v && String(v).trim() !== '');
          const progSec = durationToSeconds(progStr);
          if (isFinite(durationSec) && isFinite(effectiveSec) && durationSec > 0) {
            const toleranceSec = 10 * 60;
            let numerator = effectiveSec;
            if (isFinite(progSec) && progSec > 0) {
              const threshold = Math.max(progSec - toleranceSec, 0);
              numerator = Math.max(effectiveSec - threshold, 0);
            }
            const eficiencia = Math.min(numerator / durationSec, 1);
            const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
            for (const col of possibleEficienciaCols) {
              if (currentHeaders.includes(col)) {
                dataProcesada[index][col] = eficienciaStr;
                break;
              }
            }
          }
          
          // Mostrar notificación de éxito
          const temaZoom = sesionZoom['Tema'] || sesionZoom['Topic'] || "";
          const zoomMatch = temaZoom.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)/i);
          const seccionZoomMostrar = zoomMatch ? zoomMatch[2] : "";
          
          mostrarToast(`✅ Fila actualizada:<br>
            <b>Docente:</b> ${docente}<br>
            <b>Excel:</b> ${seccion} / <b>Zoom:</b> ${seccionZoomMostrar}<br>
            <b>Sesión:</b> ${sesion}`, "success");
          console.log(` ✓ Autocompletado: ${docente} - ${curso} - Sesión ${sesion}`);
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
        if (hasFecha) return;
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
          // ACTUALIZADO: Guardar duración de la grabación en FINALIZA LA CLASE (ZOOM)
          const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom'];
          setIfHasHeader(row, possibleFinalizaCols, extractDuration(bestZoom));
          // Calcular TIEMPO EFECTIVO DICTADO y EFICIENCIA en fallback por horario
          const possibleWaitCols = ['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE', 'Tiempo de espera antes de iniciar la clase', 'TIEMPO DE ESPERA', 'Espera antes de iniciar'];
          const possibleEffectiveCols = ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'];
          const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada'];
          const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
          const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
          const zoomDurStr = extractDuration(bestZoom);
          setIfHasHeader(row, possibleTotalCols, zoomDurStr);
          const waitStr = possibleWaitCols.map(c => row[c]).find(v => v && String(v).trim() !== '');
          const durationSec = durationToSeconds(zoomDurStr);
          const waitSec = durationToSeconds(waitStr);
          let effectiveSec = NaN;
          if (isFinite(durationSec) && isFinite(waitSec)) {
            effectiveSec = Math.max(durationSec - waitSec, 0);
            setIfHasHeader(row, possibleEffectiveCols, secondsToHHMMSS(effectiveSec));
          }
          const progStr = possibleProgramadoCols.map(c => row[c]).find(v => v && String(v).trim() !== '');
          const progSec = durationToSeconds(progStr);
          if (isFinite(durationSec) && isFinite(effectiveSec) && durationSec > 0) {
            const toleranceSec = 10 * 60; // N1 = 00:10:00
            let numerator = effectiveSec;
            if (isFinite(progSec) && progSec > 0) {
              const threshold = Math.max(progSec - toleranceSec, 0);
              numerator = Math.max(effectiveSec - threshold, 0);
            }
            const eficiencia = Math.min(numerator / durationSec, 1);
            const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
            setIfHasHeader(row, possibleEficienciaCols, eficienciaStr);
          }
          row.TURNO = row.TURNO && String(row.TURNO).trim() !== '' ? row.TURNO : detectTurno(fechaInicio);
          // Calcular SI/NO para "INICIO SESION 10 MINUTOS ANTES" usando hora programada original y el inicio Zoom
          const earlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES'];
          const timeToMinutes = (timeStr) => {
            if (!timeStr || typeof timeStr !== 'string') return NaN;
            let s = timeStr.trim();
            s = s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
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
            return NaN;
          };
          const horaProgOriginal = horaProg; // de arriba en este bloque
          const progMin = timeToMinutes(horaProgOriginal);
          const zoomMin = timeToMinutes(extractTime(fechaInicio));
          const minEarly = 5;
          const maxEarly = 10;
          const diffEarly = isFinite(progMin) && isFinite(zoomMin) ? (progMin - zoomMin) : NaN;
          const inicioAntes = isFinite(diffEarly) && diffEarly >= minEarly && diffEarly <= maxEarly;
          setIfHasHeader(row, earlyCols, inicioAntes ? 'SI' : 'NO');
          usedZoomByStart.add(bestStartStr);
          
          mostrarToast(`✅ Coincidencia por horario:<br>
            <b>Docente:</b> ${docente}<br>
            <b>Diferencia:</b> ${Math.round(bestDiff)} minutos`, "info");
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
      console.log(` Sesiones existentes: ${Array.from(sesionesExistentes).sort((a,b) => a-b).join(', ')}`);
      console.log(` Total filas existentes: ${filas.length}`);
     
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
        console.log(` 📌 Primera fila asignada como SESION 1`);
      }
     
      // Crear exactamente 16 sesiones (1-16)
      for (let sesion = 1; sesion <= 16; sesion++) {
        if (existingInRange.has(sesion)) {
          // Usar la fila ORIGINAL completa SIN MODIFICAR
          const filaExistente = existingInRange.get(sesion);
         
          // Asegurarse de que SESION sea el número correcto
          filaExistente.SESION = sesion;
         
          sesionesCompletas.push(filaExistente);
          console.log(` ○ Sesión ${sesion}: YA EXISTE (mantenida con todos sus datos)`);
        } else {
          // Crear nueva fila con METADATOS básicos copiados de la primera fila
          const nuevaFila = {};
          
          // PRIMERO: Detectar y copiar TODAS las columnas relacionadas con HORAS PROGRAMADAS
          currentHeaders.forEach(columna => {
            // Copiar SOLO los campos de horas programadas (cualquier variante de nombre)
            const columnaUpper = columna.toUpperCase();
            
            if (
              // Detectar cualquier campo que mencione HORA PROGRAMADA o variantes
              (columnaUpper.includes('HORA') && 
               (columnaUpper.includes('PROG') || 
                columnaUpper.includes('PROGRAMADA')))
            ) {
              // Copiar el valor de la primera fila
              nuevaFila[columna] = primeraFila[columna] || '';
            }
          });
          
          // LUEGO: Copiar los metadatos básicos
          nuevaFila.DOCENTE = primeraFila.DOCENTE || '';
          nuevaFila.CURSO = primeraFila.CURSO || '';
          nuevaFila.SECCION = primeraFila.SECCION || '';
          nuevaFila.MODELO = primeraFila.MODELO || '';
          nuevaFila.MODALIDAD = primeraFila.MODALIDAD || '';
          nuevaFila.CICLO = primeraFila.CICLO || '';
          nuevaFila.PERIODO = primeraFila.PERIODO || '';
         
          // Aula USS copiada TAL CUAL de la primera fila
          nuevaFila['Aula USS'] = primeraFila['Aula USS'] || primeraFila['AULA USS'] || '';
          nuevaFila['AULA USS'] = primeraFila['Aula USS'] || primeraFila['AULA USS'] || '';
         
          // Otros campos de programación que deben copiarse
          nuevaFila.DIAS = primeraFila.DIAS || '';
          
          // TURNO solo si existe en la primera fila (puede sobrescribirse con Zoom)
          nuevaFila.TURNO = primeraFila.TURNO || '';
         
          // Campo único de esta fila
          nuevaFila.SESION = sesion;
         
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
                     matchSecciones(seccion, seccionZoom) &&
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
             
              // ACTUALIZADO: Guardar duración de la grabación en FINALIZA LA CLASE (ZOOM)
              const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)'];
              for (const col of possibleFinalizaCols) {
                if (currentHeaders.includes(col)) {
                  nuevaFila[col] = extractDuration(sesionZoom);
                  break;
                }
              }
             
              // Solo actualizar TURNO si estaba vacío Y viene de Zoom
              const turnoDetectado = detectTurno(fechaInicio);
              if (turnoDetectado && (!nuevaFila.TURNO || String(nuevaFila.TURNO).trim() === '')) {
                nuevaFila.TURNO = turnoDetectado;
              }
              
              // Mostrar notificación
              const temaZoom = sesionZoom['Tema'] || sesionZoom['Topic'] || "";
              const zoomMatch = temaZoom.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)/i);
              const seccionZoomMostrar = zoomMatch ? zoomMatch[2] : "";
              
              mostrarToast(`✅ Sesión ${sesion} creada con datos Zoom:<br>
                <b>Docente:</b> ${docente}<br>
                <b>Excel:</b> ${seccion} / <b>Zoom:</b> ${seccionZoomMostrar}`, "success");
              console.log(` ✓ Sesión ${sesion}: CREADA CON DATOS ZOOM`);
            } else {
              console.log(` + Sesión ${sesion}: CREADA (solo metadatos copiados)`);
            }
          } else {
            console.log(` + Sesión ${sesion}: CREADA (solo metadatos copiados)`);
          }
         
          sesionesCompletas.push(nuevaFila);
        }
      }
     
      // Agregar el bloque completo al resultado (SOLO las 16 sesiones)
      resultadoFinal.push(...sesionesCompletas);
     
      const nuevasCreadas = 16 - existingInRange.size;
      console.log(` 📊 Total final para grupo: 16 sesiones exactas`);
      console.log(` 📊 Sesiones existentes mantenidas: ${existingInRange.size}`);
      console.log(` 📊 Sesiones nuevas creadas: ${nuevasCreadas}`);
    });
    // Agregar filas que no tienen grupo definido al final
    dataProcesada.forEach((row) => {
      const docente = row.DOCENTE || '';
      const curso = row.CURSO || '';
      const seccion = normalizeSeccion(row.SECCION || row['SECCIÓN'] || '');
     
      if (!docente || !curso || !seccion) {
        resultadoFinal.push(row);
      }
    });
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
                     matchSecciones(grupo.seccion, sz) &&
                     (sn ? parseInt(sn) : 0) === s;
            });
            if (existeZoom) conDatos++;
          }
        }
        return sum + conDatos;
      }, 0);
      
    mostrarToast(`✅ Proceso completado:<br>
      <b>${totalGrupos}</b> grupos procesados<br>
      <b>${totalSesionesCreadas}</b> sesiones creadas<br>
      <b>${totalConDatos}</b> con datos Zoom`, "success");
      
    alert(`✅ Proceso completado:\n\n${totalGrupos} grupos procesados\n${totalSesionesCreadas} sesiones nuevas creadas\n${totalConDatos} sesiones con datos de Zoom\n\n✅ Registros existentes mantenidos SIN modificar\n✅ Cada grupo ahora tiene EXACTAMENTE 16 sesiones`);
    console.log("=== PROCESO FINALIZADO ===");
     
  } catch (error) {
    mostrarToast(`❌ Error: ${error.message}`, "error");
    console.error("Error en proceso:", error);
    alert("❌ Error: " + error.message);
  } finally {
    setIsLoading(false);
  }
};


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
    
    const docentesToProcess = [...new Set(data.map(row => row.DOCENTE).filter(d => d && d.trim() !== ''))];
    if (docentesToProcess.length === 0) {
      mostrarToast(`⚠️ No hay docentes registrados en el Excel para autocompletar`, "warning");
      alert("No hay docentes registrados en el Excel para autocompletar");
      return;
    }
    
    console.log(`📋 Modo: TODOS los docentes`);
    console.log(`📋 Docentes a procesar (${docentesToProcess.length}):`, docentesToProcess);
    
    let updatedCount = 0;
    let createdCount = 0;
    const newData = [...data];
    const sesionesUsadasGlobal = new Set();
    // Agrega agregación de docentes sin PEAD detectado para una sola notificación
    const docentesSinPEAD = new Set();
    // Agrega agregación de discrepancias de sección para una sola notificación
    const seccionDiscrepancias = new Set();
    // Activar colector de discrepancias para esta ejecución
    collectSeccionDiff = (excel, zoom) => {
      seccionDiscrepancias.add(`${excel}|||${zoom}`);
    };
    
    const updateRowWithZoom = (row, zoomInfo, zoomRow) => {
      const updatedRow = { ...row };
      // Helper para convertir hora a minutos
      const timeToMinutes = (timeStr) => {
        if (!timeStr || typeof timeStr !== 'string') return NaN;
        let s = timeStr.trim();
        s = s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
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
        return NaN;
      };
      
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
      
      for (const col of possibleFinalizaCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = extractDuration(zoomRow);
          break;
        }
      }

      // Calcular TIEMPO EFECTIVO DICTADO y EFICIENCIA
      const possibleWaitCols = ['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE', 'Tiempo de espera antes de iniciar la clase', 'TIEMPO DE ESPERA', 'Espera antes de iniciar'];
      const possibleEffectiveCols = ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'];
      const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada'];
      const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
      const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];

      const zoomDurStr = extractDuration(zoomRow);
      // También llenar "Duración total clase" si existe ese encabezado
      for (const col of possibleTotalCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = zoomDurStr;
          break;
        }
      }
      const waitStr = possibleWaitCols.map(c => updatedRow[c] ?? row[c]).find(v => v && String(v).trim() !== '');
      const durationSec = durationToSeconds(zoomDurStr);
      const waitSec = durationToSeconds(waitStr);
      let effectiveSec = NaN;
      if (isFinite(durationSec) && isFinite(waitSec)) {
        effectiveSec = Math.max(durationSec - waitSec, 0);
        const effStr = secondsToHHMMSS(effectiveSec);
        for (const col of possibleEffectiveCols) {
          if (currentHeaders.includes(col)) {
            updatedRow[col] = effStr;
            break;
          }
        }
      }
      // Calcular EFICIENCIA solo si hay tiempo programado y duración
      const progStr = possibleProgramadoCols.map(c => updatedRow[c] ?? row[c]).find(v => v && String(v).trim() !== '');
      const progSec = durationToSeconds(progStr);
      if (isFinite(durationSec) && isFinite(effectiveSec) && durationSec > 0) {
        const toleranceSec = 10 * 60; // N1 = 00:10:00
        let numerator = effectiveSec;
        if (isFinite(progSec) && progSec > 0) {
          const threshold = Math.max(progSec - toleranceSec, 0);
          numerator = Math.max(effectiveSec - threshold, 0);
        }
        const eficiencia = Math.min(numerator / durationSec, 1);
        const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
        for (const col of possibleEficienciaCols) {
          if (currentHeaders.includes(col)) {
            updatedRow[col] = eficienciaStr;
            break;
          }
        }
      }
      
      const possibleEarlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES'];
      const horaProg = row['HORA INICIO'] || row['Hora Inicio'] || row['INICIO'] || row['inicio'] || '';
      const progMin = timeToMinutes(horaProg);
      const zoomMin = timeToMinutes(zoomInfo.horaInicio);
      const inferScheduledFromZoom = (zm) => {
        if (!isFinite(zm)) return NaN;
        const minute = Math.floor(zm % 60);
        if (minute >= 45) {
          return zm - minute + 60; // próxima hora en punto
        }
        return NaN;
      };
      const scheduledMin = isFinite(progMin) ? progMin : inferScheduledFromZoom(zoomMin);
      const minEarly = 5; // mínimo 5 minutos antes
      const maxEarly = 10; // máximo 10 minutos antes
      const diffEarly = isFinite(scheduledMin) && isFinite(zoomMin) ? (scheduledMin - zoomMin) : NaN;
      const inicioAntes = isFinite(diffEarly) && diffEarly >= minEarly && diffEarly <= maxEarly;
      for (const col of possibleEarlyCols) {
        if (currentHeaders.includes(col)) {
          updatedRow[col] = inicioAntes ? 'SI' : 'NO';
          break;
        }
      }

      const norm = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
      const findHeader = (aliases) => {
        for (const h of currentHeaders) {
          const hN = norm(h);
          for (const a of aliases) { if (hN === norm(a)) return h; }
        }
        return null;
      };
      const videoFlag = String((zoomRow && (zoomRow['Video'] || '')) || '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .trim().toLowerCase();
      if (videoFlag === 'no') {
        const obsHeader = findHeader(['OBSERVACIÓN', 'OBSERVACION', 'Observación', 'Observacion']);
        const recHeader = findHeader(['RECOMENDACIONES', 'Recomendaciones', 'RECOMENDACION', 'Recomendacion']);
        if (obsHeader) updatedRow[obsHeader] = 'no activo camara';
        if (recHeader) updatedRow[recHeader] = 'activar su camara';
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
          
          // Patrón estándar para encontrar PEAD
          const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
          
          if (!temaMatch) {
            // Intentar patrón más flexible para encontrar variantes de PEAD
            const patternFlexible = /(.+?)(?:(?:–|-|\/|:)\s*)(PEAD[-_ ]?[a-zA-Z0-9]*)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i;
            const temaMatchFlexible = zoomTema.match(patternFlexible);
            
            if (!temaMatchFlexible) {
              docentesSinPEAD.add(docenteActual);
              continue;
            }
            
            // Usar el formato flexible si fue encontrado, pero no spamear notificaciones
            const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatchFlexible;
            docentesSinPEAD.add(docenteActual);
            continue;
          }
          
          const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
          const cursoZoom = cursoParte.trim();
          const sesionZoom = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;
          const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;
          
          if (sesionesUsadasGlobal.has(claveZoom)) continue;
          
          const cursoMatch = row.CURSO && normalizeCursoName(row.CURSO) === normalizeCursoName(cursoZoom);
          const seccionMatch = row.SECCION && matchSecciones(row.SECCION, seccionZoom);
          const sesionMatch = row.SESION && parseInt(String(row.SESION)) === sesionZoom;
          
          if (cursoMatch && seccionMatch && sesionMatch) {
            const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
            const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
            
            const updatedRow = { ...row };
            
            const possibleDateCols = ['Columna 13', 'COLUMNA 13', 'Fecha', 'FECHA', 'DIA', 'Dia'];
            for (const col of possibleDateCols) {
              if (currentHeaders.includes(col)) {
                updatedRow[col] = extractDate(fechaInicio);
                break;
              }
            }
            
            const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
            for (const col of possibleStartCols) {
              if (currentHeaders.includes(col)) {
                updatedRow[col] = extractTime(fechaInicio);
                break;
              }
            }
            
            const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
            for (const col of possibleEndCols) {
              if (currentHeaders.includes(col)) {
                updatedRow[col] = extractTime(fechaFin);
                break;
              }
            }
            
            // Calcular SI/NO para "INICIO SESION 10 MINUTOS ANTES"
            const timeToMinutes = (timeStr) => {
              if (!timeStr || typeof timeStr !== 'string') return NaN;
              let s = timeStr.trim();
              s = s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
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
              return NaN;
            };
            const earlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES'];
            const horaProg = row['HORA INICIO'] || row['Hora Inicio'] || row['INICIO'] || row['inicio'] || '';
            const progMin = timeToMinutes(horaProg);
            const zoomMin = timeToMinutes(extractTime(fechaInicio));
            const inferScheduledFromZoom = (zm) => {
              if (!isFinite(zm)) return NaN;
              const minute = Math.floor(zm % 60);
              if (minute >= 45) {
                return zm - minute + 60;
              }
              return NaN;
            };
            const scheduledMin = isFinite(progMin) ? progMin : inferScheduledFromZoom(zoomMin);
            const minEarly = 5;
            const maxEarly = 10;
            const diffEarly = isFinite(scheduledMin) && isFinite(zoomMin) ? (scheduledMin - zoomMin) : NaN;
            const inicioAntes = isFinite(diffEarly) && diffEarly >= minEarly && diffEarly <= maxEarly;
            for (const col of earlyCols) {
              if (currentHeaders.includes(col)) {
                updatedRow[col] = inicioAntes ? 'SI' : 'NO';
                break;
              }
            }
            
            // Guardar duración de la grabación en FINALIZA LA CLASE (ZOOM)
            const possibleFinalizaCols = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom'];
            for (const col of possibleFinalizaCols) {
              if (currentHeaders.includes(col)) {
                updatedRow[col] = extractDuration(zoomRow);
                break;
              }
            }

            // Calcular TIEMPO EFECTIVO DICTADO y EFICIENCIA en coincidencia exacta
            const possibleWaitCols = ['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE', 'Tiempo de espera antes de iniciar la clase', 'TIEMPO DE ESPERA', 'Espera antes de iniciar'];
            const possibleEffectiveCols = ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'];
            const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada'];
            const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
            const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
            const zoomDurStr = extractDuration(zoomRow);
            for (const col of possibleTotalCols) {
              if (currentHeaders.includes(col)) {
                updatedRow[col] = zoomDurStr;
                break;
              }
            }
            const waitStr = possibleWaitCols.map(c => updatedRow[c] ?? row[c]).find(v => v && String(v).trim() !== '');
            const durationSec = durationToSeconds(zoomDurStr);
            const waitSec = durationToSeconds(waitStr);
            let effectiveSec = NaN;
            if (isFinite(durationSec) && isFinite(waitSec)) {
              effectiveSec = Math.max(durationSec - waitSec, 0);
              const effStr = secondsToHHMMSS(effectiveSec);
              for (const col of possibleEffectiveCols) {
                if (currentHeaders.includes(col)) {
                  updatedRow[col] = effStr;
                  break;
                }
              }
            }
            const progStr = possibleProgramadoCols.map(c => updatedRow[c] ?? row[c]).find(v => v && String(v).trim() !== '');
            const progSec = durationToSeconds(progStr);
            if (isFinite(durationSec) && isFinite(effectiveSec) && durationSec > 0) {
              const toleranceSec = 10 * 60;
              let numerator = effectiveSec;
              if (isFinite(progSec) && progSec > 0) {
                const threshold = Math.max(progSec - toleranceSec, 0);
                numerator = Math.max(effectiveSec - threshold, 0);
              }
              const eficiencia = Math.min(numerator / durationSec, 1);
              const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
              for (const col of possibleEficienciaCols) {
                if (currentHeaders.includes(col)) {
                  updatedRow[col] = eficienciaStr;
                  break;
                }
              }
            }
            {
              const norm = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
              const findHeader = (aliases) => {
                for (const h of currentHeaders) {
                  const hN = norm(h);
                  for (const a of aliases) { if (hN === norm(a)) return h; }
                }
                return null;
              };
              const videoFlagExact = String((zoomRow && (zoomRow['Video'] || '')) || '')
                .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
                .trim().toLowerCase();
              if (videoFlagExact === 'no') {
                const obsHeader = findHeader(['OBSERVACIÓN', 'OBSERVACION', 'Observación', 'Observacion']);
                const recHeader = findHeader(['RECOMENDACIONES', 'Recomendaciones', 'RECOMENDACION', 'Recomendacion']);
                if (obsHeader) updatedRow[obsHeader] = 'no activo camara';
                if (recHeader) updatedRow[recHeader] = 'activar su camara';
              }
            }
            
            if (!updatedRow.TURNO || updatedRow.TURNO.toString().trim() === '') {
              updatedRow.TURNO = detectTurno(fechaInicio);
            }
            
            newData[index] = updatedRow;
            sesionesUsadasGlobal.add(claveZoom);
            updatedCount++;
            
            mostrarToast(`✅ Fila actualizada:<br>
              <b>Docente:</b> ${docenteActual}<br>
              <b>Excel:</b> ${row.SECCION} / <b>Zoom:</b> ${seccionZoom}<br>
              <b>Sesión:</b> ${sesionZoom}`, "info");
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
          
          // Patrón estándar para encontrar PEAD
          const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
          
          if (!temaMatch) {
            // Intentar patrón más flexible para encontrar variantes de PEAD
            const patternFlexible = /(.+?)(?:(?:–|-|\/|:)\s*)(PEAD[-_ ]?[a-zA-Z0-9]*)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i;
            const temaMatchFlexible = zoomTema.match(patternFlexible);
            
            if (!temaMatchFlexible) continue;
            
            // Agregar a la lista de docentes con PEAD no estándar para notificación consolidada
            const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatchFlexible;
            docentesSinPEAD.add(zoomDocente);
            continue;
          }
          
          const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
          const cursoZoom = cursoParte.trim();
          const sesionZoom = sesionNumeroStr ? parseInt(sesionNumeroStr) : 0;
          const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;
          
          if (sesionesUsadasGlobal.has(claveZoom)) continue;

          // NUEVO: Autocompletar SOLO si la SECCIÓN del Zoom coincide con alguna SECCIÓN existente del mismo docente+curso
          const seccionesExistentesArr = [...new Set(
            newData
              .filter(r => r.DOCENTE === docenteActual && normalizeCursoName(r.CURSO || "") === normalizeCursoName(cursoZoom))
              .map(r => r.SECCION)
              .filter(Boolean)
          )];
          const coincideConAlgunaSeccion = seccionesExistentesArr.some(sec => matchSecciones(sec || "", seccionZoom));
          if (!coincideConAlgunaSeccion) {
            mostrarToast(`⚠️ Fila vacía no completada (sección distinta):<br>
              <b>Docente:</b> ${docenteActual}<br>
              <b>Secciones en Excel:</b> ${seccionesExistentesArr.join(', ') || '(vacías)'}<br>
              <b>Sección en Zoom:</b> ${seccionZoom}`, 'warning');
            continue;
          }
          
          const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
          const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
          
          newData[index] = updateRowWithZoom(row, {
            curso: cursoZoom,
            fecha: extractDate(fechaInicio),
            horaInicio: extractTime(fechaInicio),
            horaFin: extractTime(fechaFin),
            turno: detectTurno(fechaInicio)
          }, zoomRow);
          
          newData[index].CURSO = cursoZoom;
          newData[index].SECCION = seccionZoom;
          newData[index].SESION = sesionZoom;
          
          sesionesUsadasGlobal.add(claveZoom);
          updatedCount++;
          
          mostrarToast(`✅ Fila vacía completada:<br>
            <b>Docente:</b> ${docenteActual}<br>
            <b>Curso:</b> ${cursoZoom}<br>
            <b>Sección en Zoom:</b> ${seccionZoom}<br>
            <b>Sesión:</b> ${sesionZoom}`, "success");
          console.log(`✓ Fila vacía ${index} COMPLETADA con: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
          break;
        }
      });
      
      // Fallback adicional: emparejar por horario cuando el Tema no contiene PEAD-
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
          if (!horaProg || hasFecha) return;
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
            // NUEVO: Solo aplicar fallback por horario si podemos extraer PEAD y coincide la SECCIÓN
            const temaStr = bestZoom['Tema'] || bestZoom['Topic'] || "";
            const temaMatch = temaStr.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)/i);
            if (!temaMatch) return; // Sin PEAD en tema, no autocompletar
            const seccionZoom = temaMatch[2];
            const rowSeccion = row.SECCION || "";
            let permitir = false;
            if (rowSeccion && String(rowSeccion).trim() !== '') {
              permitir = matchSecciones(rowSeccion, seccionZoom);
            } else {
              const cursoBase = row.CURSO || extractCursoFromTema(temaStr);
              const seccionesExistentesArr = [...new Set(
                newData
                  .filter(r => r.DOCENTE === docenteActual && normalizeCursoName(r.CURSO || "") === normalizeCursoName(cursoBase || ""))
                  .map(r => r.SECCION)
                  .filter(Boolean)
              )];
              permitir = seccionesExistentesArr.some(sec => matchSecciones(sec || "", seccionZoom));
            }
            if (!permitir) return; // No hay coincidencia por SECCIÓN, no actualizar

            const fechaInicio = bestStartStr || "";
            const fechaFin = bestEndStr || "";
            const updatedRow = updateRowWithZoom(row, {
              curso: row.CURSO || extractCursoFromTema(bestZoom['Tema'] || bestZoom['Topic'] || ""),
              fecha: extractDate(fechaInicio),
              horaInicio: extractTime(fechaInicio),
              horaFin: extractTime(fechaFin),
              turno: detectTurno(fechaInicio)
            }, bestZoom);
            newData[index] = updatedRow;
            usedZoomByStart.add(bestStartStr);
            updatedCount++;
            
            mostrarToast(`✅ Coincidencia por horario:<br>
              <b>Docente:</b> ${docenteActual}<br>
              <b>Diferencia:</b> ${Math.round(bestDiff)} minutos`, "info");
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
        
        // Patrón estándar para encontrar PEAD
        const temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
        
        if (!temaMatch) {
          // Intentar patrón más flexible para encontrar variantes de PEAD
          const patternFlexible = /(.+?)(?:(?:–|-|\/|:)\s*)(PEAD[-_ ]?[a-zA-Z0-9]*)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i;
          const temaMatchFlexible = zoomTema.match(patternFlexible);
          
          if (temaMatchFlexible) {
            const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatchFlexible;
            docentesSinPEAD.add(docenteActual);
          }
          return;
        }
        
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
        
        // Restringir creación: solo si el curso YA existe en Excel para el docente y la sección coincide
        const filasMismoDocente = newData.filter(row => 
          row.DOCENTE === docenteActual && 
          normalizeCursoName(row.CURSO || "") === normalizeCursoName(cursoZoom)
        );

        if (filasMismoDocente.length === 0) {
          mostrarToast(`⚠️ Curso no encontrado en Excel (no se crea fila):<br>
            <b>Docente:</b> ${docenteActual}<br>
            <b>Curso en Zoom:</b> ${cursoZoom}`, 'warning');
          return; // No crear si el curso no existe para el docente
        }

        const seccionesExistentesArr = [...new Set(
          filasMismoDocente
            .map(row => row.SECCION)
            .filter(Boolean)
        )];

        const coincideConAlgunaSeccion = seccionesExistentesArr.some(sec => matchSecciones(sec || "", seccionZoom));

        if (!coincideConAlgunaSeccion) {
          const seccionesExistentes = seccionesExistentesArr.join(", ");
          mostrarToast(`⚠️ Sección distinta (no se crea fila):<br>
            <b>Docente:</b> ${docenteActual}<br>
            <b>Secciones en Excel:</b> ${seccionesExistentes || '(vacías)'}<br>
            <b>Sección en Zoom:</b> ${seccionZoom}`, 'warning');
          return; // No crear si la sección de Zoom no coincide con alguna existente
        }
        
        const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
        const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
        const newRow = {};
        currentHeaders.forEach(header => {
          newRow[header] = "";
        });
        
        newRow.DOCENTE = docenteActual;
        newRow.CURSO = cursoZoom;
        newRow.SECCION = seccionZoom;
        newRow.SESION = sesionZoom;
        newRow.TURNO = detectTurno(fechaInicio);
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
        
        for (const col of possibleFinalizaCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = extractDuration(zoomRow);
            break;
          }
        }

        // Calcular TIEMPO EFECTIVO DICTADO y EFICIENCIA al crear nueva fila (si existe espera y programado)
        const possibleWaitCols = ['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE', 'Tiempo de espera antes de iniciar la clase', 'TIEMPO DE ESPERA', 'Espera antes de iniciar'];
        const possibleEffectiveCols = ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'];
        const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada'];
        const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
        const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
        const zoomDurStr = extractDuration(zoomRow);
        for (const col of possibleTotalCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = zoomDurStr;
            break;
          }
        }
        const waitStr = possibleWaitCols.map(c => newRow[c]).find(v => v && String(v).trim() !== '');
        const durationSec = durationToSeconds(zoomDurStr);
        const waitSec = durationToSeconds(waitStr);
        let effectiveSec = NaN;
        if (isFinite(durationSec) && isFinite(waitSec)) {
          effectiveSec = Math.max(durationSec - waitSec, 0);
          for (const col of possibleEffectiveCols) {
            if (currentHeaders.includes(col)) {
              newRow[col] = secondsToHHMMSS(effectiveSec);
              break;
            }
          }
        }
        const progStr = possibleProgramadoCols.map(c => newRow[c]).find(v => v && String(v).trim() !== '');
        const progSec = durationToSeconds(progStr);
        if (isFinite(durationSec) && isFinite(effectiveSec) && durationSec > 0) {
          const toleranceSec = 10 * 60;
          let numerator = effectiveSec;
          if (isFinite(progSec) && progSec > 0) {
            const threshold = Math.max(progSec - toleranceSec, 0);
            numerator = Math.max(effectiveSec - threshold, 0);
          }
          const eficiencia = Math.min(numerator / durationSec, 1);
          const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
          for (const col of possibleEficienciaCols) {
            if (currentHeaders.includes(col)) {
              newRow[col] = eficienciaStr;
              break;
            }
          }
        }
        
        const timeToMinutes = (timeStr) => {
          if (!timeStr || typeof timeStr !== 'string') return NaN;
          let s = timeStr.trim();
          s = s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
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
          return NaN;
        };
        const earlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES'];
        const refRow = newData.find(r => r.DOCENTE === docenteActual && normalizeCursoName(r.CURSO || '') === normalizeCursoName(cursoZoom) && matchSecciones(r.SECCION || '', seccionZoom));
        const horaProg = (refRow && (refRow['HORA INICIO'] || refRow['Hora Inicio'] || refRow['INICIO'] || refRow['inicio'])) || '';
        const progMin = timeToMinutes(horaProg);
        const zoomMin = timeToMinutes(extractTime(fechaInicio));
        const inferScheduledFromZoom = (zm) => {
          if (!isFinite(zm)) return NaN;
          const minute = Math.floor(zm % 60);
          if (minute >= 45) {
            return zm - minute + 60;
          }
          return NaN;
        };
        const scheduledMin = isFinite(progMin) ? progMin : inferScheduledFromZoom(zoomMin);
        const minEarly = 5;
        const maxEarly = 10;
        const diffEarly = isFinite(scheduledMin) && isFinite(zoomMin) ? (scheduledMin - zoomMin) : NaN;
        const inicioAntes = isFinite(diffEarly) && diffEarly >= minEarly && diffEarly <= maxEarly;
        for (const col of earlyCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = inicioAntes ? 'SI' : 'NO';
            break;
          }
        }

        const videoFlagNew = String((zoomRow && (zoomRow['Video'] || '')) || '')
          .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
          .trim().toLowerCase();
        const norm = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
        const findHeader = (aliases) => {
          for (const h of currentHeaders) {
            const hN = norm(h);
            for (const a of aliases) { if (hN === norm(a)) return h; }
          }
          return null;
        };
        if (videoFlagNew === 'no') {
          const obsHeader = findHeader(['OBSERVACIÓN', 'OBSERVACION', 'Observación', 'Observacion']);
          const recHeader = findHeader(['RECOMENDACIONES', 'Recomendaciones', 'RECOMENDACION', 'Recomendacion']);
          if (obsHeader) newRow[obsHeader] = 'no activo camara';
          if (recHeader) newRow[recHeader] = 'activar su camara';
        }
        
        newData.push(newRow);
        sesionesUsadasGlobal.add(claveZoom);
        createdCount++;
        
        mostrarToast(`✅ Nueva fila creada:<br>
          <b>Docente:</b> ${docenteActual}<br>
          <b>Curso:</b> ${cursoZoom}<br>
          <b>Sección:</b> ${newRow.SECCION}<br>
          <b>Sesión:</b> ${sesionZoom}`, "success");
        console.log(`✓ Nueva fila realmente necesaria: ${cursoZoom} - ${newRow.SECCION} - Sesión ${sesionZoom}`);
      });
    });
    
    setData(newData);
    // Notificación agregada: docentes sin PEAD detectado
    if (docentesSinPEAD.size > 0) {
      const lista = Array.from(docentesSinPEAD);
      const preview = lista.slice(0, 6).join(', ');
      const mas = lista.length > 6 ? `, y ${lista.length - 6} más` : '';
      mostrarToast(`⚠️ No se detectó PEAD en ${lista.length} docente(s):<br><b>${preview}${mas}</b>`, 'warning');
    }
    // Notificación consolidada: diferencias en secciones
    if (seccionDiscrepancias.size > 0) {
      const ejemplos = Array.from(seccionDiscrepancias).slice(0, 3).map(s => {
        const [excel, zoom] = s.split('|||');
        return `Excel: <b>${excel}</b> / Zoom: <b>${zoom}</b>`;
      }).join('<br>');
      const mas = seccionDiscrepancias.size > 3 ? `<br>… y ${seccionDiscrepancias.size - 3} más` : '';
      mostrarToast(`⚠️ Diferencias en secciones detectadas: <b>${seccionDiscrepancias.size}</b><br>${ejemplos}${mas}`, 'warning');
    }
    mostrarToast(`✅ Proceso completado:<br>
      <b>${updatedCount}</b> filas actualizadas<br>
      <b>${createdCount}</b> filas nuevas creadas`, "success");
    alert(`✅ Completado:\n\n${updatedCount} filas autocompletadas\n${createdCount} filas nuevas creadas`);
  } catch (error) {
    mostrarToast(`❌ Error: ${error.message}`, "error");
    alert("❌ Error: " + error.message);
    console.error(error);
  } finally {
    setIsLoading(false);
    event.target.value = "";
    // Desactivar colector al finalizar
    collectSeccionDiff = null;
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
     
      let rawValue = cell.value;
      if (rawValue === null || rawValue === undefined) return "";
     
      if (typeof rawValue === 'string') {
        return rawValue.trim();
      }
     
      if (typeof rawValue === 'number') {
        return String(rawValue);
      }
     
      if (typeof rawValue === 'object') {
        if (rawValue.hyperlink && rawValue.text) {
          return String(rawValue.text).trim();
        }
       
        if (Array.isArray(rawValue.richText)) {
          return rawValue.richText.map(rt => rt.text || '').join('').trim();
        }
       
        if (rawValue.text !== undefined) {
          return String(rawValue.text).trim();
        }
       
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
        return '';
      }
     
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
   
    let bestRowNum = -1; let bestCount = 0; let bestKeywordRowNum = -1; let bestKeywordCount = 0;
    const headerKeywords = ['DOCENTE','CURSO','SECCION','SECCIÓN','SESION','HORA INICIO','HORA FIN','FECHA','DIA','COLUMNA 13'];
   
    // MÉTODO 1: Leer directamente las primeras filas SIN filtros
    for (let rowNum = 1; rowNum <= MAX_SCAN_ROWS; rowNum++) {
      console.log(`\n📋 === FILA ${rowNum} ===`);
     
      const row = worksheet.getRow(rowNum);
      const allCells = readAllCellsInRow(row, MAX_COLS);
     
      // Mostrar TODO lo que encuentra
      allCells.forEach((cellText, index) => {
        if (cellText && cellText.trim() !== '') {
          console.log(` Col ${index + 1}: "${cellText}"`);
        }
      });
     
      // Contar celdas con contenido real
      const nonEmptyCells = allCells.filter(cell => cell && cell.trim() !== '').length;
      console.log(` 📊 Total celdas con contenido: ${nonEmptyCells}`);
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
        console.log(` Probando fila ${i}...`);
        const row = worksheet.getRow(i);
        const cells = readAllCellsInRow(row, MAX_COLS);
        const nonEmpty = cells.filter(c => c && c.trim() !== '');
       
        if (nonEmpty.length > 0) {
          console.log(` ✅ Encontré ${nonEmpty.length} celdas en fila ${i}`);
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
    };
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
      let sheetHeaders = [];
      let sheetDataToExport = [];
      // Verificar si hay caché para esta hoja específica
      if (sheetDataCache[sheetIndex]) {
        // Usar datos cacheados (incluye modificaciones y filas agregadas)
        const cached = sheetDataCache[sheetIndex];
        sheetHeaders = cached.headers || [];
        sheetDataToExport = cached.data || [];
        console.log(`✅ Usando caché para hoja "${sheetName}" (índice ${sheetIndex}): ${sheetDataToExport.length} filas, ${sheetHeaders.length} headers`);
      } else {
        // Cargar datos originales desde el workbook si no hay caché
        if (workbookData) {
          const originalWorksheet = workbookData.worksheets[sheetIndex];
          const loaded = loadSheetData(originalWorksheet);
          sheetHeaders = loaded.headers || [];
          sheetDataToExport = loaded.data || [];
          console.log(`📥 Cargando datos originales para hoja "${sheetName}" (índice ${sheetIndex}): ${sheetDataToExport.length} filas, ${sheetHeaders.length} headers`);
        } else {
          // Fallback extremo (no debería pasar)
          sheetHeaders = [];
          sheetDataToExport = [];
          console.warn(`⚠️ No se pudo cargar hoja "${sheetName}" - sin workbookData`);
        }
      }
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
        const rowData = sheetHeaders.map(h => row[h] !== undefined ? row[h] : "");
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
    console.log(`✅ Exportado ${availableSheetsList.length} hojas con datos específicos por hoja.`);
    alert(`✅ ¡Archivo exportado exitosamente con ${availableSheetsList.length} hoja(s) y datos únicos por hoja!`);
  };
  const deleteRow = (index) => {
    // Mapear índice visible a índice real cuando MONITOREO reordena filas por DOCENTE
    let realIndex = index;
    const selectedSheetName = (availableSheets[selectedSheet]?.name || '').toString();
    const isMonitoreoView = selectedSheetName.toLowerCase().includes('monitoreo');
    // Cuando hay docente aleatorio, displayData está filtrado; mapear al índice real
    if (randomDocente) {
      const matchingIndices = [];
      data.forEach((r, idx) => {
        if ((r.DOCENTE ?? '') === randomDocente) matchingIndices.push(idx);
      });
      realIndex = matchingIndices[index] ?? realIndex;
    }
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
    // Mapear índice visible a índice real cuando MONITOREO reordena filas por DOCENTE
    let realIndex = rowIndex;
    const selectedSheetName = (availableSheets[selectedSheet]?.name || '').toString();
    const isMonitoreoView = selectedSheetName.toLowerCase().includes('monitoreo');
    // Cuando hay docente aleatorio, displayData está filtrado; mapear al índice real
    if (randomDocente) {
      const matchingIndices = [];
      data.forEach((r, idx) => {
        if ((r.DOCENTE ?? '') === randomDocente) matchingIndices.push(idx);
      });
      realIndex = matchingIndices[rowIndex] ?? realIndex;
    }
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

    // Recalcular en tiempo real: TIEMPO EFECTIVO DICTADO y EFICIENCIA
    const waitTimeAliases = [
      'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE',
      'Tiempo de espera antes de iniciar la clase',
      'TIEMPO DE ESPERA',
      'Espera antes de iniciar'
    ];
    const effectiveTimeAliases = [
      'TIEMPO EFECTIVO DICTADO',
      'Tiempo Efectivo Dictado',
      'TIEMPO EFECTIVO DOCENTE',
      'Tiempo efectivo docente'
    ];
    const zoomDurationAliases = [
      'FINALIZA LA CLASE (ZOOM)',
      'Hora Finalización Zoom',
      'DURACIÓN TOTAL CLASE',
      'Duración total clase'
    ];
    const programmedTimeAliases = [
      'TIEMPO PROGRAMADO',
      'Tiempo Programado',
      'DURACIÓN PROGRAMADA',
      'Duración Programada'
    ];
    const efficiencyAliases = [
      'EFICIENCIA',
      'Eficiencia',
      'INDICE EFICIENCIA',
      'Índice de Eficiencia'
    ];

    const currentHeadersList = activeTab?.currentHeaders || [];
    const editedIsWaitTime = waitTimeAliases.includes(columnName);
    const editedIsZoomDuration = zoomDurationAliases.includes(columnName);
    const editedIsProgrammedTime = programmedTimeAliases.includes(columnName);
    const editedIsEffectiveTime = effectiveTimeAliases.includes(columnName);

    if (editedIsWaitTime || editedIsZoomDuration || editedIsProgrammedTime || editedIsEffectiveTime) {
      const rowObj = newData[realIndex] || {};

      // Obtener duración total en segundos (desde Zoom o columna de duración)
      let durationSec = 0;
      for (const dCol of zoomDurationAliases) {
        if (currentHeadersList.includes(dCol)) {
          const raw = (editedIsZoomDuration && dCol === columnName) ? value : rowObj[dCol];
          if (raw !== undefined && raw !== '') {
            const sec = durationToSeconds(String(raw));
            if (sec && Number.isFinite(sec)) { durationSec = sec; break; }
          }
        }
      }

      let waitSec = NaN;
      let waitProvided = false;
      for (const wCol of waitTimeAliases) {
        if (currentHeadersList.includes(wCol)) {
          const wRaw = (editedIsWaitTime && wCol === columnName) ? value : rowObj[wCol];
          const hasVal = String(wRaw || '').trim() !== '';
          waitProvided = hasVal;
          if (hasVal) {
            const sec = durationToSeconds(String(wRaw));
            if (Number.isFinite(sec)) { waitSec = sec; }
          }
          break;
        }
      }

      if (editedIsWaitTime && !waitProvided) {
        for (const eCol of effectiveTimeAliases) {
          if (currentHeadersList.includes(eCol)) { rowObj[eCol] = ''; break; }
        }
        for (const col of efficiencyAliases) {
          if (currentHeadersList.includes(col)) { rowObj[col] = ''; break; }
        }
        setData(newData);
        return;
      }

      let effectiveSec = null;
      if (editedIsEffectiveTime) {
        const sec = durationToSeconds(String(value || ''));
        if (Number.isFinite(sec)) {
          effectiveSec = sec;
        } else {
          for (const col of efficiencyAliases) {
            if (currentHeadersList.includes(col)) { rowObj[col] = ''; break; }
          }
          setData(newData);
          return;
        }
      } else if (durationSec > 0 && waitProvided && Number.isFinite(waitSec)) {
        effectiveSec = Math.max(durationSec - waitSec, 0);
        const effectiveStr = secondsToHHMMSS(effectiveSec);
        for (const eCol of effectiveTimeAliases) {
          if (currentHeadersList.includes(eCol)) { rowObj[eCol] = effectiveStr; break; }
        }
      }

      // Recalcular eficiencia si hay tiempo programado y ya tenemos efectivo
      if (effectiveSec !== null) {
        let programmedSec = 0;
        for (const pCol of programmedTimeAliases) {
          if (currentHeadersList.includes(pCol)) {
            const pRaw = (editedIsProgrammedTime && pCol === columnName) ? value : rowObj[pCol];
            const hasP = String(pRaw || '').trim() !== '';
            if (hasP) {
              const sec = durationToSeconds(String(pRaw));
              if (Number.isFinite(sec)) { programmedSec = sec; }
            }
            break;
          }
        }

        const toleranceSec = 10 * 60;
        let numerator = effectiveSec;
        if (programmedSec > 0) {
          numerator = Math.max(effectiveSec - Math.max(programmedSec - toleranceSec, 0), 0);
        }
        const eficiencia = Math.min(numerator / durationSec, 1);
        const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
        for (const col of efficiencyAliases) {
          if (currentHeadersList.includes(col)) { rowObj[col] = eficienciaStr; break; }
        }
      }
    }

    setData(newData);
  };
  // ===== DATOS COMPUTADOS =====
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
  const uniqueModelos = useMemo(() => {
    const modelos = new Set();
    data.forEach(row => {
      if (row.MODELO && row.MODELO.trim() !== '') {
        modelos.add(row.MODELO.trim());
      }
    });
    return Array.from(modelos).sort();
  }, [data]);
  const uniqueModalidades = useMemo(() => {
    const modalidades = new Set();
    data.forEach(row => {
      if (row.MODALIDAD && row.MODALIDAD.trim() !== '') {
        modalidades.add(row.MODALIDAD.trim());
      }
    });
    return Array.from(modalidades).sort();
  }, [data]);
  const uniqueCiclos = useMemo(() => {
    const ciclos = new Set();
    data.forEach(row => {
      if (row.CICLO && row.CICLO.trim() !== '') {
        ciclos.add(row.CICLO.trim());
      }
    });
    return Array.from(ciclos).sort();
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
  // Si hay un docente aleatorio seleccionado, filtramos por ese docente
  if (randomDocente) {
    return data.filter(row => row.DOCENTE === randomDocente);
  }
  
  // MONITOREO sin docente seleccionado: agrupar por DOCENTE respetando el orden original del Excel
  if (isMonitoreoView) {
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
  return data;
}, [data, isMonitoreoView, randomDocente]);
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
            <div className="mt-8 flex flex-col items-center gap-2">
              <span className="font-semibold text-gray-700 mb-2">Descarga las plantillas:</span>
              <a href="/EJEMPLO (2).xlsx" download className="px-4 py-2 bg-green-600 text-white rounded hover:bg-green-700 font-bold shadow">Descargar plantilla Excel</a>
              <a href="/meetings_Docentes_CIS_2025_09_08_2025_09_21.csv" download className="px-4 py-2 bg-indigo-600 text-white rounded hover:bg-indigo-700 font-bold shadow">Descargar reporte Zoom</a>
              {/* Guía paso a paso debajo de los botones de descarga */}
              <div className="w-full max-w-4xl mt-6">
                <Guide />
              </div>
            </div>
          </div>
        )}

        {/* Modal para historial de backups - MOVIDO AQUÍ DENTRO */}
<BackupHistoryModal
  isOpen={isBackupModalOpen}
  onClose={() => setIsBackupModalOpen(false)}
  backups={backupHistory}
  onDownload={downloadBackup}
  onDelete={deleteBackup}
  onRestore={handleRestoreBackup}  // ✅ AGREGAR ESTA LÍNEA
/>


      </div>
    </div>    
  );
}
export default App;

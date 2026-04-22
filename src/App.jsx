import React, { useState, useMemo, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { collection, getDocs, limit, query, where } from 'firebase/firestore';
import ControlPanel from "./components/ControlPanel";
import DataTable from "./components/DataTable";
import { BackupHistoryModal, useBackupManager } from './components/BackupManager';
import Guide from './components/Guide';
import Login from './components/Login';
import TemplatesDownloadPanel from './components/TemplatesDownloadPanel';
import { db } from './utils/firebase';

const AUTH_STORAGE_KEY = 'monitoreo-auth-user';
const WORKSPACE_STORAGE_PREFIX = 'monitoreo-workspace';

const getWorkspaceStorageKey = (username) => `${WORKSPACE_STORAGE_PREFIX}:${String(username || '').trim()}`;

const getWorkspaceSnapshot = (username) => {
  try {
    const key = getWorkspaceStorageKey(username);
    const raw = sessionStorage.getItem(key);
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    const tabs = Array.isArray(parsed?.tabs) ? parsed.tabs : [];
    if (tabs.length === 0) return null;

    const activeExists = tabs.some((tab) => tab?.id === parsed?.activeTabId);
    const tabIds = tabs
      .map((tab) => Number(tab?.id))
      .filter((id) => Number.isInteger(id) && id > 0);
    const inferredNextTabId = (tabIds.length > 0 ? Math.max(...tabIds) : 0) + 1;

    return {
      tabs,
      activeTabId: activeExists ? parsed.activeTabId : (tabs[0]?.id ?? null),
      nextTabId: Number.isInteger(parsed?.nextTabId) && parsed.nextTabId > 0 ? parsed.nextTabId : inferredNextTabId,
      randomDocente: parsed?.randomDocente ? String(parsed.randomDocente) : null
    };
  } catch {
    return null;
  }
};

const getInitialWorkspaceState = () => {
  try {
    const rawAuth = sessionStorage.getItem(AUTH_STORAGE_KEY);
    const auth = rawAuth ? JSON.parse(rawAuth) : null;
    const username = auth?.username ? String(auth.username).trim() : '';
    if (!username) {
      return { tabs: [], activeTabId: null, nextTabId: 1, randomDocente: null };
    }

    const snapshot = getWorkspaceSnapshot(username);
    if (!snapshot) {
      return { tabs: [], activeTabId: null, nextTabId: 1, randomDocente: null };
    }

    return snapshot;
  } catch {
    return { tabs: [], activeTabId: null, nextTabId: 1, randomDocente: null };
  }
};

function App() {
  const initialWorkspace = getInitialWorkspaceState();
  const [authUser, setAuthUser] = useState(() => {
    try {
      const raw = sessionStorage.getItem(AUTH_STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  });
  const [isUserMenuOpen, setIsUserMenuOpen] = useState(false);
  const userMenuRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (userMenuRef.current && !userMenuRef.current.contains(event.target)) {
        setIsUserMenuOpen(false);
      }
    };

    if (isUserMenuOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isUserMenuOpen]);

  const handleLogout = async () => {
    sessionStorage.removeItem(AUTH_STORAGE_KEY);
    setAuthUser(null);
    setTabs([]);
    setActiveTabId(null);
    setNextTabId(1);
    setRandomDocente(null);
    hasAutoRestored.current = false;
    setIsUserMenuOpen(false);
  };

  const handleAuthenticated = (user) => {
    setAuthUser(user || null);

    const username = user?.username ? String(user.username).trim() : '';
    if (!username) {
      setTabs([]);
      setActiveTabId(null);
      setNextTabId(1);
      setRandomDocente(null);
      hasAutoRestored.current = false;
      return;
    }

    const snapshot = getWorkspaceSnapshot(username);
    if (snapshot) {
      setTabs(snapshot.tabs || []);
      setActiveTabId(snapshot.activeTabId ?? null);
      setNextTabId(snapshot.nextTabId || 1);
      setRandomDocente(snapshot.randomDocente ?? null);
    } else {
      setTabs([]);
      setActiveTabId(null);
      setNextTabId(1);
      setRandomDocente(null);
    }

    hasAutoRestored.current = false;
  };

  const getDisplayUserName = () => {
    if (!authUser) return '';
    if (authUser.displayName && String(authUser.displayName).trim() !== '') return String(authUser.displayName);
    const base = String(authUser.username || 'Usuario').trim();
    if (base.toUpperCase() === 'ADMIN') return 'Administrador';
    return base;
  };

  // Sistema de pestañas
  const [tabs, setTabs] = useState(initialWorkspace.tabs || []);
  const [activeTabId, setActiveTabId] = useState(initialWorkspace.activeTabId ?? null);
  const [nextTabId, setNextTabId] = useState(initialWorkspace.nextTabId || 1);
  // Obtener la pestaña activa
  const activeTab = tabs.find(tab => tab.id === activeTabId);
  // Estados de la pestaña activa (si existe)
  const data = activeTab?.data || [];
  const zoomData = activeTab?.zoomData || [];
  const isLoading = activeTab?.isLoading || false;
  const [isProcessing, setIsProcessing] = useState(false);
  const availableSheets = activeTab?.availableSheets || [];
  const selectedSheet = activeTab?.selectedSheet || 0;
  const workbookData = activeTab?.workbookData || null;
  const currentHeaders = activeTab?.currentHeaders || [];
  const [randomDocente, setRandomDocente] = useState(initialWorkspace.randomDocente ?? null);

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
} = useBackupManager(mostrarToast, authUser);

const [globalBackupHistory, setGlobalBackupHistory] = useState([]);

// Backups filtrados por usuario actual para panel de indicadores
useEffect(() => {
  const loadGlobalBackups = async () => {
    if (!authUser?.username) {
      setGlobalBackupHistory([]);
      return;
    }

    try {
      const q = query(
        collection(db, 'backups'),
        where('userId', '==', authUser.username),
        limit(100)
      );
      const snapshot = await getDocs(q);
      const allBackups = [];
      snapshot.forEach((docSnap) => {
        allBackups.push({ id: docSnap.id, ...docSnap.data() });
      });

      allBackups.sort((a, b) => new Date(b.date || 0).getTime() - new Date(a.date || 0).getTime());
      const currentWorkspaceBackup = allBackups.length > 0 ? [allBackups[0]] : [];
      setGlobalBackupHistory(currentWorkspaceBackup);
      console.log(`📊 Backup de trabajo cargado para usuario: ${authUser.username}`);
    } catch (error) {
      if (error.message && error.message.includes('index')) {
        console.warn('⚠️ Necesitas crear un índice en Firestore para backups (userId + date)');
      } else {
        console.warn('No se pudieron cargar backups del usuario:', error?.message || error);
      }
      setGlobalBackupHistory([]);
    }
  };

  loadGlobalBackups();
}, [authUser]);

// Función wrapper para saveBackup
const saveBackup = () => {
  saveBackupToStorage(data, currentHeaders, activeTab, setIsLoading);
};

const handleRestoreBackup = (backup) => {
  restoreBackup(backup, createNewTab);
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


const handleDocenteFilterChange = (docente) => {
  const value = typeof docente === 'string' ? docente.trim() : '';
  setRandomDocente(value || null);
};

  // Función para crear nueva pestaña
  const createNewTab = (fileName, initialData = {}) => {
    const headers = initialData.currentHeaders || [];
    const sheetData = initialData.sheetData
      ? Object.fromEntries(
          Object.entries(initialData.sheetData).map(([sheetIndex, sheet]) => [
            sheetIndex,
            {
              ...sheet,
              headers: sheet?.headers || []
            }
          ])
        )
      : { 0: { data: initialData.data || [], headers } };

    const newTab = {
      id: nextTabId,
      name: fileName || `Archivo ${nextTabId}`,
      data: initialData.data || [],
      zoomData: [],
      isLoading: false,
      availableSheets: initialData.availableSheets || [],
      selectedSheet: 0,
      workbookData: initialData.workbookData || null,
      currentHeaders: headers,
      // Caché por hoja
      sheetData: sheetData
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
      // Al cerrar la pestaña activa, volver a la vista de plantillas.
      setActiveTabId(null);
      setRandomDocente(null);
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
  const DOCENTE_ALIAS_GROUPS = {
    "HENRY CARRASCO": ["HENRY CARRASCO", "CARRASCO HENRY"],
    "MARIELLA DELGADO": ["MARIELLA DELGADO", "DELGADO MARIELLA"],
    "POLO MOGOLLON": ["POLO MOGOLLON", "MOGOLLON POLO"],
    "LUIS GARCIA": ["LUIS GARCIA", "GARCIA LUIS", "LUIS MARTIN GARCIA CABRERA"],
    "CRISTHIAM SANCHEZ": ["CRISTHIAM SANCHEZ", "SANCHEZ CRISTHIAM"],
    "JOSE TULLUME": ["JOSE TULLUME", "TULLUME JOSE", "ANDERSON TULLUME"],
    "IVONNE SALAZAR": ["IVONNE SALAZAR", "SALAZAR IVONNE"],
    "STEPHANY CRIOLLO": ["STEPHANY CRIOLLO", "CRIOLLO STEPHANY", "CROLLO STEPHANY"],
    "GINO GUERRERO": ["GINO GUERRERO", "GUERRERO GINO"],
    "CESAR DIAZ": ["CESAR DIAZ", "DIAZ CESAR"],
    "LEANDRO CISNEROS": ["LEANDRO CISNEROS", "CISNEROS LEANDRO"],
    "MIGUEL MAQUEN": ["MIGUEL MAQUEN", "MAQUEN MIGUEL"],
    "OMAR SANCHEZ": ["OMAR SANCHEZ", "SANCHEZ OMAR"],
    "CARLOS MEJIA": ["CARLOS MEJIA", "MEJIA CARLOS"],
    "CARMEN SANDOVAL": ["CARMEN SANDOVAL", "SANDOVAL CARMEN"],
    "EDWARD CASTANEDA": ["EDWARD CASTANEDA", "CASTANEDA EDWARD", "CASTAÑEDA EDWARD"],
    "NELSON NIETO": ["NELSON NIETO", "NIETO NELSON"],
    "JENNIE QUESADA": ["JENNIE QUESADA", "QUESADA JENNIE"],
    "JAIRO SALAZAR": ["JAIRO SALAZAR", "SALAZAR JAIRO"]
  };

  const normalizeDocenteRaw = (name) => String(name || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const normalizeDocenteName = (name) => {
    const normalized = normalizeDocenteRaw(name);
    if (!normalized) return "";
    return normalized
      .split(/\s+/)
      .filter(w => w.length > 1)
      .sort()
      .join(" ");
  };

  const getDocenteTokens = (name) => {
    const normalized = normalizeDocenteRaw(name);
    if (!normalized) return [];
    return normalized
      .split(/\s+/)
      .filter(w => w.length > 1);
  };

  const resolveDocenteAlias = (name) => {
    const source = normalizeDocenteRaw(name);
    if (!source) return "";
    const sourceWords = new Set(source.split(" ").filter(Boolean));

    for (const [canonical, variants] of Object.entries(DOCENTE_ALIAS_GROUPS)) {
      for (const variant of variants) {
        const vNorm = normalizeDocenteRaw(variant);
        const vWords = vNorm.split(" ").filter(Boolean);
        if (vWords.length === 0) continue;
        const matchesVariant = vWords.every(w => sourceWords.has(w));
        if (matchesVariant) return normalizeDocenteRaw(canonical);
      }
    }

    return "";
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
      .filter(word => word.length > 1 || /^\d+$/.test(word))
      .join(" ");
  };

  const CURSO_ALIAS_GROUPS = {
    "WORD 365": ["WORD 365", "OFIMATICA WORD 365"],
    "EXCEL 365": ["EXCEL 365", "OFIMATICA EXCEL 365", "RACIS202601 OFIMATICA EXCEL 365"],
    "DISENO CON CANVA": ["DISENO CON CANVA", "RACIS202601 DISENO CON CANVA"],
    "COMPUTACION 2": ["COMPUTACION 2"],
    "COMPUTACION 2 CIV": ["COMPUTACION 2 CIV"],
    "COMPUTACION 3": ["COMPUTACION 3", "CIS202601 COMPUTACION 3"],
    "COMPUTACION 3 CIV": ["COMPUTACION 3 CIV"],
    "COMPUTACION 3 ARQ": ["COMPUTACION 3 ARQ"],
    "EXCEL ASOCIADO": ["EXCEL ASOCIADO", "MICROSOFT EXCEL ASOCIADO"],
    "WORD ASOCIADO": ["WORD ASOCIADO", "MICROSOFT WORD ASOCIADO"],
    "AUTOCAD 3D": ["AUTOCAD 3D"],
    "AUTOCAD 2D": ["AUTOCAD 2D"],
    "POWER BI": ["POWER BI"],
    "BIZAGI": ["BIZAGI", "SOFTWARE BIZAGI"],
    "MS PROJECT": ["MS PROJECT"],
    "DISENO WEB": ["DISENO WEB"],
    "PROGRAMACION PARA INGENIERIA": ["PROGRAMACION PARA INGENIERIA", "ELS202601 PROGRAMACION PARA INGENIERIA"],
    "COMPETENCIAS DIGITALES": ["COMPETENCIAS DIGITALES", "MTR202601 COMPETENCIAS DIGITALES"]
  };

  const resolveCursoAlias = (name) => {
    const source = normalizeCursoName(name);
    if (!source) return "";
    const sourceWords = new Set(source.split(" ").filter(Boolean));

    for (const [canonical, variants] of Object.entries(CURSO_ALIAS_GROUPS)) {
      for (const variant of variants) {
        const vNorm = normalizeCursoName(variant);
        if (!vNorm) continue;

        if (source.includes(vNorm) || vNorm.includes(source)) {
          return normalizeCursoName(canonical);
        }

        const vWords = vNorm.split(" ").filter(Boolean);
        if (vWords.length === 0) continue;
        const matchesVariant = vWords.every(w => sourceWords.has(w));
        if (matchesVariant) return normalizeCursoName(canonical);
      }
    }

    return "";
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

    const aliasExcel = resolveDocenteAlias(docenteExcel);
    const aliasZoom = resolveDocenteAlias(docenteZoom);
    if (aliasExcel && aliasZoom && aliasExcel === aliasZoom) return true;
   
    if (normalizedExcel === normalizedZoom) return true;

    const tokensExcel = getDocenteTokens(docenteExcel);
    const tokensZoom = getDocenteTokens(docenteZoom);
    const shortTokens = tokensExcel.length <= tokensZoom.length ? tokensExcel : tokensZoom;
    const longTokensSet = new Set(tokensExcel.length <= tokensZoom.length ? tokensZoom : tokensExcel);

    // Caso clave: permitir nombre corto vs nombre completo (p.ej. "GARCIA LUIS" vs "LUIS MARTIN GARCIA CABRERA").
    if (shortTokens.length >= 2 && shortTokens.every(token => longTokensSet.has(token))) {
      return true;
    }
   
    const wordsExcel = normalizedExcel.split(" ");
    const wordsZoom = normalizedZoom.split(" ");
    const commonWords = wordsExcel.filter(word => wordsZoom.includes(word));
   
    return commonWords.length >= 2;
  };

  const matchCursos = (cursoExcel, cursoZoom) => {
    const normalizedExcel = normalizeCursoName(cursoExcel);
    const normalizedZoom = normalizeCursoName(cursoZoom);

    const aliasExcel = resolveCursoAlias(cursoExcel);
    const aliasZoom = resolveCursoAlias(cursoZoom);
    // Si ambos resuelven a un alias conocido, comparar por alias.
    // Permite que un alias base (sin carrera) coincida con un alias específico (con carrera).
    // Pero NUNCA permite que CIV coincida con ARQ.
    if (aliasExcel && aliasZoom) {
      if (aliasExcel === aliasZoom) return true;
      // Fallback: si uno es base y el otro es específico del mismo grupo (p.ej. "COMPUTACION 3" vs "COMPUTACION 3 CIV")
      if (aliasZoom.startsWith(aliasExcel + ' ') || aliasExcel.startsWith(aliasZoom + ' ')) return true;
      return false;
    }

    if (!normalizedExcel || !normalizedZoom) return false;
    if (normalizedExcel === normalizedZoom) return true;

    // Caso clave: si el nombre del curso del Excel aparece dentro del tema Zoom, aceptar
    // Ej.: "EXCEL 365" dentro de "RACIS202601 OFIMATICA EXCEL 365 PEAD E SESION 02"
    if (normalizedZoom.includes(normalizedExcel) || normalizedExcel.includes(normalizedZoom)) return true;

    const stripNoise = (text) => text
      .replace(/\bRACIS\d+\b/g, ' ')
      .replace(/\bOFIMATICA\b/g, ' ')
      .replace(/\bPEAD\b/g, ' ')
      .replace(/\bSESION\b/g, ' ')
      .replace(/\bSESSION\b/g, ' ')
      .replace(/\b\d{1,2}\b/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    const cleanedExcel = stripNoise(normalizedExcel);
    const cleanedZoom = stripNoise(normalizedZoom);

    if (cleanedExcel && cleanedZoom) {
      if (cleanedZoom.includes(cleanedExcel) || cleanedExcel.includes(cleanedZoom)) return true;
    }

    const wordsExcel = (cleanedExcel || normalizedExcel).split(" ").filter(w => w.length > 0);
    const wordsZoom = (cleanedZoom || normalizedZoom).split(" ").filter(w => w.length > 0);

    if (wordsExcel.length === 0 || wordsZoom.length === 0) return false;

    const commonWords = wordsExcel.filter(word => wordsZoom.includes(word));

    // Flexible: acepta coincidencia parcial fuerte para variantes de encabezado/tema
    return commonWords.length >= Math.max(1, Math.floor(wordsExcel.length * 0.6));
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
  const parseClockValueToSeconds = (input) => {
    const s = String(input || '').trim();
    if (!s) return NaN;

    const m12 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M|[ap]\.\s*m\.)$/i);
    if (m12) {
      let h = parseInt(m12[1], 10);
      const m = parseInt(m12[2], 10);
      const sec = parseInt(m12[3] || '0', 10);
      const p = m12[4].toUpperCase().replace(/\./g, '').replace(/\s+/g, '');
      if (p === 'PM' && h !== 12) h += 12;
      if (p === 'AM' && h === 12) h = 0;
      return h * 3600 + m * 60 + sec;
    }

    const m24 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m24) {
      const h = parseInt(m24[1], 10);
      const m = parseInt(m24[2], 10);
      const sec = parseInt(m24[3] || '0', 10);
      return h * 3600 + m * 60 + sec;
    }

    return NaN;
  };
  const getFirstAliasValue = (rowObj, aliases = []) => {
    let firstDefined = '';
    for (const col of aliases) {
      if (rowObj?.[col] !== undefined) {
        if (firstDefined === '') firstDefined = rowObj[col];
        const raw = rowObj[col];
        if (raw !== null && raw !== undefined && String(raw).trim() !== '') {
          return raw;
        }
      }
    }
    return firstDefined;
  };
  const getClassBaseStartSec = (rowObj) => {
    const scheduledStartAliases = ['HORA INICIO', 'Hora Inicio', 'INICIO', 'inicio'];
    const scheduledStartSec = parseClockValueToSeconds(getFirstAliasValue(rowObj, scheduledStartAliases));
    if (!Number.isFinite(scheduledStartSec)) return NaN;
    return ((scheduledStartSec % 86400) + 86400) % 86400;
  };

  const getRoundedClassHourSec = (rowObj) => {
    const scheduledStartSec = getClassBaseStartSec(rowObj);
    if (!Number.isFinite(scheduledStartSec)) return NaN;
    return Math.round(scheduledStartSec / 3600) * 3600;
  };

  const getRealStartSec = (rowObj) => {
    const scheduledStartAliases = ['HORA INICIO', 'Hora Inicio', 'INICIO', 'inicio'];
    const waitTimeAliases = [
      'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE',
      'Tiempo de espera antes de iniciar la clase',
      'TIEMPO DE ESPERA',
      'Espera antes de iniciar'
    ];
    const inicioRealAliases = ['INICIO REAL CLASE', 'Inicio Real Clase'];

    const rawScheduledStartSec = parseClockValueToSeconds(getFirstAliasValue(rowObj, scheduledStartAliases));
    const waitSec = durationToSeconds(String(getFirstAliasValue(rowObj, waitTimeAliases) || ''));
    if (Number.isFinite(rawScheduledStartSec) && Number.isFinite(waitSec)) {
      return ((rawScheduledStartSec + waitSec) % 86400 + 86400) % 86400;
    }

    const explicitRealStartSec = parseClockValueToSeconds(getFirstAliasValue(rowObj, inicioRealAliases));
    if (!Number.isFinite(explicitRealStartSec)) return NaN;
    return ((explicitRealStartSec % 86400) + 86400) % 86400;
  };
  const calculateEffectiveMetrics = ({ rowObj, durationSec, programmedSec, toleranceSec = 10 * 60 }) => {
    const waitTimeAliases = [
      'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE',
      'Tiempo de espera antes de iniciar la clase',
      'TIEMPO DE ESPERA',
      'Espera antes de iniciar'
    ];
    const inicioRealAliases = ['INICIO REAL CLASE', 'Inicio Real Clase'];

    const waitSec = durationToSeconds(String(getFirstAliasValue(rowObj, waitTimeAliases) || ''));

    // Regla de negocio: sin tiempo de espera positivo, no se calcula tiempo efectivo ni eficiencia.
    if (!Number.isFinite(waitSec) || waitSec <= 0) {
      return { effectiveSec: NaN, eficiencia: NaN };
    }

    const scheduledStartSec = getClassBaseStartSec(rowObj);
    const realStartSec = getRealStartSec(rowObj);

    if (!Number.isFinite(durationSec) || durationSec <= 0) {
      return { effectiveSec: NaN, eficiencia: NaN };
    }

    // Regla: la tolerancia se evalúa respecto a la hora de clase redondeada desde la hora programada.
    // Ejemplo: 07:55 -> base 08:00; iniciar antes de 08:00 nunca penaliza.
    const toMinuteFloor = (sec) => Math.floor(sec / 60) * 60;
    const classBaseHourSec = Number.isFinite(scheduledStartSec)
      ? Math.round(scheduledStartSec / 3600) * 3600
      : NaN;

    const effectiveRealStartForDelay = Number.isFinite(realStartSec)
      ? toMinuteFloor(realStartSec)
      : NaN;

    const lostSec = Number.isFinite(classBaseHourSec) && Number.isFinite(effectiveRealStartForDelay)
      ? Math.max(effectiveRealStartForDelay - classBaseHourSec - toleranceSec, 0)
      : 0;
    const effectiveSec = Math.max(durationSec - lostSec, 0);
    const denominator = Number.isFinite(programmedSec) && programmedSec > 0 ? programmedSec : durationSec;
    const eficiencia = denominator > 0 ? Math.min(effectiveSec / denominator, 1) : NaN;

    return { effectiveSec, eficiencia };
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
    const parsed = parseZoomTopic(tema);
    if (!parsed) return tema || "";
    return parsed.curso;
  };

  const parseZoomTopic = (tema) => {
    if (!tema) return null;

    const cleanTema = String(tema).replace(/\s+/g, ' ').trim();
    const seccionMatch = cleanTema.match(/\b(PEAD[-_ ]?[a-zA-Z0-9]+)\b/i);
    if (!seccionMatch) return null;

    const seccionRaw = String(seccionMatch[1] || '');
    const curso = String(cleanTema.slice(0, seccionMatch.index) || '').trim().replace(/[\-:–\s]+$/, '');
    let seccion = seccionRaw.toUpperCase().trim();
    seccion = seccion.replace(/[_\s]+/g, '-');
    if (seccion.startsWith('PEAD') && !seccion.startsWith('PEAD-')) {
      seccion = seccion.replace(/^PEAD/, 'PEAD-');
    }

    // Captura sesión en todo el tema, aunque no esté inmediatamente después de PEAD.
    // Ejemplos: "PEAD-a - PAC Ing. Civil Sesión 01", "PEAD-d - SEMANA 01".
    const sesionMatch = cleanTema.match(/(?:SESI[OÓ]N|SESSION|SEMANA)\s*(?:N[°º]\.?\s*)?(?:#\s*)?0*(\d{1,2})\b/i);
    const sesion = sesionMatch ? parseInt(sesionMatch[1], 10) : 0;

    // Extraer carrera desde la parte del tema DESPUÉS de PEAD para distinguir CIV vs ARQ, etc.
    const afterPead = cleanTema.slice(seccionMatch.index + seccionMatch[0].length);
    let carreraSuffix = '';
    if (/\bING\.?\s*CIV|\bCIVIL\b/i.test(afterPead)) carreraSuffix = 'CIV';
    else if (/\bARQUITECTURA\b|\bARQ\b/i.test(afterPead)) carreraSuffix = 'ARQ';
    else if (/\bING\.?\s*SIS|\bSISTEMAS\b/i.test(afterPead)) carreraSuffix = 'CIS';
    const cursoFinal = carreraSuffix ? `${curso} ${carreraSuffix}` : curso;

    return {
      curso: cursoFinal,
      seccion,
      sesion
    };
  };

  const buildZoomRowIdentity = (zoomRow) => {
    if (!zoomRow) return '';
    const host = zoomRow['Anfitrión'] || zoomRow['Host'] || '';
    const topic = zoomRow['Tema'] || zoomRow['Topic'] || '';
    const start = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || '';
    const end = zoomRow['Hora de finalización'] || zoomRow['End Time'] || '';
    return `${host}|||${topic}|||${start}|||${end}`;
  };

  const buildZoomSessionLookup = (rows) => {
    const zoomRows = Array.isArray(rows) ? rows : [];
    const groupToDateSet = new Map();

    const toDateKey = (dateTimeStr) => {
      const d = extractDate(dateTimeStr || '');
      const m = String(d || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!m) return '';
      return `${m[3]}-${m[2]}-${m[1]}`;
    };

    zoomRows.forEach((zoomRow) => {
      const tema = zoomRow['Tema'] || zoomRow['Topic'] || '';
      const parsed = parseZoomTopic(tema);
      if (!parsed) return;

      const host = zoomRow['Anfitrión'] || zoomRow['Host'] || '';
      const start = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || '';
      const dateKey = toDateKey(start);
      if (!host || !dateKey) return;

      const groupKey = `${normalizeDocenteName(host)}|||${normalizeCursoName(parsed.curso)}|||${normalizeSeccion(parsed.seccion)}`;
      if (!groupToDateSet.has(groupKey)) {
        groupToDateSet.set(groupKey, new Set());
      }
      groupToDateSet.get(groupKey).add(dateKey);
    });

    const groupDateToSession = new Map();
    groupToDateSet.forEach((dateSet, groupKey) => {
      const orderedDates = Array.from(dateSet).sort((a, b) => a.localeCompare(b));
      const byDate = new Map();
      orderedDates.forEach((dateKey, idx) => {
        byDate.set(dateKey, idx + 1);
      });
      groupDateToSession.set(groupKey, byDate);
    });

    const lookup = new Map();
    zoomRows.forEach((zoomRow) => {
      const tema = zoomRow['Tema'] || zoomRow['Topic'] || '';
      const parsed = parseZoomTopic(tema);
      if (!parsed) return;

      const explicitSession = parseInt(parsed.sesion, 10);
      let resolvedSession = Number.isFinite(explicitSession) ? explicitSession : 0;

      if (!resolvedSession || resolvedSession <= 0) {
        const host = zoomRow['Anfitrión'] || zoomRow['Host'] || '';
        const start = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || '';
        const dateKey = toDateKey(start);
        if (host && dateKey) {
          const groupKey = `${normalizeDocenteName(host)}|||${normalizeCursoName(parsed.curso)}|||${normalizeSeccion(parsed.seccion)}`;
          const sessionByDate = groupDateToSession.get(groupKey);
          const inferred = sessionByDate ? sessionByDate.get(dateKey) : 0;
          resolvedSession = inferred || 0;
        }
      }

      lookup.set(buildZoomRowIdentity(zoomRow), resolvedSession);
    });

    return lookup;
  };

  const parseZoomRowInfo = (zoomRow, sessionLookup) => {
    if (!zoomRow) return null;
    const tema = zoomRow['Tema'] || zoomRow['Topic'] || '';
    const parsed = parseZoomTopic(tema);
    if (!parsed) return null;

    const rowKey = buildZoomRowIdentity(zoomRow);
    const explicitSession = parseInt(parsed.sesion, 10);
    let sesion = Number.isFinite(explicitSession) ? explicitSession : 0;

    if ((!sesion || sesion <= 0) && sessionLookup instanceof Map) {
      const inferred = parseInt(sessionLookup.get(rowKey), 10);
      if (Number.isFinite(inferred) && inferred > 0) {
        sesion = inferred;
      }
    }

    return {
      curso: parsed.curso,
      seccion: parsed.seccion,
      sesion
    };
  };
  // ===== HANDLERS =====

  const handleAutocompletarConZoom = async (rowMode = 'regular') => {
  const normalizedRowMode = String(rowMode || 'regular').toLowerCase();
  const targetSessionCount = normalizedRowMode === 'regular' ? 16 : 12;
  const isLargeBatch = (data?.length || 0) > 80 || (zoomData?.length || 0) > 120;
  const showDetailedToasts = !isLargeBatch;
  const showDetailedLogs = !isLargeBatch;
  const notifyDetail = (message, type = 'info') => {
    if (!showDetailedToasts) return;
    mostrarToast(message, type);
  };
  const logDetail = (...args) => {
    if (!showDetailedLogs) return;
    console.log(...args);
  };
  if (data.length === 0) {
    mostrarToast(`⚠️ Primero carga el archivo Excel`, "warning");
    alert("⚠️ Primero carga el archivo Excel");
    return;
  }
  setIsProcessing(true);
  setIsLoading(true);
 
  try {
    logDetail("=== INICIANDO PROCESO COMPLETO ===");
    // Acumuladores para notificaciones agregadas
    const docentesSinPEAD = new Set();
   
    // PASO 1: Autocompletar filas existentes con datos de Zoom (si hay CSV cargado)
    let dataProcesada = [...data];
   
    const zoomSessionLookup = buildZoomSessionLookup(zoomData);
    if (zoomData.length > 0) {
      logDetail("\n📋 PASO 1: Autocompletando filas existentes con datos de Zoom");
     
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
         
          const parsedTema = parseZoomRowInfo(zoomRow, zoomSessionLookup);
          if (!parsedTema) return false;
         
          const cursoZoom = parsedTema.curso;
          const seccionZoom = parsedTema.seccion;
          const sesionZoom = parsedTema.sesion;
         
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
          const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada', 'HORAS PROGRAMADAS', 'Horas Programadas'];
          const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
          const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
          const zoomDurStr = extractDuration(sesionZoom);
          for (const col of possibleTotalCols) {
            if (currentHeaders.includes(col)) {
              dataProcesada[index][col] = zoomDurStr;
              break;
            }
          }
          const durationSec = durationToSeconds(zoomDurStr);
          const progStr = possibleProgramadoCols.map(c => dataProcesada[index][c]).find(v => v && String(v).trim() !== '');
          const progSec = durationToSeconds(progStr);
          const { effectiveSec, eficiencia } = calculateEffectiveMetrics({
            rowObj: dataProcesada[index],
            durationSec,
            programmedSec: progSec
          });
          if (Number.isFinite(effectiveSec)) {
            for (const col of possibleEffectiveCols) {
              if (currentHeaders.includes(col)) {
                dataProcesada[index][col] = secondsToHHMMSS(effectiveSec);
                break;
              }
            }
          }
          if (Number.isFinite(eficiencia)) {
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
          
          notifyDetail(`✅ Fila actualizada:<br>
            <b>Docente:</b> ${docente}<br>
            <b>Excel:</b> ${seccion} / <b>Zoom:</b> ${seccionZoomMostrar}<br>
            <b>Sesión:</b> ${sesion}`, "success");
          logDetail(` ✓ Autocompletado: ${docente} - ${curso} - Sesión ${sesion}`);
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
          const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada', 'HORAS PROGRAMADAS', 'Horas Programadas'];
          const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
          const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
          const zoomDurStr = extractDuration(bestZoom);
          setIfHasHeader(row, possibleTotalCols, zoomDurStr);
          const durationSec = durationToSeconds(zoomDurStr);
          const progStr = possibleProgramadoCols.map(c => row[c]).find(v => v && String(v).trim() !== '');
          const progSec = durationToSeconds(progStr);
          const { effectiveSec, eficiencia } = calculateEffectiveMetrics({
            rowObj: row,
            durationSec,
            programmedSec: progSec
          });
          if (Number.isFinite(effectiveSec)) {
            setIfHasHeader(row, possibleEffectiveCols, secondsToHHMMSS(effectiveSec));
          }
          if (Number.isFinite(eficiencia)) {
            const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
            setIfHasHeader(row, possibleEficienciaCols, eficienciaStr);
          }
          row.TURNO = row.TURNO && String(row.TURNO).trim() !== '' ? row.TURNO : detectTurno(fechaInicio);
          // Calcular SI/NO para "INICIO SESION 10 MINUTOS ANTES" usando hora programada original y el inicio Zoom
          const earlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES', 'INICIO SESION 10 a 5 MINUTOS ANTES', 'Inicio Sesion 10 a 5 minutos antes'];
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
          
          notifyDetail(`✅ Coincidencia por horario:<br>
            <b>Docente:</b> ${docente}<br>
            <b>Diferencia:</b> ${Math.round(bestDiff)} minutos`, "info");
        }
      });
    }
    // PASO 2: Detectar grupos únicos por DOCENTE+CURSO+SECCIÓN
    logDetail("\n📋 PASO 2: Detectando grupos únicos por DOCENTE+CURSO+SECCIÓN");
   
    const gruposPorSeccion = new Map();
   
    dataProcesada.forEach((row, originalIndex) => {
      if (!row) return;
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
    logDetail(`Total grupos detectados: ${gruposPorSeccion.size}`);
    // PASO 3: Crear exactamente la cantidad de sesiones configurada por cada grupo
    logDetail(`\n📋 PASO 3: Creando exactamente ${targetSessionCount} sesiones por cada grupo`);
    Array.from(gruposPorSeccion.entries()).forEach(([key, grupo]) => {
      const { docente, curso, seccion, primeraFila, filas, sesionesExistentes } = grupo;
     
      console.log(`\n--- ${docente} - ${curso} - ${seccion} ---`);
      console.log(` Sesiones existentes: ${Array.from(sesionesExistentes).sort((a,b) => a-b).join(', ')}`);
      console.log(` Total filas existentes: ${filas.length}`);
     
      const sesionesCompletas = [];
     
      // Crear Map con las filas existentes dentro del rango objetivo
      const existingInRange = new Map();
      filas.forEach(f => {
        const s = parseInt(String(f.SESION || 0));
        if (s >= 1 && s <= targetSessionCount && !existingInRange.has(s)) {
          existingInRange.set(s, f);
        }
      });
     
      // Si hay filas existentes pero ninguna tiene SESION en rango objetivo,
      // asignar la primera fila como SESION 1
      if (existingInRange.size === 0 && filas.length > 0) {
        const primeraFilaConDatos = filas[0];
        primeraFilaConDatos.SESION = 1;
        existingInRange.set(1, primeraFilaConDatos);
        console.log(` 📌 Primera fila asignada como SESION 1`);
      }
     
      // Crear exactamente la cantidad objetivo de sesiones
      for (let sesion = 1; sesion <= targetSessionCount; sesion++) {
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
          
          // Inicializar todos los campos con vacío para asegurar consistencia
          currentHeaders.forEach(header => {
            nuevaFila[header] = '';
          });
          
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
             
              const parsedTema = parseZoomRowInfo(zoomRow, zoomSessionLookup);
              if (!parsedTema) return false;
             
              const cursoZoom = parsedTema.curso;
              const seccionZoom = parsedTema.seccion;
              const sesionZoomNum = parsedTema.sesion;
             
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
              
              notifyDetail(`✅ Sesión ${sesion} creada con datos Zoom:<br>
                <b>Docente:</b> ${docente}<br>
                <b>Excel:</b> ${seccion} / <b>Zoom:</b> ${seccionZoomMostrar}`, "success");
              logDetail(` ✓ Sesión ${sesion}: CREADA CON DATOS ZOOM`);
            } else {
              logDetail(` + Sesión ${sesion}: CREADA (solo metadatos copiados)`);
            }
          } else {
            logDetail(` + Sesión ${sesion}: CREADA (solo metadatos copiados)`);
          }
         
          sesionesCompletas.push(nuevaFila);
        }
      }
     
      grupo.sesionesGeneradas = sesionesCompletas;
     
      const nuevasCreadas = 16 - existingInRange.size;
      console.log(` 📊 Total final para grupo: 16 sesiones exactas`);
      console.log(` 📊 Sesiones existentes mantenidas: ${existingInRange.size}`);
      console.log(` 📊 Sesiones nuevas creadas: ${nuevasCreadas}`);
    });

    console.log(`gruposPorSeccion.size: ${gruposPorSeccion.size}`);
    const resultadoFinal = [];
    Array.from(gruposPorSeccion.values()).forEach(grupo => {
      if (grupo.sesionesGeneradas) {
        resultadoFinal.push(...grupo.sesionesGeneradas);
      }
    });
    
    // Agregar filas que no pertenecen a ningún grupo (sin docente/curso/seccion)
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
              if (!matchDocente(grupo.docente, zd)) return false;
              const parsedTema = parseZoomRowInfo(z, zoomSessionLookup);
              if (!parsedTema) return false;
              return normalizeCursoName(parsedTema.curso) === normalizeCursoName(grupo.curso) &&
                matchSecciones(grupo.seccion, parsedTema.seccion) &&
                parsedTema.sesion === s;
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
      
    logDetail("=== PROCESO FINALIZADO ===");
     
  } catch (error) {
    mostrarToast(`❌ Error: ${error.message}`, "error");
    console.error("Error en proceso:", error);
    alert("❌ Error: " + error.message);
    setIsProcessing(false);
  } finally {
    setIsProcessing(false);
    setIsLoading(false);
  }
};


const handleZoomCsvUpload = async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  setIsLoading(true);
  const mostrarToast = (mensaje, tipo = 'info') => {
    const text = String(mensaje || '').toLowerCase();
    const isAutocompleteToast =
      text.includes('fila actualizada') ||
      text.includes('fila vacía completada') ||
      text.includes('coincidencia por horario') ||
      text.includes('proceso completado');

    if (!isAutocompleteToast) return;

    const toast = document.createElement('div');
    toast.style.position = 'fixed';
    toast.style.top = '20px';
    toast.style.right = '20px';
    toast.style.backgroundColor = tipo === 'error' ? '#f44336' : tipo === 'warning' ? '#ff9800' : '#4CAF50';
    toast.style.color = 'white';
    toast.style.padding = '12px 15px';
    toast.style.borderRadius = '5px';
    toast.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
    toast.style.zIndex = '10000';
    toast.style.maxWidth = '360px';
    toast.style.fontSize = '14px';
    toast.style.fontWeight = 'bold';
    toast.innerHTML = mensaje;
    document.body.appendChild(toast);
    setTimeout(() => {
      if (document.body.contains(toast)) document.body.removeChild(toast);
    }, 4000);
  };
  const alert = () => {};
  try {
    // Parser robusto para CSV: soporta comas dentro de fechas y textos entre comillas
    const arrayBuffer = await file.arrayBuffer();
    const csvWorkbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheet = csvWorkbook.Sheets[csvWorkbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '', raw: false });
    const parsedZoomData = rawRows.map((row) => {
      const normalizedRow = {};
      Object.keys(row || {}).forEach((key) => {
        normalizedRow[String(key || '').trim()] = row[key];
      });
      return normalizedRow;
    });
    const zoomSessionLookup = buildZoomSessionLookup(parsedZoomData);
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
    const allowCreateRows = false; // Solo autocompletar filas existentes
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

    const isSameDocente = (a, b) => matchDocente(a, b) || matchDocente(b, a);

    const timeToMinutes = (timeStr) => {
      if (!timeStr || typeof timeStr !== 'string') return NaN;
      let s = timeStr.trim();
      s = s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
      const m12 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i);
      if (m12) {
        let h = parseInt(m12[1]) || 0;
        const min = parseInt(m12[2]) || 0;
        const sec = parseInt(m12[3] || '0') || 0;
        const p = m12[4].toUpperCase();
        if (p === 'PM' && h !== 12) h += 12;
        if (p === 'AM' && h === 12) h = 0;
        return h * 60 + min + sec / 60;
      }
      const m24 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (m24) {
        const h = parseInt(m24[1]) || 0;
        const min = parseInt(m24[2]) || 0;
        const sec = parseInt(m24[3] || '0') || 0;
        return h * 60 + min + sec / 60;
      }
      return NaN;
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
      const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada', 'HORAS PROGRAMADAS', 'Horas Programadas'];
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
      const durationSec = durationToSeconds(zoomDurStr);
      const progStr = possibleProgramadoCols.map(c => updatedRow[c] ?? row[c]).find(v => v && String(v).trim() !== '');
      const progSec = durationToSeconds(progStr);
      const metricRow = { ...row, ...updatedRow };
      const { effectiveSec, eficiencia } = calculateEffectiveMetrics({
        rowObj: metricRow,
        durationSec,
        programmedSec: progSec
      });
      if (Number.isFinite(effectiveSec)) {
        const effStr = secondsToHHMMSS(effectiveSec);
        for (const col of possibleEffectiveCols) {
          if (currentHeaders.includes(col)) {
            updatedRow[col] = effStr;
            break;
          }
        }
      }
      // Calcular EFICIENCIA solo si hay tiempo programado y duración
      if (Number.isFinite(eficiencia)) {
        const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
        for (const col of possibleEficienciaCols) {
          if (currentHeaders.includes(col)) {
            updatedRow[col] = eficienciaStr;
            break;
          }
        }
      }
      
      const possibleEarlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES', 'INICIO SESION 10 a 5 MINUTOS ANTES', 'Inicio Sesion 10 a 5 minutos antes'];
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
        if (!isSameDocente(row.DOCENTE, docenteActual)) return;
        for (const zoomRow of parsedZoomData) {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
          const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
          
          if (!matchDocente(docenteActual, zoomDocente) || !zoomTema) continue;
          
          // Patrón estándar para encontrar PEAD
          const parsedTema = parseZoomRowInfo(zoomRow, zoomSessionLookup);
          
          if (!parsedTema) {
            docentesSinPEAD.add(docenteActual);
            continue;
          }
          
          const cursoZoom = parsedTema.curso;
          const seccionZoom = parsedTema.seccion;
          const sesionZoom = parsedTema.sesion;
          if (!sesionZoom || sesionZoom <= 0) continue;
          const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;
          
          if (sesionesUsadasGlobal.has(claveZoom)) continue;
          
          const cursoMatch = row.CURSO && matchCursos(row.CURSO, cursoZoom);
          const seccionMatch = row.SECCION && matchSecciones(row.SECCION, seccionZoom);
          const sesionMatch = row.SESION && parseInt(String(row.SESION)) === sesionZoom;
          
          if (cursoMatch && seccionMatch && sesionMatch) {
            const baseDate = extractDate(zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "");
            const sameGroup = parsedZoomData.filter(z => {
              const zDoc = z['Anfitrión'] || z['Host'] || "";
              if (!matchDocente(docenteActual, zDoc)) return false;
              const zTema = z['Tema'] || z['Topic'] || "";
              const parsedTema = parseZoomRowInfo(z, zoomSessionLookup);
              if (!parsedTema) return false;
              const zCursoNorm = normalizeCursoName(parsedTema.curso);
              const zSeccionZoom = parsedTema.seccion;
              const zSesion = parsedTema.sesion;
              if (!matchCursos(row.CURSO || "", zCursoNorm)) return false;
              if (!matchSecciones(row.SECCION || "", zSeccionZoom)) return false;
              if (parseInt(String(row.SESION || 0)) !== zSesion) return false;
              const zDate = extractDate(z['Hora de inicio'] || z['Start Time'] || "");
              return zDate === baseDate;
            });
            const t2m = (t) => {
              if (!t || typeof t !== 'string') return NaN;
              let s = t.trim();
              s = s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
              const m12 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i);
              if (m12) {
                let h = parseInt(m12[1])||0; const min = parseInt(m12[2])||0; const sec = parseInt(m12[3]||'0')||0; const p = m12[4].toUpperCase();
                if (p === 'PM' && h !== 12) h += 12; if (p === 'AM' && h === 12) h = 0;
                return h*60 + min + sec/60;
              }
              const m24 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
              if (m24) { const h = parseInt(m24[1])||0; const min = parseInt(m24[2])||0; const sec = parseInt(m24[3]||'0')||0; return h*60 + min + sec/60; }
              return NaN;
            };
            let earliestStartStr = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
            let latestEndStr = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
            let totalSec = 0;
            sameGroup.forEach(z => {
              const sStr = z['Hora de inicio'] || z['Start Time'] || "";
              const eStr = z['Hora de finalización'] || z['End Time'] || "";
              if (isFinite(t2m(extractTime(sStr))) && isFinite(t2m(extractTime(earliestStartStr))) && t2m(extractTime(sStr)) < t2m(extractTime(earliestStartStr))) {
                earliestStartStr = sStr;
              }
              if (isFinite(t2m(extractTime(eStr))) && isFinite(t2m(extractTime(latestEndStr))) && t2m(extractTime(eStr)) > t2m(extractTime(latestEndStr))) {
                latestEndStr = eStr;
              }
              const dStr = extractDuration(z);
              const dSec = durationToSeconds(dStr);
              if (isFinite(dSec)) totalSec += dSec;
            });
            const aggRow = {
              'Hora de inicio': earliestStartStr,
              'Start Time': earliestStartStr,
              'Hora de finalización': latestEndStr,
              'End Time': latestEndStr,
              'Duración (hh:mm:ss)': secondsToHHMMSS(totalSec)
            };
            
            const fechaInicio = earliestStartStr || "";
            const fechaFin = latestEndStr || "";
            
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
            const earlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES', 'INICIO SESION 10 a 5 MINUTOS ANTES', 'Inicio Sesion 10 a 5 minutos antes'];
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
                updatedRow[col] = extractDuration(aggRow);
                break;
              }
            }

            // Calcular TIEMPO EFECTIVO DICTADO y EFICIENCIA en coincidencia exacta
            const possibleWaitCols = ['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE', 'Tiempo de espera antes de iniciar la clase', 'TIEMPO DE ESPERA', 'Espera antes de iniciar'];
            const possibleEffectiveCols = ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'];
            const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada', 'HORAS PROGRAMADAS', 'Horas Programadas'];
            const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
            const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
            const zoomDurStr = extractDuration(aggRow);
            for (const col of possibleTotalCols) {
              if (currentHeaders.includes(col)) {
                updatedRow[col] = zoomDurStr;
                break;
              }
            }
            const durationSec = durationToSeconds(zoomDurStr);
            const progStr = possibleProgramadoCols.map(c => updatedRow[c] ?? row[c]).find(v => v && String(v).trim() !== '');
            const progSec = durationToSeconds(progStr);
            const metricRow = { ...row, ...updatedRow };
            const { effectiveSec, eficiencia } = calculateEffectiveMetrics({
              rowObj: metricRow,
              durationSec,
              programmedSec: progSec
            });
            if (Number.isFinite(effectiveSec)) {
              const effStr = secondsToHHMMSS(effectiveSec);
              for (const col of possibleEffectiveCols) {
                if (currentHeaders.includes(col)) {
                  updatedRow[col] = effStr;
                  break;
                }
              }
            }
            if (Number.isFinite(eficiencia)) {
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
            
            notifyDetail(`✅ Fila actualizada:<br>
              <b>Docente:</b> ${docenteActual}<br>
              <b>Excel:</b> ${row.SECCION} / <b>Zoom:</b> ${seccionZoom}<br>
              <b>Sesión:</b> ${sesionZoom}`, "info");
            logDetail(`✓ Fila ${index} AUTOCOMPLETADA: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
            break;
          }
        }
      });
      
      // Segunda pasada: Autocompletar filas vacías
      console.log("Buscando filas vacías para autocompletar...");
      
      newData.forEach((row, index) => {
        if (!isSameDocente(row.DOCENTE, docenteActual)) return;
        const hasEmptySession = !row.CURSO || row.CURSO.toString().trim() === '' ||
                               !row.SECCION || row.SECCION.toString().trim() === '' ||
                               !row.SESION || row.SESION.toString().trim() === '';
        if (!hasEmptySession) return;
        
        for (const zoomRow of parsedZoomData) {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
          const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
          
          if (!matchDocente(docenteActual, zoomDocente) || !zoomTema) continue;
          
          // Patrón estándar para encontrar PEAD
          const parsedTema = parseZoomRowInfo(zoomRow, zoomSessionLookup);
          
          if (!parsedTema) {
            continue;
          }
          
          const cursoZoom = parsedTema.curso;
          const seccionZoom = parsedTema.seccion;
          const sesionZoom = parsedTema.sesion;
          const claveZoom = `${normalizeCursoName(cursoZoom)}|||${seccionZoom.toUpperCase()}|||${sesionZoom}`;
          
          if (sesionesUsadasGlobal.has(claveZoom)) continue;

          // NUEVO: Autocompletar SOLO si la SECCIÓN del Zoom coincide con alguna SECCIÓN existente del mismo docente+curso
          const seccionesExistentesArr = [...new Set(
            newData
              .filter(r => r && isSameDocente(r.DOCENTE, docenteActual) && matchCursos(r.CURSO || "", cursoZoom))
              .map(r => r.SECCION)
              .filter(Boolean)
          )];
          const coincideConAlgunaSeccion = seccionesExistentesArr.some(sec => matchSecciones(sec || "", seccionZoom));
          if (!coincideConAlgunaSeccion) {
            notifyDetail(`⚠️ Fila vacía no completada (sección distinta):<br>
              <b>Docente:</b> ${docenteActual}<br>
              <b>Secciones en Excel:</b> ${seccionesExistentesArr.join(', ') || '(vacías)'}<br>
              <b>Sección en Zoom:</b> ${seccionZoom}`, 'warning');
            continue;
          }
          
          const baseDate = extractDate(zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "");
          const sameGroup = parsedZoomData.filter(z => {
            const zDoc = z['Anfitrión'] || z['Host'] || "";
            if (!matchDocente(docenteActual, zDoc)) return false;
            const zTema = z['Tema'] || z['Topic'] || "";
            const parsedTema = parseZoomRowInfo(z, zoomSessionLookup);
            if (!parsedTema) return false;
            const zCursoNorm = normalizeCursoName(parsedTema.curso);
            const zSeccionZoom = parsedTema.seccion;
            const zSesion = parsedTema.sesion;
            if (!matchCursos(cursoZoom, zCursoNorm)) return false;
            if (!matchSecciones(seccionZoom, zSeccionZoom)) return false;
            if (sesionZoom !== zSesion) return false;
            const zDate = extractDate(z['Hora de inicio'] || z['Start Time'] || "");
            return zDate === baseDate;
          });
          const t2m = (t) => {
            if (!t || typeof t !== 'string') return NaN;
            let s = t.trim();
            s = s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi, (m) => m.toLowerCase().includes('a') ? 'AM' : 'PM');
            const m12 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i);
            if (m12) { let h = parseInt(m12[1])||0; const min = parseInt(m12[2])||0; const sec = parseInt(m12[3]||'0')||0; const p = m12[4].toUpperCase(); if (p==='PM' && h!==12) h+=12; if (p==='AM' && h===12) h=0; return h*60+min+sec/60; }
            const m24 = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
            if (m24) { const h = parseInt(m24[1])||0; const min = parseInt(m24[2])||0; const sec = parseInt(m24[3]||'0')||0; return h*60+min+sec/60; }
            return NaN;
          };
          let earliestStartStr = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
          let latestEndStr = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
          let totalSec = durationToSeconds(extractDuration(zoomRow));
          sameGroup.forEach(z => {
            const sStr = z['Hora de inicio'] || z['Start Time'] || "";
            const eStr = z['Hora de finalización'] || z['End Time'] || "";
            if (isFinite(t2m(extractTime(sStr))) && isFinite(t2m(extractTime(earliestStartStr))) && t2m(extractTime(sStr)) < t2m(extractTime(earliestStartStr))) earliestStartStr = sStr;
            if (isFinite(t2m(extractTime(eStr))) && isFinite(t2m(extractTime(latestEndStr))) && t2m(extractTime(eStr)) > t2m(extractTime(latestEndStr))) latestEndStr = eStr;
            const dSec = durationToSeconds(extractDuration(z)); if (isFinite(dSec)) totalSec += dSec;
          });
          const aggRow = { 'Hora de inicio': earliestStartStr, 'Start Time': earliestStartStr, 'Hora de finalización': latestEndStr, 'End Time': latestEndStr, 'Duración (hh:mm:ss)': secondsToHHMMSS(totalSec) };
          const fechaInicio = earliestStartStr || "";
          const fechaFin = latestEndStr || "";
          
          newData[index] = updateRowWithZoom(row, {
            curso: cursoZoom,
            fecha: extractDate(fechaInicio),
            horaInicio: extractTime(fechaInicio),
            horaFin: extractTime(fechaFin),
            turno: detectTurno(fechaInicio)
          }, aggRow);
          
          newData[index].CURSO = cursoZoom;
          newData[index].SECCION = seccionZoom;
          newData[index].SESION = sesionZoom;
          
          sesionesUsadasGlobal.add(claveZoom);
          updatedCount++;
          
          notifyDetail(`✅ Fila vacía completada:<br>
            <b>Docente:</b> ${docenteActual}<br>
            <b>Curso:</b> ${cursoZoom}<br>
            <b>Sección en Zoom:</b> ${seccionZoom}<br>
            <b>Sesión:</b> ${sesionZoom}`, "success");
          logDetail(`✓ Fila vacía ${index} COMPLETADA con: ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}`);
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
          if (!isSameDocente(row.DOCENTE, docenteActual)) return;
          const rowSesion = parseInt(String(row.SESION || 0), 10);
          if (Number.isFinite(rowSesion) && rowSesion > 0) {
            // Si la fila ya tiene sesión, el fallback por horario solo debe aplicar a esa sesión.
            // Evita que un reporte de semana 1 rellene sesiones 2..16.
          }
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
            const parsedTema = parseZoomRowInfo(bestZoom, zoomSessionLookup);
            if (!parsedTema) return; // Sin PEAD en tema, no autocompletar
            const seccionZoom = parsedTema.seccion;
            const sesionZoom = parsedTema.sesion;
            if (Number.isFinite(rowSesion) && rowSesion > 0) {
              if (!sesionZoom || sesionZoom <= 0 || sesionZoom !== rowSesion) return;
            }
            const rowSeccion = row.SECCION || "";
            let permitir = false;
            if (rowSeccion && String(rowSeccion).trim() !== '') {
              permitir = matchSecciones(rowSeccion, seccionZoom);
            } else {
              const cursoBase = row.CURSO || extractCursoFromTema(temaStr);
              const seccionesExistentesArr = [...new Set(
                newData
                  .filter(r => isSameDocente(r.DOCENTE, docenteActual) && matchCursos(r.CURSO || "", cursoBase || ""))
                  .map(r => r.SECCION)
                  .filter(Boolean)
              )];
              permitir = seccionesExistentesArr.some(sec => matchSecciones(sec || "", seccionZoom));
            }
            if (!permitir) return; // No hay coincidencia por SECCIÓN, no actualizar

          const baseDate = extractDate(bestStartStr || "");
          const temaStr2 = bestZoom['Tema'] || bestZoom['Topic'] || "";
          const parsedTema2 = parseZoomRowInfo(bestZoom, zoomSessionLookup);
          const seccionZoom2 = parsedTema2 ? parsedTema2.seccion : "";
          const sesionZoom2 = parsedTema2 ? parsedTema2.sesion : 0;
          const sameGroup = parsedZoomData.filter(z => {
            const zDoc = z['Anfitrión'] || z['Host'] || "";
            if (!matchDocente(docenteActual, zDoc)) return false;
            const zTema = z['Tema'] || z['Topic'] || "";
            const parsedTema = parseZoomRowInfo(z, zoomSessionLookup);
            if (!parsedTema) return false;
            const zSeccionZoom = parsedTema.seccion;
            const zSesion = parsedTema.sesion;
            if (!matchSecciones(seccionZoom, zSeccionZoom)) return false;
            if (sesionZoom2 > 0 && zSesion > 0 && zSesion !== sesionZoom2) return false;
            const zDate = extractDate(z['Hora de inicio'] || z['Start Time'] || "");
            return zDate === baseDate;
          });
          const t2m = (t) => { if (!t || typeof t !== 'string') return NaN; let s=t.trim(); s=s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi,(m)=>m.toLowerCase().includes('a')?'AM':'PM'); const m12=s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i); if(m12){let h=parseInt(m12[1])||0; const min=parseInt(m12[2])||0; const sec=parseInt(m12[3]||'0')||0; const p=m12[4].toUpperCase(); if(p==='PM'&&h!==12) h+=12; if(p==='AM'&&h===12) h=0; return h*60+min+sec/60;} const m24=s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/); if(m24){ const h=parseInt(m24[1])||0; const min=parseInt(m24[2])||0; const sec=parseInt(m24[3]||'0')||0; return h*60+min+sec/60;} return NaN; };
          let earliestStartStr = bestStartStr || "";
          let latestEndStr = bestEndStr || "";
          let totalSec = durationToSeconds(extractDuration(bestZoom));
          sameGroup.forEach(z => {
            const sStr = z['Hora de inicio'] || z['Start Time'] || "";
            const eStr = z['Hora de finalización'] || z['End Time'] || "";
            if (isFinite(t2m(extractTime(sStr))) && isFinite(t2m(extractTime(earliestStartStr))) && t2m(extractTime(sStr)) < t2m(extractTime(earliestStartStr))) earliestStartStr = sStr;
            if (isFinite(t2m(extractTime(eStr))) && isFinite(t2m(extractTime(latestEndStr))) && t2m(extractTime(eStr)) > t2m(extractTime(latestEndStr))) latestEndStr = eStr;
            const dSec = durationToSeconds(extractDuration(z)); if (isFinite(dSec)) totalSec += dSec;
          });
          const aggRow = { 'Hora de inicio': earliestStartStr, 'Start Time': earliestStartStr, 'Hora de finalización': latestEndStr, 'End Time': latestEndStr, 'Duración (hh:mm:ss)': secondsToHHMMSS(totalSec) };
          const fechaInicio = earliestStartStr || "";
          const fechaFin = latestEndStr || "";
            const updatedRow = updateRowWithZoom(row, {
              curso: row.CURSO || extractCursoFromTema(bestZoom['Tema'] || bestZoom['Topic'] || ""),
              fecha: extractDate(fechaInicio),
              horaInicio: extractTime(fechaInicio),
              horaFin: extractTime(fechaFin),
              turno: detectTurno(fechaInicio)
            }, aggRow);
            newData[index] = updatedRow;
            usedZoomByStart.add(bestStartStr);
            updatedCount++;
            
            notifyDetail(`✅ Coincidencia por horario:<br>
              <b>Docente:</b> ${docenteActual}<br>
              <b>Diferencia:</b> ${Math.round(bestDiff)} minutos`, "info");
            logDetail(`✓ Fallback por horario aplicado en fila ${index} (dif ${Math.round(bestDiff)} min)`);
          }
        });
      }

      // Pasada final: SOLO completar filas existentes (sin crear nuevas)
      // Usa Docente + Curso + Seccion + Sesion para asegurar que las filas ya creadas se rellenen.
      newData.forEach((row, index) => {
        if (!isSameDocente(row.DOCENTE, docenteActual)) return;

        const rowCurso = row.CURSO || '';
        const rowSeccion = row.SECCION || row['SECCIÓN'] || '';
        const rowSesion = parseInt(String(row.SESION || 0), 10);
        if (!rowCurso || !rowSeccion || !rowSesion) return;

        const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
        const startCol = possibleStartCols.find(col => currentHeaders.includes(col));
        if (startCol && row[startCol] && String(row[startCol]).trim() !== '') return;

        const candidateZoomRows = parsedZoomData.filter((zoomRow) => {
          const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || '';
          if (!matchDocente(docenteActual, zoomDocente)) return false;

          const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || '';
          const parsedTema = parseZoomRowInfo(zoomRow, zoomSessionLookup);
          if (!parsedTema) return false;

          const zCurso = String(parsedTema.curso || '').trim();
          const zSeccion = parsedTema.seccion;
          const zSesion = parsedTema.sesion;

          if (!matchCursos(rowCurso, zCurso)) return false;
          if (!matchSecciones(rowSeccion, zSeccion)) return false;
          if (!zSesion || zSesion <= 0 || zSesion !== rowSesion) return false;
          return true;
        });

        if (candidateZoomRows.length === 0) return;

        // Selecciona un candidato estable por hora de inicio más temprana.
        let bestZoomRow = candidateZoomRows[0];
        let bestMin = Infinity;
        candidateZoomRows.forEach((z) => {
          const s = z['Hora de inicio'] || z['Start Time'] || '';
          const mins = timeToMinutes(extractTime(s));
          if (isFinite(mins) && mins < bestMin) {
            bestMin = mins;
            bestZoomRow = z;
          }
        });

        const fechaInicio = bestZoomRow['Hora de inicio'] || bestZoomRow['Start Time'] || '';
        const fechaFin = bestZoomRow['Hora de finalización'] || bestZoomRow['End Time'] || '';

        const updatedRow = updateRowWithZoom(row, {
          curso: rowCurso,
          fecha: extractDate(fechaInicio),
          horaInicio: extractTime(fechaInicio),
          horaFin: extractTime(fechaFin),
          turno: detectTurno(fechaInicio)
        }, bestZoomRow);

        newData[index] = updatedRow;
        updatedCount++;

        notifyDetail(`✅ Fila actualizada:<br>
          <b>Docente:</b> ${docenteActual}<br>
          <b>Curso:</b> ${rowCurso}<br>
          <b>Sección:</b> ${rowSeccion}<br>
          <b>Sesión:</b> ${rowSesion}`, 'success');
      });
      
      // Tercera pasada: Crear nuevas filas necesarias (opcional)
      if (allowCreateRows) {
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
          isSameDocente(row.DOCENTE, docenteActual) &&
          matchCursos(row.CURSO || "", cursoZoom) &&
          matchSecciones(row.SECCION || "", seccionZoom) &&
          parseInt(String(row.SESION || 0)) === sesionZoom
        );
        
        if (existingRow) {
          console.log(`⚠️ Ya existe fila para ${cursoZoom} - ${seccionZoom} - Sesión ${sesionZoom}. NO se crea duplicado.`);
          notifyDetail(`✅ Coincidencia encontrada:<br>
            <b>Docente:</b> ${docenteActual}<br>
            <b>Excel:</b> ${existingRow.SECCION}<br>
            <b>Zoom:</b> ${seccionZoom}`, 'info');
          sesionesUsadasGlobal.add(claveZoom);
          return;
        }
        
        // Restringir creación: solo si el curso YA existe en Excel para el docente y la sección coincide
        const filasMismoDocente = newData.filter(row => 
          isSameDocente(row.DOCENTE, docenteActual) && 
          matchCursos(row.CURSO || "", cursoZoom)
        );

        if (filasMismoDocente.length === 0) {
          notifyDetail(`⚠️ Curso no encontrado en Excel (no se crea fila):<br>
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
          notifyDetail(`⚠️ Sección distinta (no se crea fila):<br>
            <b>Docente:</b> ${docenteActual}<br>
            <b>Secciones en Excel:</b> ${seccionesExistentes || '(vacías)'}<br>
            <b>Sección en Zoom:</b> ${seccionZoom}`, 'warning');
          return; // No crear si la sección de Zoom no coincide con alguna existente
        }
        
        const baseDate = extractDate(zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "");
        const sameGroup = parsedZoomData.filter(z => {
          const zDoc = z['Anfitrión'] || z['Host'] || "";
          if (!matchDocente(docenteActual, zDoc)) return false;
          const zTema = z['Tema'] || z['Topic'] || "";
          const m = zTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
          if (!m) return false;
          const [, zCursoParte, zSeccionZoom, zSesionStr] = m;
          const zCursoNorm = normalizeCursoName(zCursoParte.trim());
          const zSesion = zSesionStr ? parseInt(zSesionStr) : 0;
          if (normalizeCursoName(cursoZoom) !== zCursoNorm) return false;
          if (seccionZoom.toUpperCase() !== zSeccionZoom.toUpperCase()) return false;
          if (sesionZoom !== zSesion) return false;
          const zDate = extractDate(z['Hora de inicio'] || z['Start Time'] || "");
          return zDate === baseDate;
        });
        const t2m = (t) => { if (!t || typeof t !== 'string') return NaN; let s=t.trim(); s=s.replace(/a\.?\s*m\.?|p\.?\s*m\.?/gi,(m)=>m.toLowerCase().includes('a')?'AM':'PM'); const m12=s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)/i); if(m12){ let h=parseInt(m12[1])||0; const min=parseInt(m12[2])||0; const sec=parseInt(m12[3]||'0')||0; const p=m12[4].toUpperCase(); if(p==='PM'&&h!==12) h+=12; if(p==='AM'&&h===12) h=0; return h*60+min+sec/60;} const m24=s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/); if(m24){ const h=parseInt(m24[1])||0; const min=parseInt(m24[2])||0; const sec=parseInt(m24[3]||'0')||0; return h*60+min+sec/60;} return NaN; };
        let earliestStartStr = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
        let latestEndStr = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";
        let totalSec = durationToSeconds(extractDuration(zoomRow));
        sameGroup.forEach(z => { const sStr = z['Hora de inicio'] || z['Start Time'] || ""; const eStr = z['Hora de finalización'] || z['End Time'] || ""; if (isFinite(t2m(extractTime(sStr))) && isFinite(t2m(extractTime(earliestStartStr))) && t2m(extractTime(sStr)) < t2m(extractTime(earliestStartStr))) earliestStartStr = sStr; if (isFinite(t2m(extractTime(eStr))) && isFinite(t2m(extractTime(latestEndStr))) && t2m(extractTime(eStr)) > t2m(extractTime(latestEndStr))) latestEndStr = eStr; const dSec = durationToSeconds(extractDuration(z)); if (isFinite(dSec)) totalSec += dSec; });
        const aggRow = { 'Hora de inicio': earliestStartStr, 'Start Time': earliestStartStr, 'Hora de finalización': latestEndStr, 'End Time': latestEndStr, 'Duración (hh:mm:ss)': secondsToHHMMSS(totalSec) };
        const fechaInicio = earliestStartStr || "";
        const fechaFin = latestEndStr || "";
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
        const possibleStartCols = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
        const possibleEndCols = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
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
            newRow[col] = extractDuration(aggRow);
            break;
          }
        }

        // Calcular TIEMPO EFECTIVO DICTADO y EFICIENCIA al crear nueva fila (si existe espera y programado)
        const possibleWaitCols = ['TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE', 'Tiempo de espera antes de iniciar la clase', 'TIEMPO DE ESPERA', 'Espera antes de iniciar'];
        const possibleEffectiveCols = ['TIEMPO EFECTIVO DICTADO', 'Tiempo Efectivo Dictado', 'TIEMPO EFECTIVO DOCENTE', 'Tiempo efectivo docente'];
        const possibleProgramadoCols = ['TIEMPO PROGRAMADO', 'Tiempo Programado', 'DURACIÓN PROGRAMADA', 'Duración Programada', 'HORAS PROGRAMADAS', 'Horas Programadas'];
        const possibleEficienciaCols = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
        const possibleTotalCols = ['DURACIÓN TOTAL CLASE', 'Duración total clase'];
        const zoomDurStr = extractDuration(aggRow);
        for (const col of possibleTotalCols) {
          if (currentHeaders.includes(col)) {
            newRow[col] = zoomDurStr;
            break;
          }
        }
        const durationSec = durationToSeconds(zoomDurStr);
        const progStr = possibleProgramadoCols.map(c => newRow[c]).find(v => v && String(v).trim() !== '');
        const progSec = durationToSeconds(progStr);
        const { effectiveSec, eficiencia } = calculateEffectiveMetrics({
          rowObj: newRow,
          durationSec,
          programmedSec: progSec
        });
        if (Number.isFinite(effectiveSec)) {
          for (const col of possibleEffectiveCols) {
            if (currentHeaders.includes(col)) {
              newRow[col] = secondsToHHMMSS(effectiveSec);
              break;
            }
          }
        }
        if (Number.isFinite(eficiencia)) {
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
        const earlyCols = ['INICIO SESION 10 MINUTOS ANTES', 'Inicio Sesion 10 minutos antes', 'INICIO SESIÓN 10 MINUTOS ANTES', 'INICIO SESION 10 a 5 MINUTOS ANTES', 'Inicio Sesion 10 a 5 minutos antes'];
        const refRow = newData.find(r => isSameDocente(r.DOCENTE, docenteActual) && matchCursos(r.CURSO || '', cursoZoom) && matchSecciones(r.SECCION || '', seccionZoom));
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
        
        notifyDetail(`✅ Nueva fila creada:<br>
          <b>Docente:</b> ${docenteActual}<br>
          <b>Curso:</b> ${cursoZoom}<br>
          <b>Sección:</b> ${newRow.SECCION}<br>
          <b>Sesión:</b> ${sesionZoom}`, "success");
        logDetail(`✓ Nueva fila realmente necesaria: ${cursoZoom} - ${newRow.SECCION} - Sesión ${sesionZoom}`);
        });
      }
    });

    // Refuerzo final: completar SOLO filas existentes que sigan sin HORA INICIO.
    // Esto cubre casos donde las pasadas por docente no encontraron match por formato de tema.
    const startHeaders = ['HORA INICIO', 'Hora Inicio', 'INICIO', 'inicio'];
    const endHeaders = ['HORA FIN', 'Hora Fin', 'FIN', 'fin'];
    const findHeaderInCurrent = (aliases) => aliases.find(h => currentHeaders.includes(h));
    const startHeader = findHeaderInCurrent(startHeaders);
    const endHeader = findHeaderInCurrent(endHeaders);

    newData.forEach((row, idx) => {
      if (!row) return;
      if (!row.DOCENTE || !row.CURSO || !(row.SECCION || row['SECCIÓN'])) return;

      const hasStart = startHeader && row[startHeader] && String(row[startHeader]).trim() !== '';
      if (hasStart) return;

      const rowSesion = parseInt(String(row.SESION || 0), 10);
      const rowSeccion = row.SECCION || row['SECCIÓN'] || '';

      const candidates = parsedZoomData.filter((z) => {
        const zDoc = z['Anfitrión'] || z['Host'] || '';
        if (!matchDocente(row.DOCENTE, zDoc)) return false;

        const zTema = z['Tema'] || z['Topic'] || '';
        const m = zTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD[-_ ]?[a-zA-Z0-9]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
        if (!m) return false;

        const [, zCursoPart, zSeccionRaw, zSesionStr] = m;
        const zCurso = String(zCursoPart || '').trim();
        const zSeccion = String(zSeccionRaw || '').replace(/[_\s]+/g, '-').toUpperCase();
        const zSesion = zSesionStr ? parseInt(zSesionStr, 10) : 0;

        if (!matchCursos(row.CURSO, zCurso)) return false;
        if (!matchSecciones(rowSeccion, zSeccion)) return false;
        if (!zSesion || zSesion <= 0 || !rowSesion || rowSesion <= 0 || zSesion !== rowSesion) return false;
        return true;
      });

      if (candidates.length === 0) return;

      let selected = candidates[0];
      let best = Infinity;
      candidates.forEach((z) => {
        const t = timeToMinutes(extractTime(z['Hora de inicio'] || z['Start Time'] || ''));
        if (isFinite(t) && t < best) {
          best = t;
          selected = z;
        }
      });

      const startStr = selected['Hora de inicio'] || selected['Start Time'] || '';
      const endStr = selected['Hora de finalización'] || selected['End Time'] || '';
      const updated = updateRowWithZoom(row, {
        curso: row.CURSO,
        fecha: extractDate(startStr),
        horaInicio: extractTime(startStr),
        horaFin: extractTime(endStr),
        turno: detectTurno(startStr)
      }, selected);

      if (startHeader) updated[startHeader] = extractTime(startStr);
      if (endHeader) updated[endHeader] = extractTime(endStr);

      newData[idx] = updated;
      updatedCount++;
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
    const creationNote = allowCreateRows
      ? `<b>${createdCount}</b> filas nuevas creadas`
      : `<b>0</b> filas nuevas creadas (desactivado)`;
    mostrarToast(`✅ Proceso completado:<br>
      <b>${updatedCount}</b> filas actualizadas<br>
      ${creationNote}`, "success");
    alert(`✅ Completado:\n\n${updatedCount} filas autocompletadas\n${allowCreateRows ? createdCount : 0} filas nuevas creadas`);
  } catch (error) {
    mostrarToast(`❌ Error: ${error.message}`, "error");
    alert("❌ Error: " + error.message);
    console.error(error);
  } finally {
    setIsLoading(false);
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
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });

      const sheetNames = workbook.SheetNames.map((name, index) => ({
        index,
        name
      }));

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const { data: loadedData, headers: sheetHeaders } = loadSheetData(worksheet);

      // Crear nueva pestaña con el archivo
      createNewTab(file.name, {
        data: loadedData,
        availableSheets: sheetNames,
        workbookData: workbook,
        currentHeaders: sheetHeaders,
        sheetData: { 0: { data: loadedData, headers: sheetHeaders } }
      });

      // Auto-guardar inmediatamente al subir archivo (silencioso)
      saveBackupToStorage(loadedData, sheetHeaders, { name: file.name }, () => {}, { silent: true });

    } catch (error) {
      alert("Error al cargar el archivo: " + error.message);
      console.error(error);
    } finally {
      if (activeTab) updateActiveTab({ isLoading: false });
      event.target.value = "";
    }
  };
  const loadSheetData = (worksheet) => {
    try {
      // Usar XLSX para convertir la hoja a JSON directamente
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });

      if (jsonData.length === 0) {
        console.log("No se encontraron datos en la hoja");
        return { data: [], headers: [] };
      }

      console.log("Datos crudos del Excel:", jsonData.slice(0, 5)); // Debug: primeras 5 filas

      // Detectar automáticamente dónde están los headers
      let headerRowIndex = 0;
      let headers = [];

      // Buscar la primera fila que tenga al menos 3 columnas no vacías
      for (let i = 0; i < Math.min(jsonData.length, 10); i++) {
        const row = jsonData[i];
        if (!row || !Array.isArray(row)) continue;

        const nonEmptyCells = row.filter(cell => String(cell || '').trim() !== '');
        if (nonEmptyCells.length >= 3) {
          // Verificar si parece una fila de headers (contiene texto descriptivo)
          const textCells = row.filter(cell =>
            String(cell || '').trim() !== '' &&
            isNaN(String(cell || '').trim()) &&
            String(cell || '').trim().length > 2
          );

          if (textCells.length >= 2) {
            headerRowIndex = i;
            headers = row.map(header => String(header || '').trim());
            console.log(`Headers detectados en fila ${i + 1}:`, headers);
            break;
          }
        }
      }

      // Si no se encontraron headers válidos, usar la primera fila
      if (headers.length === 0) {
        headerRowIndex = 0;
        headers = jsonData[0].map(header => String(header || '').trim());
        console.log("Usando primera fila como headers:", headers);
      }

      // Convertir las filas de datos a objetos
      const data = [];
      for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || !Array.isArray(row)) continue;

        const obj = {};
        let hasData = false;

        headers.forEach((header, index) => {
          const value = String(row[index] || '').trim();
          obj[header] = value;
          if (value !== '') hasData = true;
        });

        // Solo incluir filas que tengan al menos algún dato
        if (hasData) {
          data.push(obj);
        }
      }

      console.log(`Total de registros cargados: ${data.length}`);
      if (data.length > 0) {
        console.log("Primera fila cargada:", data[0]);
        console.log("Encabezados finales:", headers);
      } else {
        console.log("No se encontraron filas con datos");
      }

      return { data, headers };
    } catch (error) {
      console.error("Error en loadSheetData:", error);
      return { data: [], headers: [] };
    }
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
    const sheetName = workbookData.SheetNames[sheetIndex];
    const worksheet = workbookData.Sheets[sheetName];
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

    // Crear workbook con XLSX
    const wb = XLSX.utils.book_new();

    // Convertir datos a formato de hoja
    const wsData = [currentHeaders];
    data.forEach(row => {
      const rowData = currentHeaders.map(header => row[header] || '');
      wsData.push(rowData);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Monitoreo');

    // Generar y descargar el archivo
    const fileName = `Monitoreo_USS_${activeTab?.name ? activeTab.name.replace(/\.[^/.]+$/, "") : "datos"}_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, fileName);

    alert('✅ ¡Archivo exportado exitosamente!');
  };
  const dataRef = useRef(data);
  useEffect(() => { dataRef.current = data; }, [data]);

  const autoSaveTimerRef = useRef(null);
  const hasAutoRestored = useRef(false);

  // Guardar estado de trabajo actual para restaurarlo inmediatamente al recargar.
  useEffect(() => {
    if (!authUser?.username) return;

    const key = getWorkspaceStorageKey(authUser.username);
    if (!tabs.length) {
      sessionStorage.removeItem(key);
      return;
    }

    const snapshot = {
      tabs,
      activeTabId,
      nextTabId,
      randomDocente
    };

    try {
      sessionStorage.setItem(key, JSON.stringify(snapshot));
    } catch {
      // Si falla por límite de storage, continuamos con restauración remota.
    }
  }, [authUser, tabs, activeTabId, nextTabId, randomDocente]);

  // Auto-guardado 30 segundos después del último cambio en datos
  useEffect(() => {
    if (!activeTab || !data?.length) return;
    if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
    autoSaveTimerRef.current = setTimeout(() => {
      saveBackupToStorage(dataRef.current, activeTab?.currentHeaders || [], activeTab, () => {}, { silent: true });
    }, 30000);
    return () => { if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current); };
  }, [data]); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-restaurar al iniciar sesión si no hay pestañas abiertas
  useEffect(() => {
    if (!authUser || hasAutoRestored.current || backupHistory.length === 0 || tabs.length > 0) return;
    const latest = backupHistory[0];
    if (!latest?.data?.length) return;
    hasAutoRestored.current = true;
    const headers = latest.headers || Object.keys(latest.data[0] || {});
    createNewTab(latest.name, {
      data: latest.data,
      availableSheets: [{ index: 0, name: 'Monitoreo' }],
      workbookData: null,
      currentHeaders: headers,
      sheetData: { 0: { data: latest.data, headers } }
    });
    mostrarToast(`🔄 Sesión restaurada: <b>${latest.name}</b>`, 'info');
  }, [authUser, backupHistory]); // eslint-disable-line react-hooks/exhaustive-deps

  const deleteRow = useCallback((index) => {
    const currentData = dataRef.current;
    let realIndex = index;
    if (randomDocente) {
      const matchingIndices = [];
      currentData.forEach((r, idx) => {
        if (r && (r.DOCENTE ?? '') === randomDocente) matchingIndices.push(idx);
      });
      realIndex = matchingIndices[index] ?? realIndex;
    }
   
    const newData = currentData.filter((_, i) => i !== realIndex);
    setData(newData);
  }, [randomDocente]);

  const handleCellChange = useCallback((rowIndex, columnName, value) => {
    const currentData = dataRef.current;
    let realIndex = rowIndex;
    if (randomDocente) {
      const matchingIndices = [];
      currentData.forEach((r, idx) => {
        if (r && (r.DOCENTE ?? '') === randomDocente) matchingIndices.push(idx);
      });
      realIndex = matchingIndices[rowIndex] ?? realIndex;
    }
    const newData = [...currentData];
    // Asegurar que el objeto existe antes del spread
    const safeRow = newData[realIndex] || {};
    const rowObj = { ...safeRow, [columnName]: value };
    newData[realIndex] = rowObj;
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
      'Duración Programada',
      'HORAS PROGRAMADAS',
      'Horas Programadas'
    ];
    const efficiencyAliases = [
      'EFICIENCIA',
      'Eficiencia',
      'INDICE EFICIENCIA',
      'Índice de Eficiencia'
    ];
    const scheduledStartAliases = ['HORA INICIO', 'Hora Inicio', 'INICIO', 'inicio'];
    const inicioRealAliases = [
      'INICIO REAL CLASE',
      'Inicio Real Clase'
    ];
    const actualEndAliases = ['HORA FIN', 'Hora Fin', 'FIN', 'fin'];
    const hiAliases = ['H.I', 'H.I.', 'HI', 'H I', 'INICIO A LA HORA', 'Inicio a la Hora'];
    const hfAliases = ['H.F', 'H.F.', 'HF', 'H F', 'FIN A LA HORA', 'Fin a la Hora'];

    const getFirstCell = (aliases) => {
      let firstDefined = '';
      for (const col of aliases) {
        if (rowObj[col] !== undefined) {
          if (firstDefined === '') firstDefined = rowObj[col];
          const raw = rowObj[col];
          if (raw !== null && raw !== undefined && String(raw).trim() !== '') {
            return raw;
          }
        }
      }
      return firstDefined;
    };

    const setFirstCell = (aliases, v, allowCreate = true) => {
      for (const col of aliases) {
        if (currentHeadersList.includes(col) || Object.prototype.hasOwnProperty.call(rowObj, col)) {
          rowObj[col] = v;
          return;
        }
      }
      if (allowCreate) {
        rowObj[aliases[0]] = v;
      }
    };

    const parseClockToSeconds = (input) => {
      const s = String(input || '').trim();
      if (!s) return NaN;

      const m12 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M|[ap]\.\s*m\.)$/i);
      if (m12) {
        let h = parseInt(m12[1], 10);
        const m = parseInt(m12[2], 10);
        const sec = parseInt(m12[3] || '0', 10);
        const p = m12[4].toUpperCase().replace(/\./g, '').replace(/\s+/g, '');
        if (p === 'PM' && h !== 12) h += 12;
        if (p === 'AM' && h === 12) h = 0;
        return h * 3600 + m * 60 + sec;
      }

      const m24 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (m24) {
        const h = parseInt(m24[1], 10);
        const m = parseInt(m24[2], 10);
        const sec = parseInt(m24[3] || '0', 10);
        return h * 3600 + m * 60 + sec;
      }

      return NaN;
    };

    const formatClockFromSeconds = (totalSeconds, baseClockRaw) => {
      const normalized = ((Math.floor(totalSeconds) % 86400) + 86400) % 86400;
      const hours24 = Math.floor(normalized / 3600);
      const minutes = Math.floor((normalized % 3600) / 60);
      const seconds = normalized % 60;

      const use12h = /([AP]M|[ap]\.\s*m\.)/i.test(String(baseClockRaw || ''));
      if (use12h) {
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const h12 = (hours24 % 12) || 12;
        return `${String(h12).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')} ${ampm}`;
      }

      return `${String(hours24).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    };

    const currentHeadersList = activeTab?.currentHeaders || [];
    const editedIsWaitTime = waitTimeAliases.includes(columnName);
    const editedIsZoomDuration = zoomDurationAliases.includes(columnName);
    const editedIsProgrammedTime = programmedTimeAliases.includes(columnName);
    const editedIsEffectiveTime = effectiveTimeAliases.includes(columnName);
    const editedIsScheduledStart = scheduledStartAliases.includes(columnName);
    const editedIsInicioReal = inicioRealAliases.includes(columnName);
    const editedIsOnTimeRelated = [
      ...scheduledStartAliases,
      ...inicioRealAliases,
      ...waitTimeAliases,
      ...actualEndAliases,
      ...programmedTimeAliases
    ].includes(columnName);

    if (editedIsWaitTime || editedIsZoomDuration || editedIsProgrammedTime || editedIsEffectiveTime) {
      const rowObj = newData[realIndex] || {};
      const toleranceSec = 10 * 60;

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

      let waitProvided = false;
      for (const wCol of waitTimeAliases) {
        if (currentHeadersList.includes(wCol)) {
          const wRaw = (editedIsWaitTime && wCol === columnName) ? value : rowObj[wCol];
          const hasVal = String(wRaw || '').trim() !== '';
          waitProvided = hasVal;
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
        dataRef.current = newData;
        setData(newData);
        return;
      }

      const waitCurrentSec = durationToSeconds(String(getFirstCell(waitTimeAliases) || ''));
      const hasPositiveWait = Number.isFinite(waitCurrentSec) && waitCurrentSec > 0;

      let effectiveSec = null;
      if (editedIsEffectiveTime) {
        const sec = durationToSeconds(String(value || ''));
        if (Number.isFinite(sec) && hasPositiveWait) {
          effectiveSec = sec;
        } else {
          for (const eCol of effectiveTimeAliases) {
            if (currentHeadersList.includes(eCol)) { rowObj[eCol] = ''; break; }
          }
          for (const col of efficiencyAliases) {
            if (currentHeadersList.includes(col)) { rowObj[col] = ''; break; }
          }
          dataRef.current = newData;
          setData(newData);
          return;
        }
      } else if (durationSec > 0 && waitProvided && hasPositiveWait) {
        let programmedSec = NaN;
        for (const pCol of programmedTimeAliases) {
          if (currentHeadersList.includes(pCol)) {
            const pRaw = (editedIsProgrammedTime && pCol === columnName) ? value : rowObj[pCol];
            const sec = durationToSeconds(String(pRaw || ''));
            if (Number.isFinite(sec)) { programmedSec = sec; }
            break;
          }
        }

        const { effectiveSec: computedEffectiveSec } = calculateEffectiveMetrics({
          rowObj,
          durationSec,
          programmedSec
        });
        effectiveSec = computedEffectiveSec;

        if (!Number.isFinite(effectiveSec)) {
          effectiveSec = null;
        }
      } else if (waitProvided && !hasPositiveWait) {
        for (const eCol of effectiveTimeAliases) {
          if (currentHeadersList.includes(eCol)) { rowObj[eCol] = ''; break; }
        }
        for (const col of efficiencyAliases) {
          if (currentHeadersList.includes(col)) { rowObj[col] = ''; break; }
        }
      }

      if (effectiveSec !== null) {
        const effectiveStr = secondsToHHMMSS(effectiveSec);
        for (const eCol of effectiveTimeAliases) {
          if (currentHeadersList.includes(eCol)) { rowObj[eCol] = effectiveStr; break; }
        }
      }

      // Recalcular eficiencia si hay tiempo programado y ya tenemos efectivo
      if (effectiveSec !== null) {
        let programmedSec = NaN;
        for (const pCol of programmedTimeAliases) {
          if (currentHeadersList.includes(pCol)) {
            const pRaw = (editedIsProgrammedTime && pCol === columnName) ? value : rowObj[pCol];
            const sec = durationToSeconds(String(pRaw || ''));
            if (Number.isFinite(sec)) { programmedSec = sec; }
            break;
          }
        }

        const { eficiencia } = calculateEffectiveMetrics({
          rowObj,
          durationSec,
          programmedSec
        });
        const eficienciaStr = Number.isFinite(eficiencia) ? `${(eficiencia * 100).toFixed(2)}%` : '';
        for (const col of efficiencyAliases) {
          if (currentHeadersList.includes(col)) { rowObj[col] = eficienciaStr; break; }
        }
      }
    }

    // Calcular INICIO REAL CLASE desde HORA INICIO + TIEMPO DE ESPERA.
    if (editedIsScheduledStart || editedIsWaitTime || editedIsInicioReal) {
      const scheduledStartRaw = getFirstCell(scheduledStartAliases);
      const scheduledStartSec = parseClockToSeconds(scheduledStartRaw);
      const waitRaw = getFirstCell(waitTimeAliases);
      const waitSec = durationToSeconds(String(waitRaw || ''));

      if (Number.isFinite(scheduledStartSec) && Number.isFinite(waitSec)) {
        const realStart = formatClockFromSeconds(scheduledStartSec + waitSec, scheduledStartRaw);
        setFirstCell(inicioRealAliases, realStart, false);
      }
    }

    if (editedIsOnTimeRelated) {
      const toleranceSec = 10 * 60;

      const sessionStartSec = parseClockToSeconds(getFirstCell(scheduledStartAliases));
      const hiReferenceSec = getClassBaseStartSec(rowObj);

      if (Number.isFinite(sessionStartSec) && Number.isFinite(hiReferenceSec)) {
        setFirstCell(hiAliases, sessionStartSec <= (hiReferenceSec + toleranceSec) ? 'SI' : 'NO', false);
      }

      const actualEndSec = parseClockToSeconds(getFirstCell(actualEndAliases));
      const hfBaseStartSec = getRoundedClassHourSec(rowObj);
      let programmedSec = NaN;
      const programmedRaw = getFirstCell(programmedTimeAliases);
      if (String(programmedRaw || '').trim() !== '') {
        programmedSec = durationToSeconds(String(programmedRaw));
      }

      if (Number.isFinite(hfBaseStartSec) && Number.isFinite(actualEndSec) && Number.isFinite(programmedSec) && programmedSec > 0) {
        const scheduledEndSec = hfBaseStartSec + programmedSec;
        setFirstCell(hfAliases, actualEndSec >= scheduledEndSec ? 'SI' : 'NO', false);
      }
    }

    dataRef.current = newData;
    setData(newData);
  }, [randomDocente, activeTab]);

  // Recalcular columnas derivadas al cargar/restaurar datos, sin esperar edición manual.
  useEffect(() => {
    if (!activeTab || !Array.isArray(dataRef.current) || dataRef.current.length === 0) return;

    const headers = activeTab?.currentHeaders || [];
    const waitTimeAliases = [
      'TIEMPO DE ESPERA ANTES DE INICIAR LA CLASE',
      'Tiempo de espera antes de iniciar la clase',
      'TIEMPO DE ESPERA',
      'Espera antes de iniciar'
    ];
    const programmedTimeAliases = [
      'TIEMPO PROGRAMADO',
      'Tiempo Programado',
      'DURACIÓN PROGRAMADA',
      'Duración Programada',
      'HORAS PROGRAMADAS',
      'Horas Programadas'
    ];
    const effectiveTimeAliases = [
      'TIEMPO EFECTIVO DICTADO',
      'Tiempo Efectivo Dictado',
      'TIEMPO EFECTIVO DOCENTE',
      'Tiempo efectivo docente'
    ];
    const efficiencyAliases = [
      'EFICIENCIA',
      'Eficiencia',
      'INDICE EFICIENCIA',
      'Índice de Eficiencia'
    ];
    const durationAliases = ['DURACIÓN TOTAL CLASE', 'Duración total clase', 'FINALIZA LA CLASE (ZOOM)', 'Hora Finalización Zoom'];
    const scheduledStartAliases = ['HORA INICIO', 'Hora Inicio', 'INICIO', 'inicio'];
    const inicioRealAliases = ['INICIO REAL CLASE', 'Inicio Real Clase'];
    const actualEndAliases = ['HORA FIN', 'Hora Fin', 'FIN', 'fin'];
    const hiAliases = ['H.I', 'H.I.', 'HI', 'H I', 'INICIO A LA HORA', 'Inicio a la Hora'];
    const hfAliases = ['H.F', 'H.F.', 'HF', 'H F', 'FIN A LA HORA', 'Fin a la Hora'];

    const getFirstCell = (rowObj, aliases) => {
      let firstDefined = '';
      for (const col of aliases) {
        if (rowObj[col] !== undefined) {
          if (firstDefined === '') firstDefined = rowObj[col];
          const raw = rowObj[col];
          if (raw !== null && raw !== undefined && String(raw).trim() !== '') {
            return raw;
          }
        }
      }
      return firstDefined;
    };

    const setFirstCell = (rowObj, aliases, v, allowCreate = true) => {
      for (const col of aliases) {
        if (headers.includes(col) || Object.prototype.hasOwnProperty.call(rowObj, col)) {
          rowObj[col] = v;
          return true;
        }
      }
      if (allowCreate) {
        rowObj[aliases[0]] = v;
        return true;
      }
      return false;
    };

    const parseClockToSeconds = (input) => {
      const s = String(input || '').trim();
      if (!s) return NaN;

      const m12 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M|[ap]\.\s*m\.)$/i);
      if (m12) {
        let h = parseInt(m12[1], 10);
        const m = parseInt(m12[2], 10);
        const sec = parseInt(m12[3] || '0', 10);
        const p = m12[4].toUpperCase().replace(/\./g, '').replace(/\s+/g, '');
        if (p === 'PM' && h !== 12) h += 12;
        if (p === 'AM' && h === 12) h = 0;
        return h * 3600 + m * 60 + sec;
      }

      const m24 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (m24) {
        const h = parseInt(m24[1], 10);
        const m = parseInt(m24[2], 10);
        const sec = parseInt(m24[3] || '0', 10);
        return h * 3600 + m * 60 + sec;
      }

      return NaN;
    };

    const formatClockFromSeconds = (totalSeconds, baseClockRaw) => {
      const normalized = ((Math.floor(totalSeconds) % 86400) + 86400) % 86400;
      const hours24 = Math.floor(normalized / 3600);
      const minutes = Math.floor((normalized % 3600) / 60);
      const seconds = normalized % 60;

      const use12h = /([AP]M|[ap]\.\s*m\.)/i.test(String(baseClockRaw || ''));
      if (use12h) {
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const h12 = (hours24 % 12) || 12;
        return `${String(h12).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')} ${ampm}`;
      }

      return `${String(hours24).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    };

    let hasChanges = false;
    const toleranceSec = 10 * 60;
    const nextData = dataRef.current.map((row) => {
      if (!row || typeof row !== 'object') return row;

      const nextRow = { ...row };

      const scheduledStartRaw = getFirstCell(nextRow, scheduledStartAliases);
      const scheduledStartSecForReal = parseClockToSeconds(scheduledStartRaw);
      const waitSec = durationToSeconds(String(getFirstCell(nextRow, waitTimeAliases) || ''));

      if (Number.isFinite(scheduledStartSecForReal) && Number.isFinite(waitSec)) {
        const realStart = formatClockFromSeconds(scheduledStartSecForReal + waitSec, scheduledStartRaw);
        if (setFirstCell(nextRow, inicioRealAliases, realStart, false) && getFirstCell(row, inicioRealAliases) !== realStart) {
          hasChanges = true;
        }
      }

      const durationSec = durationToSeconds(String(getFirstCell(nextRow, durationAliases) || ''));
      const programmedSec = durationToSeconds(String(getFirstCell(nextRow, programmedTimeAliases) || ''));
      const { effectiveSec, eficiencia } = calculateEffectiveMetrics({
        rowObj: nextRow,
        durationSec,
        programmedSec
      });

      if (Number.isFinite(effectiveSec)) {
        const effectiveStr = secondsToHHMMSS(effectiveSec);
        if (setFirstCell(nextRow, effectiveTimeAliases, effectiveStr, false) && getFirstCell(row, effectiveTimeAliases) !== effectiveStr) {
          hasChanges = true;
        }
      } else {
        if (setFirstCell(nextRow, effectiveTimeAliases, '', false) && String(getFirstCell(row, effectiveTimeAliases) || '') !== '') {
          hasChanges = true;
        }
      }

      if (Number.isFinite(eficiencia)) {
        const eficienciaStr = `${(eficiencia * 100).toFixed(2)}%`;
        if (setFirstCell(nextRow, efficiencyAliases, eficienciaStr, false) && getFirstCell(row, efficiencyAliases) !== eficienciaStr) {
          hasChanges = true;
        }
      } else {
        if (setFirstCell(nextRow, efficiencyAliases, '', false) && String(getFirstCell(row, efficiencyAliases) || '') !== '') {
          hasChanges = true;
        }
      }

      const sessionStartSec = parseClockToSeconds(getFirstCell(nextRow, scheduledStartAliases));
      const hiReferenceSec = getClassBaseStartSec(nextRow);
      if (Number.isFinite(sessionStartSec) && Number.isFinite(hiReferenceSec)) {
        const hiValue = sessionStartSec <= (hiReferenceSec + toleranceSec) ? 'SI' : 'NO';
        if (setFirstCell(nextRow, hiAliases, hiValue, false) && getFirstCell(row, hiAliases) !== hiValue) {
          hasChanges = true;
        }
      }

      const actualEndSec = parseClockToSeconds(getFirstCell(nextRow, actualEndAliases));
      const hfBaseStartSec = getRoundedClassHourSec(nextRow);
      if (Number.isFinite(hfBaseStartSec) && Number.isFinite(actualEndSec) && Number.isFinite(programmedSec) && programmedSec > 0) {
        const scheduledEndSec = hfBaseStartSec + programmedSec;
        const hfValue = actualEndSec >= scheduledEndSec ? 'SI' : 'NO';
        if (setFirstCell(nextRow, hfAliases, hfValue, false) && getFirstCell(row, hfAliases) !== hfValue) {
          hasChanges = true;
        }
      }

      return nextRow;
    });

    if (hasChanges) {
      dataRef.current = nextData;
      setData(nextData);
    }
  }, [activeTab, data]);
  // ===== DATOS COMPUTADOS =====
  // Generar opciones dinámicas desde los datos
  const uniqueCursos = useMemo(() => {
    const cursos = new Set();
    data.forEach(row => {
      if (row && row.CURSO && typeof row.CURSO === 'string' && row.CURSO.trim() !== '') {
        cursos.add(row.CURSO.trim());
      }
    });
    return Array.from(cursos).sort();
  }, [data]);
  const uniqueDocentes = useMemo(() => {
    const docentes = new Set();
    data.forEach(row => {
      const docente = row?.DOCENTE;
      if (docente && typeof docente === 'string' && docente.trim() !== '') {
        docentes.add(docente.trim());
      }
    });
    return Array.from(docentes).sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
  }, [data]);
  const uniqueSecciones = useMemo(() => {
    const secciones = new Set();
    data.forEach(row => {
      if (row && row.SECCION && typeof row.SECCION === 'string' && row.SECCION.trim() !== '') {
        secciones.add(row.SECCION.trim());
      }
    });
    return Array.from(secciones).sort();
  }, [data]);
  const uniqueTurnos = useMemo(() => {
    const turnos = new Set();
    data.forEach(row => {
      if (row && row.TURNO && typeof row.TURNO === 'string' && row.TURNO.trim() !== '') {
        turnos.add(row.TURNO.trim());
      }
    });
    return Array.from(turnos).sort();
  }, [data]);
  const uniqueDias = useMemo(() => {
    const dias = new Set();
    data.forEach(row => {
      if (row && row.DIAS && typeof row.DIAS === 'string' && row.DIAS.trim() !== '') {
        dias.add(row.DIAS.trim());
      }
    });
    return Array.from(dias).sort();
  }, [data]);
  const uniqueModelos = useMemo(() => {
    const modelos = new Set();
    data.forEach(row => {
      if (row && row.MODELO && typeof row.MODELO === 'string' && row.MODELO.trim() !== '') {
        modelos.add(row.MODELO.trim());
      }
    });
    return Array.from(modelos).sort();
  }, [data]);
  const uniqueModalidades = useMemo(() => {
    const modalidades = new Set();
    data.forEach(row => {
      if (row && row.MODALIDAD && typeof row.MODALIDAD === 'string' && row.MODALIDAD.trim() !== '') {
        modalidades.add(row.MODALIDAD.trim());
      }
    });
    return Array.from(modalidades).sort();
  }, [data]);
  const uniqueCiclos = useMemo(() => {
    const ciclos = new Set();
    data.forEach(row => {
      if (row && row.CICLO && typeof row.CICLO === 'string' && row.CICLO.trim() !== '') {
        ciclos.add(row.CICLO.trim());
      }
    });
    return Array.from(ciclos).sort();
  }, [data]);
  const uniquePeriodos = useMemo(() => {
    const periodos = new Set();
    data.forEach(row => {
      if (row && row.PERIODO && typeof row.PERIODO === 'string' && row.PERIODO.trim() !== '') {
        periodos.add(row.PERIODO.trim());
      }
    });
    return Array.from(periodos).sort();
  }, [data]);
  const selectedSheetName = (availableSheets[selectedSheet]?.name || '').toString();
  const isMonitoreoView = selectedSheetName.toLowerCase().includes('monitoreo');
  const dropdownOptions = useMemo(() => ({
    MODELO: uniqueModelos.length > 0 ? uniqueModelos : [],
    MODALIDAD: uniqueModalidades.length > 0 ? uniqueModalidades : [],
    CURSO: uniqueCursos.length > 0 ? uniqueCursos : [],
    SECCION: uniqueSecciones.length > 0 ? uniqueSecciones : [],
    TURNO: uniqueTurnos.length > 0 ? uniqueTurnos : [],
    DIAS: uniqueDias.length > 0 ? uniqueDias : [],
    CICLO: uniqueCiclos.length > 0 ? uniqueCiclos : [],
    PERIODO: uniquePeriodos.length > 0 ? uniquePeriodos : []
  }), [uniqueModelos, uniqueModalidades, uniqueCursos, uniqueSecciones, uniqueTurnos, uniqueDias, uniqueCiclos, uniquePeriodos]);
  const displayData = useMemo(() => {
    console.log('Computing displayData - data length:', data?.length, 'randomDocente:', randomDocente);
    // Si hay un docente aleatorio seleccionado, filtramos por ese docente
    if (randomDocente) {
      const filtered = data.filter(row => row && row.DOCENTE === randomDocente);
      console.log('Filtered data length:', filtered.length);
      return filtered;
    }
    
    // Mostrar los datos tal cual del Excel, respetando los espacios vacíos y orden original
    console.log('Returning full data');
    return data;
  }, [data, randomDocente]);

  if (!authUser) {
    return <Login onAuthenticated={handleAuthenticated} />;
  }

  // ===== RENDER =====
  return (
    <div className="min-h-screen bg-[#f7fbff] p-4">
      <div className="max-w-full mx-auto">
        {/* SISTEMA DE PESTAÑAS */}
        <div className="bg-white rounded-t-xl shadow-lg mb-0 border border-[#d7f3fa]">
          <div className="flex items-center bg-gray-100 border-b-2 border-gray-300">
            <div className="flex items-center overflow-x-auto flex-1 min-w-0">
              {tabs.map((tab) => (
                <div
                  key={tab.id}
                  className={`flex items-center px-4 py-3 cursor-pointer border-r border-gray-300 transition-all whitespace-nowrap ${
                    activeTabId === tab.id
                      ? 'bg-white border-b-4 border-[#5a2290] font-bold text-[#5a2290]'
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
                    className="text-[#5a2290] hover:text-[#11acd3] font-bold text-xl ml-2"
                  >
                    ×
                  </button>
                </div>
              ))}
              <button
                onClick={() => document.getElementById('file-input-new-tab').click()}
                className="px-6 py-3 bg-[#11acd3] text-white hover:bg-[#5a2290] font-bold whitespace-nowrap text-sm"
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

            <div className="flex items-center gap-2 px-2" ref={userMenuRef}>
              <button
                type="button"
                onClick={() => setIsBackupModalOpen(true)}
                className="inline-flex items-center gap-2 bg-[#11acd3] border border-[#11acd3] rounded-lg px-3 py-2 text-sm font-semibold text-white shadow hover:bg-[#0f9bbf]"
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
                </svg>
                Ver archivos guardados
              </button>

              <div className="relative">
                <button
                  type="button"
                  onClick={() => setIsUserMenuOpen(prev => !prev)}
                  className="inline-flex items-center gap-2 bg-white border border-[#11acd3] rounded-lg px-3 py-2 text-sm font-semibold text-[#5a2290] shadow hover:bg-[#e8f6fb]"
                >
                  <span>{getDisplayUserName()}</span>
                  <svg className={`w-4 h-4 transition-transform ${isUserMenuOpen ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7" />
                  </svg>
                </button>

                {isUserMenuOpen && (
                  <div className="absolute right-0 mt-2 w-52 bg-white border border-[#11acd3] rounded-lg shadow-lg z-50 overflow-hidden">
                    <div className="px-3 py-2 border-b border-[#d7f3fa] text-xs text-[#5a2290] bg-[#f3ecfb]">Sesión activa</div>
                    <button
                      type="button"
                      onClick={handleLogout}
                      className="w-full text-left px-3 py-2 text-sm font-semibold text-[#5a2290] hover:bg-[#e8f6fb]"
                    >
                      Cerrar sesión
                    </button>
                  </div>
                )}
              </div>
            </div>
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
      docenteOptions={uniqueDocentes}
      selectedDocente={randomDocente || ''}
      onDocenteFilterChange={handleDocenteFilterChange}
      onSaveBackup={saveBackup}
    />
    <DataTable
      data={displayData}
      headers={currentHeaders.length > 0 ? currentHeaders : []}
      dropdownOptions={dropdownOptions}
      onCellChange={handleCellChange}
      onDeleteRow={deleteRow}
      isProcessing={isProcessing}
    />
  </>
) : (
  <TemplatesDownloadPanel backupHistory={globalBackupHistory} />
)}

{/* Modal para historial de backups - MOVIDO AQUÍ DENTRO */}
<BackupHistoryModal
  isOpen={isBackupModalOpen}
  onClose={() => setIsBackupModalOpen(false)}
  backups={backupHistory}
  onDownload={downloadBackup}
  onDelete={deleteBackup}
  onRestore={handleRestoreBackup}
/>


      </div>
    </div>    
  );
}
export default App;

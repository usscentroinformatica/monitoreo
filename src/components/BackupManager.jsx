// src/components/BackupManager.jsx
import React, { useState, useEffect } from 'react';
import { db } from '../utils/firebase';
import { collection, addDoc, deleteDoc, doc, getDocs, query, orderBy, limit } from 'firebase/firestore';

// ========================================
// COMPONENTE: Modal de Historial de Backups
// ========================================
export const BackupHistoryModal = ({ isOpen, onClose, backups, onDownload, onDelete, onRestore }) => {
  if (!isOpen) return null;
  
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-6xl max-h-[90vh] flex flex-col">
        {/* Header */}
        <div className="p-6 bg-gradient-to-r from-blue-900 to-blue-700 text-white rounded-t-xl flex justify-between items-center">
          <div>
            <h2 className="text-2xl font-bold">📂 Historial de Archivos Guardados</h2>
            <p className="text-blue-200 text-sm mt-1">
              Gestiona tus copias de seguridad y restaura versiones anteriores
            </p>
          </div>
          <button 
            onClick={onClose}
            className="text-white hover:text-red-300 transition-colors"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
        
        {/* Content */}
        <div className="overflow-y-auto p-6 flex-1 bg-gray-50">
          {backups.length === 0 ? (
            <div className="text-center text-gray-500 py-12">
              <div className="text-6xl mb-4">📦</div>
              <p className="text-xl font-semibold mb-2">No hay archivos guardados</p>
              <p className="text-sm text-gray-400">
                Los archivos que guardes aparecerán aquí para que puedas restaurarlos cuando quieras
              </p>
            </div>
          ) : (
            <div className="grid gap-6">
              {backups.map((backup) => {
                const date = new Date(backup.date);
                const formattedDate = date.toLocaleDateString('es-ES', {
                  year: 'numeric',
                  month: 'long',
                  day: 'numeric'
                });
                const formattedTime = date.toLocaleTimeString('es-ES', {
                  hour: '2-digit',
                  minute: '2-digit'
                });
                
                return (
                  <div 
                    key={backup.id}
                    className="border-2 border-gray-200 rounded-xl overflow-hidden hover:shadow-xl transition-all duration-300 bg-white"
                  >
                    {/* Header de la tarjeta */}
                    <div className="bg-gradient-to-r from-blue-50 to-indigo-50 p-4 border-b-2 border-blue-100">
                      <div className="flex justify-between items-start">
                        <div className="flex-1">
                          <h3 className="font-bold text-xl text-blue-900 mb-1">
                            📄 {backup.name}
                          </h3>
                          <div className="flex gap-4 text-sm text-gray-600">
                            <span className="flex items-center gap-1">
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
                              </svg>
                              {formattedDate}
                            </span>
                            <span className="flex items-center gap-1">
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
                              </svg>
                              {formattedTime}
                            </span>
                            <span className="flex items-center gap-1 font-semibold text-blue-600">
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                              </svg>
                              {backup.rowCount} filas
                            </span>
                          </div>
                        </div>
                        
                        {/* Botones de acción */}
                        <div className="flex gap-2">
                          <button
                            onClick={() => onRestore(backup)}
                            className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors flex items-center gap-2 font-semibold shadow-md"
                            title="Restaurar este archivo"
                          >
                            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                            </svg>
                            Abrir
                          </button>
                          
                          <button
                            onClick={() => onDownload(backup)}
                            className="bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 transition-colors flex items-center gap-2 font-semibold shadow-md"
                            title="Descargar archivo"
                          >
                            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                            </svg>
                            Descargar
                          </button>
                          
                          <button
                            onClick={() => {
                              if (window.confirm('¿Estás seguro de eliminar esta copia de seguridad?')) {
                                onDelete(backup.id);
                              }
                            }}
                            className="bg-red-500 text-white px-4 py-2 rounded-lg hover:bg-red-600 transition-colors flex items-center gap-2 font-semibold shadow-md"
                            title="Eliminar"
                          >
                            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                            </svg>
                            Eliminar
                          </button>
                        </div>
                      </div>
                    </div>
                    
                    {/* Vista previa de datos */}
                    {backup.preview && backup.preview.length > 0 && (
                      <div className="p-4 bg-gray-50">
                        <div className="text-xs text-gray-500 mb-2 font-semibold uppercase tracking-wide">
                          📊 Vista previa (primeras 3 filas):
                        </div>
                        <div className="overflow-x-auto rounded-lg border border-gray-200">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="bg-gray-200">
                                {Object.keys(backup.preview[0]).slice(0, 6).map((key, i) => (
                                  <th key={i} className="p-2 text-left text-xs text-gray-700 font-bold uppercase tracking-wide">
                                    {key}
                                  </th>
                                ))}
                                {Object.keys(backup.preview[0]).length > 6 && (
                                  <th className="p-2 text-left text-xs text-gray-700 font-bold">...</th>
                                )}
                              </tr>
                            </thead>
                            <tbody>
                              {backup.preview.map((row, rowIndex) => (
                                <tr key={`${backup.id}-preview-row-${rowIndex}`} className={rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                                  {Object.entries(row).slice(0, 6).map(([key, value], i) => (
                                    <td key={i} className="p-2 text-xs border-t border-gray-200">
                                      <div className="truncate max-w-xs" title={String(value || '')}>
                                        {String(value || '').substring(0, 30)}
                                        {String(value || '').length > 30 ? '...' : ''}
                                      </div>
                                    </td>
                                  ))}
                                  {Object.keys(row).length > 6 && (
                                    <td className="p-2 text-xs border-t border-gray-200 text-gray-400">...</td>
                                  )}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
        
        {/* Footer */}
        <div className="p-6 border-t-2 border-gray-200 bg-white rounded-b-xl">
          <div className="flex justify-between items-center">
            <span className="text-sm text-gray-600 font-semibold">
              📦 Total: {backups.length} {backups.length === 1 ? 'archivo guardado' : 'archivos guardados'}
            </span>
            <button
              onClick={onClose}
              className="px-6 py-2 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded-lg font-semibold transition-colors"
            >
              Cerrar
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ========================================
// HOOK: useBackupManager
// ========================================
export const useBackupManager = (ExcelJS, mostrarToast) => {
  const [backupHistory, setBackupHistory] = useState([]);
  const [isBackupModalOpen, setIsBackupModalOpen] = useState(false);

  // Cargar backups desde Firebase al inicio
  useEffect(() => {
    const loadBackups = async () => {
      try {
        const backupsCollection = collection(db, "backups");
        const backupsQuery = query(backupsCollection, orderBy("date", "desc"), limit(50));
        const querySnapshot = await getDocs(backupsQuery);
        
        const loadedBackups = [];
        querySnapshot.forEach((doc) => {
          loadedBackups.push({
            id: doc.id,
            ...doc.data()
          });
        });
        
        // ✅ Filtrar duplicados por ID para evitar warning de keys
        const uniqueBackups = loadedBackups.filter((backup, index, self) =>
          index === self.findIndex(b => b.id === backup.id)
        );
        
        setBackupHistory(uniqueBackups);
        console.log(`✅ ${uniqueBackups.length} backups únicos cargados desde Firebase`);
      } catch (error) {
        if (
          error.code === 'permission-denied' ||
          (error.message && error.message.includes('Missing or insufficient permissions'))
        ) {
          console.warn('⚠️ No tienes permisos para cargar copias de seguridad.');
        } else {
          console.error("❌ Error loading backups:", error);
          mostrarToast("❌ Error al cargar copias de seguridad", "error");
        }
      }
    };
    
    loadBackups();
  }, [mostrarToast]);

  // ========================================
  // FUNCIÓN: Guardar Backup en Firebase (solo Firestore, sin Storage)
  // ========================================
  const saveBackup = async (data, currentHeaders, activeTab, setIsLoading) => {
    if (!data || data.length === 0) {
      mostrarToast('❌ No hay datos para guardar', 'error');
      return;
    }

    try {
      setIsLoading(true);
      mostrarToast('📤 Guardando archivo en Firebase...', 'info');
      
      // 1. Preparar datos completos para guardar (incluyendo headers como referencia)
      const backupData = {
        name: activeTab?.name || `Archivo ${Date.now()}`,
        date: new Date().toISOString(),
        rowCount: data.length,
        headers: currentHeaders,  // Guardar headers para reconstruir Excel
        data: data,  // Guardar TODOS los datos (array de objetos)
        preview: data.slice(0, 3).map(row => {
          const preview = {};
          const keyColumns = [
            'DOCENTE', 'CURSO', 'SECCION', 'SESION', 'FECHA', 
            'DIA', 'HORA INICIO', 'HORA FIN', 'MODALIDAD', 'MODELO',
            'TURNO'
          ];
          
          keyColumns.forEach(col => {
            if (row[col] !== undefined) {
              preview[col] = row[col];
            }
          });
          
          if (Object.keys(preview).length < 6) {
            Object.keys(row).slice(0, 6 - Object.keys(preview).length).forEach(key => {
              if (!preview[key]) {
                preview[key] = row[key];
              }
            });
          }
          
          return preview;
        })
      };
      
      console.log('📝 Guardando datos en Firestore...');
      const docRef = await addDoc(collection(db, "backups"), backupData);
      console.log('✅ Backup guardado con ID:', docRef.id);
      
      // 2. Actualizar estado local
      const newBackup = {
        id: docRef.id,
        ...backupData
      };
      
      setBackupHistory(prev => [newBackup, ...prev]);
      
      mostrarToast(`✅ ¡Copia guardada exitosamente!<br><b>${backupData.name}</b><br>${data.length} filas guardadas en Firebase`, 'success');
    } catch (error) {
      console.error('❌ Error detallado al guardar:', error);
      mostrarToast(`❌ Error al guardar: ${error.message}`, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // ========================================
  // FUNCIÓN: Descargar Backup (generar Excel localmente)
  // ========================================
  const downloadBackup = async (backup) => {
    try {
      if (!backup.data || backup.data.length === 0) {
        mostrarToast('❌ No hay datos para descargar', 'error');
        return;
      }

      mostrarToast(`📥 Generando descarga: ${backup.name}...`, 'info');
      
      // 1. Crear workbook Excel desde datos guardados
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Monitoreo');
      
      // Añadir encabezados con estilo
      if (backup.headers && backup.headers.length > 0) {
        const headerRow = worksheet.addRow(backup.headers);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF203864' }
        };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
      }
      
      // Añadir datos
      backup.data.forEach(row => {
        const rowValues = backup.headers ? backup.headers.map(header => row[header] || '') : Object.values(row);
        worksheet.addRow(rowValues);
      });
      
      // Ajustar ancho de columnas
      worksheet.columns = backup.headers ? backup.headers.map(() => ({ width: 15 })) : [{ width: 15 }];
      
      // 2. Generar buffer y blob
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      // 3. Descargar
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${backup.name.replace(/[^a-z0-9]/gi, '_')}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      mostrarToast(`✅ Descarga completada: ${backup.name}`, 'success');
    } catch (error) {
      mostrarToast(`❌ Error al descargar: ${error.message}`, 'error');
      console.error('Error al descargar backup:', error);
    }
  };

  // ========================================
  // FUNCIÓN: Restaurar Backup (Nueva Pestaña)
  // ========================================
  const restoreBackup = async (backup, createNewTab, ExcelJS) => {
    try {
      if (!backup.data || backup.data.length === 0) {
        mostrarToast('❌ No hay datos para restaurar', 'error');
        return;
      }

      mostrarToast(`🔄 Restaurando archivo: ${backup.name}...`, 'info');
      
      // 1. Usar datos directamente desde Firestore (ya cargados en backup)
      const loadedData = backup.data;
      const sheetHeaders = backup.headers || Object.keys(loadedData[0] || {});
      
      // 2. Crear nueva pestaña con los datos restaurados
      const sheetNames = [{ index: 0, name: 'Monitoreo' }];
      
      createNewTab(`${backup.name} (Restaurado)`, {
        data: loadedData,
        availableSheets: sheetNames,
        workbookData: null,  // No necesitamos workbook completo aquí
        currentHeaders: sheetHeaders,
        sheetData: { 0: { data: loadedData, headers: sheetHeaders } }
      });
      
      mostrarToast(`✅ Archivo restaurado: ${backup.name}<br>${loadedData.length} filas cargadas`, 'success');
    } catch (error) {
      console.error('❌ Error al restaurar backup:', error);
      mostrarToast(`❌ Error al restaurar: ${error.message}`, 'error');
    }
  };

  // ========================================
  // FUNCIÓN: Eliminar Backup (solo Firestore)
  // ========================================
  const deleteBackup = async (backupId) => {
    try {
      // Eliminar de Firestore
      await deleteDoc(doc(db, "backups", backupId));
      console.log('✅ Backup eliminado de Firestore');
      
      // Actualizar estado local
      setBackupHistory(prev => prev.filter(b => b.id !== backupId));
      
      mostrarToast('🗑️ Copia de seguridad eliminada exitosamente', 'info');
    } catch (error) {
      mostrarToast(`❌ Error al eliminar: ${error.message}`, 'error');
      console.error('Error al eliminar backup:', error);
    }
  };

  return {
    backupHistory,
    isBackupModalOpen,
    setIsBackupModalOpen,
    saveBackup,
    downloadBackup,
    restoreBackup,
    deleteBackup
  };
};

export default { BackupHistoryModal, useBackupManager };
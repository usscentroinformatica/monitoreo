// src/components/BackupManager.jsx
import React, { useState, useEffect } from 'react';
import { db, storage } from '../utils/firebase';
import { collection, addDoc, deleteDoc, doc, getDocs, query, orderBy, limit } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL, deleteObject } from 'firebase/storage';

// Componente modal para mostrar el historial de backups
export const BackupHistoryModal = ({ isOpen, onClose, backups, onDownload, onDelete }) => {
  if (!isOpen) return null;
  
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl max-h-[80vh] flex flex-col">
        <div className="p-4 bg-blue-900 text-white rounded-t-xl flex justify-between items-center">
          <h2 className="text-xl font-bold">Historial de Archivos Guardados</h2>
          <button 
            onClick={onClose}
            className="text-white hover:text-red-300"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
        
        <div className="overflow-y-auto p-4 flex-1">
          {backups.length === 0 ? (
            <div className="text-center text-gray-500 py-8">
              <p className="text-2xl mb-2">🗃️</p>
              <p>No hay archivos guardados</p>
              <p className="text-sm mt-2">Los archivos que guardes aparecerán aquí</p>
            </div>
          ) : (
            <div className="grid gap-4">
              {backups.map((backup) => {
                const date = new Date(backup.date);
                const formattedDate = `${date.toLocaleDateString()} ${date.toLocaleTimeString()}`;
                
                return (
                  <div 
                    key={backup.id}
                    className="border border-gray-200 rounded-lg overflow-hidden hover:shadow-md transition-shadow"
                  >
                    <div className="bg-gray-100 p-3 flex justify-between items-start">
                      <div>
                        <h3 className="font-bold text-blue-900">{backup.name}</h3>
                        <p className="text-sm text-gray-600">
                          Guardado: {formattedDate} • {backup.rowCount} filas
                        </p>
                      </div>
                      <div className="flex gap-2">
                        <button
                          onClick={() => onDownload(backup)}
                          className="bg-green-600 text-white p-2 rounded-lg hover:bg-green-700"
                          title="Descargar archivo"
                        >
                          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                          </svg>
                        </button>
                        <button
                          onClick={() => {
                            if (window.confirm('¿Estás seguro de que deseas eliminar esta copia de seguridad?')) {
                              onDelete(backup.id);
                            }
                          }}
                          className="bg-red-500 text-white p-2 rounded-lg hover:bg-red-600"
                          title="Eliminar"
                        >
                          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                          </svg>
                        </button>
                      </div>
                    </div>
                    
                    {backup.preview && backup.preview.length > 0 && (
                      <div className="p-3 border-t border-gray-200">
                        <div className="text-xs text-gray-500 mb-1">Vista previa (3 primeras filas):</div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="bg-gray-100">
                                {Object.keys(backup.preview[0]).slice(0, 5).map((key, i) => (
                                  <th key={i} className="p-1 text-left text-xs text-gray-600 font-medium">
                                    {key}
                                  </th>
                                ))}
                                {Object.keys(backup.preview[0]).length > 5 && (
                                  <th className="p-1 text-left text-xs text-gray-600 font-medium">...</th>
                                )}
                              </tr>
                            </thead>
                            <tbody>
                              {backup.preview.map((row, rowIndex) => (
                                <tr key={rowIndex} className={rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                                  {Object.entries(row).slice(0, 5).map(([key, value], i) => (
                                    <td key={i} className="p-1 text-xs border-t border-gray-100">
                                      {String(value || '').substring(0, 20)}
                                      {String(value || '').length > 20 ? '...' : ''}
                                    </td>
                                  ))}
                                  {Object.keys(row).length > 5 && (
                                    <td className="p-1 text-xs border-t border-gray-100">...</td>
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
        
        <div className="p-4 border-t border-gray-200 bg-gray-50 rounded-b-xl">
          <div className="flex justify-between items-center">
            <span className="text-sm text-gray-500">
              {backups.length} {backups.length === 1 ? 'archivo guardado' : 'archivos guardados'}
            </span>
            <button
              onClick={onClose}
              className="px-4 py-2 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded-lg"
            >
              Cerrar
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// Hook personalizado para manejar backups con Firebase
export const useBackupManager = (ExcelJS, mostrarToast) => {
  const [backupHistory, setBackupHistory] = useState([]);
  const [isBackupModalOpen, setIsBackupModalOpen] = useState(false);

  // Cargar backups solo cuando el modal está abierto
  useEffect(() => {
    if (!isBackupModalOpen) return;
    const loadBackups = async () => {
      try {
        const backupsCollection = collection(db, "backups");
        const backupsQuery = query(backupsCollection, orderBy("date", "desc"), limit(20));
        const querySnapshot = await getDocs(backupsQuery);
        const loadedBackups = [];
        querySnapshot.forEach((doc) => {
          loadedBackups.push({
            id: doc.id,
            ...doc.data()
          });
        });
        setBackupHistory(loadedBackups);
      } catch (error) {
        if (isBackupModalOpen) {
          console.error("Error loading backups:", error);
          mostrarToast("❌ Error al cargar copias de seguridad", "error");
        }
      }
    };
    loadBackups();
  }, [isBackupModalOpen, mostrarToast]);

  // Guardar backup en Firebase
  const saveBackup = async (data, currentHeaders, activeTab, setIsLoading) => {
    if (!data || data.length === 0) {
      mostrarToast('❌ No hay datos para guardar', 'error');
      return;
    }

    try {
      setIsLoading(true);
      
      // Crear un workbook para guardar
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Monitoreo');
      
      // Añadir encabezados
      if (currentHeaders.length > 0) {
        worksheet.addRow(currentHeaders);
      }
      
      // Añadir datos
      data.forEach(row => {
        const rowValues = currentHeaders.map(header => row[header] || '');
        worksheet.addRow(rowValues);
      });
      
      // Generar buffer
      const buffer = await workbook.xlsx.writeBuffer();
      
      // Crear un Blob del Excel
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      // Generar nombre de archivo único
      const timestamp = Date.now();
      const fileName = `backup_${timestamp}.xlsx`;
      
      // Subir archivo a Firebase Storage
      const storageRef = ref(storage, `backups/${fileName}`);
      await uploadBytes(storageRef, blob);
      
      // Obtener URL de descarga
      const downloadURL = await getDownloadURL(storageRef);
      
      // Crear registro en Firestore
      const backupData = {
        name: activeTab?.name || `Archivo ${timestamp}`,
        date: new Date().toISOString(),
        rowCount: data.length,
        fileURL: downloadURL,
        fileName: fileName,
        // Guardar solo información esencial para la vista previa
        preview: data.slice(0, 3).map(row => {
          const preview = {};
          // Limitar a columnas clave para no sobrecargar Firestore
          const keyColumns = [
            'DOCENTE', 'CURSO', 'SECCION', 'SESION', 'FECHA', 
            'DIA', 'HORA INICIO', 'HORA FIN', 'MODALIDAD', 'MODELO',
            'TURNO'
          ];
          
          // Filtrar columnas clave que existen en el objeto
          keyColumns.forEach(col => {
            if (row[col] !== undefined) {
              preview[col] = row[col];
            }
          });
          
          // Añadir algunas columnas adicionales si hay espacio
          if (Object.keys(preview).length < 5) {
            Object.keys(row).slice(0, 5 - Object.keys(preview).length).forEach(key => {
              if (!preview[key]) {
                preview[key] = row[key];
              }
            });
          }
          
          return preview;
        })
      };
      
      const docRef = await addDoc(collection(db, "backups"), backupData);
      
      // Actualizar el estado local
      const newBackup = {
        id: docRef.id,
        ...backupData
      };
      
      setBackupHistory(prev => [newBackup, ...prev]);
      
      mostrarToast(`✅ Copia de seguridad guardada:<br><b>${backupData.name}</b><br>${data.length} filas`, 'success');
    } catch (error) {
      mostrarToast(`❌ Error al guardar: ${error.message}`, 'error');
      console.error('Error al guardar backup:', error);
    } finally {
      setIsLoading(false);
    }
  };

  // Descargar backup
  const downloadBackup = async (backup) => {
    try {
      window.open(backup.fileURL, '_blank');
      mostrarToast(`✅ Descargando archivo: ${backup.name}`, 'success');
    } catch (error) {
      mostrarToast(`❌ Error al descargar: ${error.message}`, 'error');
      console.error('Error al descargar backup:', error);
    }
  };

  // Eliminar backup
  const deleteBackup = async (backupId) => {
    try {
      // Obtener la referencia al documento
      const backupRef = doc(db, "backups", backupId);
      
      // Buscar el backup en el estado local para obtener la información del archivo
      const backup = backupHistory.find(b => b.id === backupId);
      
      if (backup && backup.fileName) {
        // Eliminar el archivo de Storage
        const storageRef = ref(storage, `backups/${backup.fileName}`);
        await deleteObject(storageRef);
      }
      
      // Eliminar el documento de Firestore
      await deleteDoc(backupRef);
      
      // Actualizar estado local
      setBackupHistory(prev => prev.filter(b => b.id !== backupId));
      
      mostrarToast('🗑️ Copia de seguridad eliminada', 'info');
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
    deleteBackup
  };
};

export default { BackupHistoryModal, useBackupManager };

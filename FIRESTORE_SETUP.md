# Configuración de Firestore - Reglas y Índices

## 1. REGLAS DE SEGURIDAD (Security Rules)

Ve a Firebase Console → Firestore Database → Rules y reemplaza con esto:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    
    // Colección de autenticación (usuarios)
    match /auth_users/{document=**} {
      allow read: if true;
      allow write: if request.auth != null;
    }
    
    // Colección de usuarios (para lookup)
    match /usuarios/{document=**} {
      allow read: if true;
      allow write: if request.auth != null;
    }
    
    // Colección de backups - FILTRO POR USUARIO
    match /backups/{document=**} {
      // Solo lectura del propietario
      allow read: if request.auth != null && resource.data.userId == request.auth.uid;
      
      // Permitir leer sin autenticación si el documento tiene userId (para login)
      allow read: if resource.data.userId == request.query.get('userId');
      
      // Crear nuevos backups (cualquier usuario autenticado)
      allow create: if request.auth != null && request.resource.data.userId == request.auth.uid;
      
      // Actualizar/eliminar solo el propietario
      allow update, delete: if request.auth != null && resource.data.userId == request.auth.uid;
    }
    
    // Otros
    match /{document=**} {
      allow read, write: if false;
    }
  }
}
```

## 2. ÍNDICE COMPUESTO REQUERIDO

**Colección:** `backups`

**Campos del índice (en este orden):**
1. `userId` (Ascending)
2. `date` (Descending)
3. `__name__` (Descending)

**Pasos en Firebase Console:**
1. Firestore Database → Indexes → Create index
2. Colección: `backups`
3. Agrega campos:
   - `userId` → Ascending
   - `date` → Descending
4. Crear index

O Firebase te proporcionará un link automático cuando intentes hacer la consulta la primera vez.

## 3. ESTRUCTURA DE DOCUMENTO BACKUP

Cada documento en `backups` debe tener esta estructura:

```json
{
  "name": "Archivo_202604",
  "date": "2026-04-22T10:30:00.000Z",
  "userId": "USUARIO_NOMBRE",        // 🔑 IMPORTANTE: Campo requerido
  "rowCount": 150,
  "headers": ["DOCENTE", "CURSO", ...],
  "data": [
    { "DOCENTE": "...", "CURSO": "...", ... },
    ...
  ],
  "preview": [...]
}
```

## 4. NOTAS IMPORTANTES

- ⚠️ Si tienes backups ANTIGUOS sin campo `userId`, necesitas limpiarlos o agregarles el campo manualmente.
- 📌 El código ahora guarda automáticamente `userId: authUser.username` en cada backup.
- 🔐 Solo cada usuario verá sus propios backups (filtro en lectura).
- ✅ Backups nuevos aparecerán inmediatamente en el historial del usuario.

## 5. TROUBLESHOOTING

Si ves error **"Missing or insufficient permissions"**:
1. Verifica que `userId` esté guardado en los documentos
2. Comprueba que el índice se haya creado completamente
3. Recarga la página (puede tardar hasta 1 minuto)

Si ves error **"Index not found"**:
1. Firestore te dará un link en la consola del navegador
2. Haz clic en el link o crea el índice manualmente (paso 3 arriba)
3. Espera a que el índice se construya (algunos minutos)

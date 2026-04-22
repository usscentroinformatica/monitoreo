import React, { useState } from 'react';
import { db } from '../utils/firebase';
import { addDoc, collection, getDocs, limit, query } from 'firebase/firestore';

function Login({ onAuthenticated }) {
  const [loginUser, setLoginUser] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const [needsInitialSetup, setNeedsInitialSetup] = useState(false);
  const [setupUser, setSetupUser] = useState('ADMIN');
  const [setupPassword, setSetupPassword] = useState('');
  const [setupDisplayName, setSetupDisplayName] = useState('Administrador');
  const [setupError, setSetupError] = useState('');
  const [isCreatingSetupUser, setIsCreatingSetupUser] = useState(false);

  const handleLoginSubmit = async (event) => {
    event.preventDefault();
    const username = String(loginUser || '').trim();
    const password = String(loginPassword || '').trim();

    if (!username || !password) {
      setLoginError('Completa usuario y contraseña.');
      return;
    }

    setIsLoggingIn(true);
    setLoginError('');

    try {
      const usernameUpper = username.toUpperCase();
      const candidateCollections = ['auth_users', 'usuarios', 'users', 'credenciales'];
      let matchedUser = null;
      let totalUsers = 0;

      for (const colName of candidateCollections) {
        const snapshot = await getDocs(query(collection(db, colName), limit(100)));
        if (snapshot.empty) continue;

        totalUsers += snapshot.size;

        snapshot.forEach((docSnap) => {
          if (matchedUser) return;
          const userData = docSnap.data() || {};

          const dbUserRaw =
            userData.username ??
            userData.usuario ??
            userData.user ??
            userData.email ??
            '';

          const dbPasswordRaw =
            userData.password ??
            userData['contraseña'] ??
            userData.contrasena ??
            userData.clave ??
            userData.pass ??
            '';

          const dbStatus = String(userData.estado ?? userData.status ?? 'activo').toLowerCase();
          const isActive = !['inactivo', 'inactive', 'disabled', 'deshabilitado', 'false', '0'].includes(dbStatus);
          if (!isActive) return;

          const dbUser = String(dbUserRaw).trim();
          const dbUserUpper = dbUser.toUpperCase();
          const dbEmailLocal = dbUser.includes('@') ? dbUser.split('@')[0].toUpperCase() : dbUserUpper;
          const passOk = String(dbPasswordRaw).trim() === password;
          const userOk = dbUserUpper === usernameUpper || dbEmailLocal === usernameUpper;

          if (userOk && passOk) {
            const displayName =
              userData.displayName ||
              userData.nombre ||
              userData.name ||
              (dbUserUpper === 'ADMIN' || dbEmailLocal === 'ADMIN' ? 'Administrador' : dbUser);

            matchedUser = {
              id: docSnap.id,
              username: dbUser,
              displayName: String(displayName || 'Usuario')
            };
          }
        });

        if (matchedUser) break;
      }

      if (totalUsers === 0) {
        setNeedsInitialSetup(true);
        setLoginError('No hay usuarios configurados en la base de datos. Crea el usuario inicial abajo.');
        return;
      }

      if (!matchedUser) {
        setLoginError('Usuario o contraseña inválidos. Verifica que el usuario exista en Firestore.');
        return;
      }

      setNeedsInitialSetup(false);
      sessionStorage.setItem('monitoreo-auth-user', JSON.stringify(matchedUser));
      onAuthenticated(matchedUser);
      setLoginError('');
      setLoginPassword('');
    } catch (error) {
      setLoginError('No se pudo iniciar sesión. Revisa permisos/reglas de Firestore e inténtalo otra vez.');
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleCreateInitialUser = async (event) => {
    event.preventDefault();
    const username = String(setupUser || '').trim().toUpperCase();
    const password = String(setupPassword || '').trim();
    const displayName = String(setupDisplayName || '').trim() || 'Administrador';

    if (!username || !password) {
      setSetupError('Debes ingresar usuario y contraseña para crear el acceso inicial.');
      return;
    }

    setIsCreatingSetupUser(true);
    setSetupError('');

    try {
      await addDoc(collection(db, 'auth_users'), {
        username,
        password,
        displayName,
        estado: 'activo',
        createdAt: new Date().toISOString()
      });

      setNeedsInitialSetup(false);
      setLoginUser(username);
      setLoginPassword('');
      setLoginError('Usuario inicial creado. Ahora inicia sesión.');
      setSetupPassword('');
    } catch (error) {
      setSetupError('No se pudo crear el usuario inicial. Verifica reglas de Firestore para permitir escritura en auth_users.');
    } finally {
      setIsCreatingSetupUser(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#11acd3] p-4 flex items-center justify-center">
      <div className="w-full max-w-md bg-white shadow-2xl rounded-2xl p-8 border-2 border-[#5a2290]">
        <div className="h-2 w-full bg-[#63ed12] rounded-full mb-5"></div>
        <div className="text-center mb-5">
          <div className="inline-flex items-center justify-center w-14 h-14 rounded-full bg-[#5a2290] text-white font-black text-xl mb-3">USS</div>
        </div>
        <h1 className="text-2xl font-extrabold text-[#5a2290] mb-6 text-center">
          Acceso
        </h1>

        <form onSubmit={handleLoginSubmit} className="space-y-4">
          <div>
            <label htmlFor="login-user" className="block text-sm font-semibold text-slate-700 mb-1">
              Usuario
            </label>
            <input
              id="login-user"
              type="text"
              value={loginUser}
              onChange={(e) => setLoginUser(e.target.value)}
                className="w-full rounded-lg border border-slate-300 px-3 py-2 focus:outline-none focus:ring-2 focus:ring-[#11acd3]"
              placeholder="admin o admin@uss.edu.pe"
              autoComplete="username"
            />
          </div>

          <div>
            <label htmlFor="login-password" className="block text-sm font-semibold text-slate-700 mb-1">
              Contraseña
            </label>
            <input
              id="login-password"
              type="password"
              value={loginPassword}
              onChange={(e) => setLoginPassword(e.target.value)}
                className="w-full rounded-lg border border-slate-300 px-3 py-2 focus:outline-none focus:ring-2 focus:ring-[#11acd3]"
              placeholder="Tu contraseña"
              autoComplete="current-password"
            />
          </div>

          {loginError && (
            <p className="text-sm text-red-600 font-semibold">
              {loginError}
            </p>
          )}

          <button
            type="submit"
            disabled={isLoggingIn}
            className="w-full rounded-lg bg-[#11acd3] hover:bg-[#0f9bbf] text-white font-bold py-2.5 transition-colors disabled:opacity-60"
          >
            {isLoggingIn ? 'Ingresando...' : 'Ingresar'}
          </button>
        </form>

        {needsInitialSetup && (
          <div className="mt-5 p-4 rounded-xl border border-[#63ed12] bg-[#f4fde9]">
            <h3 className="text-sm font-extrabold text-[#5a2290] mb-2">Configuración inicial</h3>
            <p className="text-xs text-slate-700 mb-3">No se detectaron usuarios. Crea el primer acceso para habilitar el login.</p>

            <form onSubmit={handleCreateInitialUser} className="space-y-3">
              <input
                type="text"
                value={setupUser}
                onChange={(e) => setSetupUser(e.target.value)}
                className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#11acd3]"
                placeholder="Usuario"
              />
              <input
                type="password"
                value={setupPassword}
                onChange={(e) => setSetupPassword(e.target.value)}
                className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#11acd3]"
                placeholder="Contraseña"
              />
              <input
                type="text"
                value={setupDisplayName}
                onChange={(e) => setSetupDisplayName(e.target.value)}
                className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#11acd3]"
                placeholder="Nombre para mostrar"
              />

              {setupError && <p className="text-xs font-semibold text-red-600">{setupError}</p>}

              <button
                type="submit"
                disabled={isCreatingSetupUser}
                className="w-full rounded-lg bg-[#63ed12] hover:bg-[#54cb0f] text-[#103b07] font-bold py-2 text-sm disabled:opacity-60"
              >
                {isCreatingSetupUser ? 'Creando usuario...' : 'Crear usuario inicial'}
              </button>
            </form>
          </div>
        )}

        <p className="text-xs text-slate-500 text-center mt-4">
          Las credenciales se validan desde Firestore y no están hardcodeadas en el código.
        </p>
      </div>
    </div>
  );
}

export default Login;

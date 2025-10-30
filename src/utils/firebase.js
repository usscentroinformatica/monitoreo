// src/utils/firebase.js
import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';
import { getStorage } from 'firebase/storage';

// Configuración de Firebase
const firebaseConfig = {
  apiKey: "AIzaSyCz4hPVVhSLYjTMRc-IoQxbLjdJ2xc-QdI",
  authDomain: "backupmonitoreo.firebaseapp.com",
  projectId: "backupmonitoreo",
  storageBucket: "backupmonitoreo.appspot.com", // Corrige esta línea a .appspot.com
  messagingSenderId: "529641473535",
  appId: "1:529641473535:web:88ef807a017fdef72bd3ca"
};

// Inicializar Firebase
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const storage = getStorage(app);

import { db } from "./firebase";
import { doc, getDoc, setDoc, deleteDoc } from "firebase/firestore";

const COLLECTION = "mandal";

// Mirrors the shape of the Claude-artifact window.storage API this app was
// originally built against: get/set(key, shared). shared=true -> everyone
// using this deployment shares the same data (Firestore). shared=false ->
// stays local to this browser (localStorage), used only for "which account
// is this device currently logged in as".
export const storage = {
  async get(key, shared) {
    if (!shared) {
      const v = localStorage.getItem(key);
      if (v === null) throw new Error("not found");
      return { key, value: v, shared };
    }
    const snap = await getDoc(doc(db, COLLECTION, key));
    if (!snap.exists()) throw new Error("not found");
    return { key, value: snap.data().value, shared };
  },

  async set(key, value, shared) {
    if (!shared) {
      localStorage.setItem(key, value);
      return { key, value, shared };
    }
    await setDoc(doc(db, COLLECTION, key), { value });
    return { key, value, shared };
  },

  async delete(key, shared) {
    if (!shared) {
      localStorage.removeItem(key);
      return { key, deleted: true, shared };
    }
    await deleteDoc(doc(db, COLLECTION, key));
    return { key, deleted: true, shared };
  },
};

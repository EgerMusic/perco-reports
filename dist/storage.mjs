// Keep each GitHub Pages project in its own database on the shared github.io origin.
const scope = location.pathname.replace(/\/index\.html$/i, '').replace(/\/+$/, '') || '/';
const databaseName = `perco-workspace-v1:${scope}`;
let databasePromise;
export class StorageConflictError extends Error {
  constructor() { super('Файлы изменены в другой вкладке. Обновите страницу перед продолжением.'); this.name='StorageConflictError'; }
}
function openDatabase() {
  if (!globalThis.indexedDB) return Promise.reject(new Error('Хранилище браузера недоступно.'));
  if (!databasePromise) databasePromise = new Promise((resolve,reject)=>{
    const request=indexedDB.open(databaseName,1);
    let failed=false;
    request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains('workspace'))request.result.createObjectStore('workspace');};
    request.onsuccess=()=>{
      if(failed){request.result.close();return;}
      const db=request.result;
      db.onversionchange=()=>{db.close();databasePromise=null;};
      resolve(db);
    };
    request.onerror=()=>{failed=true;reject(request.error||new Error('Не удалось открыть хранилище браузера.'));};
    request.onblocked=()=>{failed=true;reject(new Error('Закройте другие вкладки сервиса и повторите попытку.'));};
  }).catch(error=>{databasePromise=null;throw error;});
  return databasePromise;
}
export async function loadWorkspace() {
  const db=await openDatabase();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction('workspace','readonly'),request=tx.objectStore('workspace').get('current');
    tx.oncomplete=()=>resolve(request.result||null);
    tx.onerror=tx.onabort=()=>reject(tx.error||new Error('Не удалось прочитать сохранённые файлы.'));
  });
}
async function writeWorkspace(payload,expectedToken) {
  const db=await openDatabase();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction('workspace','readwrite'),store=tx.objectStore('workspace'),read=store.get('current');
    let problem=null,record;
    read.onsuccess=()=>{
      if((read.result?.token??null)!==expectedToken){problem=new StorageConflictError();tx.abort();return;}
      record={...payload,version:1,token:crypto.randomUUID(),savedAt:Date.now()};
      try{store.put(record,'current');}catch(error){problem=error;tx.abort();}
    };
    tx.oncomplete=()=>resolve(record);
    tx.onerror=tx.onabort=()=>reject(problem||tx.error||new Error('Не удалось сохранить файлы.'));
  });
}
export const saveWorkspace=(files,settings,expectedToken)=>writeWorkspace({files,settings,cleared:false},expectedToken);
// A small tombstone prevents an old tab from resurrecting deleted files.
export const clearWorkspace=expectedToken=>writeWorkspace({files:{events:[],schedule:null,presence:null},settings:{},cleared:true},expectedToken);

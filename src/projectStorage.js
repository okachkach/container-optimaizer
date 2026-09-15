let connection;

function newId() {
  return window.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function database() {
  if (!connection) connection = new Promise((resolve, reject) => {
    const request = indexedDB.open('containerOptimizerProjects', 2);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains('workspace')) db.createObjectStore('workspace');
      const history = db.objectStoreNames.contains('history')
        ? request.transaction.objectStore('history')
        : db.createObjectStore('history', { keyPath: 'id' });
      if (!history.indexNames.contains('projectId')) history.createIndex('projectId', 'projectId');
      if (!db.objectStoreNames.contains('projects')) db.createObjectStore('projects', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta');
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  return connection;
}

function result(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function done(tx, message) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = resolve;
    tx.onabort = tx.onerror = () => reject(tx.error || new Error(message));
  });
}

export function projectSummary(project) {
  return { id: project.id, projectName: project.projectName || 'Untitled project',
    createdAt: project.createdAt || project.lastSaved, updatedAt: project.updatedAt || project.lastSaved,
    containerCount: project.containers?.length || 0,
    cartonCount: project.containers?.reduce((sum, c) => sum + (c.totalCartons || 0), 0) || 0 };
}

export async function readProjects() {
  const db = await database();
  const tx = db.transaction(['workspace', 'projects', 'history', 'meta']);
  const [legacy, projects, allHistory, activeId] = await Promise.all([
    result(tx.objectStore('workspace').get('current')), result(tx.objectStore('projects').getAll()),
    result(tx.objectStore('history').getAll()), result(tx.objectStore('meta').get('activeProjectId'))
  ]);
  if (!projects.length && legacy) {
    const id = newId();
    const migrated = { ...legacy, id, projectName: legacy.projectName || 'My shipment',
      createdAt: legacy.lastSaved || new Date().toISOString(), updatedAt: legacy.lastSaved || new Date().toISOString() };
    const migratedHistory = allHistory.map(entry => ({ ...entry, projectId: id }));
    const write = db.transaction(['projects', 'history', 'meta'], 'readwrite');
    write.objectStore('projects').put(migrated);
    migratedHistory.forEach(entry => write.objectStore('history').put(entry));
    write.objectStore('meta').put(id, 'activeProjectId');
    await done(write, 'Could not migrate the existing project');
    return { current: migrated, projects: [projectSummary(migrated)], history: migratedHistory };
  }
  const current = projects.find(project => project.id === activeId) || projects[0] || null;
  return { current, projects: projects.map(projectSummary),
    history: current ? allHistory.filter(entry => entry.projectId === current.id) : [] };
}

export async function openProject(projectId) {
  const db = await database();
  const tx = db.transaction(['projects', 'history', 'meta'], 'readwrite');
  const projectRequest = tx.objectStore('projects').get(projectId);
  const historyRequest = tx.objectStore('history').index('projectId').getAll(projectId);
  tx.objectStore('meta').put(projectId, 'activeProjectId');
  const [project, history] = await Promise.all([result(projectRequest), result(historyRequest)]);
  await done(tx, 'Could not open the project');
  if (!project) throw new Error('Project was not found.');
  return { project, history };
}

export async function saveProject(current, snapshot) {
  if (!current.id) throw new Error('Project ID is missing.');
  const db = await database();
  const tx = db.transaction(['projects', 'history', 'meta'], 'readwrite');
  tx.objectStore('projects').put(current);
  tx.objectStore('meta').put(current.id, 'activeProjectId');
  if (snapshot) tx.objectStore('history').put({ ...snapshot, projectId: current.id });
  await done(tx, 'Storage transaction failed');
}

export async function updateHistory(entry, remove = false) {
  const db = await database();
  const tx = db.transaction('history', 'readwrite');
  if (remove) tx.objectStore('history').delete(entry.id);
  else tx.objectStore('history').put(entry);
  await done(tx, 'History update failed');
}

export async function deleteProject(projectId, nextProjectId) {
  const db = await database();
  const tx = db.transaction(['projects', 'history', 'meta'], 'readwrite');
  tx.objectStore('projects').delete(projectId);
  const store = tx.objectStore('history');
  store.index('projectId').openKeyCursor(IDBKeyRange.only(projectId)).onsuccess = event => {
    const cursor = event.target.result;
    if (cursor) { store.delete(cursor.primaryKey); cursor.continue(); }
  };
  if (nextProjectId) tx.objectStore('meta').put(nextProjectId, 'activeProjectId');
  else tx.objectStore('meta').delete('activeProjectId');
  await done(tx, 'Could not delete the project');
}

export async function exportLibrary() {
  const db = await database();
  const tx = db.transaction(['projects', 'history']);
  const [projects, history] = await Promise.all([
    result(tx.objectStore('projects').getAll()), result(tx.objectStore('history').getAll())
  ]);
  return { type: 'container-optimizer-library', version: 1, exportedAt: new Date().toISOString(), projects, history };
}

export async function importLibrary(library) {
  const db = await database();
  const tx = db.transaction(['projects', 'history', 'meta'], 'readwrite');
  library.projects.forEach(project => tx.objectStore('projects').put(project));
  library.history.forEach(entry => tx.objectStore('history').put(entry));
  if (library.projects[0]) tx.objectStore('meta').put(library.projects[0].id, 'activeProjectId');
  await done(tx, 'Could not restore the projects backup');
}

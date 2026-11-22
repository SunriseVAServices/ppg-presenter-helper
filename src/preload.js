const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('ppg', {
	send: (a) => ipcRenderer.send('ui', a),
	on: (cb) => ipcRenderer.on('status', (e, s) => cb(s)),
});

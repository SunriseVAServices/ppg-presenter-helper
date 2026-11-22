// PPG Presenter Helper: a small tray app for the presentation laptop.
// It connects to theppg.llc with a 6-digit code (typed into the Remote once), then plays presenters'
// PowerPoint / Keynote files in the real apps, and Canva links full screen, while their phone drives them.

const { app, BrowserWindow, Tray, Menu, ipcMain, screen, safeStorage, nativeImage, shell } = require('electron');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Helper } = require('./core');

const MAC = process.platform === 'darwin';
const DEFAULT_SITE = 'https://theppg.llc';

if (!app.requestSingleInstanceLock()) { app.quit(); }

let win = null, tray = null, helper = null, show = null, driverRef = null;
const cfgFile = () => path.join(app.getPath('userData'), 'settings.json');

/* ---------- settings (the connection key is stored encrypted by Windows / the Mac keychain) ---------- */
function load() {
	let c = {};
	try { c = JSON.parse(fs.readFileSync(cfgFile(), 'utf8')); } catch (e) { /* first run */ }
	let token = '';
	try {
		if (c.tok && safeStorage.isEncryptionAvailable()) { token = safeStorage.decryptString(Buffer.from(c.tok, 'base64')); }
	} catch (e) { token = ''; }
	return { site: c.site || DEFAULT_SITE, display: c.display || 0, token, launched: !!c.launched };
}
function save(patch) {
	const c = Object.assign(load(), patch);
	const out = { site: c.site, display: c.display, launched: c.launched };
	if (c.token && safeStorage.isEncryptionAvailable()) { out.tok = safeStorage.encryptString(c.token).toString('base64'); }
	fs.mkdirSync(path.dirname(cfgFile()), { recursive: true });
	fs.writeFileSync(cfgFile(), JSON.stringify(out, null, 1));
}

/* ---------- Canva (and any web link): our own full-screen window on the TV ---------- */
function tvDisplay() {
	const all = screen.getAllDisplays();
	const pick = load().display;
	const chosen = all.find((d) => d.id === pick);
	if (chosen) { return chosen; }
	const prim = screen.getPrimaryDisplay();
	return all.find((d) => d.id !== prim.id) || prim; // The TV is usually the second screen.
}

class Web {
	constructor() { this.w = null; this.n = 1; }
	async open(a) {
		this.close();
		const d = tvDisplay();
		this.w = new BrowserWindow({ x: d.bounds.x, y: d.bounds.y, width: d.bounds.width, height: d.bounds.height, fullscreen: true, frame: false, backgroundColor: '#000000', autoHideMenuBar: true, webPreferences: { partition: 'persist:canva', sandbox: true } });
		this.w.on('closed', () => { this.w = null; });
		this.w.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
		this.n = 1;
		await this.w.loadURL(a.url);
		this.w.focus();
	}
	key(k) {
		if (!this.w) { return; }
		this.w.focus();
		this.w.webContents.sendInputEvent({ type: 'keyDown', keyCode: k });
		this.w.webContents.sendInputEvent({ type: 'keyUp', keyCode: k });
	}
	next() { this.key('Right'); this.n++; }
	prev() { this.key('Left'); this.n = Math.max(1, this.n - 1); }
	close() { if (this.w) { this.w.destroy(); this.w = null; } }
	state() { return this.w ? { mode: 'web', slide: this.n, total: 0 } : { mode: 'idle', slide: 0, total: 0 }; }
}

/* ---------- one driver that sends each file to the right app ---------- */
function makeDriver(dir) {
	const web = new Web();
	let ppt = null, mac = null, cur = null;
	const win32 = () => { if (!ppt) { const { PowerPointWin } = require('./ppt-win'); ppt = new PowerPointWin(dir); } return ppt; };
	const macs = () => { if (!mac) { mac = require('./mac'); } return mac; };
	return {
		async open(a, file) {
			if (a.type === 'web') { cur = web; return web.open(a); }
			if (MAC) {
				const m = macs();
				if (a.type === 'key' || !(await m.hasApp('com.microsoft.Powerpoint'))) {
					if (!(await m.hasApp('com.apple.iWork.Keynote'))) { throw new Error(a.type === 'key' ? 'Keynote isn’t installed on this Mac' : 'Neither PowerPoint nor Keynote is installed on this Mac'); }
					cur = new m.Keynote();
				} else {
					cur = new m.PowerPointMac();
				}
				return cur.open(file);
			}
			if (a.type === 'key') { throw new Error('Keynote files need a Mac'); }
			cur = win32();
			return cur.open(file);
		},
		next() { return cur ? cur.next() : null; },
		prev() { return cur ? cur.prev() : null; },
		async close() { const c = cur; cur = null; if (c) { await c.close(); } },
		async state() { return cur ? cur.state() : { mode: 'idle', slide: 0, total: 0 }; },
		quit() { web.close(); if (ppt) { ppt.quit(); } },
	};
}

/* ---------- the little window ---------- */
function openWindow() {
	if (win) { win.show(); win.focus(); return; }
	win = new BrowserWindow({ width: 420, height: 560, resizable: false, maximizable: false, fullscreenable: false, title: 'PPG Presenter Helper', backgroundColor: '#0A0B0D', autoHideMenuBar: true, show: false, webPreferences: { preload: path.join(__dirname, 'preload.js'), sandbox: true } });
	win.loadFile(path.join(__dirname, 'ui.html'));
	win.once('ready-to-show', () => win.show());
	win.on('close', (e) => { if (!app.quitting) { e.preventDefault(); win.hide(); } }); // Keep running in the tray.
	win.on('closed', () => { win = null; });
	win.webContents.setWindowOpenHandler(({ url }) => { shell.openExternal(url); return { action: 'deny' }; });
}

function push() {
	if (!helper) { return; }
	const c = load();
	const s = Object.assign({}, helper.status, {
		site: c.site,
		siteName: helper.status.site || '',
		name: helper.name,
		ver: app.getVersion(),
		login: app.getLoginItemSettings().openAtLogin,
		displays: screen.getAllDisplays().map((d, i) => ({ id: d.id, label: (d.label || ('Screen ' + (i + 1))) + ' · ' + d.size.width + '×' + d.size.height + (d.id === screen.getPrimaryDisplay().id ? ' (main)' : '') })),
		display: tvDisplay().id,
	});
	if (win) { win.webContents.send('status', s); }
	if (tray) {
		const what = s.phase === 'connected' ? (s.mode === 'opening' ? 'Opening…' : (s.key ? 'Presenting: ' + (s.title || '') : 'Connected · ready')) : (s.phase === 'offline' ? 'Can’t reach the website' : 'Not connected');
		tray.setToolTip('PPG Presenter Helper · ' + what);
	}
}

app.whenReady().then(() => {
	if (MAC && app.dock) { app.dock.hide(); }
	const c = load();
	if (!c.launched) { app.setLoginItemSettings({ openAtLogin: true }); save({ launched: true }); } // Starts with the laptop.
	const dir = path.join(os.tmpdir(), 'ppg-presenter-helper');
	const driver = makeDriver(dir);
	driverRef = driver;
	const start = () => {
		const cc = load();
		helper = new Helper({ site: cc.site, token: cc.token, name: os.hostname().replace(/\.local$/, ''), os: MAC ? 'mac' : 'win', ver: app.getVersion(), driver, dir });
		helper.on('status', push);
		helper.on('token', (t) => save({ token: t }));
		helper.on('log', (m) => console.log(m));
		helper.start();
	};
	start();

	const img = nativeImage.createFromPath(path.join(__dirname, MAC ? 'trayTemplate.png' : 'tray-win.png'));
	if (MAC) { img.setTemplateImage(true); }
	tray = new Tray(img);
	tray.setContextMenu(Menu.buildFromTemplate([
		{ label: 'Open PPG Presenter Helper', click: openWindow },
		{ type: 'separator' },
		{ label: 'Quit', click: () => { app.quitting = true; app.quit(); } },
	]));
	tray.on('click', openWindow);
	openWindow();

	ipcMain.on('ui', (e, a) => {
		if (a.do === 'ready') { push(); }
		if (a.do === 'disconnect') { helper.forget(); save({ token: '' }); push(); }
		if (a.do === 'site') {
			let u = String(a.v || '').trim();
			if (u && !/^https?:\/\//i.test(u)) { u = 'https://' + u; }
			save({ site: u || DEFAULT_SITE, token: '' });
			helper.stop(); helper.removeAllListeners(); start(); push();
		}
		if (a.do === 'display') { save({ display: Number(a.v) || 0 }); push(); }
		if (a.do === 'login') { app.setLoginItemSettings({ openAtLogin: !!a.v }); push(); }
		if (a.do === 'test') { // Show a test card on the TV screen for 4 seconds.
			if (show) { show.destroy(); }
			const d = tvDisplay();
			show = new BrowserWindow({ x: d.bounds.x, y: d.bounds.y, width: d.bounds.width, height: d.bounds.height, fullscreen: true, frame: false, backgroundColor: '#0A0B0D' });
			show.loadURL('data:text/html,' + encodeURIComponent('<body style="margin:0;height:100vh;display:flex;align-items:center;justify-content:center;background:#0A0B0D;color:#F4EC32;font:700 8vw system-ui">Canva shows here</body>'));
			setTimeout(() => { if (show) { show.destroy(); show = null; } }, 4000);
		}
	});
	screen.on('display-added', push);
	screen.on('display-removed', push);
});

app.on('second-instance', openWindow);
app.on('window-all-closed', (e) => e.preventDefault && e.preventDefault()); // Stay in the tray.
app.on('before-quit', () => { app.quitting = true; });
app.on('will-quit', () => { try { helper && helper.stop(); driverRef && driverRef.quit(); } catch (e) { /* fine */ } });

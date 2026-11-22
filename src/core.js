// The part of the helper that talks to the website. Plain Node, no Electron, so it can be tested on its own.
//
//   const h = new Helper({ site, token, name, os, ver, driver, dir })
//   h.on('status', s => ...)   // { phase: 'pairing'|'connected'|'offline'|'revoked', code, mode, key, slide, total, err, title }
//   h.on('token', t => ...)    // save it (treat it like a password)
//   h.start()
//
// The driver plays things: { open(cmd, file) -> Promise, next(), prev(), close(), state() -> {mode, slide, total} }.

const EventEmitter = require('events');
const fs = require('fs');
const path = require('path');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

class Helper extends EventEmitter {
	constructor(o) {
		super();
		this.site = String(o.site || 'https://theppg.llc').replace(/\/+$/, '');
		this.token = o.token || '';
		this.name = o.name || 'Presentation laptop';
		this.os = o.os === 'mac' ? 'mac' : 'win';
		this.ver = o.ver || '1.0.0';
		this.driver = o.driver;
		this.dir = o.dir; // Where downloaded presentations go (emptied as we go).
		this.since = null;
		this.cur = { key: '', title: '', err: '' };
		this.errAt = 0;
		this.queue = Promise.resolve();
		this.running = false;
		this.status = { phase: this.token ? 'connected' : 'pairing' };
	}

	api(p) { return this.site + '/wp-json/ppg/v1/' + p; }

	async post(p, body) {
		const headers = { 'Content-Type': 'application/json' };
		if (this.token) { headers['X-PPG-Helper'] = this.token; }
		const r = await fetch(this.api(p), { method: 'POST', headers, body: JSON.stringify(body || {}), signal: AbortSignal.timeout(15000) });
		const j = await r.json().catch(() => null);
		return { status: r.status, j };
	}

	set(s) {
		this.status = Object.assign({}, this.status, s);
		this.emit('status', this.status);
	}

	start() {
		if (this.running) { return; }
		this.running = true;
		this.loop();
	}

	stop() { this.running = false; }

	/** Forget the connection (the Disconnect button, or the website disconnected us). */
	forget(phase) {
		this.token = '';
		this.since = null;
		this.emit('token', '');
		this.set({ phase: phase || 'pairing', code: '' });
	}

	async loop() {
		while (this.running) {
			try {
				if (!this.token) { await this.pair(); } else { await this.poll(); }
			} catch (e) {
				this.set({ phase: this.token ? 'offline' : 'pairing', code: this.token ? '' : this.status.code, net: String(e && e.message || e) });
				await sleep(3000);
			}
		}
	}

	/** Show a 6-digit code until someone types it into the Remote. Codes last 20 minutes. */
	async pair() {
		const r = await this.post('rm/h/pair', { os: this.os, name: this.name });
		if (r.status !== 200 || !r.j || !r.j.id) { throw new Error((r.j && r.j.message) || 'The website didn’t answer (' + r.status + ').'); }
		this.set({ phase: this.status.phase === 'revoked' ? 'revoked' : 'pairing', code: r.j.code, net: '' });
		const id = r.j.id;
		while (this.running && !this.token) {
			await sleep(2000);
			const w = await this.post('rm/h/pair', { id });
			if (!w.j || w.j.gone) { return; } // Expired: get a new code.
			if (w.j.ok && w.j.token) {
				this.token = w.j.token;
				this.since = null;
				this.emit('token', this.token);
				this.set({ phase: 'connected', code: '', site: w.j.site || '' });
				return;
			}
		}
	}

	async poll() {
		let st = { mode: 'idle', slide: 0, total: 0 };
		try { st = Object.assign(st, await this.driver.state()); } catch (e) { /* keep idle */ }
		if (st.mode === 'idle' && this.cur.key && !this.busy) {
			this.cur.key = ''; // They closed it on the laptop (Esc): hand back to the Stage.
		}
		if (this.cur.err && Date.now() - this.errAt > 8000) { this.cur.err = ''; }
		const state = { mode: this.busy ? 'opening' : st.mode, key: this.cur.key || (this.cur.err ? this.cur.errKey : ''), slide: st.slide || 0, total: st.total || 0, err: this.cur.err, os: this.os, ver: this.ver };
		const body = { state };
		if (this.since !== null) { body.since = this.since; }
		const r = await this.post('rm/h/poll', body);
		if (r.status === 401 || r.status === 403) { this.forget('revoked'); return; }
		if (r.status !== 200 || !r.j) { throw new Error('The website didn’t answer (' + r.status + ').'); }
		this.set({ phase: 'connected', net: '', site: r.j.site || this.status.site, mode: state.mode, key: state.key, slide: state.slide, total: state.total, err: state.err, title: this.cur.key ? this.cur.title : '' });
		const first = this.since === null;
		this.since = r.j.last;
		if (!first) {
			for (const c of r.j.cmds || []) { this.run(c); }
		}
		await sleep(this.busy ? 400 : 700);
	}

	/** Commands run one at a time, in order (Next waits for a file that's still opening). */
	run(c) {
		this.queue = this.queue.then(() => this.exec(c)).catch((e) => this.emit('log', 'command failed: ' + (e && e.message || e)));
		return this.queue;
	}

	async exec(c) {
		const a = c.arg || {};
		this.emit('log', 'command: ' + c.cmd + (a.key ? ' ' + a.key : ''));
		switch (c.cmd) {
			case 'open': return this.open(a);
			case 'next': if (this.cur.key && (!a.key || a.key === this.cur.key)) { await this.driver.next(); } return;
			case 'prev': if (this.cur.key && (!a.key || a.key === this.cur.key)) { await this.driver.prev(); } return;
			case 'close': return this.close();
		}
	}

	async open(a) {
		await this.close();
		this.busy = true;
		this.cur = { key: a.key, title: a.title || '', err: '' };
		try {
			let file = '';
			if (a.file) { file = await this.download(a); }
			await this.driver.open(a, file);
		} catch (e) {
			this.cur = { key: '', title: '', err: String(e && e.message || e).slice(0, 180), errKey: a.key };
			this.errAt = Date.now();
			this.emit('log', 'couldn’t open: ' + this.cur.err);
			try { await this.driver.close(); } catch (x) { /* nothing open */ }
		} finally {
			this.busy = false;
		}
	}

	async close() {
		if (!this.cur.key && !this.busy) { try { await this.driver.close(); } catch (e) { /* fine */ } return; }
		this.cur = { key: '', title: '', err: '' };
		try { await this.driver.close(); } catch (e) { /* already closed */ }
		this.tidy();
	}

	async download(a) {
		fs.mkdirSync(this.dir, { recursive: true });
		const r = await fetch(a.file, { headers: { 'X-PPG-Helper': this.token }, signal: AbortSignal.timeout(180000) });
		if (r.status !== 200) {
			const j = await r.json().catch(() => null);
			throw new Error((j && j.message) || 'Couldn’t download the file (' + r.status + ').');
		}
		const buf = Buffer.from(await r.arrayBuffer());
		const ext = a.ext === 'key' ? 'key' : 'pptx';
		const safe = String(a.title || a.key).replace(/[^A-Za-z0-9 _-]+/g, '').trim().slice(0, 40) || 'presentation';
		const file = path.join(this.dir, safe + ' ' + Date.now() + '.' + ext);
		fs.writeFileSync(file, buf);
		return file;
	}

	/** Delete downloaded presentations (PowerPoint may still hold the newest one open; that's fine). */
	tidy() {
		try {
			for (const f of fs.readdirSync(this.dir)) {
				try { fs.unlinkSync(path.join(this.dir, f)); } catch (e) { /* in use */ }
			}
		} catch (e) { /* no folder yet */ }
	}
}

module.exports = { Helper };

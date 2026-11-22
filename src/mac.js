// Mac: drives Keynote or PowerPoint with AppleScript. The first time, macOS asks
// "PPG Presenter Helper wants to control Keynote" — click OK.

const { execFile } = require('child_process');

function osa(script, ms) {
	return new Promise((ok, no) => {
		execFile('osascript', ['-e', script], { timeout: ms || 10000 }, (err, out, errOut) => {
			if (err) {
				const m = String(errOut || err.message || '');
				if (/-1743|not allowed/i.test(m)) { return no(new Error('Allow the helper to control the app: System Settings › Privacy & Security › Automation')); }
				if (/-1728|-600|isn.t running|Can.t get application/i.test(m)) { return no(new Error('The app isn’t installed on this Mac')); }
				return no(new Error(m.replace(/^.*execution error: /, '').trim().slice(0, 160) || 'AppleScript error'));
			}
			ok(String(out).trim());
		});
	});
}

const q = (s) => '"' + String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';

function hasApp(id) {
	return osa('id of application id ' + q(id), 5000).then(() => true, () => false);
}

class Keynote {
	async open(file) {
		await osa('tell application "Keynote"\nactivate\nset d to open (POSIX file ' + q(file) + ')\ndelay 0.5\nstart d from first slide of d\nend tell', 90000);
	}
	next() { return osa('tell application "Keynote" to show next'); }
	prev() { return osa('tell application "Keynote" to show previous'); }
	async close() {
		await osa('if application "Keynote" is running then\ntell application "Keynote"\nif playing then stop front document\nif (count of documents) > 0 then close front document saving no\nend tell\nend if');
	}
	async state() {
		const r = await osa('if application "Keynote" is running then\ntell application "Keynote"\nif playing then return ((slide number of current slide of front document) as text) & "," & ((count of slides of front document) as text)\nend tell\nend if\nreturn ""', 4000);
		if (!r) { return { mode: 'idle', slide: 0, total: 0 }; }
		const [s, t] = r.split(',').map(Number);
		return { mode: 'key', slide: s, total: t };
	}
}

class PowerPointMac {
	async open(file) {
		await osa('tell application "Microsoft PowerPoint"\nactivate\nopen (POSIX file ' + q(file) + ')\ndelay 1\nset p to active presentation\nrun slide show slide show settings of p\nend tell', 90000);
	}
	next() { return osa('tell application "Microsoft PowerPoint" to go to next slide slide show view of slide show window 1'); }
	prev() { return osa('tell application "Microsoft PowerPoint" to go to previous slide slide show view of slide show window 1'); }
	async close() {
		await osa('if application "Microsoft PowerPoint" is running then\ntell application "Microsoft PowerPoint"\ntry\nexit slide show slide show view of slide show window 1\nend try\ntry\nclose active presentation saving no\nend try\nend tell\nend if');
	}
	async state() {
		const r = await osa('if application "Microsoft PowerPoint" is running then\ntell application "Microsoft PowerPoint"\nif (count of slide show windows) > 0 then return ((slide index of slide of slide show view of slide show window 1) as text) & "," & ((count of slides of active presentation) as text)\nend tell\nend if\nreturn ""', 4000);
		if (!r) { return { mode: 'idle', slide: 0, total: 0 }; }
		const [s, t] = r.split(',').map(Number);
		return { mode: 'ppt', slide: s, total: t };
	}
}

module.exports = { Keynote, PowerPointMac, hasApp };

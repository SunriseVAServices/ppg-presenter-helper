// Windows: drives the installed PowerPoint through its own automation (COM), from one PowerShell
// that stays open. We send it one JSON line per command and it answers with one JSON line.

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const PS = String.raw`
$ErrorActionPreference = 'Stop'
$ppt = $null; $pres = $null; $mine = $false
function Say($o) { [Console]::Out.WriteLine(($o | ConvertTo-Json -Compress)); [Console]::Out.Flush() }
function View() { if ($pres -eq $null) { return $null }; try { if ($pres.SlideShowWindow -ne $null) { return $pres.SlideShowWindow.View } } catch {}; return $null }
function State() {
  $v = View
  if ($v -eq $null) { return @{ ok = $true; mode = 'idle'; slide = 0; total = 0 } }
  $n = $pres.Slides.Count; $s = $n
  try { if ($v.State -ne 5) { $s = $v.Slide.SlideIndex } } catch {}
  return @{ ok = $true; mode = 'ppt'; slide = $s; total = $n }
}
while ($true) {
  $line = [Console]::In.ReadLine()
  if ($line -eq $null) { break }
  try {
    $c = $line | ConvertFrom-Json
    switch ($c.cmd) {
      'open' {
        if ($ppt -eq $null) {
          try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $ppt = New-Object -ComObject PowerPoint.Application; $mine = $true }
        }
        $pres = $ppt.Presentations.Open($c.path, -1, 0, 0)
        $set = $pres.SlideShowSettings
        $set.ShowType = 1
        $set.StartingSlide = 1
        $set.EndingSlide = $pres.Slides.Count
        $set.RangeType = 1
        $null = $set.Run()
        Start-Sleep -Milliseconds 300
        try { $pres.SlideShowWindow.Activate() } catch {}
        Say (State)
      }
      'next' { $v = View; if ($v -ne $null) { $v.Next() }; Say (State) }
      'prev' { $v = View; if ($v -ne $null) { $v.Previous() }; Say (State) }
      'close' {
        if ($pres -ne $null) {
          try { $v = View; if ($v -ne $null) { $v.Exit() } } catch {}
          try { $pres.Saved = -1; $pres.Close() } catch {}
          $pres = $null
        }
        if ($mine -and $ppt -ne $null) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {}; $ppt = $null; $mine = $false }
        Say @{ ok = $true; mode = 'idle'; slide = 0; total = 0 }
      }
      'state' { Say (State) }
      default { Say @{ ok = $false; err = 'Unknown command' } }
    }
  } catch {
    $m = $_.Exception.Message
    if ($m -match '80040154|Invalid class string|not registered') { $m = 'PowerPoint isn''t installed on this computer' }
    Say @{ ok = $false; err = $m }
  }
}
`;

class PowerPointWin {
	constructor(dir) {
		this.dir = dir;
		this.ps = null;
		this.wait = [];
		this.buf = '';
		this.chain = Promise.resolve();
	}

	boot() {
		if (this.ps && this.ps.exitCode === null) { return; }
		fs.mkdirSync(this.dir, { recursive: true });
		const file = path.join(this.dir, 'ppg-powerpoint.ps1');
		fs.writeFileSync(file, '﻿' + PS, 'utf8');
		this.ps = spawn('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', file], { windowsHide: true });
		this.buf = '';
		this.ps.stdout.setEncoding('utf8');
		this.ps.stdout.on('data', (d) => {
			this.buf += d;
			let i;
			while ((i = this.buf.indexOf('\n')) >= 0) {
				const line = this.buf.slice(0, i).trim();
				this.buf = this.buf.slice(i + 1);
				if (!line) { continue; }
				const w = this.wait.shift();
				if (w) { try { w.ok(JSON.parse(line)); } catch (e) { w.no(new Error('PowerPoint answered strangely.')); } }
			}
		});
		this.ps.on('exit', () => { this.wait.splice(0).forEach((w) => w.no(new Error('PowerPoint control stopped.'))); });
	}

	/** One command at a time; each gets one answer. */
	send(cmd, extra, ms) {
		const p = this.chain.then(() => new Promise((ok, no) => {
			this.boot();
			const t = setTimeout(() => { no(new Error('PowerPoint didn’t respond.')); try { this.ps.kill(); } catch (e) { /* gone */ } }, ms || 8000);
			this.wait.push({ ok: (j) => { clearTimeout(t); j.ok === false ? no(new Error(j.err || 'PowerPoint error')) : ok(j); }, no: (e) => { clearTimeout(t); no(e); } });
			this.ps.stdin.write(JSON.stringify(Object.assign({ cmd }, extra || {})) + '\n');
		}));
		this.chain = p.catch(() => {});
		return p;
	}

	open(file) { return this.send('open', { path: file }, 60000); }
	next() { return this.send('next'); }
	prev() { return this.send('prev'); }
	close() { return this.ps ? this.send('close', {}, 15000) : Promise.resolve(); }
	state() { return this.ps ? this.send('state', {}, 5000) : Promise.resolve({ mode: 'idle', slide: 0, total: 0 }); }
	quit() { try { this.ps && this.ps.kill(); } catch (e) { /* gone */ } }
}

module.exports = { PowerPointWin };

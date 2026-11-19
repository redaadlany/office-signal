#==APP==#
# Office Signal app - downloaded and run by the Office Signal launcher.
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic
try {
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ---- who am I (asked once, saved) ----
$global:dir = Join-Path $env:APPDATA 'OfficeSignal'
New-Item -ItemType Directory -Force -Path $global:dir | Out-Null
$cfg = Join-Path $global:dir 'me.txt'
$global:myName = $null; $global:myCid = $null
if (Test-Path $cfg) { $lines = @(Get-Content $cfg -Encoding UTF8); if ($lines.Count -ge 2) { $global:myName = $lines[0].Trim(); $global:myCid = $lines[1].Trim() } }
if (-not $global:myName) {
  $n = [Microsoft.VisualBasic.Interaction]::InputBox('Your name (your colleague will see it):', 'Office Signal', $env:USERNAME)
  if (-not $n -or -not $n.Trim()) { return }
  $global:myName = $n.Trim()
  $global:myCid = [guid]::NewGuid().ToString('N').Substring(0, 12)
  [IO.File]::WriteAllLines($cfg, [string[]]@($global:myName, $global:myCid), (New-Object Text.UTF8Encoding $false))
}

# ---- the window page ----
$global:htmlPath = Join-Path $global:dir 'mail.html'
$html = @'
<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Messages</title>
<style>
  *{box-sizing:border-box;min-width:0}
  html,body{margin:0;height:100%}
  body{background:#000;color:#e7e9ea;font:15px/1.4 "Segoe UI",system-ui,sans-serif;overflow:hidden}
  svg{width:20px;height:20px;flex:none}
  button{font:inherit;color:inherit}
  #panel{position:fixed;inset:0;display:flex;flex-direction:column;background:#16181c}
  header{display:flex;align-items:center;gap:10px;padding:14px 12px 12px 18px;border-bottom:1px solid #2f3336;flex:none}
  .ttl{flex:1}
  .ttl b{font-size:19px;font-weight:700;display:block;line-height:1.2}
  .ttl small{font-size:13px;color:#71767b;display:flex;align-items:center;gap:6px}
  .dot{width:8px;height:8px;border-radius:50%;background:#71767b}.dot.on{background:#00ba7c}
  .ic{width:38px;height:38px;border-radius:8px;border:1px solid transparent;background:transparent;display:flex;align-items:center;justify-content:center;cursor:pointer;color:#e7e9ea}
  .ic:hover{background:#2a2e33}.ic.out{border-color:#1d9bf0}
  .sigs{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;padding:12px;border-bottom:1px solid #2f3336;flex:none}
  .sigs button{height:44px;border:0;border-radius:999px;font-weight:700;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px}
  .sigs button:active{transform:scale(.96)}
  .s-red{background:#f4212e;color:#fff}.s-yel{background:#ffd400;color:#0f1419}.s-grn{background:#00ba7c;color:#fff}
  #thread{flex:1;overflow-y:auto;padding:14px 14px 6px;display:flex;flex-direction:column;gap:4px}
  #thread::-webkit-scrollbar{width:8px}#thread::-webkit-scrollbar-thumb{background:#2f3336;border-radius:4px}
  .row{display:flex;flex-direction:column;max-width:78%}
  .row.me{align-self:flex-end;align-items:flex-end}.row.them{align-self:flex-start;align-items:flex-start}
  .bub{padding:9px 14px;border-radius:20px;white-space:pre-wrap;overflow-wrap:anywhere}
  .me .bub{background:#1d9bf0;color:#fff;border-bottom-right-radius:6px}
  .them .bub{background:#2f3336;border-bottom-left-radius:6px}
  .meta{font-size:12px;color:#71767b;margin:3px 6px 8px}
  .sig{align-self:center;display:flex;flex-direction:column;align-items:center;margin:8px 0}
  .chip{padding:7px 16px;border-radius:999px;font-weight:700;font-size:13px;text-align:center}
  .chip.red{background:#f4212e;color:#fff}.chip.yellow{background:#ffd400;color:#0f1419}.chip.green{background:#00ba7c;color:#fff}
  .empty{margin:auto;color:#71767b;text-align:center}
  .comp{padding:10px 12px 14px;border-top:1px solid #2f3336;flex:none}
  .box{display:flex;align-items:flex-end;gap:6px;background:#202327;border-radius:16px;padding:6px 6px 6px 14px}
  .box textarea{flex:1;background:transparent;border:0;outline:0;color:#e7e9ea;resize:none;font:15px/1.4 "Segoe UI",sans-serif;height:24px;max-height:110px;padding:6px 0}
  .box textarea::placeholder{color:#71767b}
  .send{width:36px;height:36px;border:0;border-radius:50%;background:transparent;color:#1d9bf0;cursor:pointer;display:flex;align-items:center;justify-content:center}
  .send:hover{background:#1d9bf01a}
  #status{font-size:12px;color:#f4212e;margin:6px 4px 0}
  #status:empty{display:none}
  .av{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;color:#fff;border:2px solid #202327}
  #pill{position:fixed;inset:0;display:none;align-items:center;justify-content:center;background:#000;padding:10px}
  #pill button{width:100%;max-width:310px;height:62px;border-radius:999px;background:#202327;border:1px solid #2f3336;display:flex;align-items:center;gap:12px;padding:0 14px 0 22px;cursor:pointer;box-shadow:0 0 18px rgba(255,255,255,.08)}
  #pill b{font-size:18px;flex:1;text-align:left}
  .stack{display:flex}.stack .av{margin-left:-10px}
  .badge{background:#1d9bf0;color:#fff;font-size:12px;font-weight:700;border-radius:999px;padding:1px 7px;display:none}
  .toast{position:fixed;left:50%;bottom:90px;transform:translateX(-50%);background:#e7e9ea;color:#0f1419;font-weight:600;padding:8px 16px;border-radius:999px;font-size:13px;display:none;z-index:7;white-space:nowrap}
  .dlg{position:fixed;inset:0;background:rgba(91,112,131,.4);display:none;align-items:center;justify-content:center;z-index:8;padding:20px}
  .dlg .box2{background:#000;border-radius:16px;padding:24px;width:min(320px,100%);text-align:center}
  .dlg h3{margin:0 0 6px;font-size:19px}.dlg p{margin:0 0 18px;color:#71767b}
  .dlg button{display:block;width:100%;height:44px;border-radius:999px;font-weight:700;cursor:pointer;margin-top:10px}
  #yes{background:#f4212e;border:0;color:#fff}#no{background:transparent;border:1px solid #536471;color:#e7e9ea}
  #ov{position:fixed;inset:0;display:none;flex-direction:column;align-items:center;justify-content:center;gap:10px;text-align:center;z-index:9;cursor:pointer;padding:20px}
  #ov h1{font-size:40px;margin:0;line-height:1.1}#ov p{font-size:18px;margin:0}
  #ov.red{background:#f4212e;color:#fff;animation:fl .25s infinite alternate}#ov.yellow{background:#ffd400;color:#0f1419}#ov.green{background:#00ba7c;color:#fff}
  @keyframes fl{to{background:#7a0000}}
  @keyframes bl{0%,100%{background:#f4212e;border-color:#f4212e}50%{background:#7a0000;border-color:#7a0000}}
  #pill.alert button{animation:bl .9s infinite;color:#fff}
  body.wide .row{max-width:62%}
</style></head><body>
<div id="panel">
  <header>
    <div class="ttl"><b>Messages</b><small><span class="dot" id="dot"></span><span id="conn">Connecting...</span></small></div>
    <button class="ic out" id="expand" type="button" title="Bigger window"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 3h5v5M8 17H3v-5"/></svg></button>
    <button class="ic" id="close" type="button" title="Minimize to bar"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m4 4 12 12M16 4 4 16"/></svg></button>
  </header>
  <div class="sigs">
    <button class="s-red" data-c="red" type="button" title="A red bar blinks on his screen">Emergency</button>
    <button class="s-yel" data-c="yellow" type="button" title="His screen turns yellow">Come</button>
    <button class="s-grn" data-c="green" type="button" title="His screen turns green">Call off</button>
  </div>
  <div id="thread"></div>
  <div class="comp">
    <form id="f" class="box"><textarea id="txt" rows="1" placeholder="Start a new message"></textarea>
      <button class="send" title="Send"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M2.5 9.5 17.5 3l-6.5 15-2.2-6.3z"/></svg></button></form>
    <div id="status"></div>
  </div>
</div>
<div id="pill"><button type="button" id="pillBtn"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M2.5 9.5 17.5 3l-6.5 15-2.2-6.3z"/></svg><b>Messages</b><span class="badge" id="badge"></span><span class="stack" id="stack"></span></button></div>
<div class="toast" id="toast"></div>
<div class="dlg" id="dlg"><div class="box2"><h3>Send emergency?</h3><p>A red Messages bar will blink on your colleague's screen.</p><button id="yes" type="button">Send emergency</button><button id="no" type="button">Cancel</button></div></div>
<div id="ov"><h1 id="ovt"></h1><p id="ovm"></p></div>
<script>
{
const TOPIC = '__TOPIC__';
const SERVERS = ['https://ntfy.sh/' + TOPIC, 'https://ntfy.envs.net/' + TOPIC];
const LABEL = {red:'EMERGENCY!', yellow:'COME PLEASE', green:'CALL OFF - BREAK'};
const CHIP = {red:'Emergency', yellow:'Come please', green:'Call off - break'};
const DOT = ' - ';
const $ = id => document.getElementById(id);
const hp = new URLSearchParams(location.hash.slice(1));
const DESKTOP = !!hp.get('c');
let name = hp.get('n') || localStorage.getItem('sigName');
while(!name){ name = (prompt('Your name?')||'').trim(); }
localStorage.setItem('sigName', name);
let cid = hp.get('c') || localStorage.getItem('sigCid');
if(!cid){ cid = Math.random().toString(36).slice(2,14); localStorage.setItem('sigCid', cid); }

const COLORS = ['#7856ff','#00ba7c','#f91880','#ff7a00','#1d9bf0','#d4a500'];
const colorOf = s => { let h=0; for(const ch of s) h=(h*31+ch.charCodeAt(0))|0; return COLORS[Math.abs(h)%COLORS.length]; };
const ini = s => s.trim().split(/\s+/).map(w=>w[0]).join('').slice(0,2).toUpperCase() || '?';
const thread = $('thread'), txt = $('txt'), st = $('status');
const seen = new Set(), people = new Set();
let unread = 0;
const isMine = m => m.c ? m.c === cid : m.name === name;
const fmt = t => { const d = new Date(t*1000), now = new Date();
  const hm = d.toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'});
  return d.toDateString()===now.toDateString() ? hm : d.toLocaleDateString([], {weekday:'short'}) + ' ' + hm; };

function renderStack(){
  const s = $('stack'); s.innerHTML = '';
  [...people].slice(-3).forEach(p => { const a = document.createElement('span'); a.className='av'; a.style.background = colorOf(p); a.textContent = ini(p); a.title = p; s.appendChild(a); });
}
function add(raw, live){
  if(!raw || raw.event !== 'message') return;
  let m; try{ m = JSON.parse(raw.message); }catch(e){ return; }
  if(!m || typeof m.name !== 'string') return;
  const key = m.k || raw.id;
  if(seen.has(key)) return;
  seen.add(key);
  const mine = isMine(m);
  if(!mine && !people.has(m.name)){ people.add(m.name); renderStack(); }
  const e = thread.querySelector('.empty'); if(e) e.remove();
  const atBottom = thread.scrollHeight - thread.scrollTop - thread.clientHeight < 80;
  const meta = document.createElement('div'); meta.className = 'meta';
  meta.textContent = (mine ? '' : m.name + DOT) + fmt(raw.time);
  let el;
  if(CHIP[m.type]){
    el = document.createElement('div'); el.className = 'sig';
    const c = document.createElement('div'); c.className = 'chip ' + m.type;
    c.textContent = (mine ? 'You sent: ' : '') + CHIP[m.type] + (m.text ? DOT + m.text : '');
    el.append(c, meta);
  } else {
    el = document.createElement('div'); el.className = 'row ' + (mine ? 'me' : 'them');
    const b = document.createElement('div'); b.className = 'bub'; b.textContent = m.text || '';
    el.append(b, meta);
  }
  thread.appendChild(el);
  if(!live || atBottom || mine) thread.scrollTop = thread.scrollHeight;
  if(live && !mine && Date.now()/1000 - raw.time < 600){
    if(document.hidden || $('pill').style.display === 'flex'){ unread++; $('badge').textContent = unread; $('badge').style.display = 'inline-block'; document.title = '(' + unread + ') Messages'; }
    if(m.type === 'red'){ $('panel').style.display='none'; $('pill').style.display='flex'; $('pill').classList.add('alert'); if(!unread){ unread=1; $('badge').textContent='!'; $('badge').style.display='inline-block'; } }
    else if(CHIP[m.type] && !DESKTOP) flash(m.type, 'From ' + m.name + (m.text ? ': ' + m.text : ''));
  }
}
function clearUnread(){ unread = 0; $('badge').style.display = 'none'; document.title = 'Messages'; }
document.addEventListener('visibilitychange', () => { if(!document.hidden && $('pill').style.display !== 'flex') clearUnread(); });

let ovTimer = null;
function flash(type, line){
  const ov = $('ov'); ov.className = type; ov.style.display = 'flex';
  $('ovt').textContent = LABEL[type]; $('ovm').textContent = line;
  clearTimeout(ovTimer); ovTimer = setTimeout(() => { ov.style.display='none'; }, 2000);
}
$('ov').addEventListener('click', () => { $('ov').style.display='none'; });
let toastT = null;
function toast(s){ const t = $('toast'); t.textContent = s; t.style.display = 'block'; clearTimeout(toastT); toastT = setTimeout(()=>{ t.style.display='none'; }, 1800); }

function sizeWin(w, h){ try{ window.resizeTo(w, h); window.moveTo(screen.availLeft + screen.availWidth - w - 16, screen.availTop + screen.availHeight - h - 16); }catch(e){} }
let wide = false;
$('expand').addEventListener('click', () => { wide = !wide; document.body.classList.toggle('wide', wide); sizeWin(wide ? 760 : 440, wide ? 820 : 680); });
$('close').addEventListener('click', () => { $('panel').style.display='none'; $('pill').style.display='flex'; sizeWin(360, 120); });
$('pillBtn').addEventListener('click', () => { $('pill').classList.remove('alert'); $('pill').style.display='none'; $('panel').style.display='flex'; clearUnread(); sizeWin(wide ? 760 : 440, wide ? 820 : 680); thread.scrollTop = thread.scrollHeight; txt.focus(); });

const lastIds = {}, live = {};
function timeout(ms){ const c = new AbortController(); setTimeout(() => c.abort(), ms); return c.signal; }
async function loadHistory(){
  const all = [];
  await Promise.all(SERVERS.map(async s => {
    try{
      const r = await fetch(s + '/json?poll=1&since=12h', {signal: timeout(8000)});
      (await r.text()).split('\n').forEach(l => { if(l.trim()){ try{ const j = JSON.parse(l); all.push(j); lastIds[s] = j.id; }catch(e){} } });
    }catch(e){}
  }));
  all.sort((a,b) => a.time - b.time).forEach(j => add(j, false));
  if(!thread.children.length){ thread.innerHTML = '<div class="empty">No messages yet.<br>Say hello below.</div>'; }
}
function setConn(){
  const n = Object.values(live).filter(Boolean).length;
  $('conn').textContent = n ? (people.size ? [...people].slice(-1)[0] + DOT + 'connected' : 'Connected') : 'Offline - retrying...';
  $('dot').className = 'dot' + (n ? ' on' : '');
}
function connect(s){
  const es = new EventSource(s + '/sse?since=' + (lastIds[s] || 'now'));
  es.onopen = () => { live[s] = true; setConn(); };
  es.onmessage = e => { try{ const j = JSON.parse(e.data); if(j.id) lastIds[s] = j.id; add(j, true); setConn(); }catch(err){} };
  es.onerror = () => { live[s] = false; setConn(); es.close(); setTimeout(() => connect(s), 4000); };
}
function post(o){
  o.name = name; o.c = cid; o.k = cid + '-' + Date.now() + '-' + Math.random().toString(36).slice(2,7);
  const body = JSON.stringify(o);
  return Promise.any(SERVERS.map(s => fetch(s, {method:'POST', body, signal: timeout(8000)}).then(r => { if(!r.ok) throw new Error(r.status); })));
}
function grow(){ txt.style.height = '24px'; txt.style.height = Math.min(txt.scrollHeight, 110) + 'px'; }
txt.addEventListener('input', grow);
async function send(){
  const text = txt.value.trim(); if(!text) return;
  txt.value = ''; grow();
  try{ await post({text}); st.textContent=''; }catch(e){ st.textContent = 'Not sent - check the internet and try again.'; txt.value = text; grow(); }
}
$('f').addEventListener('submit', e => { e.preventDefault(); send(); });
txt.addEventListener('keydown', e => { if(e.key==='Enter' && !e.shiftKey){ e.preventDefault(); send(); } });
async function fire(c){
  const note = txt.value.trim(); txt.value = ''; grow();
  try{ await post({type:c, text:note}); toast('Sent: ' + CHIP[c]); }
  catch(e){ st.textContent = 'SIGNAL NOT SENT - check the internet!'; txt.value = note; grow(); }
}
document.querySelectorAll('.sigs button').forEach(b => b.addEventListener('click', () => {
  if(b.dataset.c === 'red'){ $('dlg').style.display = 'flex'; $('yes').focus(); } else fire(b.dataset.c);
}));
$('no').addEventListener('click', () => { $('dlg').style.display='none'; });
$('yes').addEventListener('click', () => { $('dlg').style.display='none'; fire('red'); });
setConn(); loadHistory().then(() => { setConn(); SERVERS.forEach(connect); });
}
</script></body></html>
'@
$html = $html.Replace('__TOPIC__', $global:topic)
[IO.File]::WriteAllText($global:htmlPath, $html, (New-Object Text.UTF8Encoding $false))

function Open-Window {
  $u = ([Uri]$global:htmlPath).AbsoluteUri + '#n=' + [Uri]::EscapeDataString($global:myName) + '&c=' + $global:myCid
  $wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
  $x = $wa.Right - 440 - 16; $y = $wa.Bottom - 680 - 16
  $a = @("--app=$u", '--window-size=440,680', "--window-position=$x,$y")
  try { Start-Process 'msedge' -ArgumentList $a -ErrorAction Stop; return } catch {}
  try { Start-Process 'chrome' -ArgumentList $a -ErrorAction Stop; return } catch {}
  Start-Process $u
}

# ---- only one copy running: if already running, just open the window ----
$created = $false
$global:mutex = New-Object System.Threading.Mutex($true, 'Local\OfficeSignalApp', [ref]$created)
if (-not $created) { Open-Window; return }

# ---- full-screen colour flash (click-through, does not steal focus) ----
$cs = @"
using System.Windows.Forms;
public class SignalFlashForm : Form {
  protected override bool ShowWithoutActivation { get { return true; } }
  protected override CreateParams CreateParams { get { CreateParams cp = base.CreateParams; cp.ExStyle |= 0x00080000 | 0x00000020 | 0x00000080 | 0x00000008; return cp; } }
}
public class AlertBarForm : Form {
  protected override bool ShowWithoutActivation { get { return true; } }
  protected override CreateParams CreateParams { get { CreateParams cp = base.CreateParams; cp.ExStyle |= 0x00000080 | 0x00000008; return cp; } }
}
"@
$global:flashType = 'System.Windows.Forms.Form'; $global:barType = 'System.Windows.Forms.Form'
try { Add-Type -TypeDefinition $cs -ReferencedAssemblies System.Windows.Forms -ErrorAction Stop; $global:flashType = 'SignalFlashForm'; $global:barType = 'AlertBarForm' } catch {}

$global:titles = @{ red = 'EMERGENCY!'; yellow = 'COME PLEASE'; green = 'CALL OFF - BREAK' }
$global:colors = @{ red = [Drawing.Color]::FromArgb(214, 40, 40); yellow = [Drawing.Color]::FromArgb(242, 194, 0); green = [Drawing.Color]::FromArgb(31, 157, 85) }
$global:darkRed = [Drawing.Color]::FromArgb(122, 0, 0)
$global:flash = $null; $global:flashKind = ''; $global:flashStart = Get-Date

$global:flashTimer = New-Object Windows.Forms.Timer
$global:flashTimer.Interval = 125
$global:flashTimer.Add_Tick({
  if (-not $global:flash) { $global:flashTimer.Stop(); return }
  $el = ((Get-Date) - $global:flashStart).TotalMilliseconds
  if ($el -ge 2000) { $global:flashTimer.Stop(); try { $global:flash.Close(); $global:flash.Dispose() } catch {}; $global:flash = $null; return }
  if ($global:flashKind -eq 'red') {
    if ([int][Math]::Floor($el / 250) % 2 -eq 1) { $global:flash.BackColor = $global:darkRed } else { $global:flash.BackColor = $global:colors['red'] }
  }
})

function Show-Flash($kind, $line) {
  if ($global:flash) { try { $global:flash.Close(); $global:flash.Dispose() } catch {} }
  $fm = New-Object $global:flashType
  $fm.FormBorderStyle = 'None'; $fm.ShowInTaskbar = $false; $fm.TopMost = $true
  $fm.StartPosition = 'Manual'; $fm.Bounds = [Windows.Forms.SystemInformation]::VirtualScreen
  $fm.BackColor = $global:colors[$kind]; $fm.Opacity = 0.8
  $lb = New-Object Windows.Forms.Label
  $lb.Dock = 'Fill'; $lb.TextAlign = 'MiddleCenter'; $lb.BackColor = [Drawing.Color]::Transparent
  $lb.Font = New-Object Drawing.Font('Segoe UI', 40, [Drawing.FontStyle]::Bold)
  if ($kind -eq 'yellow') { $lb.ForeColor = [Drawing.Color]::FromArgb(30, 30, 30) } else { $lb.ForeColor = [Drawing.Color]::White }
  $lb.Text = $global:titles[$kind] + "`r`n" + $line
  $fm.Controls.Add($lb)
  $global:flash = $fm; $global:flashKind = $kind; $global:flashStart = Get-Date
  $fm.Show()
  $global:flashTimer.Start()
}

# ---- emergency: small red Messages bar that blinks until clicked ----
$global:bar = $null; $global:barOn = $false
$global:barTimer = New-Object Windows.Forms.Timer
$global:barTimer.Interval = 450
$global:barTimer.Add_Tick({
  if (-not $global:bar) { $global:barTimer.Stop(); return }
  $global:barOn = -not $global:barOn
  if ($global:barOn) { $global:bar.BackColor = $global:darkRed } else { $global:bar.BackColor = $global:colors['red'] }
  $global:bar.TopMost = $true
})
function Close-AlertBar {
  $global:barTimer.Stop()
  if ($global:bar) { try { $global:bar.Close(); $global:bar.Dispose() } catch {}; $global:bar = $null }
}
function Show-AlertBar($from, $text) {
  Close-AlertBar
  $fm = New-Object $global:barType
  $fm.FormBorderStyle = 'None'; $fm.ShowInTaskbar = $false; $fm.TopMost = $true; $fm.StartPosition = 'Manual'
  $w = 330; $h = 58
  $wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
  $fm.Bounds = New-Object Drawing.Rectangle(($wa.Right - $w - 16), ($wa.Bottom - $h - 16), $w, $h)
  $fm.BackColor = $global:colors['red']
  $gp = New-Object Drawing.Drawing2D.GraphicsPath
  $gp.AddArc(0, 0, $h, $h, 90, 180); $gp.AddArc($w - $h, 0, $h, $h, 270, 180); $gp.CloseFigure()
  $fm.Region = New-Object Drawing.Region($gp)
  $lb = New-Object Windows.Forms.Label
  $lb.Dock = 'Fill'; $lb.TextAlign = 'MiddleCenter'; $lb.AutoEllipsis = $true
  $lb.ForeColor = [Drawing.Color]::White; $lb.BackColor = [Drawing.Color]::Transparent
  $lb.Font = New-Object Drawing.Font('Segoe UI', 12, [Drawing.FontStyle]::Bold)
  $lb.Cursor = [Windows.Forms.Cursors]::Hand
  $t = "Messages  -  Emergency from $from"; if ($text) { $t += ": $text" }
  $lb.Text = $t
  $lb.Add_Click({ Close-AlertBar; Open-Window })
  $fm.Controls.Add($lb)
  $global:bar = $fm; $global:barOn = $false
  $fm.Show()
  $global:barTimer.Start()
}

# ---- tray icon ----
$global:ni = New-Object Windows.Forms.NotifyIcon
$global:ni.Icon = [Drawing.SystemIcons]::Information
$global:ni.Text = 'Office Signal'
$global:ni.Visible = $true
$menu = New-Object Windows.Forms.ContextMenuStrip
[void]$menu.Items.Add('Open', $null, { Open-Window })
[void]$menu.Items.Add('Quit', $null, { $global:ni.Visible = $false; [Windows.Forms.Application]::Exit() })
$global:ni.ContextMenuStrip = $menu
$global:ni.Add_DoubleClick({ Open-Window })
$global:ni.Add_BalloonTipClicked({ Open-Window })

# ---- listen to the channel in the background (two servers, in case one is blocked) ----
$global:q = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
$global:seenKeys = New-Object 'System.Collections.Generic.HashSet[string]'
$global:bgs = @()
$listener = {
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
  $last = $null
  while ($true) {
    try {
      $u = $base; if ($last) { $u = "$($base)?since=$last" }
      $req = [Net.HttpWebRequest]::Create($u)
      $req.Timeout = 30000; $req.ReadWriteTimeout = 120000
      $resp = $req.GetResponse()
      $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
      while ($true) {
        $line = $sr.ReadLine()
        if ($null -eq $line) { break }
        if ($line.Contains('"event":"message"')) {
          if ($line -match '"id":"([^"]+)"') { $last = $matches[1] }
          $q.Enqueue($line)
        }
      }
      $resp.Close()
    } catch {}
    Start-Sleep -Seconds 3
  }
}
foreach ($srv in @('https://ntfy.sh', 'https://ntfy.envs.net')) {
  $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
  $rs.SessionStateProxy.SetVariable('q', $global:q)
  $rs.SessionStateProxy.SetVariable('base', "$srv/$($global:topic)/json")
  $bg = [powershell]::Create(); $bg.Runspace = $rs
  [void]$bg.AddScript($listener)
  [void]$bg.BeginInvoke()
  $global:bgs += $bg
}

function Handle-Line($line) {
  try { $e = $line | ConvertFrom-Json } catch { return }
  if ($e.event -ne 'message') { return }
  try { $m = $e.message | ConvertFrom-Json } catch { return }
  $key = '' + $m.k; if (-not $key) { $key = '' + $e.id }
  if (-not $global:seenKeys.Add($key)) { return }
  $from = ('' + $m.name).Trim(); if (-not $from) { $from = 'Colleague' }
  if ($m.c) { $mine = ('' + $m.c) -eq $global:myCid } else { $mine = $from -eq $global:myName }
  if ($mine) { return }
  $kind = '' + $m.type; $text = ('' + $m.text).Trim()
  if ($kind -eq 'red') {
    Show-AlertBar $from $text
  } elseif ($global:titles.ContainsKey($kind)) {
    $line2 = 'From ' + $from; if ($text) { $line2 += ': ' + $text }
    Show-Flash $kind $line2
    $global:ni.ShowBalloonTip(6000, $global:titles[$kind], $line2, [Windows.Forms.ToolTipIcon]::Warning)
  } elseif ($text) {
    $global:ni.ShowBalloonTip(5000, "New message from $from", $text, [Windows.Forms.ToolTipIcon]::Info)
  }
}

$global:poll = New-Object Windows.Forms.Timer
$global:poll.Interval = 250
$global:poll.Add_Tick({
  $item = $null
  while ($global:q.TryDequeue([ref]$item)) { Handle-Line $item }
})
$global:poll.Start()

# ---- automatic updates: check the online copy every 20 minutes ----
function Restart-App {
  try { $global:ni.Visible = $false } catch {}
  try { $global:mutex.ReleaseMutex() } catch {}
  if ($global:launcher -and (Test-Path $global:launcher)) { Start-Process -FilePath $global:launcher }
  [Environment]::Exit(0)
}
$global:updTimer = New-Object Windows.Forms.Timer
$global:updTimer.Interval = 20 * 60 * 1000
$global:updTimer.Add_Tick({
  try {
    $wc = New-Object Net.WebClient; $wc.Encoding = [Text.Encoding]::UTF8
    $new = $wc.DownloadString($global:appUrl + '?t=' + [DateTime]::UtcNow.Ticks)
    if ($new -match '#==APP==#' -and (Test-Path $global:appCache)) {
      $cur = [IO.File]::ReadAllText($global:appCache)
      if ($new -ne $cur) {
        [IO.File]::WriteAllText($global:appCache, $new, (New-Object Text.UTF8Encoding $false))
        Restart-App
      }
    }
  } catch {}
})
if ($global:appUrl) { $global:updTimer.Start() }

Open-Window
$global:ni.ShowBalloonTip(4000, 'Office Signal is running', 'It stays here next to the clock. Double-click the icon to open Messages.', [Windows.Forms.ToolTipIcon]::Info)
[Windows.Forms.Application]::Run((New-Object Windows.Forms.ApplicationContext))
$global:ni.Dispose()
}
catch {
  [Windows.Forms.MessageBox]::Show("Office Signal could not start:`r`n`r`n$($_.Exception.Message)", 'Office Signal') | Out-Null
}
[Environment]::Exit(0)

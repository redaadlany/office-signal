#==APP==#
# Office Signal app - downloaded and run by the Office Signal launcher.
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic
try {
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$global:appVersion = '2026-11-23.1'

function Get-VerNum([string]$v) { try { $q = $v.Trim().Split('.'); return ([int64]($q[0].Replace('-', '')) * 1000 + [int64]$q[1]) } catch { return [int64]0 } }
function Get-TextVer([string]$t) { if ($t -match "appVersion = '([^']+)'") { return $matches[1] } else { return '' } }
$global:goodFile = Join-Path (Join-Path $env:APPDATA 'OfficeSignal') 'app.good.ps1'
# never run an older copy than the newest one that already worked
if (-not $global:fallbackTried -and (Test-Path $global:goodFile)) {
  $goodText = [IO.File]::ReadAllText($global:goodFile)
  if ((Get-VerNum (Get-TextVer $goodText)) -gt (Get-VerNum $global:appVersion)) { $global:fallbackTried = $true; iex $goodText; return }
}

# ---- only one copy running: if already running, ask it to show the window ----
$created = $false
$global:mutex = New-Object System.Threading.Mutex($true, 'Local\OfficeSignalApp', [ref]$created)
$verFile = Join-Path (Join-Path $env:APPDATA 'OfficeSignal') 'running.txt'
if (-not $created) {
  $runningVer = ''; try { $runningVer = ([IO.File]::ReadAllText($verFile)).Trim() } catch {}
  if ($runningVer -eq $global:appVersion) {
    try { ([System.Threading.EventWaitHandle]::OpenExisting('Local\OfficeSignalShow')).Set() | Out-Null } catch {}
    return
  }
  # a newer version was just downloaded: ask the old one to close, then take over
  try { ([System.Threading.EventWaitHandle]::OpenExisting('Local\OfficeSignalQuit')).Set() | Out-Null } catch {}
  $got = $false
  try { $got = $global:mutex.WaitOne(8000) } catch [System.Threading.AbandonedMutexException] { $got = $true }
  if (-not $got) {
    try { ([System.Threading.EventWaitHandle]::OpenExisting('Local\OfficeSignalShow')).Set() | Out-Null } catch {}
    return
  }
}
try { [IO.Directory]::CreateDirectory((Split-Path $verFile)) | Out-Null; [IO.File]::WriteAllText($verFile, $global:appVersion) } catch {}

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

# ---- the app itself (native Windows window) ----
$cs = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

public class OsMsg {
  public string Key = "", Name = "", C = "", Text = "", Type = "";
  public long Time, SeenT;
  public bool Mine;
  public string Url = "";
  public Image Img;
  public bool ImgFailed, ImgLoading;
}

public static class OsNative {
  [DllImport("user32.dll")] public static extern bool ReleaseCapture();
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);
  [DllImport("shell32.dll", CharSet = CharSet.Unicode)] public static extern int SetCurrentProcessExplicitAppUserModelID(string id);
  [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr ctx);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern int GetDeviceCaps(IntPtr dc, int index);
  public static void Drag(Form f) { ReleaseCapture(); SendMessage(f.Handle, 0xA1, (IntPtr)2, IntPtr.Zero); }
}

public static class OsUi {
  public static bool Light;
  public static Color Bg = Color.FromArgb(22, 24, 28);
  public static Color Box = Color.FromArgb(32, 35, 39);
  public static Color Bubble = Color.FromArgb(47, 51, 54);
  public static Color Line = Color.FromArgb(47, 51, 54);
  public static Color Ink = Color.FromArgb(231, 233, 234);
  public static Color Mute = Color.FromArgb(113, 118, 123);
  public static Color PillBg = Color.FromArgb(32, 35, 39);
  public static Color PillLine = Color.FromArgb(47, 51, 54);
  public static Color PillInk = Color.FromArgb(231, 233, 234);
  public static void SetTheme(bool light) {
    Light = light;
    if (light) {
      Bg = Color.FromArgb(250, 249, 245); Box = Color.White; Bubble = Color.FromArgb(240, 238, 230);
      Line = Color.FromArgb(229, 226, 216); Ink = Color.FromArgb(61, 57, 41); Mute = Color.FromArgb(135, 134, 127);
    } else {
      Bg = Color.FromArgb(38, 38, 36); Box = Color.FromArgb(48, 48, 46); Bubble = Color.FromArgb(20, 20, 19);
      Line = Color.FromArgb(62, 62, 58); Ink = Color.FromArgb(240, 238, 230); Mute = Color.FromArgb(166, 163, 154);
    }
  }
  public static Color Clay = Color.FromArgb(217, 119, 87);
  public static Color Hover() { return Mix(Bg, Ink, 0.09f); }
  public static void DrawCentered(Graphics g, string text, Font f, RectangleF r, Color c) {
    using (SolidBrush b = new SolidBrush(c)) using (StringFormat sf = new StringFormat()) {
      sf.Alignment = StringAlignment.Center; sf.LineAlignment = StringAlignment.Center;
      TextRenderingHint old = g.TextRenderingHint; g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
      g.DrawString(text, f, b, new RectangleF(r.X, r.Y + 0.5f, r.Width, r.Height), sf);
      g.TextRenderingHint = old;
    }
  }
  public static void DrawAvatar(Graphics g, string name, RectangleF r, Color ring, Font f) {
    Color c = AvatarColor(name);
    using (SolidBrush b = new SolidBrush(c)) g.FillEllipse(b, r);
    if (ring != Color.Empty) { using (Pen pn = new Pen(ring, 2f)) g.DrawEllipse(pn, r); }
    DrawCentered(g, Initial(name), f, r, Color.White);
  }
  public static Color Blue = Color.FromArgb(29, 155, 240);
  public static Color Red = Color.FromArgb(244, 33, 46);
  public static Color Yellow = Color.FromArgb(255, 212, 0);
  public static Color Green = Color.FromArgb(0, 186, 124);
  public static Color DarkText = Color.FromArgb(15, 20, 25);
  public static Color[] Avatars = { Color.FromArgb(120, 86, 255), Color.FromArgb(0, 186, 124), Color.FromArgb(249, 24, 128), Color.FromArgb(255, 122, 0), Color.FromArgb(29, 155, 240), Color.FromArgb(212, 165, 0) };

  public static GraphicsPath Round(RectangleF r, float rad) {
    GraphicsPath p = new GraphicsPath();
    float d = rad * 2;
    if (d > r.Height) d = r.Height;
    if (d > r.Width) d = r.Width;
    if (d < 1) { p.AddRectangle(r); return p; }
    p.AddArc(r.X, r.Y, d, d, 180, 90);
    p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
    p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
    p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
    p.CloseFigure();
    return p;
  }
  public static bool IsSignal(string k) { return k == "red" || k == "yellow" || k == "green"; }
  public static bool IsChip(string k) { return IsSignal(k) || k == "live" || k == "livestop"; }
  public static Color KindColor(string k) { if (k == "red") return Red; if (k == "yellow") return Yellow; if (k == "msg" || k == "live") return Blue; if (k == "livestop") return Bubble; return Green; }
  public static Color KindDark(string k) { if (k == "red") return Color.FromArgb(122, 0, 0); if (k == "yellow") return Color.FromArgb(122, 96, 0); if (k == "msg" || k == "live") return Color.FromArgb(11, 79, 128); return Color.FromArgb(0, 84, 56); }
  public static Color KindText(string k) { if (k == "livestop") return Ink; return k == "yellow" ? DarkText : Color.White; }
  public static string KindLabel(string k) { if (k == "red") return "Emergency"; if (k == "yellow") return "Come please"; return "Call off - break"; }
  public static Color AvatarColor(string s) {
    int h = 0;
    foreach (char ch in s) { h = unchecked(h * 31 + ch); }
    return Avatars[(int)((uint)h % (uint)Avatars.Length)];
  }
  public static string Initial(string s) { s = (s ?? "").Trim(); return s.Length > 0 ? s.Substring(0, 1).ToUpper() : "?"; }
  public static Color Mix(Color a, Color b, float t) {
    return Color.FromArgb((int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t), (int)(a.B + (b.B - a.B) * t));
  }
  public static Bitmap MakeBitmap(int size) {
    Bitmap bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
    using (Graphics g = Graphics.FromImage(bmp)) {
      g.SmoothingMode = SmoothingMode.AntiAlias;
      g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
      g.Clear(Color.Transparent);
      using (GraphicsPath p = Round(new RectangleF(0, 0, size - 1, size - 1), size * 0.2f))
      using (SolidBrush b = new SolidBrush(Color.FromArgb(51, 6, 10))) { g.FillPath(b, p); }
      using (Font f = new Font("Segoe UI Semibold", Math.Max(6f, size * 0.42f), FontStyle.Regular, GraphicsUnit.Pixel))
      using (SolidBrush tb = new SolidBrush(Color.FromArgb(255, 92, 92)))
      using (StringFormat sf = new StringFormat()) {
        sf.Alignment = StringAlignment.Center; sf.LineAlignment = StringAlignment.Center;
        g.DrawString("Os", f, tb, new RectangleF(0, size * 0.02f, size, size), sf);
      }
    }
    return bmp;
  }
  public const string JokeText = "nta 9090lian";
  static Image joke;
  public static Image JokeImage() {
    if (joke == null) { try { joke = Image.FromStream(new MemoryStream(Convert.FromBase64String(JokeB64))); } catch { } }
    return joke;
  }
  public const string JokeB64 = "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";
  public static Icon MakeIcon(int size) { Bitmap b = MakeBitmap(size); return Icon.FromHandle(b.GetHicon()); }
  public static string Clock(long unix) {
    DateTime d = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(unix).ToLocalTime();
    if (d.Date == DateTime.Now.Date) return d.ToString("HH:mm");
    return d.ToString("ddd HH:mm");
  }
  public static long Now() { return (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds; }
}

public class RoundBtn : Control {
  public Color Fill, Fore;
  public Color Dot = Color.Empty;
  bool hover;
  public RoundBtn(string text, Color fill, Color fore) {
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
    Text = text; Fill = fill; Fore = fore; Cursor = Cursors.Hand;
    Font = new Font("Segoe UI Semibold", 10.5f);
    BackColor = OsUi.Bg;
  }
  protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
  protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
    g.Clear(Parent != null ? Parent.BackColor : OsUi.Bg);
    if (Dot != Color.Empty) {
      Color bgc = Parent != null ? Parent.BackColor : OsUi.Bg;
      Color tf = OsUi.Mix(bgc, Dot, hover ? 0.26f : 0.15f);
      using (GraphicsPath p = OsUi.Round(new RectangleF(0.5f, 0.5f, Width - 2, Height - 2), Height / 2f)) {
        using (SolidBrush b = new SolidBrush(tf)) g.FillPath(b, p);
        using (Pen pn = new Pen(OsUi.Mix(bgc, Dot, 0.45f), 1f)) g.DrawPath(pn, p);
      }
      Size ts = TextRenderer.MeasureText(Text, Font);
      int tx = (Width - ts.Width - 14) / 2;
      using (SolidBrush db = new SolidBrush(Dot)) g.FillEllipse(db, tx, Height / 2f - 4, 8, 8);
      TextRenderer.DrawText(g, Text, Font, new Point(tx + 14, (Height - ts.Height) / 2), OsUi.Ink, tf);
      return;
    }
    Color c = hover ? OsUi.Mix(Fill, Color.White, 0.12f) : Fill;
    using (GraphicsPath p = OsUi.Round(new RectangleF(0, 0, Width - 1, Height - 1), Height / 2f))
    using (SolidBrush b = new SolidBrush(c)) { g.FillPath(b, p); }
    TextRenderer.DrawText(g, Text, Font, new Rectangle(0, 0, Width, Height), Fore, Fill, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
  }
}

public class IconBtn : Control {
  public string Kind;
  public bool Outline;
  bool hover;
  public IconBtn(string kind) {
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    Kind = kind; Cursor = Cursors.Hand; BackColor = OsUi.Bg;
  }
  protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
  protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
    g.Clear(BackColor);
    RectangleF r = new RectangleF(0.5f, 0.5f, Width - 2, Height - 2);
    if (Kind == "send") {
      RectangleF sr = new RectangleF(3, 3, Width - 7, Height - 7);
      using (GraphicsPath p = OsUi.Round(sr, 8)) using (SolidBrush b = new SolidBrush(hover ? OsUi.Mix(OsUi.Clay, Color.Black, 0.12f) : OsUi.Clay)) g.FillPath(b, p);
      using (Pen ap = new Pen(Color.White, 2f)) {
        ap.StartCap = LineCap.Round; ap.EndCap = LineCap.Round; ap.LineJoin = LineJoin.Round;
        float ccx = Width / 2f - 0.5f, ccy = Height / 2f - 0.5f;
        g.DrawLine(ap, ccx, ccy + 6, ccx, ccy - 6);
        g.DrawLines(ap, new PointF[] { new PointF(ccx - 5, ccy - 1), new PointF(ccx, ccy - 6), new PointF(ccx + 5, ccy - 1) });
      }
      return;
    }
    if (hover) { using (GraphicsPath p = OsUi.Round(r, 8)) using (SolidBrush b = new SolidBrush(OsUi.Hover())) g.FillPath(b, p); }
    if (Outline) { using (GraphicsPath p = OsUi.Round(r, 8)) using (Pen pn = new Pen(OsUi.Blue, 1)) g.DrawPath(pn, p); }
    float cx = Width / 2f, cy = Height / 2f;
    Color col = (Kind == "screen" || Kind == "gear" || Kind == "min" || Kind == "expand" || Kind == "close") ? OsUi.Mute : OsUi.Ink;
    using (Pen pen = new Pen(col, 1.8f)) {
      pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round; pen.LineJoin = LineJoin.Round;
      if (Kind == "close") {
        g.DrawLine(pen, cx - 7, cy - 7, cx + 7, cy + 7);
        g.DrawLine(pen, cx + 7, cy - 7, cx - 7, cy + 7);
      } else if (Kind == "expand") {
        g.DrawLines(pen, new PointF[] { new PointF(cx + 1, cy - 7), new PointF(cx + 7, cy - 7), new PointF(cx + 7, cy - 1) });
        g.DrawLines(pen, new PointF[] { new PointF(cx - 1, cy + 7), new PointF(cx - 7, cy + 7), new PointF(cx - 7, cy + 1) });
      } else if (Kind == "min") {
        g.DrawLine(pen, cx - 7, cy + 1, cx + 7, cy + 1);
      } else if (Kind == "screen") {
        g.DrawRectangle(pen, cx - 8, cy - 7, 16, 11);
        g.DrawLine(pen, cx - 4, cy + 8, cx + 4, cy + 8);
        g.DrawLine(pen, cx, cy + 4, cx, cy + 8);
      } else if (Kind == "joke") {
        using (Pen jp = new Pen(OsUi.Yellow, 1.8f)) {
          g.DrawEllipse(jp, cx - 8, cy - 8, 16, 16);
          g.DrawArc(jp, cx - 4.5f, cy - 3, 9, 7, 20, 140);
          g.FillEllipse(Brushes.Gold, cx - 4, cy - 4, 2.4f, 2.4f);
          g.FillEllipse(Brushes.Gold, cx + 1.6f, cy - 4, 2.4f, 2.4f);
        }
      } else if (Kind == "gear") {
        g.DrawEllipse(pen, cx - 3, cy - 3, 6, 6);
        for (int i = 0; i < 8; i++) {
          double a = i * Math.PI / 4;
          g.DrawLine(pen, (float)(cx + Math.Cos(a) * 5.5), (float)(cy + Math.Sin(a) * 5.5), (float)(cx + Math.Cos(a) * 8), (float)(cy + Math.Sin(a) * 8));
        }
      } else {
        g.DrawPolygon(pen, new PointF[] { new PointF(cx - 8, cy), new PointF(cx + 8, cy - 7), new PointF(cx + 1, cy + 8), new PointF(cx - 1.5f, cy + 1.5f) });
      }
    }
  }
}

public class ChatView : Control {
  public List<OsMsg> Items = new List<OsMsg>();
  public long SeenUpTo;
  public bool Loaded;
  int scroll;
  Font fBody = new Font("Segoe UI", 10.5f);
  Font fMeta = new Font("Segoe UI", 8.25f);
  Font fChip = new Font("Segoe UI Semibold", 9.5f);
  const TextFormatFlags WrapFlags = TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl | TextFormatFlags.NoPadding;

  public ChatView() {
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
    BackColor = OsUi.Bg;
  }
  int MaxBubble() { return Math.Max(120, (int)(Width * 0.78)); }
  public OsApp App;
  List<KeyValuePair<Rectangle, OsMsg>> chipRects = new List<KeyValuePair<Rectangle, OsMsg>>();
  string ChipText(OsMsg m) {
    if (m.Type == "live") return m.Mine ? "You are sharing your screen" : m.Name + " is sharing the screen - click to watch";
    if (m.Type == "livestop") return m.Mine ? "You stopped sharing your screen" : m.Name + " stopped sharing the screen";
    return (m.Mine ? "You sent: " : "") + OsUi.KindLabel(m.Type) + (m.Text != "" ? " - " + m.Text : "");
  }
  Font fSerif = new Font("Georgia", 11f);
  Font MsgFont(OsMsg m) { return m.Mine ? fBody : fSerif; }
  int MaxFor(OsMsg m) { return m.Mine ? MaxBubble() : Math.Max(120, Width - 20); }
  Size BubbleSize(OsMsg m) {
    Size s = TextRenderer.MeasureText(m.Text == "" ? " " : m.Text, MsgFont(m), new Size(MaxFor(m) - 28, int.MaxValue), WrapFlags);
    return new Size(Math.Min(MaxFor(m), s.Width + 28), s.Height + 18);
  }
  public string Peer = "";
  Font fVu = new Font("Segoe UI", 8.25f);
  Font fAv = new Font("Segoe UI Semibold", 6.5f);
  int LastSeenIndex() {
    int idx = -1;
    if (SeenUpTo <= 0) return -1;
    for (int i = 0; i < Items.Count; i++) { if (Items[i].Mine && Items[i].Time <= SeenUpTo) idx = i; }
    return idx;
  }
  Size JokeSize() {
    Image im = OsUi.JokeImage();
    if (im == null) return new Size(0, 0);
    int w = Math.Min(180, MaxBubble());
    return new Size(w, w * im.Height / im.Width);
  }
  List<KeyValuePair<Rectangle, OsMsg>> shotRects = new List<KeyValuePair<Rectangle, OsMsg>>();
  Size ShotSize() { int w = Math.Min(240, MaxBubble()); return new Size(w, w * 9 / 16); }
  protected override void OnMouseMove(MouseEventArgs e) {
    base.OnMouseMove(e);
    bool over = false;
    foreach (KeyValuePair<Rectangle, OsMsg> kv in shotRects) if (kv.Key.Contains(e.Location)) over = true;
    foreach (KeyValuePair<Rectangle, OsMsg> kv in chipRects) if (kv.Key.Contains(e.Location)) over = true;
    Cursor = over ? Cursors.Hand : Cursors.Default;
  }
  protected override void OnMouseClick(MouseEventArgs e) {
    base.OnMouseClick(e);
    foreach (KeyValuePair<Rectangle, OsMsg> kv in chipRects) {
      if (kv.Key.Contains(e.Location) && App != null) { App.OpenViewer(kv.Value.Name); return; }
    }
    foreach (KeyValuePair<Rectangle, OsMsg> kv in shotRects) {
      if (kv.Key.Contains(e.Location) && kv.Value.Img != null) {
        try {
          string p = Path.Combine(Path.GetTempPath(), "OfficeSignal-" + Math.Abs(kv.Value.Key.GetHashCode()).ToString() + ".jpg");
          kv.Value.Img.Save(p, System.Drawing.Imaging.ImageFormat.Jpeg);
          System.Diagnostics.Process.Start(p);
        } catch { }
        return;
      }
    }
  }
  int ItemHeight(OsMsg m, bool vu) {
    if (m.Type == "shot") return ShotSize().Height + 4 + 16 + 8 + (vu ? 18 : 0);
    int h = OsUi.IsChip(m.Type) ? 30 + 4 + 16 + 12 : BubbleSize(m).Height + 4 + 16 + 8;
    if (m.Type == "joke") h += JokeSize().Height + 6;
    if (vu) h += 18;
    return h;
  }
  int ContentHeight() { int h = 14; int ls = LastSeenIndex(); for (int i = 0; i < Items.Count; i++) h += ItemHeight(Items[i], i == ls); return h + 6; }
  int MaxScroll() { return Math.Max(0, ContentHeight() - Height); }
  public void ScrollToEnd() { scroll = MaxScroll(); Invalidate(); }
  public void Add(OsMsg m) {
    bool atBottom = scroll >= MaxScroll() - 40;
    Items.Add(m);
    if (atBottom || m.Mine) scroll = MaxScroll();
    Invalidate();
  }
  public void UpdateTime(string key, long t) {
    foreach (OsMsg m in Items) { if (m.Key == key) { m.Time = t; Invalidate(); return; } }
  }
  protected override void OnResize(EventArgs e) { base.OnResize(e); scroll = Math.Min(scroll, MaxScroll()); }
  protected override void OnMouseWheel(MouseEventArgs e) {
    scroll -= e.Delta / 120 * 60;
    if (scroll < 0) scroll = 0;
    int mx = MaxScroll(); if (scroll > mx) scroll = mx;
    Invalidate();
  }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
    g.Clear(BackColor);
    if (Items.Count == 0) {
      TextRenderer.DrawText(g, Loaded ? "No messages yet.\nSay hello below." : "Loading...", fBody, new Rectangle(0, 0, Width, Height), OsUi.Mute, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
      return;
    }
    int lastSeenMine = LastSeenIndex();
    shotRects.Clear(); chipRects.Clear();
    int y = 14 - scroll;
    for (int i = 0; i < Items.Count; i++) {
      OsMsg m = Items[i];
      int h = ItemHeight(m, i == lastSeenMine);
      if (y + h < 0 || y > Height) { y += h; continue; }
      string meta = (m.Mine ? "" : m.Name + " - ") + OsUi.Clock(m.Time);
      if (m.Type == "shot") {
        Size ss = ShotSize();
        int sx = m.Mine ? Width - 14 - ss.Width : 14;
        Rectangle sr = new Rectangle(sx, y, ss.Width, ss.Height);
        using (GraphicsPath p = OsUi.Round(sr, 14)) {
          using (SolidBrush b = new SolidBrush(OsUi.Bubble)) g.FillPath(b, p);
          if (m.Img != null) {
            Region old = g.Clip; g.SetClip(p);
            float sc = Math.Min(ss.Width / (float)m.Img.Width, ss.Height / (float)m.Img.Height);
            int iw = (int)(m.Img.Width * sc), ih = (int)(m.Img.Height * sc);
            g.DrawImage(m.Img, new Rectangle(sx + (ss.Width - iw) / 2, y + (ss.Height - ih) / 2, iw, ih));
            g.Clip = old;
          } else {
            string st = m.ImgFailed ? "Screenshot expired" : "Loading screenshot...";
            TextRenderer.DrawText(g, st, fMeta, sr, OsUi.Mute, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
          }
        }
        shotRects.Add(new KeyValuePair<Rectangle, OsMsg>(sr, m));
        TextRenderer.DrawText(g, meta + "  -  Screenshot", fMeta, new Rectangle(m.Mine ? 0 : sx + 6, y + ss.Height + 4, m.Mine ? Width - 20 : Width, 16), OsUi.Mute, m.Mine ? TextFormatFlags.Right : TextFormatFlags.Left);
      } else if (OsUi.IsChip(m.Type)) {
        string t = ChipText(m);
        Size ts = TextRenderer.MeasureText(t, fChip);
        int w = Math.Min(Width - 40, ts.Width + 32);
        Rectangle r = new Rectangle((Width - w) / 2, y, w, 30);
        using (GraphicsPath p = OsUi.Round(r, 15)) using (SolidBrush b = new SolidBrush(OsUi.KindColor(m.Type))) g.FillPath(b, p);
        TextRenderer.DrawText(g, t, fChip, r, OsUi.KindText(m.Type), OsUi.KindColor(m.Type), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        if (m.Type == "live" && !m.Mine) chipRects.Add(new KeyValuePair<Rectangle, OsMsg>(r, m));
        TextRenderer.DrawText(g, meta, fMeta, new Rectangle(0, y + 34, Width, 16), OsUi.Mute, TextFormatFlags.HorizontalCenter);
      } else {
        int y0 = y;
        if (m.Type == "joke") {
          Size js = JokeSize();
          if (js.Width > 0) {
            int jx = m.Mine ? Width - 14 - js.Width : 14;
            using (GraphicsPath p = OsUi.Round(new Rectangle(jx, y, js.Width, js.Height), 16)) {
              Region old = g.Clip; g.SetClip(p); g.DrawImage(OsUi.JokeImage(), new Rectangle(jx, y, js.Width, js.Height)); g.Clip = old;
            }
            y += js.Height + 6;
          }
        }
        Size bs = BubbleSize(m);
        int x = m.Mine ? Width - 14 - bs.Width : 0;
        Rectangle r = new Rectangle(x, y, bs.Width, bs.Height);
        Color fill = m.Mine ? OsUi.Bubble : BackColor;
        if (m.Mine) { using (GraphicsPath p = OsUi.Round(r, Math.Min(16, bs.Height / 2f))) using (SolidBrush b = new SolidBrush(fill)) g.FillPath(b, p); }
        TextRenderer.DrawText(g, m.Text, MsgFont(m), new Rectangle(x + 14, y + 9, bs.Width - 28, bs.Height - 18), OsUi.Ink, fill, WrapFlags);
        TextRenderer.DrawText(g, meta, fMeta, new Rectangle(m.Mine ? 0 : x + 6, y + bs.Height + 4, m.Mine ? Width - 20 : Width, 16), OsUi.Mute, m.Mine ? TextFormatFlags.Right : TextFormatFlags.Left);
        y = y0;
      }
      if (i == lastSeenMine) {
        int vy = y + h - 18 - (OsUi.IsChip(m.Type) ? 10 : 6);
        Size vs = TextRenderer.MeasureText("Vu", fVu);
        int vx = OsUi.IsChip(m.Type) ? (Width + 18 - vs.Width) / 2 : Width - 20 - vs.Width;
        TextRenderer.DrawText(g, "Vu", fVu, new Point(vx, vy), OsUi.Mute);
        if (Peer != "") OsUi.DrawAvatar(g, Peer, new RectangleF(vx - 18, vy + 1, 14, 14), Color.Empty, fAv);
      }
      y += h;
    }
    int ch = ContentHeight();
    if (ch > Height) {
      int th = Math.Max(30, Height * Height / ch);
      int ty = (int)((Height - th) * (scroll / (float)Math.Max(1, ch - Height)));
      using (GraphicsPath p = OsUi.Round(new Rectangle(Width - 7, ty + 2, 4, th - 4), 2)) using (SolidBrush b = new SolidBrush(OsUi.Line)) g.FillPath(b, p);
    }
  }
}

public class ConfirmForm : Form {
  public ConfirmForm() : this("Send emergency?", "Your colleague's Messages bar will blink red.", "Send emergency", OsUi.Red) { }
  public ConfirmForm(string title, string sub, string yesText, Color yesColor) {
    FormBorderStyle = FormBorderStyle.None; StartPosition = FormStartPosition.CenterParent; ShowInTaskbar = false;
    BackColor = Color.Black; Size = new Size(320, 190); TopMost = true;
    Region = new Region(OsUi.Round(new RectangleF(0, 0, 320, 190), 16));
    Label t = new Label(); t.Text = title; t.ForeColor = OsUi.Ink; t.Font = new Font("Segoe UI Semibold", 13f); t.AutoSize = false; t.TextAlign = ContentAlignment.MiddleCenter; t.Bounds = new Rectangle(0, 18, 320, 30);
    Label s = new Label(); s.Text = sub; s.ForeColor = OsUi.Mute; s.Font = new Font("Segoe UI", 9.5f); s.AutoSize = false; s.TextAlign = ContentAlignment.MiddleCenter; s.Bounds = new Rectangle(10, 50, 300, 22);
    RoundBtn yes = new RoundBtn(yesText, yesColor, Color.White); yes.BackColor = Color.Black; yes.Bounds = new Rectangle(24, 84, 272, 40);
    RoundBtn no = new RoundBtn("Cancel", Color.FromArgb(32, 35, 39), OsUi.Ink); no.BackColor = Color.Black; no.Bounds = new Rectangle(24, 132, 272, 40);
    yes.Click += delegate { DialogResult = DialogResult.Yes; Close(); };
    no.Click += delegate { DialogResult = DialogResult.No; Close(); };
    Controls.Add(t); Controls.Add(s); Controls.Add(yes); Controls.Add(no);
    KeyPreview = true;
    KeyDown += delegate(object o, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.No; Close(); } };
  }
}

public class Header : Control {
  ChatForm form;
  public IconBtn Expand, Close, Min, Gear;
  Font fTitle = new Font("Segoe UI Semibold", 14f);
  Font fSub = new Font("Segoe UI", 9f);
  public Header(ChatForm f) {
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    form = f; BackColor = OsUi.Bg;
    Gear = new IconBtn("gear"); Min = new IconBtn("min"); Expand = new IconBtn("expand"); Close = new IconBtn("close");
    Controls.Add(Gear); Controls.Add(Min); Controls.Add(Expand); Controls.Add(Close);
  }
  protected override void OnResize(EventArgs e) {
    base.OnResize(e);
    Close.Bounds = new Rectangle(Width - 50, 16, 38, 38);
    Expand.Bounds = new Rectangle(Width - 94, 16, 38, 38);
    Min.Bounds = new Rectangle(Width - 138, 16, 38, 38);
    Gear.Bounds = new Rectangle(Width - 182, 16, 38, 38);
  }
  protected override void OnMouseDown(MouseEventArgs e) { base.OnMouseDown(e); if (e.Button == MouseButtons.Left) OsNative.Drag(form); }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
    g.Clear(BackColor);
    TextRenderer.DrawText(g, "Messages", fTitle, new Point(16, 10), OsUi.Ink);
    bool on = form.App.AnyConnected();
    using (SolidBrush b = new SolidBrush(on ? OsUi.Green : OsUi.Mute)) g.FillEllipse(b, 19, 45, 8, 8);
    TextRenderer.DrawText(g, form.App.StatusText(), fSub, new Point(32, 41), OsUi.Mute);
    using (Pen p = new Pen(OsUi.Line)) g.DrawLine(p, 0, Height - 1, Width, Height - 1);
  }
}

public class ChatForm : Form {
  public OsApp App;
  public ChatView View;
  public Header Head;
  Panel sig, bottom;
  RoundBtn bRed, bYel, bGrn;
  TextBox input;
  IconBtn sendBtn, jokeBtn, shotBtn;
  public bool Quitting;
  bool wide;
  SettingsView settings;

  public ChatForm(OsApp app) {
    App = app;
    Text = "Messages"; Icon = app.AppIcon;
    FormBorderStyle = FormBorderStyle.None; StartPosition = FormStartPosition.Manual;
    BackColor = OsUi.Line; ShowInTaskbar = true; KeyPreview = true;
    Size = new Size(440, 680);
    Place();
    Head = new Header(this);
    sig = new Panel(); sig.BackColor = OsUi.Bg;
    sig.Paint += delegate(object o, PaintEventArgs e) { using (Pen p = new Pen(OsUi.Line)) e.Graphics.DrawLine(p, 0, sig.Height - 1, sig.Width, sig.Height - 1); };
    bRed = new RoundBtn("Emergency", OsUi.Red, Color.White); bRed.Dot = OsUi.Red;
    bYel = new RoundBtn("Come", OsUi.Yellow, OsUi.DarkText); bYel.Dot = Color.FromArgb(230, 180, 0);
    bGrn = new RoundBtn("Call off", OsUi.Green, Color.White); bGrn.Dot = OsUi.Green;
    bRed.Font = bYel.Font = bGrn.Font = new Font("Segoe UI", 10f);
    sig.Controls.Add(bRed); sig.Controls.Add(bYel); sig.Controls.Add(bGrn);
    sig.Resize += delegate { LayoutSignals(); };
    View = new ChatView(); View.App = app;
    bottom = new Panel(); bottom.BackColor = OsUi.Bg;
    bottom.Paint += delegate(object o, PaintEventArgs e) {
      Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
      using (GraphicsPath gp = OsUi.Round(new RectangleF(12, 10, bottom.Width - 24, 52), 18)) {
        using (SolidBrush b = new SolidBrush(OsUi.Box)) g.FillPath(b, gp);
        using (Pen p = new Pen(OsUi.Line, 1f)) g.DrawPath(p, gp);
      }
    };
    input = new TextBox(); input.BorderStyle = BorderStyle.None; input.BackColor = OsUi.Box; input.ForeColor = OsUi.Ink; input.Font = new Font("Segoe UI", 11f);
    input.MaxLength = 2000;
    sendBtn = new IconBtn("send"); sendBtn.BackColor = OsUi.Box;
    jokeBtn = new IconBtn("joke"); jokeBtn.BackColor = OsUi.Box;
    shotBtn = new IconBtn("screen"); shotBtn.BackColor = OsUi.Box;
    bottom.Controls.Add(input); bottom.Controls.Add(sendBtn); bottom.Controls.Add(jokeBtn); bottom.Controls.Add(shotBtn);
    shotBtn.Click += delegate {
      ContextMenuStrip cm = new ContextMenuStrip();
      cm.Font = new Font("Segoe UI", 10f);
      if (App.Sharing) cm.Items.Add("Stop sharing my screen", null, delegate { App.StopShare(); });
      else cm.Items.Add("Share my screen (live)", null, delegate { StartLive(); });
      cm.Items.Add("Send a screenshot", null, delegate { SendScreenshot(); });
      cm.Show(shotBtn, new Point(0, -cm.PreferredSize.Height));
    };
    jokeBtn.Click += delegate { App.Send(OsUi.JokeText, "joke"); FocusInput(); };
    bottom.Resize += delegate { LayoutBottom(); };
    Controls.Add(Head); Controls.Add(sig); Controls.Add(View); Controls.Add(bottom);

    Head.Close.Click += delegate { Hide(); };
    Head.Min.Click += delegate { WindowState = FormWindowState.Minimized; };
    Head.Expand.Click += delegate { wide = !wide; Size = wide ? new Size(760, 820) : new Size(440, 680); Place(); };
    sendBtn.Click += delegate { SendInput(); };
    input.KeyDown += delegate(object o, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; SendInput(); } };
    bRed.Click += delegate { Signal("red"); };
    bYel.Click += delegate { Signal("yellow"); };
    bGrn.Click += delegate { Signal("green"); };
    KeyDown += delegate(object o, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) { if (settings.Visible) settings.Visible = false; else Hide(); } };
    settings = new SettingsView(this); settings.Visible = false;
    Controls.Add(settings);
    Head.Gear.Click += delegate { settings.Open(); };
    DoLayout(); LayoutSignals(); LayoutBottom();
  }
  protected override CreateParams CreateParams {
    get { CreateParams cp = base.CreateParams; cp.ClassStyle |= 0x20000; cp.Style |= 0x20000; return cp; }
  }
  protected override void OnHandleCreated(EventArgs e) {
    base.OnHandleCreated(e);
    try { OsNative.SendMessage(input.Handle, 0x1501, (IntPtr)1, "Reply..."); } catch { }
  }
  protected override void OnFormClosing(FormClosingEventArgs e) {
    if (!Quitting && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); return; }
    base.OnFormClosing(e);
  }
  protected override void OnResize(EventArgs e) {
    base.OnResize(e);
    DoLayout();
  }
  void DoLayout() {
    if (WindowState == FormWindowState.Minimized || Width < 50) return;
    Region = new Region(OsUi.Round(new RectangleF(0, 0, Width, Height), 14));
    int w = Width - 2, h = Height - 2;
    if (Head == null || sig == null || View == null || bottom == null) return;
    Head.Bounds = new Rectangle(1, 1, w, 70);
    sig.Bounds = new Rectangle(1, 71, w, 68);
    bottom.Bounds = new Rectangle(1, h + 1 - 72, w, 72);
    View.Bounds = new Rectangle(1, 139, w, h - 138 - 72);
    if (settings != null) settings.Bounds = new Rectangle(1, 1, w, h);
  }
  public void ApplyTheme() {
    BackColor = OsUi.Line;
    foreach (Control c in new Control[] { Head, sig, View, bottom, Head.Gear, Head.Min, Head.Expand, Head.Close, bRed, bYel, bGrn }) { c.BackColor = OsUi.Bg; c.Invalidate(); }
    input.BackColor = OsUi.Box; input.ForeColor = OsUi.Ink; sendBtn.BackColor = OsUi.Box; sendBtn.Invalidate(); jokeBtn.BackColor = OsUi.Box; jokeBtn.Invalidate(); shotBtn.BackColor = OsUi.Box; shotBtn.Invalidate();
    settings.ApplyTheme();
    Invalidate(true);
  }
  void LayoutSignals() {
    int gap = 8, pad = 12, bw = (sig.Width - pad * 2 - gap * 2) / 3;
    bRed.Bounds = new Rectangle(pad, 12, bw, 44);
    bYel.Bounds = new Rectangle(pad + bw + gap, 12, bw, 44);
    bGrn.Bounds = new Rectangle(pad + (bw + gap) * 2, 12, bw, 44);
  }
  void LayoutBottom() {
    int boxW = bottom.Width - 24;
    input.Bounds = new Rectangle(28, 12 + (48 - input.PreferredHeight) / 2, boxW - 16 - 124, input.PreferredHeight);
    shotBtn.Bounds = new Rectangle(12 + boxW - 124, 18, 36, 36);
    jokeBtn.Bounds = new Rectangle(12 + boxW - 84, 18, 36, 36);
    sendBtn.Bounds = new Rectangle(12 + boxW - 44, 18, 36, 36);
  }
  public void Place() {
    Rectangle wa = Screen.PrimaryScreen.WorkingArea;
    Location = new Point(wa.Right - Width - 16, wa.Bottom - Height - 16);
  }
  public void FocusInput() { try { input.Focus(); } catch { } }
  void StartLive() {
    using (ConfirmForm c = new ConfirmForm("Share your screen?", "Your colleague sees your screen (every 5 s) until you press Stop.", "Share screen", OsUi.Blue)) {
      if (c.ShowDialog(this) != DialogResult.Yes) return;
    }
    App.StartShare();
  }
  void SendScreenshot() {
    using (ConfirmForm c = new ConfirmForm("Send a screenshot?", "A picture of your whole screen goes to your colleague.", "Send screenshot", OsUi.Blue)) {
      if (c.ShowDialog(this) != DialogResult.Yes) return;
    }
    bool pillWas = App.Pill.Visible;
    Hide(); App.Pill.Hide();
    Application.DoEvents(); Thread.Sleep(350);
    Bitmap shot = null;
    shot = OsApp.CaptureScreen();
    Show(); Activate();
    if (shot != null) App.SendShot(shot);
  }
  void SendInput() {
    string t = input.Text.Trim();
    if (t == "") return;
    input.Text = "";
    App.Send(t, null);
  }
  void Signal(string kind) {
    if (kind == "red") {
      using (ConfirmForm c = new ConfirmForm()) { if (c.ShowDialog(this) != DialogResult.Yes) return; }
    }
    string note = input.Text.Trim(); input.Text = "";
    App.Send(note, kind);
  }
}

public class SettingsView : Control {
  ChatForm form;
  IconBtn close;
  TextBox nameBox;
  RoundBtn save;
  Font fTitle = new Font("Segoe UI Semibold", 14f);
  Font fLbl = new Font("Segoe UI", 9f);
  Font fSeg = new Font("Segoe UI Semibold", 10f);
  Rectangle segRect;
  public SettingsView(ChatForm f) {
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    form = f; BackColor = OsUi.Bg; Cursor = Cursors.Default;
    close = new IconBtn("close");
    nameBox = new TextBox(); nameBox.BorderStyle = BorderStyle.None; nameBox.Font = new Font("Segoe UI", 11f); nameBox.MaxLength = 40;
    save = new RoundBtn("Save", OsUi.Blue, Color.White);
    Controls.Add(close); Controls.Add(nameBox); Controls.Add(save);
    close.Click += delegate { Visible = false; form.FocusInput(); };
    save.Click += delegate { form.App.SetName(nameBox.Text); Visible = false; form.FocusInput(); };
    nameBox.KeyDown += delegate(object o, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; form.App.SetName(nameBox.Text); Visible = false; form.FocusInput(); } };
    ApplyTheme();
  }
  public void Open() { nameBox.Text = form.App.Name; Visible = true; BringToFront(); nameBox.Focus(); nameBox.SelectAll(); }
  public void ApplyTheme() {
    BackColor = OsUi.Bg; close.BackColor = OsUi.Bg; save.BackColor = OsUi.Bg;
    nameBox.BackColor = OsUi.Box; nameBox.ForeColor = OsUi.Ink;
    Invalidate(true);
  }
  protected override void OnResize(EventArgs e) {
    base.OnResize(e);
    close.Bounds = new Rectangle(Width - 50, 16, 38, 38);
    segRect = new Rectangle(16, 112, Width - 32, 40);
    nameBox.Bounds = new Rectangle(30, 200 + (44 - nameBox.PreferredHeight) / 2, Width - 60, nameBox.PreferredHeight);
    save.Bounds = new Rectangle(16, 262, Width - 32, 42);
  }
  protected override void OnMouseDown(MouseEventArgs e) {
    base.OnMouseDown(e);
    if (e.Y < 70 && e.Button == MouseButtons.Left) { OsNative.Drag(form); return; }
    if (segRect.Contains(e.Location)) {
      bool light = e.X > segRect.X + segRect.Width / 2;
      if (light != OsUi.Light) form.App.SetTheme(light);
    }
  }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
    g.Clear(BackColor);
    TextRenderer.DrawText(g, "Settings", fTitle, new Point(16, 18), OsUi.Ink);
    using (Pen p = new Pen(OsUi.Line)) g.DrawLine(p, 0, 69, Width, 69);
    TextRenderer.DrawText(g, "Theme", fLbl, new Point(16, 88), OsUi.Mute);
    using (GraphicsPath p = OsUi.Round(segRect, 20)) using (SolidBrush b = new SolidBrush(OsUi.Box)) g.FillPath(b, p);
    int half = segRect.Width / 2;
    Rectangle left = new Rectangle(segRect.X + 3, segRect.Y + 3, half - 3, segRect.Height - 6);
    Rectangle right = new Rectangle(segRect.X + half, segRect.Y + 3, half - 3, segRect.Height - 6);
    Rectangle on = OsUi.Light ? right : left;
    using (GraphicsPath p = OsUi.Round(on, 17)) using (SolidBrush b = new SolidBrush(OsUi.Blue)) g.FillPath(b, p);
    TextRenderer.DrawText(g, "Night", fSeg, left, OsUi.Light ? OsUi.Ink : Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    TextRenderer.DrawText(g, "Light", fSeg, right, OsUi.Light ? Color.White : OsUi.Ink, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    TextRenderer.DrawText(g, "Your name", fLbl, new Point(16, 176), OsUi.Mute);
    using (GraphicsPath p = OsUi.Round(new RectangleF(16, 200, Width - 32, 44), 10)) using (SolidBrush b = new SolidBrush(OsUi.Box)) g.FillPath(b, p);
  }
}

public class ShareBar : Form {
  OsApp app;
  Font f = new Font("Segoe UI Semibold", 10.5f);
  bool on;
  System.Windows.Forms.Timer t;
  public ShareBar(OsApp a) {
    app = a;
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
    FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; TopMost = true; StartPosition = FormStartPosition.Manual;
    Cursor = Cursors.Hand; Text = "Office Signal sharing";
    int w = 300, h = 40;
    Rectangle wa = Screen.PrimaryScreen.WorkingArea;
    Bounds = new Rectangle(wa.Left + (wa.Width - w) / 2, wa.Top + 10, w, h);
    using (GraphicsPath p = OsUi.Round(new RectangleF(0, 0, w, h), h / 2f)) Region = new Region(p);
    t = new System.Windows.Forms.Timer(); t.Interval = 700; t.Tick += delegate { on = !on; Invalidate(); }; t.Start();
  }
  protected override bool ShowWithoutActivation { get { return true; } }
  protected override CreateParams CreateParams { get { CreateParams cp = base.CreateParams; cp.ExStyle |= 0x00000080 | 0x00000008; return cp; } }
  protected override void OnClick(EventArgs e) { base.OnClick(e); app.StopShare(); }
  protected override void OnFormClosed(FormClosedEventArgs e) { t.Stop(); base.OnFormClosed(e); }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
    g.Clear(Color.FromArgb(32, 35, 39));
    using (SolidBrush b = new SolidBrush(on ? Color.FromArgb(122, 0, 0) : OsUi.Red)) g.FillEllipse(b, 16, 14, 12, 12);
    TextRenderer.DrawText(g, "Sharing your screen", f, new Rectangle(36, 0, 170, Height), Color.White, TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
    Rectangle stop = new Rectangle(Width - 84, 6, 74, 28);
    using (GraphicsPath p = OsUi.Round(stop, 14)) using (SolidBrush b = new SolidBrush(OsUi.Red)) g.FillPath(b, p);
    TextRenderer.DrawText(g, "Stop", f, stop, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
  }
}

public class LiveViewer : Form {
  OsApp app;
  public string Who;
  Image img;
  long lastTime;
  DateTime lastFrame = DateTime.MinValue;
  string status = "Waiting for the screen...";
  bool stopped;
  volatile bool open = true;
  List<HttpWebRequest> reqs = new List<HttpWebRequest>();
  Font fInfo = new Font("Segoe UI Semibold", 10f);
  Font fBig = new Font("Segoe UI", 12f);
  System.Windows.Forms.Timer tick;
  public LiveViewer(OsApp a, string who, string[] servers) {
    app = a; Who = who;
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    Text = who + "'s screen"; Icon = a.AppIcon; BackColor = Color.Black;
    StartPosition = FormStartPosition.CenterScreen;
    Rectangle wa = Screen.PrimaryScreen.WorkingArea;
    Size = new Size(Math.Min(1280, wa.Width - 80), Math.Min(760, wa.Height - 80));
    tick = new System.Windows.Forms.Timer(); tick.Interval = 1000; tick.Tick += delegate { Invalidate(); }; tick.Start();
    foreach (string s in servers) {
      string srv = s;
      Thread th = new Thread(delegate() { Listen(srv); });
      th.IsBackground = true; th.Start();
    }
  }
  public void SetStopped() { stopped = true; status = Who + " stopped sharing the screen"; Invalidate(); }
  public void SetStarted() { stopped = false; Invalidate(); }
  protected override void OnFormClosed(FormClosedEventArgs e) {
    open = false; tick.Stop();
    lock (reqs) { foreach (HttpWebRequest r in reqs) { try { r.Abort(); } catch { } } }
    base.OnFormClosed(e);
  }
  void Listen(string server) {
    string since = "30s";
    while (open) {
      HttpWebRequest req = null;
      try {
        req = (HttpWebRequest)WebRequest.Create(server + "/" + app.Topic + "-live/json?since=" + since);
        req.Timeout = 30000; req.ReadWriteTimeout = 120000;
        lock (reqs) reqs.Add(req);
        using (WebResponse resp = req.GetResponse())
        using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) {
          string line;
          while (open && (line = sr.ReadLine()) != null) {
            try {
              Dictionary<string, object> ev = new JavaScriptSerializer().DeserializeObject(line) as Dictionary<string, object>;
              if (ev == null) continue;
              object evt; ev.TryGetValue("event", out evt);
              if (Convert.ToString(evt) != "message") continue;
              object idv; if (ev.TryGetValue("id", out idv)) since = Convert.ToString(idv);
              long tm = 0; object tv; if (ev.TryGetValue("time", out tv)) tm = Convert.ToInt64(tv);
              object att; if (!ev.TryGetValue("attachment", out att)) continue;
              Dictionary<string, object> ad = att as Dictionary<string, object>;
              object uv; if (ad == null || !ad.TryGetValue("url", out uv)) continue;
              string url = Convert.ToString(uv);
              if (tm < lastTime) continue;
              byte[] data;
              using (WebClient wc = new WebClient()) data = wc.DownloadData(url);
              Image im = Image.FromStream(new MemoryStream(data));
              long tm2 = tm;
              if (IsHandleCreated) BeginInvoke((MethodInvoker)delegate {
                if (tm2 < lastTime) { im.Dispose(); return; }
                lastTime = tm2;
                Image old = img; img = im; lastFrame = DateTime.Now; stopped = false;
                if (old != null) old.Dispose();
                Invalidate();
              });
            } catch { }
          }
        }
      } catch { }
      lock (reqs) reqs.Remove(req);
      if (open) Thread.Sleep(3000);
    }
  }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.Clear(Color.Black);
    g.InterpolationMode = InterpolationMode.HighQualityBilinear;
    if (img != null) {
      float sc = Math.Min(ClientSize.Width / (float)img.Width, ClientSize.Height / (float)img.Height);
      int iw = (int)(img.Width * sc), ih = (int)(img.Height * sc);
      g.DrawImage(img, new Rectangle((ClientSize.Width - iw) / 2, (ClientSize.Height - ih) / 2, iw, ih));
    }
    if (img == null || stopped) {
      Rectangle r = new Rectangle(0, ClientSize.Height / 2 - 30, ClientSize.Width, 60);
      using (SolidBrush b = new SolidBrush(Color.FromArgb(180, 0, 0, 0))) g.FillRectangle(b, r);
      TextRenderer.DrawText(g, status, fBig, r, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }
    string info = stopped ? "Stopped" : (img == null ? "Connecting..." : "Live - updated " + Math.Max(0, (int)(DateTime.Now - lastFrame).TotalSeconds) + " s ago");
    Size isz = TextRenderer.MeasureText(info, fInfo);
    Rectangle ir = new Rectangle(12, 12, isz.Width + 34, 28);
    g.SmoothingMode = SmoothingMode.AntiAlias;
    using (GraphicsPath p = OsUi.Round(ir, 14)) using (SolidBrush b = new SolidBrush(Color.FromArgb(200, 22, 24, 28))) g.FillPath(b, p);
    using (SolidBrush b = new SolidBrush(stopped ? OsUi.Mute : OsUi.Red)) g.FillEllipse(b, ir.X + 12, ir.Y + 10, 8, 8);
    TextRenderer.DrawText(g, info, fInfo, new Point(ir.X + 26, ir.Y + 5), Color.White);
  }
}

public class PillForm : Form {
  OsApp app;
  Font fBig = new Font("Segoe UI Semibold", 13.5f);
  Font fMid = new Font("Segoe UI Semibold", 11f);
  Font fSmall = new Font("Segoe UI Semibold", 8.5f);
  Font fAvatar = new Font("Segoe UI Semibold", 15f, FontStyle.Regular, GraphicsUnit.Pixel);
  public PillForm(OsApp a) {
    app = a;
    SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
    FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; TopMost = true; StartPosition = FormStartPosition.Manual;
    BackColor = Color.Black; Cursor = Cursors.Hand; Text = "Office Signal bar";
    int w = 282, h = 58;
    tip = new ToolTip();
    Rectangle wa = Screen.PrimaryScreen.WorkingArea;
    Bounds = new Rectangle(wa.Right - w - 16, wa.Bottom - h - 16, w, h);
    using (GraphicsPath p = new GraphicsPath()) {
      p.AddArc(0, 0, h, h, 90, 180); p.AddArc(w - h, 0, h, h, 270, 180); p.CloseFigure();
      Region = new Region(p);
    }
    LoadPos();
  }
  protected override bool ShowWithoutActivation { get { return true; } }
  protected override CreateParams CreateParams {
    get { CreateParams cp = base.CreateParams; cp.ExStyle |= 0x00000080 | 0x00000008; return cp; }
  }
  ToolTip tip;
  string lastTip = "";
  string[] miniKinds = { "red", "yellow", "green" };
  RectangleF[] miniRects = new RectangleF[3];
  string sentKind; DateTime sentAt;
  int MiniHit(Point p) { for (int i = 0; i < 3; i++) { RectangleF r = miniRects[i]; r.Inflate(4, 2); if (r.Contains(p)) return i; } return -1; }
  Point downPt, startLoc;
  bool pressing, moved;
  static string PosFile() { return Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficeSignal"), "pill.txt"); }
  void LoadPos() {
    try {
      if (!File.Exists(PosFile())) return;
      string[] p = File.ReadAllText(PosFile()).Trim().Split(',');
      Point pt = new Point(int.Parse(p[0]), int.Parse(p[1]));
      foreach (Screen sc in Screen.AllScreens) {
        if (sc.WorkingArea.Contains(new Rectangle(pt, Size))) { Location = pt; return; }
      }
    } catch { }
  }
  protected override void OnMouseDown(MouseEventArgs e) {
    base.OnMouseDown(e);
    if (e.Button != MouseButtons.Left) return;
    pressing = true; moved = false; downPt = Cursor.Position; startLoc = Location;
  }
  protected override void OnMouseMove(MouseEventArgs e) {
    base.OnMouseMove(e);
    CheckHover(e.Location);
    if (!pressing) return;
    Point c = Cursor.Position;
    int dx = c.X - downPt.X, dy = c.Y - downPt.Y;
    if (!moved && Math.Abs(dx) + Math.Abs(dy) > 5) { moved = true; Cursor = Cursors.SizeAll; }
    if (moved) Location = new Point(startLoc.X + dx, startLoc.Y + dy);
  }
  int lastHov = -2;
  protected override void OnMouseHover(EventArgs e) { base.OnMouseHover(e); Invalidate(); }
  void CheckHover(Point p) { int hv = MiniHit(p); if (hv != lastHov) { lastHov = hv; Invalidate(); } }
  protected override void OnMouseUp(MouseEventArgs e) {
    base.OnMouseUp(e);
    if (!pressing) return;
    pressing = false;
    Cursor = Cursors.Hand;
    if (!moved && e.Button == MouseButtons.Left) {
      int mi = MiniHit(e.Location);
      if (mi >= 0) {
        string kd = miniKinds[mi];
        app.Send("", kd);
        sentKind = kd; sentAt = DateTime.Now; Invalidate();
        System.Windows.Forms.Timer ft = new System.Windows.Forms.Timer(); ft.Interval = 900;
        ft.Tick += delegate { ft.Stop(); ft.Dispose(); sentKind = null; Invalidate(); };
        ft.Start();
        return;
      }
    }
    if (moved) {
      Rectangle wa = Screen.FromPoint(Cursor.Position).WorkingArea;
      int x = Math.Max(wa.Left, Math.Min(Left, wa.Right - Width));
      int y = Math.Max(wa.Top, Math.Min(Top, wa.Bottom - Height));
      Location = new Point(x, y);
      try { Directory.CreateDirectory(Path.GetDirectoryName(PosFile())); File.WriteAllText(PosFile(), x + "," + y); } catch { }
    } else if (e.Button == MouseButtons.Left) {
      app.PillClicked();
    }
  }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
    int w = Width, h = Height;
    Color bg = OsUi.PillBg;
    g.Clear(bg);
    using (GraphicsPath p = new GraphicsPath()) {
      p.AddArc(0, 0, h - 1, h - 1, 90, 180); p.AddArc(w - h, 0, h - 1, h - 1, 270, 180); p.CloseFigure();
      using (Pen pn = new Pen(OsUi.PillLine, 1.5f)) g.DrawPath(pn, p);
    }
    float cy = h / 2f;
    using (Pen pen = new Pen(OsUi.PillInk, 2f)) {
      pen.LineJoin = LineJoin.Round;
      g.DrawPolygon(pen, new PointF[] { new PointF(20, cy - 1), new PointF(38, cy - 9), new PointF(30, cy + 9), new PointF(27.5f, cy + 1.5f) });
    }
    string[] kinds = { "msg", "red", "yellow", "green" };
    float dx = 56, ds = 20;
    for (int i = 0; i < kinds.Length; i++) {
      string kd = kinds[i];
      bool lit = app.Lit.Contains(kd) || (kd == "msg" && app.Lit.Contains("live"));
      RectangleF r = new RectangleF(dx + i * 30, cy - ds / 2, ds, ds);
      if (lit) {
        Color c = app.BlinkOn ? OsUi.KindDark(kd) : OsUi.KindColor(kd);
        if (!app.BlinkOn) { using (SolidBrush gl = new SolidBrush(Color.FromArgb(70, OsUi.KindColor(kd)))) g.FillEllipse(gl, r.X - 4, r.Y - 4, r.Width + 8, r.Height + 8); }
        using (SolidBrush b = new SolidBrush(c)) g.FillEllipse(b, r);
      } else {
        using (Pen pn = new Pen(Color.FromArgb(74, 80, 86), 2f)) g.DrawEllipse(pn, r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2);
      }
    }
    int ax = w - 46, ay = (h - 32) / 2;
    if (app.Other != "") OsUi.DrawAvatar(g, app.Other, new RectangleF(ax, ay, 32, 32), bg, fAvatar);
    float mx = w - 70, msz = 13;
    using (Pen sep = new Pen(Color.FromArgb(58, 63, 68), 1f)) g.DrawLine(sep, mx - 10, 12, mx - 10, h - 12);
    for (int i = 0; i < 3; i++) {
      RectangleF r = new RectangleF(mx, cy - 21.5f + i * 15, msz, msz);
      miniRects[i] = r;
      Color mc = OsUi.KindColor(miniKinds[i]);
      if (sentKind == miniKinds[i]) { using (SolidBrush gl = new SolidBrush(Color.FromArgb(110, mc))) g.FillEllipse(gl, r.X - 4, r.Y - 4, r.Width + 8, r.Height + 8); }
      using (SolidBrush b = new SolidBrush(mc)) g.FillEllipse(b, r);
    }
    Point mp = PointToClient(Cursor.Position);
    int hov = ClientRectangle.Contains(mp) ? MiniHit(mp) : -1;
    string t = hov >= 0 ? "Send: " + OsUi.KindLabel(miniKinds[hov]) : app.TipText();
    if (t != lastTip) { lastTip = t; tip.SetToolTip(this, t); }
  }
}

public class OsApp : ApplicationContext {
  public string Name, Cid, Topic;
  string[] servers = { "https://ntfy.sh", "https://ntfy.envs.net" };
  public ChatForm Chat;
  public PillForm Pill;
  public Icon AppIcon;
  NotifyIcon tray;
  System.Windows.Forms.Timer watch, blink;
  HashSet<string> seenKeys = new HashSet<string>();
  Dictionary<string, bool> conn = new Dictionary<string, bool>();
  public string Alert;
  public HashSet<string> Lit = new HashSet<string>();
  Dictionary<string, string> litText = new Dictionary<string, string>();
  Dictionary<string, DateTime> litAt = new Dictionary<string, DateTime>();
  const int LitSeconds = 10;
  public string TipText() {
    if (Lit.Count == 0) return "Messages";
    List<string> parts = new List<string>();
    foreach (string k in new string[] { "red", "yellow", "green", "live", "msg" }) { if (Lit.Contains(k) && litText.ContainsKey(k)) parts.Add(litText[k]); }
    return string.Join("\n", parts.ToArray());
  }
  public string AlertFrom = "";
  public string AlertText = "";
  string cfgDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficeSignal");
  public bool BlinkOn;
  public int Unread;
  public string Other = "";
  long lastOtherTime, seenSent;
  volatile bool running = true;
  EventWaitHandle showEvent;
  Random rnd = new Random();

  public OsApp(string name, string cid, string topic) {
    Name = name; Cid = cid; Topic = topic;
    try { ServicePointManager.DefaultConnectionLimit = 64; ServicePointManager.Expect100Continue = false; } catch { }
    try { OsNative.SetCurrentProcessExplicitAppUserModelID("RoboTalent.OfficeSignal"); } catch { }
    try { OsUi.SetTheme(!(File.Exists(Path.Combine(cfgDir, "theme.txt")) && File.ReadAllText(Path.Combine(cfgDir, "theme.txt")).Trim() == "dark")); } catch { OsUi.SetTheme(true); }
    AppIcon = OsUi.MakeIcon(64);
    Chat = new ChatForm(this);
    IntPtr h1 = Chat.Handle;
    Pill = new PillForm(this);
    IntPtr h2 = Pill.Handle;
    tray = new NotifyIcon();
    tray.Icon = OsUi.MakeIcon(32); tray.Text = "Office Signal"; tray.Visible = true;
    ContextMenuStrip menu = new ContextMenuStrip();
    menu.Items.Add("Open Messages", null, delegate { ShowChat(); });
    menu.Items.Add("Quit", null, delegate { Shutdown(); });
    tray.ContextMenuStrip = menu;
    tray.DoubleClick += delegate { ShowChat(); };
    watch = new System.Windows.Forms.Timer(); watch.Interval = 500; watch.Tick += delegate { Watch(); }; watch.Start();
    blink = new System.Windows.Forms.Timer(); blink.Interval = 450;
    blink.Tick += delegate {
      BlinkOn = !BlinkOn;
      List<string> old = new List<string>();
      foreach (KeyValuePair<string, DateTime> kv in litAt) if ((DateTime.Now - kv.Value).TotalSeconds > LitSeconds) old.Add(kv.Key);
      foreach (string k in old) ClearKind(k);
      Pill.TopMost = true; Pill.Invalidate();
    };
    bool createdNew;
    showEvent = new EventWaitHandle(false, EventResetMode.AutoReset, "Local\\OfficeSignalShow", out createdNew);
    Thread waiter = new Thread(delegate() {
      while (running) { try { if (showEvent.WaitOne(1000)) Ui(delegate { ShowChat(); }); } catch { Thread.Sleep(1000); } }
    });
    waiter.IsBackground = true; waiter.Start();
    Thread loader = new Thread(LoadAndListen); loader.IsBackground = true; loader.Start();
    ShowChat();
  }

  public void Ui(MethodInvoker a) {
    try { if (Chat != null && Chat.IsHandleCreated && !Chat.IsDisposed) Chat.BeginInvoke(a); } catch { }
  }
  public bool AnyConnected() { lock (conn) { foreach (bool b in conn.Values) if (b) return true; } return false; }
  public string StatusText() {
    if (!AnyConnected()) return "Offline - retrying...";
    return Other != "" ? Other + " - connected" : "Connected";
  }
  void SetConn(string s, bool on) {
    bool changed;
    lock (conn) { bool old; changed = !conn.TryGetValue(s, out old) || old != on; conn[s] = on; }
    if (changed) Ui(delegate { Chat.Head.Invalidate(); });
  }
  public bool ChatActive() {
    return Chat.Visible && Chat.WindowState != FormWindowState.Minimized && Form.ActiveForm == Chat;
  }
  public void ShowChat() {
    ClearAlert();
    Unread = 0;
    if (!Chat.Visible) Chat.Show();
    if (Chat.WindowState == FormWindowState.Minimized) Chat.WindowState = FormWindowState.Normal;
    Chat.Activate(); Chat.BringToFront();
    Chat.FocusInput();
    Pill.Invalidate();
    Watch();
  }
  void ClearAlert() { Alert = null; AlertFrom = ""; AlertText = ""; Lit.Clear(); litText.Clear(); litAt.Clear(); blink.Stop(); Pill.Invalidate(); }
  void ClearKind(string k) {
    Lit.Remove(k); litText.Remove(k); litAt.Remove(k);
    if (Lit.Count == 0) { Alert = null; blink.Stop(); }
    Pill.Invalidate();
  }
  void StartAlert(string kind, string from, string text) {
    Alert = kind; AlertFrom = from; AlertText = text ?? ""; BlinkOn = false;
    if (OsUi.IsSignal(kind)) { foreach (string o in new string[] { "red", "yellow", "green" }) { if (o != kind) { Lit.Remove(o); litText.Remove(o); litAt.Remove(o); } } }
    Lit.Add(kind);
    litAt[kind] = DateTime.Now;
    string desc;
    if (kind == "msg") desc = from + ": " + (text ?? "");
    else if (kind == "live") desc = from + " is sharing the screen";
    else desc = OsUi.KindLabel(kind) + " - " + from + ((text ?? "") != "" ? ": " + text : "");
    litText[kind] = desc;
    if (!Pill.Visible) Pill.Show();
    Pill.TopMost = true; Pill.Invalidate();
    blink.Start();
  }
  void Watch() {
    bool shown = Chat.Visible && Chat.WindowState != FormWindowState.Minimized;
    bool wantPill = Lit.Count > 0 || !shown;
    if (wantPill && !Pill.Visible) { Pill.Show(); Pill.TopMost = true; }
    else if (!wantPill && Pill.Visible) Pill.Hide();
    if (ChatActive()) {
      if (Lit.Contains("msg")) ClearKind("msg");
      if (Unread != 0) { Unread = 0; Pill.Invalidate(); }
      if (lastOtherTime > seenSent) { seenSent = lastOtherTime; Send("", "seen"); }
    }
  }
  public bool Sharing;
  ShareBar shareBar;
  LiveViewer viewer;
  public static Bitmap CaptureScreen() {
    Bitmap shot = null;
    IntPtr oldCtx = IntPtr.Zero;
    try { oldCtx = OsNative.SetThreadDpiAwarenessContext((IntPtr)(-4)); } catch { }
    try {
      int pw = 0, ph = 0;
      try { IntPtr dc = OsNative.GetDC(IntPtr.Zero); pw = OsNative.GetDeviceCaps(dc, 118); ph = OsNative.GetDeviceCaps(dc, 117); OsNative.ReleaseDC(IntPtr.Zero, dc); } catch { }
      Rectangle b = Screen.PrimaryScreen.Bounds;
      if (pw < b.Width || ph < b.Height) { pw = b.Width; ph = b.Height; }
      shot = new Bitmap(pw, ph);
      using (Graphics g = Graphics.FromImage(shot)) g.CopyFromScreen(0, 0, 0, 0, new Size(pw, ph));
    } catch { }
    try { if (oldCtx != IntPtr.Zero) OsNative.SetThreadDpiAwarenessContext(oldCtx); } catch { }
    return shot;
  }
  public static byte[] Jpeg(Bitmap src, int maxW, long quality) {
    Bitmap small = src;
    if (src.Width > maxW) {
      int h = src.Height * maxW / src.Width;
      small = new Bitmap(maxW, h);
      using (Graphics g = Graphics.FromImage(small)) { g.InterpolationMode = InterpolationMode.HighQualityBicubic; g.DrawImage(src, 0, 0, maxW, h); }
    }
    using (MemoryStream ms = new MemoryStream()) {
      System.Drawing.Imaging.ImageCodecInfo codec = null;
      foreach (System.Drawing.Imaging.ImageCodecInfo ci in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()) if (ci.MimeType == "image/jpeg") codec = ci;
      System.Drawing.Imaging.EncoderParameters ep = new System.Drawing.Imaging.EncoderParameters(1);
      ep.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
      if (codec != null) small.Save(ms, codec, ep); else small.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
      if (small != src) small.Dispose();
      return ms.ToArray();
    }
  }
  string PutFile(string topicUrl, byte[] data) {
    try {
      HttpWebRequest req = (HttpWebRequest)WebRequest.Create(topicUrl);
      req.Method = "PUT"; req.Timeout = 30000; req.ReadWriteTimeout = 30000;
      req.Headers["Filename"] = "screen.jpg";
      req.ContentLength = data.Length;
      using (Stream rs = req.GetRequestStream()) rs.Write(data, 0, data.Length);
      using (WebResponse resp = req.GetResponse()) using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) {
        Dictionary<string, object> d = new JavaScriptSerializer().DeserializeObject(sr.ReadToEnd()) as Dictionary<string, object>;
        object att;
        if (d != null && d.TryGetValue("attachment", out att)) { string u = S(att as Dictionary<string, object>, "url"); if (u != "") return u; }
      }
    } catch { }
    return null;
  }
  public void StartShare() {
    if (Sharing) return;
    Sharing = true;
    Send("", "live");
    shareBar = new ShareBar(this); shareBar.Show();
    Thread t = new Thread(ShareLoop); t.IsBackground = true; t.Start();
  }
  public void StopShare() {
    if (!Sharing) return;
    Sharing = false;
    try { if (shareBar != null) shareBar.Close(); } catch { }
    shareBar = null;
    Send("", "livestop");
  }
  void ShareLoop() {
    int idx = 0, fails = 0; long lastHash = 0; DateTime lastSent = DateTime.MinValue;
    while (Sharing && running) {
      DateTime start = DateTime.UtcNow;
      try {
        Bitmap shot = CaptureScreen();
        if (shot != null) {
          byte[] jpg = Jpeg(shot, 1280, 55L);
          shot.Dispose();
          long hsh = 1469598103934665603L;
          for (int i = 0; i < jpg.Length; i += 7) hsh = unchecked((hsh ^ jpg[i]) * 1099511628211L);
          if (hsh != lastHash || (DateTime.UtcNow - lastSent).TotalSeconds > 25) {
            string u = PutFile(servers[idx % servers.Length] + "/" + Topic + "-live", jpg); idx++;
            if (u == null) { u = PutFile(servers[idx % servers.Length] + "/" + Topic + "-live", jpg); idx++; }
            if (u != null) { lastHash = hsh; lastSent = DateTime.UtcNow; fails = 0; } else fails++;
            if (fails >= 6) {
              Ui(delegate { StopShare(); MessageBox.Show(Chat, "Screen sharing stopped: the free server is not accepting more pictures right now. Try again later.", "Office Signal"); });
              break;
            }
          }
        }
      } catch { }
      int wait = 5000 - (int)(DateTime.UtcNow - start).TotalMilliseconds;
      if (wait > 0) Thread.Sleep(wait);
    }
  }
  public void OpenViewer(string who) {
    if (viewer != null && !viewer.IsDisposed) { viewer.SetStarted(); if (viewer.WindowState == FormWindowState.Minimized) viewer.WindowState = FormWindowState.Normal; viewer.Activate(); return; }
    viewer = new LiveViewer(this, who, servers);
    viewer.Show(); viewer.Activate();
  }
  public void PillClicked() {
    if (Lit.Contains("live") && !Lit.Contains("red") && !Lit.Contains("yellow") && !Lit.Contains("green")) {
      string who = litText["live"].Replace(" is sharing the screen", ""); ClearAlert(); OpenViewer(who); return;
    }
    ShowChat();
  }
  public void SetTheme(bool light) {
    OsUi.SetTheme(light);
    try { Directory.CreateDirectory(cfgDir); File.WriteAllText(Path.Combine(cfgDir, "theme.txt"), light ? "light" : "dark"); } catch { }
    Chat.ApplyTheme();
  }
  public void SetName(string n) {
    n = (n ?? "").Trim();
    if (n == "" || n == Name) return;
    Name = n;
    try { Directory.CreateDirectory(cfgDir); File.WriteAllLines(Path.Combine(cfgDir, "me.txt"), new string[] { Name, Cid }, new UTF8Encoding(false)); } catch { }
  }
  public void Shutdown() {
    try { StopShare(); } catch { }
    running = false;
    try { tray.Visible = false; tray.Dispose(); } catch { }
    try { Pill.Close(); } catch { }
    try { Chat.Quitting = true; Chat.Close(); } catch { }
    ExitThread();
  }

  static string S(Dictionary<string, object> d, string k) {
    object o;
    if (d != null && d.TryGetValue(k, out o) && o != null) return Convert.ToString(o, CultureInfo.InvariantCulture);
    return "";
  }
  static long L(Dictionary<string, object> d, string k) {
    object o;
    if (d != null && d.TryGetValue(k, out o) && o != null) { try { return Convert.ToInt64(o, CultureInfo.InvariantCulture); } catch { } }
    return 0;
  }
  OsMsg Parse(string line, out string id) {
    id = null;
    try {
      JavaScriptSerializer js = new JavaScriptSerializer();
      Dictionary<string, object> ev = js.DeserializeObject(line) as Dictionary<string, object>;
      if (ev == null || S(ev, "event") != "message") return null;
      id = S(ev, "id");
      Dictionary<string, object> inner = js.DeserializeObject(S(ev, "message")) as Dictionary<string, object>;
      if (inner == null) return null;
      OsMsg m = new OsMsg();
      m.Name = S(inner, "name").Trim(); if (m.Name == "") return null;
      m.C = S(inner, "c"); m.Text = S(inner, "text"); m.Type = S(inner, "type");
      m.Key = S(inner, "k"); if (m.Key == "") m.Key = id;
      m.Time = L(ev, "time"); m.SeenT = L(inner, "t"); m.Url = S(inner, "url");
      m.Mine = m.C != "" ? m.C == Cid : m.Name == Name;
      return m;
    } catch { return null; }
  }

  void LoadAndListen() {
    List<OsMsg> all = new List<OsMsg>();
    Dictionary<string, string> lastIds = new Dictionary<string, string>();
    foreach (string s in servers) {
      try {
        HttpWebRequest req = (HttpWebRequest)WebRequest.Create(s + "/" + Topic + "/json?poll=1&since=12h");
        req.Timeout = 8000; req.ReadWriteTimeout = 8000;
        using (WebResponse resp = req.GetResponse())
        using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) {
          string line;
          while ((line = sr.ReadLine()) != null) {
            string id; OsMsg m = Parse(line, out id);
            if (id != null) lastIds[s] = id;
            if (m != null) all.Add(m);
          }
        }
      } catch { }
    }
    all.Sort(delegate(OsMsg a, OsMsg b) { return a.Time.CompareTo(b.Time); });
    Ui(delegate {
      foreach (OsMsg m in all) Deliver(m, false);
      Chat.View.Loaded = true;
      Chat.View.ScrollToEnd();
    });
    foreach (string s in servers) {
      string srv = s;
      string li = lastIds.ContainsKey(s) ? lastIds[s] : null;
      Thread t = new Thread(delegate() { Listen(srv, li); });
      t.IsBackground = true; t.Start();
    }
  }
  void Listen(string server, string last) {
    while (running) {
      try {
        string u = server + "/" + Topic + "/json" + (last != null ? "?since=" + last : "");
        HttpWebRequest req = (HttpWebRequest)WebRequest.Create(u);
        req.Timeout = 30000; req.ReadWriteTimeout = 120000;
        using (WebResponse resp = req.GetResponse())
        using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) {
          SetConn(server, true);
          string line;
          while (running && (line = sr.ReadLine()) != null) {
            string id; OsMsg m = Parse(line, out id);
            if (id != null) last = id;
            if (m != null) { OsMsg mm = m; Ui(delegate { Deliver(mm, true); }); }
          }
        }
      } catch { }
      SetConn(server, false);
      Thread.Sleep(3000);
    }
  }

  public void Deliver(OsMsg m, bool live) {
    if (!seenKeys.Add(m.Key)) { if (m.Mine) Chat.View.UpdateTime(m.Key, m.Time); return; }
    if (m.Type == "seen") {
      if (!m.Mine && m.SeenT > Chat.View.SeenUpTo) { Chat.View.SeenUpTo = m.SeenT; Chat.View.Invalidate(); }
      return;
    }
    if (!m.Mine) {
      Other = m.Name; Chat.View.Peer = m.Name;
      if (m.Time > lastOtherTime) lastOtherTime = m.Time;
      Chat.Head.Invalidate();
    }
    if (m.Type == "shot" && m.Img == null && m.Url != "") LoadShot(m);
    Chat.View.Add(m);
    if (live && !m.Mine && m.Type == "live") { StartAlert("live", m.Name, ""); return; }
    if (!m.Mine && m.Type == "livestop") {
      if (viewer != null && !viewer.IsDisposed) viewer.SetStopped();
      if (Lit.Contains("live")) ClearKind("live");
      return;
    }
    if (live && !m.Mine && m.Type == "shot") {
      if (!ChatActive()) { Unread++; StartAlert("msg", m.Name, "sent a screenshot"); }
      Pill.Invalidate();
    } else if (live && !m.Mine) {
      if (OsUi.IsSignal(m.Type)) StartAlert(m.Type, m.Name, m.Text);
      else if (!ChatActive()) { Unread++; StartAlert("msg", m.Name, m.Text); }
      Pill.Invalidate();
    }
  }

  void LoadShot(OsMsg m) {
    if (m.ImgLoading) return;
    m.ImgLoading = true;
    string url = m.Url;
    ThreadPool.QueueUserWorkItem(delegate {
      Image img = null;
      try { using (WebClient wc = new WebClient()) { byte[] data = wc.DownloadData(url); img = Image.FromStream(new MemoryStream(data)); } } catch { }
      Ui(delegate { if (img != null) m.Img = img; else m.ImgFailed = true; Chat.View.Invalidate(); });
    });
  }
  public void SendShot(Bitmap shot) {
    byte[] jpg = Jpeg(shot, 1920, 90L);
    Image preview = new Bitmap(shot);
    ThreadPool.QueueUserWorkItem(delegate {
      string url = null;
      foreach (string s in servers) { url = PutFile(s + "/" + Topic + "-files", jpg); if (url != null) break; }
      string u2 = url;
      Ui(delegate {
        if (string.IsNullOrEmpty(u2)) { MessageBox.Show(Chat, "The screenshot could not be sent. Check the internet and try again.", "Office Signal"); return; }
        Send("", "shot", u2, preview);
      });
    });
  }
  public void Send(string text, string type) { Send(text, type, null, null); }
  public void Send(string text, string type, string url, Image img) {
    Dictionary<string, object> d = new Dictionary<string, object>();
    string key = Cid + "-" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture) + "-" + rnd.Next(100000).ToString(CultureInfo.InvariantCulture);
    d["name"] = Name; d["c"] = Cid; d["k"] = key; d["text"] = text ?? "";
    if (type != null) d["type"] = type;
    if (type == "seen") d["t"] = seenSent;
    if (url != null) d["url"] = url;
    string body = new JavaScriptSerializer().Serialize(d);
    if (type != "seen") {
      OsMsg m = new OsMsg();
      m.Name = Name; m.C = Cid; m.Text = text ?? ""; m.Type = type ?? ""; m.Key = key; m.Time = OsUi.Now(); m.Mine = true;
      m.Url = url ?? ""; m.Img = img;
      Deliver(m, false);
      Chat.View.ScrollToEnd();
    }
    foreach (string s in servers) {
      string postUrl = s + "/" + Topic;
      ThreadPool.QueueUserWorkItem(delegate {
        try { using (WebClient wc = new WebClient()) { wc.Encoding = Encoding.UTF8; wc.UploadString(postUrl, "POST", body); } } catch { }
      });
    }
  }
}
'@
$goodFile = $global:goodFile
$built = $false
try { Add-Type -TypeDefinition $cs -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.Web.Extensions -IgnoreWarnings -ErrorAction Stop; $built = $true }
catch {
  $det = (@($Error) | Select-Object -First 6 | ForEach-Object { $_.ToString() }) -join "`r`n"
  if (-not $global:fallbackTried -and (Test-Path $goodFile)) {
    # this version does not work: run the last version that worked instead
    $global:fallbackTried = $true
    try { $global:mutex.ReleaseMutex() } catch {}
    iex ([IO.File]::ReadAllText($goodFile))
    return
  }
  throw "Build failed:`r`n$det"
}
if ($built -and $global:appCache -and (Test-Path $global:appCache)) {
  try {
    $cacheText = [IO.File]::ReadAllText($global:appCache)
    $goodVer = ''; if (Test-Path $goodFile) { $goodVer = Get-TextVer ([IO.File]::ReadAllText($goodFile)) }
    if ((Get-TextVer $cacheText) -eq $global:appVersion -and (Get-VerNum $global:appVersion) -ge (Get-VerNum $goodVer)) { Copy-Item -LiteralPath $global:appCache -Destination $goodFile -Force }
  } catch {}
}

# ---- Os icon file + desktop shortcut ----
$global:icoPath = Join-Path $global:dir 'OfficeSignal.ico'
try {
  $big = [OsUi]::MakeBitmap(256); $ms = New-Object IO.MemoryStream
  $big.Save($ms, [Drawing.Imaging.ImageFormat]::Png); $png = $ms.ToArray()
  $fsI = [IO.File]::Create($global:icoPath); $bw = New-Object IO.BinaryWriter($fsI)
  $bw.Write([UInt16]0); $bw.Write([UInt16]1); $bw.Write([UInt16]1)
  $bw.Write([Byte]0); $bw.Write([Byte]0); $bw.Write([Byte]0); $bw.Write([Byte]0)
  $bw.Write([UInt16]1); $bw.Write([UInt16]32); $bw.Write([UInt32]$png.Length); $bw.Write([UInt32]22)
  $bw.Write($png); $bw.Close()
} catch {}
try {
  $stable = Join-Path $global:dir 'Office Signal.bat'
  $launcherText = @"
@echo off
start "" powershell -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -Command "`$f='%~f0'; `$s=[IO.File]::ReadAllText(`$f); `$i=`$s.LastIndexOf('#==PS==#'); iex `$s.Substring(`$i)"
exit /b
#==PS==#
Add-Type -AssemblyName System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
`$global:launcher = `$f
`$global:topic = '$($global:topic)'
`$global:appUrl = 'https://raw.githubusercontent.com/redaadlany/office-signal/main/app.ps1'
`$d = Join-Path `$env:APPDATA 'OfficeSignal'
New-Item -ItemType Directory -Force -Path `$d | Out-Null
`$global:appCache = Join-Path `$d 'app.ps1'
function VerOf([string]`$t) { if (`$t -match "appVersion = '([^']+)'") { try { `$q = `$matches[1].Split('.'); return ([int64](`$q[0].Replace('-','')) * 1000 + [int64]`$q[1]) } catch {} }; return [int64]0 }
try {
  `$url = `$global:appUrl + '?t=' + [DateTime]::UtcNow.Ticks
  try { `$w0 = New-Object Net.WebClient; `$w0.Headers['User-Agent'] = 'OfficeSignal'; `$info = `$w0.DownloadString('https://api.github.com/repos/redaadlany/office-signal/commits/main'); if (`$info -match '"sha":\s*"([0-9a-f]{40})"') { `$url = 'https://raw.githubusercontent.com/redaadlany/office-signal/' + `$matches[1] + '/app.ps1' } } catch {}
  `$wc = New-Object Net.WebClient; `$wc.Encoding = [Text.Encoding]::UTF8
  `$code = `$wc.DownloadString(`$url)
  `$cur = ''; if (Test-Path `$global:appCache) { `$cur = [IO.File]::ReadAllText(`$global:appCache) }
  if (`$code -match '#==APP==#' -and (VerOf `$code) -ge (VerOf `$cur)) { [IO.File]::WriteAllText(`$global:appCache, `$code, (New-Object Text.UTF8Encoding `$false)) }
} catch {}
if (Test-Path `$global:appCache) {
  iex ([IO.File]::ReadAllText(`$global:appCache))
} else {
  [Windows.Forms.MessageBox]::Show("Office Signal needs the internet the first time it starts.``r``nCheck the connection and try again.", 'Office Signal') | Out-Null
}
"@
  try { [IO.File]::WriteAllText($stable, ($launcherText -replace "`r?`n", "`r`n"), [Text.Encoding]::ASCII) } catch {}
  if (Test-Path $stable) {
    $global:launcher = $stable
    $desk = [Environment]::GetFolderPath('Desktop')
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut((Join-Path $desk 'Office Signal.lnk'))
    $lnk.TargetPath = $stable; $lnk.WorkingDirectory = $global:dir
    $lnk.IconLocation = "$($global:icoPath),0"; $lnk.WindowStyle = 7; $lnk.Description = 'Office Signal'
    $lnk.Save()
    # start automatically when Windows starts (unless nostartup.txt exists)
    $startupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Office Signal.lnk'
    if (Test-Path (Join-Path $global:dir 'nostartup.txt')) {
      if (Test-Path $startupLnk) { Remove-Item -LiteralPath $startupLnk -Force -ErrorAction SilentlyContinue }
    } else {
      $sl = $sh.CreateShortcut($startupLnk)
      $sl.TargetPath = $stable; $sl.WorkingDirectory = $global:dir
      $sl.IconLocation = "$($global:icoPath),0"; $sl.WindowStyle = 7; $sl.Description = 'Office Signal'
      $sl.Save()
    }
  }
} catch {}

$global:app = New-Object OsApp($global:myName, $global:myCid, $global:topic)

# ---- automatic updates: check the online copy every 20 minutes ----
function Restart-App {
  try { $global:app.Shutdown() } catch {}
  try { $global:mutex.ReleaseMutex() } catch {}
  if ($global:launcher -and (Test-Path $global:launcher)) { Start-Process -FilePath $global:launcher -WindowStyle Hidden }
  [Environment]::Exit(0)
}
function Get-LatestApp {
  $wc = New-Object Net.WebClient; $wc.Encoding = [Text.Encoding]::UTF8
  $wc.Headers['User-Agent'] = 'OfficeSignal'
  $url = $global:appUrl + '?t=' + [DateTime]::UtcNow.Ticks
  try {
    $info = $wc.DownloadString('https://api.github.com/repos/redaadlany/office-signal/commits/main')
    if ($info -match '"sha":\s*"([0-9a-f]{40})"') { $url = 'https://raw.githubusercontent.com/redaadlany/office-signal/' + $matches[1] + '/app.ps1' }
  } catch {}
  $wc2 = New-Object Net.WebClient; $wc2.Encoding = [Text.Encoding]::UTF8
  $t = $wc2.DownloadString($url)
  if ($t -match '#==APP==#') { return $t } else { return $null }
}
$global:updTimer = New-Object Windows.Forms.Timer
$global:updTimer.Interval = 5 * 60 * 1000
$global:updTimer.Add_Tick({
  try {
    $new = Get-LatestApp
    if ($new -and (Get-VerNum (Get-TextVer $new)) -gt (Get-VerNum $global:appVersion)) {
      [IO.File]::WriteAllText($global:appCache, $new, (New-Object Text.UTF8Encoding $false))
      Restart-App
    }
  } catch {}
})
if ($global:appUrl) { $global:updTimer.Start() }

# another copy with a newer version asks this one to close
$global:quitEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Local\OfficeSignalQuit')
$global:quitTimer = New-Object Windows.Forms.Timer
$global:quitTimer.Interval = 1000
$global:quitTimer.Add_Tick({
  if ($global:quitEvent.WaitOne(0)) {
    try { $global:app.Shutdown() } catch {}
    try { $global:mutex.ReleaseMutex() } catch {}
    [Environment]::Exit(0)
  }
})
$global:quitTimer.Start()

[Windows.Forms.Application]::Run($global:app)
}
catch {
  [Windows.Forms.MessageBox]::Show("Office Signal could not start:`r`n`r`n$($_.Exception.Message)", 'Office Signal') | Out-Null
}
[Environment]::Exit(0)

#==APP==#
# Office Signal app - downloaded and run by the Office Signal launcher.
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic
try {
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$global:appVersion = '2026-11-19.4'

# ---- only one copy running: if already running, ask it to show the window ----
$created = $false
$global:mutex = New-Object System.Threading.Mutex($true, 'Local\OfficeSignalApp', [ref]$created)
if (-not $created) {
  try { ([System.Threading.EventWaitHandle]::OpenExisting('Local\OfficeSignalShow')).Set() | Out-Null } catch {}
  return
}

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
}

public static class OsNative {
  [DllImport("user32.dll")] public static extern bool ReleaseCapture();
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);
  [DllImport("shell32.dll", CharSet = CharSet.Unicode)] public static extern int SetCurrentProcessExplicitAppUserModelID(string id);
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
      Bg = Color.White; Box = Color.FromArgb(239, 243, 244); Bubble = Color.FromArgb(239, 243, 244);
      Line = Color.FromArgb(207, 217, 222); Ink = Color.FromArgb(15, 20, 25); Mute = Color.FromArgb(83, 100, 113);
    } else {
      Bg = Color.FromArgb(22, 24, 28); Box = Color.FromArgb(32, 35, 39); Bubble = Color.FromArgb(47, 51, 54);
      Line = Color.FromArgb(47, 51, 54); Ink = Color.FromArgb(231, 233, 234); Mute = Color.FromArgb(113, 118, 123);
    }
  }
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
  public static Color KindColor(string k) { if (k == "red") return Red; if (k == "yellow") return Yellow; if (k == "msg") return Blue; return Green; }
  public static Color KindDark(string k) { if (k == "red") return Color.FromArgb(122, 0, 0); if (k == "yellow") return Color.FromArgb(122, 96, 0); if (k == "msg") return Color.FromArgb(11, 79, 128); return Color.FromArgb(0, 84, 56); }
  public static Color KindText(string k) { return k == "yellow" ? DarkText : Color.White; }
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
    if (hover) { using (GraphicsPath p = OsUi.Round(r, Kind == "send" ? Height / 2f : 8)) using (SolidBrush b = new SolidBrush(Kind == "send" ? Color.FromArgb(30, 29, 155, 240) : OsUi.Hover())) g.FillPath(b, p); }
    if (Outline) { using (GraphicsPath p = OsUi.Round(r, 8)) using (Pen pn = new Pen(OsUi.Blue, 1)) g.DrawPath(pn, p); }
    float cx = Width / 2f, cy = Height / 2f;
    Color col = Kind == "send" ? OsUi.Blue : OsUi.Ink;
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
  string ChipText(OsMsg m) { return (m.Mine ? "You sent: " : "") + OsUi.KindLabel(m.Type) + (m.Text != "" ? " - " + m.Text : ""); }
  Size BubbleSize(OsMsg m) {
    Size s = TextRenderer.MeasureText(m.Text == "" ? " " : m.Text, fBody, new Size(MaxBubble() - 28, int.MaxValue), WrapFlags);
    return new Size(Math.Min(MaxBubble(), s.Width + 28), s.Height + 18);
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
  int ItemHeight(OsMsg m, bool vu) {
    int h = OsUi.IsSignal(m.Type) ? 30 + 4 + 16 + 12 : BubbleSize(m).Height + 4 + 16 + 8;
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
    int y = 14 - scroll;
    for (int i = 0; i < Items.Count; i++) {
      OsMsg m = Items[i];
      int h = ItemHeight(m, i == lastSeenMine);
      if (y + h < 0 || y > Height) { y += h; continue; }
      string meta = (m.Mine ? "" : m.Name + " - ") + OsUi.Clock(m.Time);
      if (OsUi.IsSignal(m.Type)) {
        string t = ChipText(m);
        Size ts = TextRenderer.MeasureText(t, fChip);
        int w = Math.Min(Width - 40, ts.Width + 32);
        Rectangle r = new Rectangle((Width - w) / 2, y, w, 30);
        using (GraphicsPath p = OsUi.Round(r, 15)) using (SolidBrush b = new SolidBrush(OsUi.KindColor(m.Type))) g.FillPath(b, p);
        TextRenderer.DrawText(g, t, fChip, r, OsUi.KindText(m.Type), OsUi.KindColor(m.Type), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        TextRenderer.DrawText(g, meta, fMeta, new Rectangle(0, y + 34, Width, 16), OsUi.Mute, TextFormatFlags.HorizontalCenter);
      } else {
        Size bs = BubbleSize(m);
        int x = m.Mine ? Width - 14 - bs.Width : 14;
        Rectangle r = new Rectangle(x, y, bs.Width, bs.Height);
        Color fill = m.Mine ? OsUi.Blue : OsUi.Bubble;
        using (GraphicsPath p = OsUi.Round(r, Math.Min(18, bs.Height / 2f))) using (SolidBrush b = new SolidBrush(fill)) g.FillPath(b, p);
        TextRenderer.DrawText(g, m.Text, fBody, new Rectangle(x + 14, y + 9, bs.Width - 28, bs.Height - 18), m.Mine ? Color.White : OsUi.Ink, fill, WrapFlags);
        TextRenderer.DrawText(g, meta, fMeta, new Rectangle(m.Mine ? 0 : x + 6, y + bs.Height + 4, m.Mine ? Width - 20 : Width, 16), OsUi.Mute, m.Mine ? TextFormatFlags.Right : TextFormatFlags.Left);
      }
      if (i == lastSeenMine) {
        int vy = y + h - 18 - (OsUi.IsSignal(m.Type) ? 10 : 6);
        Size vs = TextRenderer.MeasureText("Vu", fVu);
        int vx = OsUi.IsSignal(m.Type) ? (Width + 18 - vs.Width) / 2 : Width - 20 - vs.Width;
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
  public ConfirmForm() {
    FormBorderStyle = FormBorderStyle.None; StartPosition = FormStartPosition.CenterParent; ShowInTaskbar = false;
    BackColor = Color.Black; Size = new Size(320, 190); TopMost = true;
    Region = new Region(OsUi.Round(new RectangleF(0, 0, 320, 190), 16));
    Label t = new Label(); t.Text = "Send emergency?"; t.ForeColor = OsUi.Ink; t.Font = new Font("Segoe UI Semibold", 13f); t.AutoSize = false; t.TextAlign = ContentAlignment.MiddleCenter; t.Bounds = new Rectangle(0, 18, 320, 30);
    Label s = new Label(); s.Text = "Your colleague's Messages bar will blink red."; s.ForeColor = OsUi.Mute; s.Font = new Font("Segoe UI", 9.5f); s.AutoSize = false; s.TextAlign = ContentAlignment.MiddleCenter; s.Bounds = new Rectangle(10, 50, 300, 22);
    RoundBtn yes = new RoundBtn("Send emergency", OsUi.Red, Color.White); yes.BackColor = Color.Black; yes.Bounds = new Rectangle(24, 84, 272, 40);
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
  IconBtn sendBtn;
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
    bRed = new RoundBtn("Emergency", OsUi.Red, Color.White);
    bYel = new RoundBtn("Come", OsUi.Yellow, OsUi.DarkText);
    bGrn = new RoundBtn("Call off", OsUi.Green, Color.White);
    sig.Controls.Add(bRed); sig.Controls.Add(bYel); sig.Controls.Add(bGrn);
    sig.Resize += delegate { LayoutSignals(); };
    View = new ChatView();
    bottom = new Panel(); bottom.BackColor = OsUi.Bg;
    bottom.Paint += delegate(object o, PaintEventArgs e) {
      Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
      using (Pen p = new Pen(OsUi.Line)) g.DrawLine(p, 0, 0, bottom.Width, 0);
      using (GraphicsPath gp = OsUi.Round(new RectangleF(12, 12, bottom.Width - 24, 48), 16)) using (SolidBrush b = new SolidBrush(OsUi.Box)) g.FillPath(b, gp);
    };
    input = new TextBox(); input.BorderStyle = BorderStyle.None; input.BackColor = OsUi.Box; input.ForeColor = OsUi.Ink; input.Font = new Font("Segoe UI", 11f);
    input.MaxLength = 2000;
    sendBtn = new IconBtn("send"); sendBtn.BackColor = OsUi.Box;
    bottom.Controls.Add(input); bottom.Controls.Add(sendBtn);
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
    try { OsNative.SendMessage(input.Handle, 0x1501, (IntPtr)1, "Start a new message"); } catch { }
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
    input.BackColor = OsUi.Box; input.ForeColor = OsUi.Ink; sendBtn.BackColor = OsUi.Box; sendBtn.Invalidate();
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
    input.Bounds = new Rectangle(28, 12 + (48 - input.PreferredHeight) / 2, boxW - 16 - 44, input.PreferredHeight);
    sendBtn.Bounds = new Rectangle(12 + boxW - 44, 18, 36, 36);
  }
  public void Place() {
    Rectangle wa = Screen.PrimaryScreen.WorkingArea;
    Location = new Point(wa.Right - Width - 16, wa.Bottom - Height - 16);
  }
  public void FocusInput() { try { input.Focus(); } catch { } }
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
    int w = 320, h = 62;
    Rectangle wa = Screen.PrimaryScreen.WorkingArea;
    Bounds = new Rectangle(wa.Right - w - 16, wa.Bottom - h - 16, w, h);
    using (GraphicsPath p = new GraphicsPath()) {
      p.AddArc(0, 0, h, h, 90, 180); p.AddArc(w - h, 0, h, h, 270, 180); p.CloseFigure();
      Region = new Region(p);
    }
  }
  protected override bool ShowWithoutActivation { get { return true; } }
  protected override CreateParams CreateParams {
    get { CreateParams cp = base.CreateParams; cp.ExStyle |= 0x00000080 | 0x00000008; return cp; }
  }
  protected override void OnClick(EventArgs e) { base.OnClick(e); app.ShowChat(); }
  protected override void OnPaint(PaintEventArgs e) {
    Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
    int w = Width, h = Height;
    string k = app.Alert;
    Color bg = k != null ? (app.BlinkOn ? OsUi.KindDark(k) : OsUi.KindColor(k)) : OsUi.PillBg;
    g.Clear(bg);
    using (GraphicsPath p = new GraphicsPath()) {
      p.AddArc(0, 0, h - 1, h - 1, 90, 180); p.AddArc(w - h, 0, h - 1, h - 1, 270, 180); p.CloseFigure();
      using (Pen pn = new Pen(k != null ? bg : OsUi.PillLine, 1.5f)) g.DrawPath(pn, p);
    }
    Color fg = (k == "yellow" && !app.BlinkOn) ? OsUi.DarkText : (k != null ? Color.White : OsUi.PillInk);
    float cy = h / 2f;
    using (Pen pen = new Pen(fg, 2f)) {
      pen.LineJoin = LineJoin.Round;
      g.DrawPolygon(pen, new PointF[] { new PointF(24, cy - 1), new PointF(42, cy - 9), new PointF(34, cy + 9), new PointF(31.5f, cy + 1.5f) });
    }
    string title;
    if (k == "msg") title = app.AlertFrom + ": " + app.AlertText.Replace("\r", " ").Replace("\n", " ");
    else if (k != null) title = OsUi.KindLabel(k) + (app.AlertFrom != "" ? " - " + app.AlertFrom : "");
    else title = "Messages";
    TextRenderer.DrawText(g, title, k != null ? fMid : fBig, new Rectangle(56, 0, w - 56 - 64, h), fg, bg, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
    int ax = w - 52, ay = (h - 34) / 2;
    if (app.Other != "") OsUi.DrawAvatar(g, app.Other, new RectangleF(ax, ay, 34, 34), bg, fAvatar);
    if (app.Unread > 0 && k == null) {
      int bx = ax - 30, by = (h - 22) / 2;
      using (SolidBrush b = new SolidBrush(OsUi.Blue)) g.FillEllipse(b, bx, by, 22, 22);
      string n = app.Unread > 9 ? "9+" : app.Unread.ToString();
      OsUi.DrawCentered(g, n, fSmall, new RectangleF(bx, by, 22, 22), Color.White);
    }
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
    try { OsNative.SetCurrentProcessExplicitAppUserModelID("RoboTalent.OfficeSignal"); } catch { }
    try { OsUi.SetTheme(File.Exists(Path.Combine(cfgDir, "theme.txt")) && File.ReadAllText(Path.Combine(cfgDir, "theme.txt")).Trim() == "light"); } catch { }
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
    blink.Tick += delegate { BlinkOn = !BlinkOn; Pill.TopMost = true; Pill.Invalidate(); };
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
  void ClearAlert() { Alert = null; AlertFrom = ""; AlertText = ""; blink.Stop(); Pill.Invalidate(); }
  void StartAlert(string kind, string from, string text) {
    Alert = kind; AlertFrom = from; AlertText = text ?? ""; BlinkOn = false;
    if (!Pill.Visible) Pill.Show();
    Pill.TopMost = true; Pill.Invalidate();
    blink.Start();
  }
  void Watch() {
    bool shown = Chat.Visible && Chat.WindowState != FormWindowState.Minimized;
    bool wantPill = Alert != null || !shown;
    if (wantPill && !Pill.Visible) { Pill.Show(); Pill.TopMost = true; }
    else if (!wantPill && Pill.Visible) Pill.Hide();
    if (ChatActive()) {
      if (Alert == "msg") ClearAlert();
      if (Unread != 0) { Unread = 0; Pill.Invalidate(); }
      if (lastOtherTime > seenSent) { seenSent = lastOtherTime; Send("", "seen"); }
    }
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
      m.Time = L(ev, "time"); m.SeenT = L(inner, "t");
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
    Chat.View.Add(m);
    if (live && !m.Mine) {
      if (OsUi.IsSignal(m.Type)) StartAlert(m.Type, m.Name, m.Text);
      else if (!ChatActive()) { Unread++; if (Alert == null || Alert == "msg") StartAlert("msg", m.Name, m.Text); }
      Pill.Invalidate();
    }
  }

  public void Send(string text, string type) {
    Dictionary<string, object> d = new Dictionary<string, object>();
    string key = Cid + "-" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture) + "-" + rnd.Next(100000).ToString(CultureInfo.InvariantCulture);
    d["name"] = Name; d["c"] = Cid; d["k"] = key; d["text"] = text ?? "";
    if (type != null) d["type"] = type;
    if (type == "seen") d["t"] = seenSent;
    string body = new JavaScriptSerializer().Serialize(d);
    if (type != "seen") {
      OsMsg m = new OsMsg();
      m.Name = Name; m.C = Cid; m.Text = text ?? ""; m.Type = type ?? ""; m.Key = key; m.Time = OsUi.Now(); m.Mine = true;
      Deliver(m, false);
      Chat.View.ScrollToEnd();
    }
    foreach (string s in servers) {
      string url = s + "/" + Topic;
      ThreadPool.QueueUserWorkItem(delegate {
        try { using (WebClient wc = new WebClient()) { wc.Encoding = Encoding.UTF8; wc.UploadString(url, "POST", body); } } catch { }
      });
    }
  }
}
'@
try { Add-Type -TypeDefinition $cs -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.Web.Extensions -IgnoreWarnings -ErrorAction Stop }
catch { $det = (@($Error) | Select-Object -First 6 | ForEach-Object { $_.ToString() }) -join "`r`n"; throw "Build failed:`r`n$det" }

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
  if ($global:launcher -and (Test-Path $global:launcher) -and ($global:launcher -ne $stable)) { Copy-Item -LiteralPath $global:launcher -Destination $stable -Force }
  if (Test-Path $stable) {
    $global:launcher = $stable
    $desk = [Environment]::GetFolderPath('Desktop')
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut((Join-Path $desk 'Office Signal.lnk'))
    $lnk.TargetPath = $stable; $lnk.WorkingDirectory = $global:dir
    $lnk.IconLocation = "$($global:icoPath),0"; $lnk.WindowStyle = 7; $lnk.Description = 'Office Signal'
    $lnk.Save()
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

[Windows.Forms.Application]::Run($global:app)
}
catch {
  [Windows.Forms.MessageBox]::Show("Office Signal could not start:`r`n`r`n$($_.Exception.Message)", 'Office Signal') | Out-Null
}
[Environment]::Exit(0)

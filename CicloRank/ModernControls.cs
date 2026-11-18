// ============================================================================
// File: ModernControls.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Implements the visual theme engine (Light & Dark modes), persistent theme
//   settings, Windows DWM/UxTheme native dark-mode P/Invoke interop, and custom
//   anti-aliased owner-drawn WinForms controls (RoundedButton, RoundedCardPanel,
//   and RoundedInputHost).
// ============================================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CicloRank
{
    /// <summary>
    /// Immutable value struct encapsulating background, foreground, and border colors
    /// for themed <see cref="RoundedButton"/> instances.
    /// </summary>
    public readonly struct ButtonColorScheme
    {
        /// <summary>Gets the button fill color.</summary>
        public Color Back { get; }
        /// <summary>Gets the button text or icon color.</summary>
        public Color Fore { get; }
        /// <summary>Gets the button outline border color.</summary>
        public Color Border { get; }

        /// <summary>
        /// Initializes a new <see cref="ButtonColorScheme"/> with explicit fill, text, and border colors.
        /// </summary>
        public ButtonColorScheme(Color back, Color fore, Color border)
        {
            Back = back;
            Fore = fore;
            Border = border;
        }
    }

    /// <summary>
    /// Centralizes Light and Dark mode color palettes, persistent theme storage (tema.cfg),
    /// rounded GDI+ path geometry helpers, and Windows DWM / UxTheme P/Invoke interop
    /// for dark title bars and native scrollbars.
    /// </summary>
    public static class ThemeManager
    {
        /// <summary>
        /// Gets (and creates if necessary) the per-user application data folder under %LocalAppData%\CicloRank.
        /// </summary>
        public static string CartellaDatiUtente
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "CicloRank");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        private static string FilePreferenzaTema => Path.Combine(CartellaDatiUtente, "tema.cfg");
        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

        private static bool _isDark = InizializzaTema();

        /// <summary>
        /// Gets a value indicating whether Dark Mode is currently active.
        /// </summary>
        public static bool IsDark => _isDark;

        /// <summary>
        /// Occurs when the application theme switches between Light Mode and Dark Mode.
        /// </summary>
        public static event Action? TemaCambiato;

        private static readonly ConditionalWeakTable<Control, object> _controlliMonitorati = new ConditionalWeakTable<Control, object>();

        // Windows UxTheme & DWM P/Invoke declarations used to apply native dark mode to window title bars and scrollbars.
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string? pszSubAppName, string? pszSubIdList);

        [DllImport("uxtheme.dll", EntryPoint = "#133")]
        private static extern bool AllowDarkModeForWindow(IntPtr hWnd, bool allow);

        [DllImport("uxtheme.dll", EntryPoint = "#135")]
        private static extern int SetPreferredAppMode(int appMode);

        [DllImport("uxtheme.dll", EntryPoint = "#136")]
        private static extern void FlushMenuThemes();

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_THEMECHANGED = 0x031A;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        /// <summary>
        /// Configures the process-wide Windows preferred app theme mode (ForceDark vs. ForceLight)
        /// on Windows 10 1903+ and Windows 11.
        /// </summary>
        public static void AggiornaPreferenzaProcesso()
        {
            try
            {
                // 2 = ForceDark, 3 = ForceLight (Windows 10 1903+ / Windows 11)
                SetPreferredAppMode(_isDark ? 2 : 3);
                FlushMenuThemes();
            }
            catch { }
        }

        /// <summary>
        /// Applies native dark or light window chrome (including the DWM immersive dark title bar)
        /// and recursively themes child controls on the specified <see cref="Form"/>.
        /// </summary>
        /// <param name="form">The target Windows Forms window.</param>
        public static void ApplicaTemaFinestra(Form form)
        {
            if (form == null) return;
            AggiornaPreferenzaProcesso();
            ApplicaTemaControlloNativo(form);

            void ApplicaBarraTitolo()
            {
                if (!form.IsHandleCreated) return;
                int useDark = _isDark ? 1 : 0;
                try
                {
                    if (DwmSetWindowAttribute(form.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref useDark, sizeof(int)) != 0)
                    {
                        DwmSetWindowAttribute(form.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref useDark, sizeof(int));
                    }
                }
                catch { }
            }

            if (form.IsHandleCreated)
                ApplicaBarraTitolo();
            else
                form.HandleCreated += (s, e) => ApplicaBarraTitolo();
        }

        /// <summary>
        /// Attaches theme listeners and applies native Explorer / DarkMode_Explorer visual styles
        /// to a control and all of its nested child controls.
        /// </summary>
        public static void ApplicaTemaControlloNativo(Control control)
        {
            if (control == null) return;

            if (!_controlliMonitorati.TryGetValue(control, out _))
            {
                _controlliMonitorati.Add(control, new object());
                control.HandleCreated += (s, e) => ApplicaHandleNativo(control);
                control.ControlAdded += (s, e) =>
                {
                    if (e.Control != null)
                        ApplicaTemaControlloNativo(e.Control);
                };
            }

            if (control.IsHandleCreated)
            {
                ApplicaHandleNativo(control);
            }

            foreach (Control child in control.Controls)
            {
                ApplicaTemaControlloNativo(child);
            }
        }

        /// <summary>
        /// Sends native Win32 theme messages and updates the UxTheme window style on a created handle.
        /// </summary>
        private static void ApplicaHandleNativo(Control control)
        {
            if (!control.IsHandleCreated) return;
            try { AllowDarkModeForWindow(control.Handle, _isDark); } catch { }
            try { SetWindowTheme(control.Handle, _isDark ? "DarkMode_Explorer" : "Explorer", null); } catch { }
            try { SendMessage(control.Handle, WM_THEMECHANGED, IntPtr.Zero, IntPtr.Zero); } catch { }
            control.Invalidate(true);
        }

        /// <summary>
        /// Reads the saved theme preference ("dark" or "light") from %LocalAppData%\CicloRank\tema.cfg.
        /// </summary>
        private static bool InizializzaTema()
        {
            try
            {
                if (!File.Exists(FilePreferenzaTema) && File.Exists("tema.cfg"))
                {
                    File.Copy("tema.cfg", FilePreferenzaTema, true);
                }
                if (File.Exists(FilePreferenzaTema))
                {
                    string val = File.ReadAllText(FilePreferenzaTema, Utf8NoBom).Trim().ToLowerInvariant();
                    return val == "dark";
                }
            }
            catch { }

            return false;
        }

        /// <summary>
        /// Toggles the active theme between Light Mode and Dark Mode.
        /// </summary>
        public static void ToggleTema()
        {
            ImpostaTema(!_isDark);
        }

        /// <summary>
        /// Sets the active theme mode, persists it to disk, and raises <see cref="TemaCambiato"/>.
        /// </summary>
        /// <param name="dark"><c>true</c> for Dark Mode; <c>false</c> for Light Mode.</param>
        public static void ImpostaTema(bool dark)
        {
            if (_isDark == dark) return;
            _isDark = dark;

            try
            {
                File.WriteAllText(FilePreferenzaTema, _isDark ? "dark" : "light", Utf8NoBom);
            }
            catch { }

            TemaCambiato?.Invoke();
        }

        // Window and card surfaces (macOS-inspired Light vs. Dark Monochrome Grayscale)
        // Window, card, input, typography, grid, and button color palette definitions for Light and Dark modes.
        public static Color WindowBack => _isDark ? Color.FromArgb(24, 24, 26) : Color.FromArgb(245, 245, 247);
        public static Color CardBack => _isDark ? Color.FromArgb(34, 34, 37) : Color.FromArgb(255, 255, 255);
        public static Color CardBorder => _isDark ? Color.FromArgb(56, 56, 61) : Color.FromArgb(224, 224, 230);

        public static Color InputBack => _isDark ? Color.FromArgb(44, 44, 48) : Color.FromArgb(250, 250, 252);
        public static Color InputBorder => _isDark ? Color.FromArgb(72, 72, 78) : Color.FromArgb(206, 206, 214);
        public static Color InputFocusBorder => _isDark ? Color.FromArgb(150, 150, 158) : Color.FromArgb(0, 122, 255);

        public static Color TextPrimary => _isDark ? Color.FromArgb(232, 232, 236) : Color.FromArgb(28, 28, 30);
        public static Color TextSecondary => _isDark ? Color.FromArgb(156, 156, 164) : Color.FromArgb(110, 110, 118);
        public static Color TextAccent => _isDark ? Color.FromArgb(216, 216, 222) : Color.FromArgb(0, 82, 180);

        public static Color SuccessText => _isDark ? Color.FromArgb(224, 224, 230) : Color.FromArgb(24, 116, 58);
        public static Color WarningText => _isDark ? Color.FromArgb(242, 85, 95) : Color.FromArgb(196, 48, 58);
        public static Color WarningBack => _isDark ? Color.FromArgb(56, 34, 38) : Color.FromArgb(254, 240, 241);

        // DataGridView table styling
        public static Color GridBack => _isDark ? Color.FromArgb(30, 30, 33) : Color.FromArgb(255, 255, 255);
        public static Color GridAltBack => _isDark ? Color.FromArgb(36, 36, 40) : Color.FromArgb(248, 248, 250);
        public static Color GridHeaderBack => _isDark ? Color.FromArgb(44, 44, 48) : Color.FromArgb(242, 242, 246);
        public static Color GridHeaderText => _isDark ? Color.FromArgb(220, 220, 226) : Color.FromArgb(58, 58, 64);
        public static Color GridLine => _isDark ? Color.FromArgb(50, 50, 55) : Color.FromArgb(232, 232, 238);
        public static Color GridSelectionBack => _isDark ? Color.FromArgb(64, 64, 70) : Color.FromArgb(224, 236, 252);
        public static Color GridSelectionText => _isDark ? Color.FromArgb(245, 245, 248) : Color.FromArgb(20, 20, 24);

        // Pulsanti (in modalità scura rigorosamente in scala di grigi riposante)
        public static ButtonColorScheme BtnDatiGara => _isDark
            ? new ButtonColorScheme(Color.FromArgb(46, 46, 50), Color.FromArgb(230, 230, 234), Color.FromArgb(72, 72, 78))
            : new ButtonColorScheme(Color.FromArgb(247, 232, 184), Color.FromArgb(56, 44, 14), Color.FromArgb(226, 208, 152));

        public static ButtonColorScheme BtnCaricaDati => _isDark
            ? new ButtonColorScheme(Color.FromArgb(52, 52, 56), Color.FromArgb(230, 230, 234), Color.FromArgb(78, 78, 84))
            : new ButtonColorScheme(Color.FromArgb(214, 228, 246), Color.FromArgb(24, 48, 78), Color.FromArgb(186, 206, 232));

        public static ButtonColorScheme BtnAzzeraDati => _isDark
            ? new ButtonColorScheme(Color.FromArgb(62, 62, 67), Color.FromArgb(236, 236, 240), Color.FromArgb(90, 90, 96))
            : new ButtonColorScheme(Color.FromArgb(220, 68, 78), Color.FromArgb(255, 255, 255), Color.FromArgb(196, 52, 62));

        public static ButtonColorScheme BtnStart => _isDark
            ? new ButtonColorScheme(Color.FromArgb(54, 54, 58), Color.FromArgb(234, 234, 238), Color.FromArgb(82, 82, 88))
            : new ButtonColorScheme(Color.FromArgb(192, 234, 204), Color.FromArgb(20, 64, 34), Color.FromArgb(160, 212, 176));

        public static ButtonColorScheme BtnPausa => _isDark
            ? new ButtonColorScheme(Color.FromArgb(44, 44, 48), Color.FromArgb(220, 220, 226), Color.FromArgb(70, 70, 76))
            : new ButtonColorScheme(Color.FromArgb(245, 228, 172), Color.FromArgb(64, 48, 12), Color.FromArgb(224, 204, 142));

        public static ButtonColorScheme BtnAzzeraTimer => _isDark
            ? new ButtonColorScheme(Color.FromArgb(60, 60, 65), Color.FromArgb(230, 230, 234), Color.FromArgb(88, 88, 94))
            : new ButtonColorScheme(Color.FromArgb(246, 188, 188), Color.FromArgb(84, 24, 28), Color.FromArgb(226, 160, 160));

        public static ButtonColorScheme BtnExportExcel => _isDark
            ? new ButtonColorScheme(Color.FromArgb(50, 50, 54), Color.FromArgb(232, 232, 236), Color.FromArgb(78, 78, 84))
            : new ButtonColorScheme(Color.FromArgb(44, 128, 78), Color.FromArgb(255, 255, 255), Color.FromArgb(36, 108, 64));

        public static ButtonColorScheme BtnExportPdf => _isDark
            ? new ButtonColorScheme(Color.FromArgb(58, 58, 63), Color.FromArgb(232, 232, 236), Color.FromArgb(86, 86, 92))
            : new ButtonColorScheme(Color.FromArgb(196, 68, 68), Color.FromArgb(255, 255, 255), Color.FromArgb(172, 54, 54));

        public static ButtonColorScheme BtnIconNeutral => _isDark
            ? new ButtonColorScheme(Color.FromArgb(42, 42, 46), Color.FromArgb(224, 224, 230), Color.FromArgb(66, 66, 72))
            : new ButtonColorScheme(Color.FromArgb(255, 255, 255), Color.FromArgb(58, 58, 64), Color.FromArgb(214, 214, 222));

        /// <summary>
        /// Creates a closed GDI+ <see cref="GraphicsPath"/> representing a rectangle with smooth rounded corners.
        /// </summary>
        /// <param name="bounds">Bounding rectangle coordinates.</param>
        /// <param name="radius">Corner radius in pixels.</param>
        /// <returns>A <see cref="GraphicsPath"/> ready for filling or stroking.</returns>
        public static GraphicsPath CreateRoundedRectPath(RectangleF bounds, float radius)
        {
            var path = new GraphicsPath();
            float d = radius * 2f;
            if (d <= 0f)
            {
                path.AddRectangle(bounds);
                return path;
            }
            if (d > bounds.Width) d = bounds.Width;
            if (d > bounds.Height) d = bounds.Height;

            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    /// <summary>
    /// Custom owner-drawn button control featuring anti-aliased rounded corners,
    /// hover and press brightness transitions, word-wrapped centered text, and optional
    /// custom vector icon rendering via <see cref="CustomIconPainter"/>.
    /// </summary>
    public class RoundedButton : Button
    {
        private bool _isHovered;
        private bool _isPressed;

        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public int CornerRadius { get; set; } = 10;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color CustomBackColor { get; set; } = Color.White;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color CustomForeColor { get; set; } = Color.Black;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color CustomBorderColor { get; set; } = Color.LightGray;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Action<Graphics, Rectangle, Color>? CustomIconPainter { get; set; }

        /// <summary>
        /// Initializes a new double-buffered, owner-drawn <see cref="RoundedButton"/>.
        /// </summary>
        public RoundedButton()
        {
            SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw,
                true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
        }

        /// <summary>
        /// Applies a <see cref="ButtonColorScheme"/> (fill, foreground, and border colors) and repaints the button.
        /// </summary>
        public void ApplyScheme(ButtonColorScheme scheme)
        {
            CustomBackColor = scheme.Back;
            CustomForeColor = scheme.Fore;
            CustomBorderColor = scheme.Border;
            Invalidate();
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _isHovered = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _isHovered = false;
            _isPressed = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            if (mevent.Button == MouseButtons.Left)
            {
                _isPressed = true;
                Invalidate();
            }
            base.OnMouseDown(mevent);
        }

        protected override void OnMouseUp(MouseEventArgs mevent)
        {
            _isPressed = false;
            Invalidate();
            base.OnMouseUp(mevent);
        }

        /// <summary>
        /// Renders the anti-aliased rounded button background, border, and either custom vector icon or text.
        /// </summary>
        protected override void OnPaint(PaintEventArgs pevent)
        {
            var g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            Color parentBack = Parent?.BackColor ?? ThemeManager.WindowBack;
            g.Clear(parentBack);

            Color fill = CustomBackColor;
            if (_isPressed)
            {
                fill = AdjustBrightness(fill, ThemeManager.IsDark ? 1.18f : 0.90f);
            }
            else if (_isHovered)
            {
                fill = AdjustBrightness(fill, ThemeManager.IsDark ? 1.10f : 0.95f);
            }

            var rect = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            using (var path = ThemeManager.CreateRoundedRectPath(rect, CornerRadius))
            using (var brush = new SolidBrush(fill))
            using (var pen = new Pen(CustomBorderColor, 1f))
            {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }

            if (CustomIconPainter != null)
            {
                CustomIconPainter(g, ClientRectangle, CustomForeColor);
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                TextRenderer.DrawText(
                    g,
                    Text,
                    Font,
                    ClientRectangle,
                    CustomForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
            }
        }

        /// <summary>
        /// Multiplies RGB channels by a brightness factor to produce subtle hover and active press states.
        /// </summary>
        private static Color AdjustBrightness(Color c, float factor)
        {
            int r = Math.Clamp((int)(c.R * factor), 0, 255);
            int g = Math.Clamp((int)(c.G * factor), 0, 255);
            int b = Math.Clamp((int)(c.B * factor), 0, 255);
            return Color.FromArgb(c.A, r, g, b);
        }
    }

    /// <summary>
    /// Custom owner-drawn container panel that renders an anti-aliased rounded card surface
    /// with configurable background fill and border colors.
    /// </summary>
    public class RoundedCardPanel : Panel
    {
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public int CornerRadius { get; set; } = 12;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color CardFillColor { get; set; } = Color.White;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color CardBorderColor { get; set; } = Color.LightGray;

        /// <summary>
        /// Initializes a new double-buffered <see cref="RoundedCardPanel"/>.
        /// </summary>
        public RoundedCardPanel()
        {
            SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw,
                true);
        }

        /// <summary>
        /// Paints the rounded card surface and border over the parent background color.
        /// </summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Color parentBack = Parent?.BackColor ?? ThemeManager.WindowBack;
            g.Clear(parentBack);

            var rect = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            using (var path = ThemeManager.CreateRoundedRectPath(rect, CornerRadius))
            using (var brush = new SolidBrush(CardFillColor))
            using (var pen = new Pen(CardBorderColor, 1f))
            {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }
        }
    }

    /// <summary>
    /// Custom owner-drawn rounded border container hosting a borderless <see cref="TextBox"/>.
    /// Highlights its border when focused and vertically centers the inner input control.
    /// </summary>
    public class RoundedInputHost : Panel
    {
        private bool _focused;
        public TextBox InnerTextBox { get; }
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public int CornerRadius { get; set; } = 8;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color FillColor { get; set; } = Color.White;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color BorderColor { get; set; } = Color.LightGray;
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color FocusBorderColor { get; set; } = Color.DodgerBlue;

        /// <summary>
        /// Initializes a new <see cref="RoundedInputHost"/> wrapping the specified <see cref="TextBox"/>.
        /// </summary>
        public RoundedInputHost(TextBox textBox)
        {
            InnerTextBox = textBox;
            InnerTextBox.BorderStyle = BorderStyle.None;

            SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw,
                true);

            Controls.Add(InnerTextBox);

            InnerTextBox.GotFocus += (s, e) => { _focused = true; Invalidate(); };
            InnerTextBox.LostFocus += (s, e) => { _focused = false; Invalidate(); };
            Click += (s, e) => InnerTextBox.Focus();
            Resize += (s, e) => CenterInnerTextBox();
        }

        /// <summary>
        /// Synchronizes fill, text, idle border, and focus border colors across the host and inner TextBox.
        /// </summary>
        public void SyncColors(Color fill, Color fore, Color border, Color focusBorder)
        {
            FillColor = fill;
            BorderColor = border;
            FocusBorderColor = focusBorder;
            InnerTextBox.BackColor = fill;
            InnerTextBox.ForeColor = fore;
            Invalidate();
        }

        /// <summary>
        /// Horizontally pads and vertically centers the inner <see cref="TextBox"/> within the rounded host.
        /// </summary>
        public void CenterInnerTextBox()
        {
            int padX = 10;
            int w = Math.Max(20, Width - padX * 2);
            int y = Math.Max(2, (Height - InnerTextBox.Height) / 2);
            InnerTextBox.Location = new Point(padX, y);
            InnerTextBox.Width = w;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Color parentBack = Parent?.BackColor ?? ThemeManager.CardBack;
            g.Clear(parentBack);

            var rect = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            Color activeBorder = _focused ? FocusBorderColor : BorderColor;
            float borderWidth = _focused ? 1.4f : 1f;

            using (var path = ThemeManager.CreateRoundedRectPath(rect, CornerRadius))
            using (var brush = new SolidBrush(FillColor))
            using (var pen = new Pen(activeBorder, borderWidth))
            {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }
        }
    }
}

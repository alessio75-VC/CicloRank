// ============================================================================
// File: MainForm.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Main application window and race timing controller. Manages the high-precision
//   stopwatch, rapid bib number entry with live rider preview, editable recent-lap
//   cards, interactive lap-by-lap timing grid, live category-filtered standings,
//   automatic crash-recovery persistence, and Excel/PDF export workflows.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace CicloRank
{
    /// <summary>
    /// Primary user interface and race controller for CicloRank.
    /// Coordinates the live race stopwatch, bib entry with real-time rider lookup,
    /// editable last-3-passages cards, editable lap-by-lap timing grid, live category-filtered
    /// standings table, automatic crash-recovery persistence, and Excel/PDF exports.
    /// </summary>
    public partial class MainForm : Form
    {
        private static string FileSessioneTemp => Path.Combine(ThemeManager.CartellaDatiUtente, "sessione_temp.json");
        private static readonly Font FontCellaGriglia = new Font("Segoe UI", 9.5F, FontStyle.Regular);
        private static readonly Font FontHeaderGriglia = new Font("Segoe UI", 9.5F, FontStyle.Bold);
        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

        private static string _sportCorrente = "";
        /// <summary>
        /// Gets or sets the current race sport/discipline name shared with report title builders.
        /// </summary>
        public static string SportCorrente
        {
            get => _sportCorrente;
            set => _sportCorrente = (value ?? "").Trim();
        }

        private List<Ciclista> listaCiclisti = new List<Ciclista>();
        private Dictionary<int, Ciclista> mappaCiclisti = new Dictionary<int, Ciclista>();
        private List<Passaggio> listaPassaggi = new List<Passaggio>();
        private DatiGara datiGara = new DatiGara();

        private System.Windows.Forms.Timer timerGara;
        private TimeSpan tempoTrascorso = TimeSpan.Zero;
        private DateTime orarioInizioGara;
        private bool timerInEsecuzione = false;

        private RoundedCardPanel cardInput;
        private Label lblHeaderInput;
        private RoundedInputHost hostInserisciNumero;
        private TextBox txtInserisciNumero;
        private Label lblNomeAnteprima;

        private RoundedCardPanel cardUltimo;
        private RoundedInputHost hostUltimo;
        private TextBox txtUltimo;
        private Label lblAnteprimaUltimo;

        private RoundedCardPanel cardPenultimo;
        private RoundedInputHost hostPenultimo;
        private TextBox txtPenultimo;
        private Label lblAnteprimaPenultimo;

        private RoundedCardPanel cardTerzultimo;
        private RoundedInputHost hostTerzultimo;
        private TextBox txtTerzultimo;
        private Label lblAnteprimaTerzultimo;

        private RoundedButton btnDatiGara;
        private RoundedButton btnCaricaDati;
        private RoundedButton btnAzzeraDati;

        private RoundedCardPanel cardGrigliaLeft;
        private RoundedCardPanel cardGrigliaRight;

        private RoundedButton btnEsportaSessione;
        private RoundedButton btnImportaSessione;

        private Label lblFiltroCategoria;
        private ComboBox cmbFiltroCategoria;
        private Label lblLingua;
        private ComboBox cmbLingua;
        private RoundedButton btnToggleTema;

        private DataGridView dgvCronometraggio;
        private DataGridView dgvClassifiche;

        private Label lblCronometro;
        private RoundedButton btnStart;
        private RoundedButton btnPausa;
        private RoundedButton btnAzzeraTimer;

        private RoundedButton btnKoFi;
        private RoundedButton btnEsportaExcel;
        private RoundedButton btnEsportaPdf;

        private ToolTip toolTipGlobale;

        /// <summary>
        /// Initializes the main application window, builds the responsive UI layout,
        /// applies saved language and theme preferences, configures the 200ms race timer,
        /// and checks for any unsaved crash-recovery session on startup.
        /// </summary>
        public MainForm()
        {
            ThemeManager.AggiornaPreferenzaProcesso();
            toolTipGlobale = new ToolTip();
            InitializeUI();
            ApplicaLingua();
            ApplicaTema();

            timerGara = new System.Windows.Forms.Timer();
            timerGara.Interval = 200;
            timerGara.Tick += TimerGara_Tick;

            if (File.Exists("icona.ico"))
            {
                try { Icon = new Icon("icona.ico"); } catch { }
            }
            else
            {
                try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            }

            RipristinaSessioneSeEsistente();
        }

        /// <summary>
        /// Programmatically constructs the entire responsive TableLayoutPanel UI hierarchy,
        /// including top bib-entry &amp; quick-edit cards, dual DataGridViews, and bottom stopwatch/export bars.
        /// </summary>
        private void InitializeUI()
        {
            Size = new Size(1420, 840);
            MinimumSize = new Size(1120, 680);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1,
                Padding = new Padding(14, 12, 14, 12)
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 114F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 62F));

            var topPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 7,
                RowCount = 1,
                Margin = new Padding(0, 0, 0, 10)
            };
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 126F));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 126F));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 126F));

            txtInserisciNumero = new TextBox
            {
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center
            };
            txtInserisciNumero.TextChanged += TxtInserisciNumero_TextChanged;
            txtInserisciNumero.KeyDown += TxtInserisciNumero_KeyDown;

            lblNomeAnteprima = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 20,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
                AutoEllipsis = true,
                Text = ""
            };

            cardInput = CreaInputCard(out lblHeaderInput, txtInserisciNumero, out hostInserisciNumero, lblNomeAnteprima, new Padding(0, 0, 8, 0));

            txtUltimo = new TextBox
            {
                Font = new Font("Segoe UI", 15F, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center
            };
            txtUltimo.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AggiornaPassaggioDaTextBox(1, txtUltimo.Text); }
                else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; RipristinaTextBoxPassaggio(1); }
            };
            txtUltimo.Leave += (s, e) => RipristinaTextBoxPassaggio(1);
            cardUltimo = CreaSchedaRiepilogoRapido(txtUltimo, out hostUltimo, out lblAnteprimaUltimo, new Padding(0, 0, 8, 0));

            txtPenultimo = new TextBox
            {
                Font = new Font("Segoe UI", 15F, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center
            };
            txtPenultimo.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AggiornaPassaggioDaTextBox(2, txtPenultimo.Text); }
                else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; RipristinaTextBoxPassaggio(2); }
            };
            txtPenultimo.Leave += (s, e) => RipristinaTextBoxPassaggio(2);
            cardPenultimo = CreaSchedaRiepilogoRapido(txtPenultimo, out hostPenultimo, out lblAnteprimaPenultimo, new Padding(0, 0, 8, 0));

            txtTerzultimo = new TextBox
            {
                Font = new Font("Segoe UI", 15F, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center
            };
            txtTerzultimo.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AggiornaPassaggioDaTextBox(3, txtTerzultimo.Text); }
                else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; RipristinaTextBoxPassaggio(3); }
            };
            txtTerzultimo.Leave += (s, e) => RipristinaTextBoxPassaggio(3);
            cardTerzultimo = CreaSchedaRiepilogoRapido(txtTerzultimo, out hostTerzultimo, out lblAnteprimaTerzultimo, new Padding(0, 0, 10, 0));

            btnDatiGara = new RoundedButton
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 8, 0),
                CornerRadius = 11,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold)
            };
            btnDatiGara.Click += BtnDatiGara_Click;

            btnCaricaDati = new RoundedButton
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 8, 0),
                CornerRadius = 11,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold)
            };
            btnCaricaDati.Click += BtnCaricaDati_Click;

            btnAzzeraDati = new RoundedButton
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                CornerRadius = 11,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold)
            };
            btnAzzeraDati.Click += BtnAzzeraDati_Click;

            topPanel.Controls.Add(cardInput, 0, 0);
            topPanel.Controls.Add(cardUltimo, 1, 0);
            topPanel.Controls.Add(cardPenultimo, 2, 0);
            topPanel.Controls.Add(cardTerzultimo, 3, 0);
            topPanel.Controls.Add(btnDatiGara, 4, 0);
            topPanel.Controls.Add(btnCaricaDati, 5, 0);
            topPanel.Controls.Add(btnAzzeraDati, 6, 0);

            var centerPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 0, 0, 10)
            };
            centerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55F));
            centerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));

            cardGrigliaLeft = new RoundedCardPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 10, 0),
                Padding = new Padding(10, 10, 10, 8),
                CornerRadius = 12
            };

            var leftGridLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            leftGridLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            leftGridLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

            dgvCronometraggio = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ScrollBars = ScrollBars.Both,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                ReadOnly = false,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                SelectionMode = DataGridViewSelectionMode.CellSelect
            };
            dgvCronometraggio.RowTemplate.Height = 30;
            AbilitaDoubleBuffering(dgvCronometraggio);
            CollegaAngoloScrollbarGriglia(dgvCronometraggio);

            for (int i = 1; i <= 8; i++)
            {
                dgvCronometraggio.Columns.Add($"Giro{i}", LinguaManager.FormatGiroHeader(i));
            }
            dgvCronometraggio.Resize += (s, e) => AggiornaLarghezzaColonneCronometraggio();
            dgvCronometraggio.CellBeginEdit += DgvCronometraggio_CellBeginEdit;
            dgvCronometraggio.CellEndEdit += DgvCronometraggio_CellEndEdit;

            var pnlSubGrid = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 4, 0, 0),
                Margin = new Padding(0)
            };

            btnEsportaSessione = new RoundedButton
            {
                Size = new Size(30, 30),
                CornerRadius = 8,
                Margin = new Padding(6, 0, 0, 0),
                CustomIconPainter = DisegnaIconaEsportaSessione
            };
            btnEsportaSessione.Click += BtnEsportaSessione_Click;

            btnImportaSessione = new RoundedButton
            {
                Size = new Size(30, 30),
                CornerRadius = 8,
                Margin = new Padding(0),
                CustomIconPainter = DisegnaIconaImportaSessione
            };
            btnImportaSessione.Click += BtnImportaSessione_Click;

            pnlSubGrid.Controls.Add(btnEsportaSessione);
            pnlSubGrid.Controls.Add(btnImportaSessione);

            leftGridLayout.Controls.Add(dgvCronometraggio, 0, 0);
            leftGridLayout.Controls.Add(pnlSubGrid, 0, 1);
            cardGrigliaLeft.Controls.Add(leftGridLayout);

            cardGrigliaRight = new RoundedCardPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(10, 8, 10, 10),
                CornerRadius = 12
            };

            var rightPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            rightPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            rightPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var topRightBar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 0, 0, 4)
            };
            topRightBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            topRightBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            var filterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0)
            };
            lblFiltroCategoria = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Padding = new Padding(0, 6, 6, 0)
            };
            cmbFiltroCategoria = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Width = 148,
                Font = new Font("Segoe UI", 9.5F),
                Margin = new Padding(0, 2, 0, 0)
            };
            cmbFiltroCategoria.Items.Add(LinguaManager.T("FiltroTutte"));
            cmbFiltroCategoria.SelectedIndex = 0;
            cmbFiltroCategoria.SelectedIndexChanged += (s, e) => AggiornaGrigliaClassifiche();
            filterPanel.Controls.Add(lblFiltroCategoria);
            filterPanel.Controls.Add(cmbFiltroCategoria);

            var langAndThemePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Margin = new Padding(0)
            };

            btnToggleTema = new RoundedButton
            {
                Size = new Size(30, 30),
                CornerRadius = 8,
                Margin = new Padding(6, 1, 0, 0),
                CustomIconPainter = DisegnaIconaTema
            };
            btnToggleTema.Click += (s, e) =>
            {
                ThemeManager.ToggleTema();
                ApplicaTema();
            };

            cmbLingua = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Width = 126,
                Font = new Font("Segoe UI", 9.5F),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 0, 0)
            };
            foreach (var opz in LinguaManager.LingueDisponibili)
            {
                cmbLingua.Items.Add(opz);
                if (opz.Codice.Equals(LinguaManager.LinguaCorrente, StringComparison.OrdinalIgnoreCase))
                {
                    cmbLingua.SelectedItem = opz;
                }
            }
            if (cmbLingua.SelectedIndex < 0 && cmbLingua.Items.Count > 0)
            {
                cmbLingua.SelectedIndex = 0;
            }
            cmbLingua.SelectedIndexChanged += CmbLingua_SelectedIndexChanged;

            lblLingua = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Padding = new Padding(0, 6, 6, 0)
            };

            langAndThemePanel.Controls.Add(btnToggleTema);
            langAndThemePanel.Controls.Add(cmbLingua);
            langAndThemePanel.Controls.Add(lblLingua);

            topRightBar.Controls.Add(filterPanel, 0, 0);
            topRightBar.Controls.Add(langAndThemePanel, 1, 0);

            dgvClassifiche = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            dgvClassifiche.RowTemplate.Height = 30;
            AbilitaDoubleBuffering(dgvClassifiche);
            CollegaAngoloScrollbarGriglia(dgvClassifiche);

            dgvClassifiche.Columns.Add("Pos", "Pos");
            dgvClassifiche.Columns.Add("Pett", "Pett.");
            dgvClassifiche.Columns.Add("Nome", "Nome");
            dgvClassifiche.Columns.Add("Cat", "Cat.");
            dgvClassifiche.Columns.Add("Giri", "Giri");
            dgvClassifiche.Columns.Add("Tempo", "Tempo");

            dgvClassifiche.Columns[0].FillWeight = 12;
            dgvClassifiche.Columns[1].FillWeight = 14;
            dgvClassifiche.Columns[2].FillWeight = 34;
            dgvClassifiche.Columns[3].FillWeight = 14;
            dgvClassifiche.Columns[4].FillWeight = 10;
            dgvClassifiche.Columns[5].FillWeight = 16;

            rightPanel.Controls.Add(topRightBar, 0, 0);
            rightPanel.Controls.Add(dgvClassifiche, 0, 1);
            cardGrigliaRight.Controls.Add(rightPanel);

            centerPanel.Controls.Add(cardGrigliaLeft, 0, 0);
            centerPanel.Controls.Add(cardGrigliaRight, 1, 0);

            var bottomPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0)
            };
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55F));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));

            var timerPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(2, 4, 0, 0)
            };
            lblCronometro = new Label
            {
                Text = "00:00:00",
                Font = new Font("Segoe UI", 24F, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 0, 18, 0)
            };

            btnStart = new RoundedButton
            {
                Width = 96,
                Height = 42,
                CornerRadius = 10,
                Margin = new Padding(0, 2, 8, 0),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
            };
            btnStart.Click += BtnStart_Click;

            btnPausa = new RoundedButton
            {
                Width = 96,
                Height = 42,
                CornerRadius = 10,
                Margin = new Padding(0, 2, 8, 0),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
            };
            btnPausa.Click += BtnPausa_Click;

            btnAzzeraTimer = new RoundedButton
            {
                Width = 96,
                Height = 42,
                CornerRadius = 10,
                Margin = new Padding(0, 2, 0, 0),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Visible = false
            };
            btnAzzeraTimer.Click += BtnAzzeraTimer_Click;

            timerPanel.Controls.Add(lblCronometro);
            timerPanel.Controls.Add(btnStart);
            timerPanel.Controls.Add(btnPausa);
            timerPanel.Controls.Add(btnAzzeraTimer);

            var exportPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 6, 0, 0)
            };

            btnEsportaPdf = new RoundedButton
            {
                Width = 142,
                Height = 42,
                CornerRadius = 10,
                Margin = new Padding(8, 0, 0, 0),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            btnEsportaPdf.Click += BtnEsportaPdf_Click;

            btnEsportaExcel = new RoundedButton
            {
                Width = 142,
                Height = 42,
                CornerRadius = 10,
                Margin = new Padding(0),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            btnEsportaExcel.Click += BtnEsportaExcel_Click;

            btnKoFi = new RoundedButton
            {
                Width = 42,
                Height = 42,
                CornerRadius = 10,
                Margin = new Padding(0, 0, 8, 0),
                CustomIconPainter = DisegnaIconaCuore
            };
            btnKoFi.Click += (s, e) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "https://ko-fi.com/alessio75vc",
                        UseShellExecute = true
                    });
                }
                catch { }
            };

            exportPanel.Controls.Add(btnEsportaPdf);
            exportPanel.Controls.Add(btnEsportaExcel);
            exportPanel.Controls.Add(btnKoFi);

            bottomPanel.Controls.Add(timerPanel, 0, 0);
            bottomPanel.Controls.Add(exportPanel, 1, 0);

            mainLayout.Controls.Add(topPanel, 0, 0);
            mainLayout.Controls.Add(centerPanel, 0, 1);
            mainLayout.Controls.Add(bottomPanel, 0, 2);

            Controls.Add(mainLayout);
            AggiornaLarghezzaColonneCronometraggio();
        }

        /// <summary>
        /// Factory helper that builds the primary bib-entry card with a header label,
        /// rounded text input host, and live rider preview footer label.
        /// </summary>
        private RoundedCardPanel CreaInputCard(
            out Label headerLabel,
            TextBox textBox,
            out RoundedInputHost inputHost,
            Label? footerLabel,
            Padding margin)
        {
            var card = new RoundedCardPanel
            {
                Dock = DockStyle.Fill,
                Margin = margin,
                Padding = new Padding(12, 8, 12, 6),
                CornerRadius = 12
            };

            headerLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 20,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold)
            };

            inputHost = new RoundedInputHost(textBox)
            {
                Dock = DockStyle.Top,
                Height = 40,
                CornerRadius = 8
            };

            var bottomSpacer = footerLabel ?? new Label
            {
                Dock = DockStyle.Bottom,
                Height = 20,
                Text = ""
            };

            card.Controls.Add(inputHost);
            card.Controls.Add(headerLabel);
            card.Controls.Add(bottomSpacer);

            return card;
        }

        /// <summary>
        /// Factory helper that builds one of the three quick-edit cards displaying the
        /// last, second-last, or third-last recorded bib numbers.
        /// </summary>
        private RoundedCardPanel CreaSchedaRiepilogoRapido(
            TextBox textBox,
            out RoundedInputHost inputHost,
            out Label previewLabel,
            Padding margin)
        {
            var card = new RoundedCardPanel
            {
                Dock = DockStyle.Fill,
                Margin = margin,
                Padding = new Padding(12, 14, 12, 6),
                CornerRadius = 12
            };

            inputHost = new RoundedInputHost(textBox)
            {
                Dock = DockStyle.Top,
                Height = 46,
                CornerRadius = 8
            };

            previewLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 22,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
                AutoEllipsis = true,
                Text = ""
            };

            card.Controls.Add(inputHost);
            card.Controls.Add(previewLabel);

            return card;
        }

        /// <summary>
        /// Vector icon painter that draws a sun icon in Dark Mode or a crescent moon icon in Light Mode.
        /// </summary>
        private static void DisegnaIconaTema(Graphics g, Rectangle bounds, Color color)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = bounds.Width / 2f;
            float cy = bounds.Height / 2f;

            if (ThemeManager.IsDark)
            {
                using var pen = new Pen(color, 1.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                using var brush = new SolidBrush(color);
                float rCore = 4.2f;
                g.FillEllipse(brush, cx - rCore, cy - rCore, rCore * 2, rCore * 2);

                for (int i = 0; i < 8; i++)
                {
                    double angle = i * Math.PI / 4.0;
                    float x1 = cx + (float)(Math.Cos(angle) * 6.5);
                    float y1 = cy + (float)(Math.Sin(angle) * 6.5);
                    float x2 = cx + (float)(Math.Cos(angle) * 9.0);
                    float y2 = cy + (float)(Math.Sin(angle) * 9.0);
                    g.DrawLine(pen, x1, y1, x2, y2);
                }
            }
            else
            {
                using var path = new GraphicsPath();
                path.AddArc(cx - 7.5f, cy - 7.5f, 15f, 15f, 40f, 260f);
                path.AddArc(cx - 3.5f, cy - 8.5f, 12.5f, 12.5f, 290f, -210f);
                path.CloseFigure();

                using var brush = new SolidBrush(color);
                g.FillPath(brush, path);
            }
        }

        /// <summary>
        /// Vector icon painter that draws a downward tray-export arrow for saving race sessions (.json).
        /// </summary>
        private static void DisegnaIconaEsportaSessione(Graphics g, Rectangle bounds, Color color)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var pen = new Pen(color, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
            g.DrawLines(pen, new PointF[] { new PointF(7, 18), new PointF(7, 22), new PointF(22, 22), new PointF(22, 18) });
            g.DrawLine(pen, 14.5f, 6, 14.5f, 17);
            g.DrawLine(pen, 10.5f, 13, 14.5f, 17);
            g.DrawLine(pen, 18.5f, 13, 14.5f, 17);
        }

        /// <summary>
        /// Vector icon painter that draws an upward tray-import arrow for loading race sessions (.json).
        /// </summary>
        private static void DisegnaIconaImportaSessione(Graphics g, Rectangle bounds, Color color)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var pen = new Pen(color, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
            g.DrawLines(pen, new PointF[] { new PointF(7, 18), new PointF(7, 22), new PointF(22, 22), new PointF(22, 18) });
            g.DrawLine(pen, 14.5f, 18, 14.5f, 7);
            g.DrawLine(pen, 10.5f, 11, 14.5f, 7);
            g.DrawLine(pen, 18.5f, 11, 14.5f, 7);
        }

        /// <summary>
        /// Vector icon painter that draws a stylized glossy heart icon for the Ko-fi support button.
        /// </summary>
        private static void DisegnaIconaCuore(Graphics g, Rectangle bounds, Color color)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = bounds.Width / 2f;
            float cy = bounds.Height / 2f + 0.5f;

            float topDipY = cy - 4.0f;
            float topPeakY = cy - 9.4f;
            float sideX = 10.4f;
            float sideY = cy - 3.0f;
            float bottomY = cy + 8.8f;

            using var path = new GraphicsPath();
            path.AddBezier(new PointF(cx, topDipY), new PointF(cx - 2.3f, topPeakY), new PointF(cx - 8.2f, topPeakY), new PointF(cx - sideX, sideY));
            path.AddBezier(new PointF(cx - sideX, sideY), new PointF(cx - 11.4f, cy + 2.4f), new PointF(cx - 4.0f, cy + 6.5f), new PointF(cx, bottomY));
            path.AddBezier(new PointF(cx, bottomY), new PointF(cx + 4.0f, cy + 6.5f), new PointF(cx + 11.4f, cy + 2.4f), new PointF(cx + sideX, sideY));
            path.AddBezier(new PointF(cx + sideX, sideY), new PointF(cx + 8.2f, topPeakY), new PointF(cx + 2.3f, topPeakY), new PointF(cx, topDipY));
            path.CloseFigure();

            using (var brush = new SolidBrush(color))
            {
                g.FillPath(brush, path);
            }

            using (var borderPen = new Pen(Color.FromArgb(214, 72, 0), 1.1f) { LineJoin = LineJoin.Round })
            {
                g.DrawPath(borderPen, path);
            }

            using (var shinePen = new Pen(Color.FromArgb(185, 255, 255, 255), 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawArc(shinePen, cx - 7.3f, cy - 6.3f, 4.6f, 4.4f, 195f, 75f);
            }

            using (var dotBrush = new SolidBrush(Color.FromArgb(185, 255, 255, 255)))
            {
                g.FillEllipse(dotBrush, cx - 7.6f, cy - 1.1f, 1.8f, 1.8f);
            }
        }

        /// <summary>
        /// Applies the active <see cref="ThemeManager"/> color scheme across the window, cards,
        /// inputs, buttons, combo boxes, and both data grids, then refreshes all views.
        /// </summary>
        private void ApplicaTema()
        {
            ThemeManager.AggiornaPreferenzaProcesso();
            ThemeManager.ApplicaTemaFinestra(this);
            BackColor = ThemeManager.WindowBack;
            ForeColor = ThemeManager.TextPrimary;

            foreach (var card in new[] { cardInput, cardUltimo, cardPenultimo, cardTerzultimo, cardGrigliaLeft, cardGrigliaRight })
            {
                card.CardFillColor = ThemeManager.CardBack;
                card.CardBorderColor = ThemeManager.CardBorder;
                card.Invalidate();
            }

            lblHeaderInput.ForeColor = ThemeManager.TextAccent;

            hostInserisciNumero.SyncColors(ThemeManager.InputBack, ThemeManager.TextPrimary, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
            hostUltimo.SyncColors(ThemeManager.InputBack, ThemeManager.SuccessText, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
            hostPenultimo.SyncColors(ThemeManager.InputBack, ThemeManager.SuccessText, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
            hostTerzultimo.SyncColors(ThemeManager.InputBack, ThemeManager.SuccessText, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);

            btnDatiGara.ApplyScheme(ThemeManager.BtnDatiGara);
            btnCaricaDati.ApplyScheme(ThemeManager.BtnCaricaDati);
            btnAzzeraDati.ApplyScheme(ThemeManager.BtnAzzeraDati);

            btnStart.ApplyScheme(ThemeManager.BtnStart);
            btnPausa.ApplyScheme(ThemeManager.BtnPausa);
            btnAzzeraTimer.ApplyScheme(ThemeManager.BtnAzzeraTimer);

            btnKoFi.ApplyScheme(new ButtonColorScheme(
                Color.FromArgb(237, 230, 221), // #ede6dd
                Color.FromArgb(250, 91, 0),    // #fa5b00
                Color.FromArgb(216, 206, 194)
            ));
            btnEsportaExcel.ApplyScheme(ThemeManager.BtnExportExcel);
            btnEsportaPdf.ApplyScheme(ThemeManager.BtnExportPdf);

            btnEsportaSessione.ApplyScheme(ThemeManager.BtnIconNeutral);
            btnImportaSessione.ApplyScheme(ThemeManager.BtnIconNeutral);
            btnToggleTema.ApplyScheme(ThemeManager.BtnIconNeutral);

            lblFiltroCategoria.ForeColor = ThemeManager.TextPrimary;
            lblLingua.ForeColor = ThemeManager.TextPrimary;

            cmbFiltroCategoria.BackColor = ThemeManager.InputBack;
            cmbFiltroCategoria.ForeColor = ThemeManager.TextPrimary;
            ThemeManager.ApplicaTemaControlloNativo(cmbFiltroCategoria);
            cmbLingua.BackColor = ThemeManager.InputBack;
            cmbLingua.ForeColor = ThemeManager.TextPrimary;
            ThemeManager.ApplicaTemaControlloNativo(cmbLingua);

            AggiornaColoreCronometro();

            ApplicaTemaGriglia(dgvCronometraggio);
            ApplicaTemaGriglia(dgvClassifiche);

            toolTipGlobale.SetToolTip(
                btnToggleTema,
                ThemeManager.IsDark ? LinguaManager.T("TipTemaChiaro") : LinguaManager.T("TipTemaScuro"));

            AggiornaSchedaRiepilogo(hostUltimo, txtUltimo, lblAnteprimaUltimo);
            AggiornaSchedaRiepilogo(hostPenultimo, txtPenultimo, lblAnteprimaPenultimo);
            AggiornaSchedaRiepilogo(hostTerzultimo, txtTerzultimo, lblAnteprimaTerzultimo);
            TxtInserisciNumero_TextChanged(txtInserisciNumero, EventArgs.Empty);
            AggiornaGrigliaCronometraggio();
            AggiornaGrigliaClassifiche();
        }

        /// <summary>
        /// Configures colors, fonts, selection highlights, and native scrollbar theming on a <see cref="DataGridView"/>.
        /// </summary>
        private static void ApplicaTemaGriglia(DataGridView dgv)
        {
            dgv.BackgroundColor = ThemeManager.GridBack;
            dgv.GridColor = ThemeManager.GridLine;

            dgv.DefaultCellStyle.BackColor = ThemeManager.GridBack;
            dgv.DefaultCellStyle.ForeColor = ThemeManager.TextPrimary;
            dgv.DefaultCellStyle.SelectionBackColor = ThemeManager.GridSelectionBack;
            dgv.DefaultCellStyle.SelectionForeColor = ThemeManager.GridSelectionText;
            dgv.DefaultCellStyle.Font = FontCellaGriglia;

            dgv.AlternatingRowsDefaultCellStyle.BackColor = ThemeManager.GridAltBack;
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = ThemeManager.TextPrimary;
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = ThemeManager.GridSelectionBack;
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = ThemeManager.GridSelectionText;

            dgv.ColumnHeadersDefaultCellStyle.BackColor = ThemeManager.GridHeaderBack;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = ThemeManager.GridHeaderText;
            dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = ThemeManager.GridHeaderBack;
            dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = ThemeManager.GridHeaderText;
            dgv.ColumnHeadersDefaultCellStyle.Font = FontHeaderGriglia;
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ThemeManager.ApplicaTemaControlloNativo(dgv);
        }

        /// <summary>
        /// Paints the bottom-right intersection square when both horizontal and vertical scrollbars
        /// are visible on a <see cref="DataGridView"/> so it blends seamlessly with the grid header.
        /// </summary>
        private static void CollegaAngoloScrollbarGriglia(DataGridView dgv)
        {
            dgv.Paint += (s, e) =>
            {
                HScrollBar? hBar = null;
                VScrollBar? vBar = null;
                foreach (Control c in dgv.Controls)
                {
                    if (c is HScrollBar h) hBar = h;
                    else if (c is VScrollBar v) vBar = v;
                }

                if (hBar != null && vBar != null && hBar.Visible && vBar.Visible)
                {
                    var corner = new Rectangle(vBar.Left, hBar.Top, vBar.Width, hBar.Height);
                    using var br = new SolidBrush(ThemeManager.GridHeaderBack);
                    e.Graphics.FillRectangle(br, corner);
                }
            };
        }

        /// <summary>
        /// Updates the stopwatch label foreground color to match the active theme.
        /// </summary>
        private void AggiornaColoreCronometro()
        {
            lblCronometro.ForeColor = ThemeManager.TextPrimary;
        }

        /// <summary>
        /// Enables protected <c>DoubleBuffered</c> rendering on a <see cref="DataGridView"/> via reflection
        /// to eliminate flicker during rapid lap updates.
        /// </summary>
        private static void AbilitaDoubleBuffering(DataGridView dgv)
        {
            try
            {
                typeof(DataGridView)
                    .GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?
                    .SetValue(dgv, true, null);
            }
            catch { }
        }

        /// <summary>
        /// Handles user selection in the language dropdown and refreshes all UI strings.
        /// </summary>
        private void CmbLingua_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbLingua.SelectedItem is OpzioneLingua selezionata)
            {
                LinguaManager.ImpostaLingua(selezionata.Codice);
                ApplicaLingua();
            }
        }

        /// <summary>
        /// Updates all window titles, button captions, tooltips, grid column headers,
        /// and status previews using <see cref="LinguaManager"/> translations.
        /// </summary>
        private void ApplicaLingua()
        {
            Text = LinguaManager.T("FormTitle");

            lblHeaderInput.Text = LinguaManager.T("GrpInserisciNumero");

            btnDatiGara.Text = LinguaManager.T("BtnDatiGara");
            btnCaricaDati.Text = LinguaManager.T("BtnCaricaDati");
            btnAzzeraDati.Text = LinguaManager.T("BtnAzzeraDati");

            toolTipGlobale.SetToolTip(btnEsportaSessione, LinguaManager.T("TipEsportaSessione"));
            toolTipGlobale.SetToolTip(btnImportaSessione, LinguaManager.T("TipImportaSessione"));
            toolTipGlobale.SetToolTip(
                btnToggleTema,
                ThemeManager.IsDark ? LinguaManager.T("TipTemaChiaro") : LinguaManager.T("TipTemaScuro"));

            lblFiltroCategoria.Text = LinguaManager.T("LblFiltroCategoria");
            lblLingua.Text = LinguaManager.T("LblLingua");

            if (cmbFiltroCategoria.Items.Count > 0)
            {
                int idxCorrente = cmbFiltroCategoria.SelectedIndex;
                cmbFiltroCategoria.Items[0] = LinguaManager.T("FiltroTutte");
                cmbFiltroCategoria.SelectedIndex = idxCorrente >= 0 ? idxCorrente : 0;
            }

            for (int i = 0; i < dgvCronometraggio.Columns.Count; i++)
            {
                dgvCronometraggio.Columns[i].HeaderText = LinguaManager.FormatGiroHeader(i + 1);
            }

            if (dgvClassifiche.Columns.Count >= 6)
            {
                dgvClassifiche.Columns[0].HeaderText = LinguaManager.T("ColPos");
                dgvClassifiche.Columns[1].HeaderText = LinguaManager.T("ColPett");
                dgvClassifiche.Columns[2].HeaderText = LinguaManager.T("ColNome");
                dgvClassifiche.Columns[3].HeaderText = LinguaManager.T("ColCat");
                dgvClassifiche.Columns[4].HeaderText = LinguaManager.T("ColGiri");
                dgvClassifiche.Columns[5].HeaderText = LinguaManager.T("ColTempo");
            }

            btnStart.Text = LinguaManager.T("BtnStart");
            btnPausa.Text = LinguaManager.T("BtnPausa");
            btnAzzeraTimer.Text = LinguaManager.T("BtnAzzeraTimer");

            toolTipGlobale.SetToolTip(btnKoFi, LinguaManager.T("TipSupporta"));
            btnEsportaExcel.Text = LinguaManager.T("BtnEsportaExcel");
            btnEsportaPdf.Text = LinguaManager.T("BtnEsportaPdf");

            TxtInserisciNumero_TextChanged(txtInserisciNumero, EventArgs.Empty);
            AggiornaSchedaRiepilogo(hostUltimo, txtUltimo, lblAnteprimaUltimo);
            AggiornaSchedaRiepilogo(hostPenultimo, txtPenultimo, lblAnteprimaPenultimo);
            AggiornaSchedaRiepilogo(hostTerzultimo, txtTerzultimo, lblAnteprimaTerzultimo);
            AggiornaGrigliaClassifiche();
        }

        /// <summary>
        /// Prepares a lap-grid cell for inline editing by replacing its formatted display text
        /// ("n.12 (00:15:30)") with the raw editable bib number, or cancels editing if the cell is empty.
        /// </summary>
        private void DgvCronometraggio_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            var cell = dgvCronometraggio.Rows[e.RowIndex].Cells[e.ColumnIndex];
            if (cell.Tag is Passaggio p)
            {
                cell.Value = p.Pettorale.ToString();
            }
            else
            {
                e.Cancel = true;
            }
        }

        /// <summary>
        /// Validates an edited bib number in the lap grid, updates the underlying <see cref="Passaggio"/>,
        /// recalculates standings, and persists the updated session.
        /// </summary>
        private void DgvCronometraggio_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || e.RowIndex >= dgvCronometraggio.Rows.Count) return;
            var cell = dgvCronometraggio.Rows[e.RowIndex].Cells[e.ColumnIndex];
            if (cell.Tag is Passaggio p)
            {
                string val = cell.Value?.ToString() ?? "";
                if (int.TryParse(val.Trim(), out int nuovoPettorale) && nuovoPettorale > 0)
                {
                    p.Pettorale = nuovoPettorale;
                    BeginInvoke(new Action(() =>
                    {
                        RicalcolaEAggiorna();
                        SalvaSessioneTemporanea();
                    }));
                }
                else
                {
                    cell.Value = $"n.{p.Pettorale} ({p.Tempo:hh\\:mm\\:ss})";
                }
            }
        }

        /// <summary>
        /// Serializes the current race state (metadata, riders, crossings, elapsed time) to
        /// %LocalAppData%\CicloRank\sessione_temp.json for automatic crash recovery.
        /// </summary>
        private void SalvaSessioneTemporanea()
        {
            try
            {
                var sessione = new SessioneGara
                {
                    DatiGara = this.datiGara,
                    ListaCiclisti = this.listaCiclisti,
                    ListaPassaggi = this.listaPassaggi,
                    TempoTrascorsoSecondi = this.tempoTrascorso.TotalSeconds
                };

                string json = JsonSerializer.Serialize(sessione, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(FileSessioneTemp, json, Utf8NoBom);
            }
            catch { }
        }

        /// <summary>
        /// Deletes the temporary auto-recovery session file after an intentional full data reset.
        /// </summary>
        private void EliminaSessioneTemporanea()
        {
            try
            {
                try
                {
                    if (!File.Exists(FileSessioneTemp) && File.Exists("sessione_temp.json"))
                    {
                        File.Copy("sessione_temp.json", FileSessioneTemp, true);
                        File.Delete("sessione_temp.json");
                    }
                }
                catch { }

                if (File.Exists(FileSessioneTemp)) File.Delete(FileSessioneTemp);
            }
            catch { }
        }

        /// <summary>
        /// Prompts the user on startup to restore race data if an unexported temporary session file exists.
        /// </summary>
        private void RipristinaSessioneSeEsistente()
        {
            try
            {
                if (!File.Exists(FileSessioneTemp) && File.Exists("sessione_temp.json"))
                {
                    File.Copy("sessione_temp.json", FileSessioneTemp, true);
                    File.Delete("sessione_temp.json");
                }
            }
            catch { }

            if (File.Exists(FileSessioneTemp))
            {
                var result = MessageBox.Show(
                    LinguaManager.T("MsgRipristino"),
                    LinguaManager.T("TitleRipristino"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    CaricaSessioneDaFile(FileSessioneTemp);
                }
                else
                {
                    EliminaSessioneTemporanea();
                }
            }
        }

        /// <summary>
        /// Prompts for a destination path and exports the complete race session to a JSON file.
        /// </summary>
        private void BtnEsportaSessione_Click(object? sender, EventArgs e)
        {
            string nomeDefault = $"{LinguaManager.T("SessionBaseFileName")}_{DateTime.Now:yyyyMMdd_HHmm}.json";
            using (var sfd = new SaveFileDialog
            {
                Filter = LinguaManager.T("FilterSessioneSave"),
                FileName = nomeDefault
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var sessione = new SessioneGara
                        {
                            DatiGara = this.datiGara,
                            ListaCiclisti = this.listaCiclisti,
                            ListaPassaggi = this.listaPassaggi,
                            TempoTrascorsoSecondi = this.tempoTrascorso.TotalSeconds
                        };

                        string json = JsonSerializer.Serialize(sessione, new JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(sfd.FileName, json, Utf8NoBom);
                        MessageBox.Show(
                            LinguaManager.T("MsgSessioneSalvata"),
                            LinguaManager.T("TitleEsportaSessione"),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            LinguaManager.T("MsgErroreEsportaSessione") + ex.Message,
                            LinguaManager.T("TitleErroreSalvaSessione"),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Opens a file dialog to load a previously saved CicloRank JSON race session.
        /// </summary>
        private void BtnImportaSessione_Click(object? sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog { Filter = LinguaManager.T("FilterSessioneOpen") })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    CaricaSessioneDaFile(ofd.FileName);
                }
            }
        }

        /// <summary>
        /// Deserializes a <see cref="SessioneGara"/> JSON file, restores stopwatch elapsed time,
        /// race metadata, riders, and lap crossings, and refreshes all grids.
        /// </summary>
        private void CaricaSessioneDaFile(string filePath)
        {
            try
            {
                string json = File.ReadAllText(filePath, Encoding.UTF8);
                var sessione = JsonSerializer.Deserialize<SessioneGara>(json);
                if (sessione != null)
                {
                    timerGara.Stop();
                    timerInEsecuzione = false;

                    this.datiGara = sessione.DatiGara ?? new DatiGara();
                    SportCorrente = this.datiGara.Sport;
                    this.listaCiclisti = sessione.ListaCiclisti ?? new List<Ciclista>();
                    this.listaPassaggi = sessione.ListaPassaggi ?? new List<Passaggio>();
                    this.tempoTrascorso = TimeSpan.FromSeconds(sessione.TempoTrascorsoSecondi);

                    btnAzzeraTimer.Visible = tempoTrascorso > TimeSpan.Zero;

                    lblCronometro.Text = tempoTrascorso.ToString(@"hh\:mm\:ss");
                    AggiornaColoreCronometro();

                    PopolaMenuCategorie();
                    RicalcolaEAggiorna();
                    MessageBox.Show(
                        LinguaManager.T("MsgSessioneCaricata"),
                        LinguaManager.T("TitleImportaSessione"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    LinguaManager.T("MsgErroreImportaSessione") + ex.Message,
                    LinguaManager.T("TitleErroreCaricamentoSessione"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Periodic timer tick handler (every 200ms) that computes drift-free elapsed time
        /// from <see cref="DateTime.UtcNow"/> and updates the main stopwatch display.
        /// </summary>
        private void TimerGara_Tick(object? sender, EventArgs e)
        {
            tempoTrascorso = DateTime.UtcNow - orarioInizioGara;
            string testoCrono = tempoTrascorso.ToString(@"hh\:mm\:ss");
            if (lblCronometro.Text != testoCrono)
            {
                lblCronometro.Text = testoCrono;
                AggiornaColoreCronometro();
            }
        }

        /// <summary>
        /// Starts or resumes the race stopwatch without losing previously accumulated elapsed time.
        /// </summary>
        private void BtnStart_Click(object? sender, EventArgs e)
        {
            if (!timerInEsecuzione)
            {
                orarioInizioGara = DateTime.UtcNow - tempoTrascorso;
                timerGara.Start();
                timerInEsecuzione = true;
                btnAzzeraTimer.Visible = false;
            }
        }

        /// <summary>
        /// Pauses the race stopwatch, reveals the stopwatch reset button, and saves a recovery snapshot.
        /// </summary>
        private void BtnPausa_Click(object? sender, EventArgs e)
        {
            if (timerInEsecuzione)
            {
                timerGara.Stop();
                timerInEsecuzione = false;
                btnAzzeraTimer.Visible = true;
                SalvaSessioneTemporanea();
            }
        }

        /// <summary>
        /// Prompts for confirmation and resets the stopwatch clock back to 00:00:00.
        /// </summary>
        private void BtnAzzeraTimer_Click(object? sender, EventArgs e)
        {
            var res = MessageBox.Show(
                LinguaManager.T("MsgConfermaAzzeraTimer"),
                LinguaManager.T("TitleConfermaAzzeraTimer"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (res == DialogResult.Yes)
            {
                timerGara.Stop();
                timerInEsecuzione = false;
                tempoTrascorso = TimeSpan.Zero;
                lblCronometro.Text = "00:00:00";
                AggiornaColoreCronometro();
                btnAzzeraTimer.Visible = false;
                SalvaSessioneTemporanea();
            }
        }

        /// <summary>
        /// Captures the Enter key in the main bib input box to immediately log a new lap crossing.
        /// </summary>
        private void TxtInserisciNumero_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (int.TryParse(txtInserisciNumero.Text.Trim(), out int pettorale) && pettorale > 0)
                {
                    RegistraPassaggio(pettorale);
                    txtInserisciNumero.Clear();
                }
            }
        }

        /// <summary>
        /// Appends a new <see cref="Passaggio"/> at the current stopwatch timestamp, recalculates
        /// all standings and grids, and writes a crash-recovery snapshot.
        /// </summary>
        private void RegistraPassaggio(int pettorale)
        {
            listaPassaggi.Add(new Passaggio(pettorale, tempoTrascorso));
            RicalcolaEAggiorna();
            SalvaSessioneTemporanea();
        }

        /// <summary>
        /// Updates the bib number of one of the last three recorded crossings directly from
        /// the quick-edit summary cards at the top of the window.
        /// </summary>
        private void AggiornaPassaggioDaTextBox(int posizioneDaFondo, string nuovoTesto)
        {
            if (listaPassaggi.Count >= posizioneDaFondo && int.TryParse(nuovoTesto.Trim(), out int nuovoPettorale) && nuovoPettorale > 0)
            {
                int index = listaPassaggi.Count - posizioneDaFondo;
                listaPassaggi[index].Pettorale = nuovoPettorale;
                RicalcolaEAggiorna();
                SalvaSessioneTemporanea();
            }
            else
            {
                RipristinaTextBoxPassaggio(posizioneDaFondo);
            }
        }

        /// <summary>
        /// Reverts one of the quick-edit summary cards back to its saved bib number if editing is canceled.
        /// </summary>
        private void RipristinaTextBoxPassaggio(int posizioneDaFondo)
        {
            string valoreSalvato = listaPassaggi.Count >= posizioneDaFondo
                ? listaPassaggi[listaPassaggi.Count - posizioneDaFondo].Pettorale.ToString()
                : "";

            if (posizioneDaFondo == 1)
            {
                txtUltimo.Text = valoreSalvato;
                AggiornaSchedaRiepilogo(hostUltimo, txtUltimo, lblAnteprimaUltimo);
            }
            else if (posizioneDaFondo == 2)
            {
                txtPenultimo.Text = valoreSalvato;
                AggiornaSchedaRiepilogo(hostPenultimo, txtPenultimo, lblAnteprimaPenultimo);
            }
            else if (posizioneDaFondo == 3)
            {
                txtTerzultimo.Text = valoreSalvato;
                AggiornaSchedaRiepilogo(hostTerzultimo, txtTerzultimo, lblAnteprimaTerzultimo);
            }
        }

        /// <summary>
        /// Rebuilds the O(1) bib-number-to-<see cref="Ciclista"/> lookup dictionary from <see cref="listaCiclisti"/>.
        /// </summary>
        private void RicostruisciMappaCiclisti()
        {
            mappaCiclisti.Clear();
            foreach (var c in listaCiclisti)
            {
                mappaCiclisti[c.Pettorale] = c;
            }
        }

        /// <summary>
        /// Rebuilds every participant's lap list from the master <see cref="listaPassaggi"/> log,
        /// creates placeholder entries for unknown bib numbers, updates the last-3-passages cards,
        /// and refreshes both the timing grid and standings grid.
        /// </summary>
        private void RicalcolaEAggiorna()
        {
            RicostruisciMappaCiclisti();
            foreach (var c in listaCiclisti)
            {
                c.Passaggi.Clear();
            }

            foreach (var p in listaPassaggi)
            {
                if (!mappaCiclisti.TryGetValue(p.Pettorale, out var ciclista))
                {
                    ciclista = new Ciclista
                    {
                        Pettorale = p.Pettorale,
                        Nome = LinguaManager.T("SconosciutoNome"),
                        Categoria = "N.D.",
                        NomeSocieta = "",
                        Ente = "",
                        NonInElenco = true
                    };
                    listaCiclisti.Add(ciclista);
                    mappaCiclisti[p.Pettorale] = ciclista;
                }
                ciclista.Passaggi.Add(p.Tempo);
            }

            txtUltimo.Text = listaPassaggi.Count >= 1 ? listaPassaggi[listaPassaggi.Count - 1].Pettorale.ToString() : "";
            txtPenultimo.Text = listaPassaggi.Count >= 2 ? listaPassaggi[listaPassaggi.Count - 2].Pettorale.ToString() : "";
            txtTerzultimo.Text = listaPassaggi.Count >= 3 ? listaPassaggi[listaPassaggi.Count - 3].Pettorale.ToString() : "";

            AggiornaSchedaRiepilogo(hostUltimo, txtUltimo, lblAnteprimaUltimo);
            AggiornaSchedaRiepilogo(hostPenultimo, txtPenultimo, lblAnteprimaPenultimo);
            AggiornaSchedaRiepilogo(hostTerzultimo, txtTerzultimo, lblAnteprimaTerzultimo);

            AggiornaGrigliaCronometraggio();
            AggiornaGrigliaClassifiche();
        }

        /// <summary>
        /// Updates the visual state and rider name/category preview label of a quick-edit card
        /// based on whether its bib number exists in the loaded start list.
        /// </summary>
        private void AggiornaSchedaRiepilogo(RoundedInputHost host, TextBox txt, Label previewLabel)
        {
            string raw = txt.Text.Trim();
            if (int.TryParse(raw, out int pettorale) && pettorale > 0)
            {
                mappaCiclisti.TryGetValue(pettorale, out var ciclista);
                if (ciclista != null && !ciclista.NonInElenco)
                {
                    host.SyncColors(ThemeManager.InputBack, ThemeManager.SuccessText, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
                    previewLabel.ForeColor = ThemeManager.SuccessText;
                    bool hasValidCat = !string.IsNullOrWhiteSpace(ciclista.Categoria) &&
                                       !ciclista.Categoria.Equals("N.D.", StringComparison.OrdinalIgnoreCase) &&
                                       ciclista.Categoria != "-";
                    previewLabel.Text = hasValidCat
                        ? $"{ciclista.Nome.ToUpperInvariant()} ({ciclista.Categoria.ToUpperInvariant()})"
                        : ciclista.Nome.ToUpperInvariant();
                }
                else
                {
                    host.SyncColors(ThemeManager.WarningBack, ThemeManager.WarningText, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
                    previewLabel.ForeColor = ThemeManager.WarningText;
                    previewLabel.Text = LinguaManager.T("SconosciutoNome");
                }
            }
            else
            {
                host.SyncColors(ThemeManager.InputBack, ThemeManager.SuccessText, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
                previewLabel.Text = "";
            }
        }

        /// <summary>
        /// Dynamically expands or trims lap columns ("Lap 1" .. "Lap N", minimum 8) in the
        /// timing grid to match the maximum number of laps completed by any rider.
        /// </summary>
        private void AssicuraColonneCronometraggio(int numeroColonneRichieste)
        {
            int targetCols = Math.Max(8, numeroColonneRichieste);

            while (dgvCronometraggio.Columns.Count < targetCols)
            {
                int idx = dgvCronometraggio.Columns.Count + 1;
                dgvCronometraggio.Columns.Add($"Giro{idx}", LinguaManager.FormatGiroHeader(idx));
            }

            while (dgvCronometraggio.Columns.Count > targetCols)
            {
                dgvCronometraggio.Columns.RemoveAt(dgvCronometraggio.Columns.Count - 1);
            }

            for (int i = 0; i < dgvCronometraggio.Columns.Count; i++)
            {
                dgvCronometraggio.Columns[i].HeaderText = LinguaManager.FormatGiroHeader(i + 1);
                dgvCronometraggio.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dgvCronometraggio.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            AggiornaLarghezzaColonneCronometraggio();
        }

        /// <summary>
        /// Calculates uniform column widths so 8 lap columns fit cleanly within the timing grid viewport.
        /// </summary>
        private void AggiornaLarghezzaColonneCronometraggio()
        {
            if (dgvCronometraggio == null || dgvCronometraggio.Columns.Count == 0) return;

            int availWidth = dgvCronometraggio.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 2;
            int colWidth = Math.Max(102, availWidth / 8);

            foreach (DataGridViewColumn col in dgvCronometraggio.Columns)
            {
                col.Width = colWidth;
            }
        }

        /// <summary>
        /// Populates the left-hand lap-by-lap timing grid, grouping crossings by each rider's
        /// lap index, sorting each lap column chronologically, and highlighting unknown riders.
        /// </summary>
        private void AggiornaGrigliaCronometraggio()
        {
            int savedFirstDisplayedCol = dgvCronometraggio.FirstDisplayedScrollingColumnIndex;
            dgvCronometraggio.SuspendLayout();
            try
            {
                dgvCronometraggio.Rows.Clear();

                // Group all recorded crossings by each rider's personal lap index (0 = Lap 1, 1 = Lap 2, etc.).
                var conteggioGiriCiclista = new Dictionary<int, int>();
                var passaggiPerGiro = new List<List<(Ciclista Ciclista, Passaggio Passaggio, TimeSpan Tempo)>>();

                foreach (var p in listaPassaggi)
                {
                    if (!mappaCiclisti.TryGetValue(p.Pettorale, out var ciclista)) continue;

                    if (!conteggioGiriCiclista.ContainsKey(ciclista.Pettorale))
                    {
                        conteggioGiriCiclista[ciclista.Pettorale] = 0;
                    }

                    int giroIndex = conteggioGiriCiclista[ciclista.Pettorale];
                    while (passaggiPerGiro.Count <= giroIndex)
                    {
                        passaggiPerGiro.Add(new List<(Ciclista Ciclista, Passaggio Passaggio, TimeSpan Tempo)>());
                    }

                    passaggiPerGiro[giroIndex].Add((ciclista, p, p.Tempo));
                    conteggioGiriCiclista[ciclista.Pettorale]++;
                }

                int numColonne = Math.Max(8, passaggiPerGiro.Count);
                AssicuraColonneCronometraggio(numColonne);

                while (passaggiPerGiro.Count < numColonne)
                {
                    passaggiPerGiro.Add(new List<(Ciclista Ciclista, Passaggio Passaggio, TimeSpan Tempo)>());
                }

                // Sort crossings chronologically within each lap column so faster riders appear near the top.
                for (int g = 0; g < numColonne; g++)
                {
                    passaggiPerGiro[g] = passaggiPerGiro[g]
                        .OrderBy(x => x.Tempo)
                        .ToList();
                }

                int maxRighe = passaggiPerGiro.Max(l => l.Count);
                if (maxRighe == 0) return;

                for (int r = 0; r < maxRighe; r++)
                {
                    var rowValues = new object[numColonne];
                    var passaggiRow = new Passaggio?[numColonne];
                    var isNonInElenco = new bool[numColonne];

                    for (int g = 0; g < numColonne; g++)
                    {
                        if (r < passaggiPerGiro[g].Count)
                        {
                            var item = passaggiPerGiro[g][r];
                            rowValues[g] = $"n.{item.Ciclista.Pettorale} ({item.Tempo:hh\\:mm\\:ss})";
                            passaggiRow[g] = item.Passaggio;
                            isNonInElenco[g] = item.Ciclista.NonInElenco;
                        }
                        else
                        {
                            rowValues[g] = "";
                            passaggiRow[g] = null;
                            isNonInElenco[g] = false;
                        }
                    }

                    int rowIndex = dgvCronometraggio.Rows.Add(rowValues);

                    for (int g = 0; g < numColonne; g++)
                    {
                        dgvCronometraggio.Rows[rowIndex].Cells[g].Tag = passaggiRow[g];

                        if (isNonInElenco[g])
                        {
                            dgvCronometraggio.Rows[rowIndex].Cells[g].Style.ForeColor = ThemeManager.WarningText;
                            dgvCronometraggio.Rows[rowIndex].Cells[g].Style.SelectionForeColor = ThemeManager.WarningText;
                        }
                    }
                }

                if (savedFirstDisplayedCol >= 0 && savedFirstDisplayedCol < dgvCronometraggio.Columns.Count)
                {
                    try { dgvCronometraggio.FirstDisplayedScrollingColumnIndex = savedFirstDisplayedCol; } catch { }
                }
            }
            finally
            {
                dgvCronometraggio.ResumeLayout();
            }
        }

        /// <summary>
        /// Populates the right-hand live standings grid, applying any active category filter
        /// and ranking participants by descending lap count and ascending total race time.
        /// </summary>
        private void AggiornaGrigliaClassifiche()
        {
            dgvClassifiche.SuspendLayout();
            try
            {
                dgvClassifiche.Rows.Clear();

                var query = listaCiclisti.Where(c => c.Giri > 0);

                if (cmbFiltroCategoria.SelectedIndex > 0)
                {
                    string catFiltro = cmbFiltroCategoria.SelectedItem?.ToString() ?? "";
                    query = query.Where(c => c.Categoria.Equals(catFiltro, StringComparison.OrdinalIgnoreCase));
                }

                var classifica = query.OrderByDescending(c => c.Giri).ThenBy(c => c.TempoTotale).ToList();

                int pos = 1;
                foreach (var c in classifica)
                {
                    string nomeVisualizzato = c.NonInElenco ? LinguaManager.T("SconosciutoNome") : c.Nome;
                    int rowIndex = dgvClassifiche.Rows.Add(pos + "\u00B0", c.Pettorale, nomeVisualizzato, c.Categoria, c.Giri, c.TempoTotale.ToString(@"hh\:mm\:ss"));

                    if (c.NonInElenco)
                    {
                        dgvClassifiche.Rows[rowIndex].DefaultCellStyle.ForeColor = ThemeManager.WarningText;
                        dgvClassifiche.Rows[rowIndex].DefaultCellStyle.SelectionForeColor = ThemeManager.WarningText;
                    }
                    pos++;
                }
            }
            finally
            {
                dgvClassifiche.ResumeLayout();
            }
        }

        /// <summary>
        /// Displays a modal dialog allowing the operator to edit race header metadata
        /// (federation, organizer, event title, location, date, and sport discipline).
        /// </summary>
        private void BtnDatiGara_Click(object? sender, EventArgs e)
        {
            using (var dlg = new Form())
            {
                dlg.Text = LinguaManager.T("DlgDatiTitle");
                dlg.Size = new Size(510, 365);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                ThemeManager.ApplicaTemaFinestra(dlg);
                dlg.BackColor = ThemeManager.WindowBack;
                dlg.ForeColor = ThemeManager.TextPrimary;
                dlg.Font = new Font("Segoe UI", 9.5F);

                Label CreaLabelDialog(string text, int top) => new Label
                {
                    Text = text,
                    Left = 24,
                    Top = top + 6,
                    Width = 175,
                    ForeColor = ThemeManager.TextPrimary
                };

                RoundedInputHost CreaInputDialog(string initialValue, int top, out TextBox tb)
                {
                    tb = new TextBox { Text = initialValue, Font = new Font("Segoe UI", 10F) };
                    var host = new RoundedInputHost(tb)
                    {
                        Left = 205,
                        Top = top,
                        Width = 265,
                        Height = 32,
                        CornerRadius = 7
                    };
                    host.SyncColors(ThemeManager.CardBack, ThemeManager.TextPrimary, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
                    host.CenterInnerTextBox();
                    return host;
                }

                var lblComitato = CreaLabelDialog(LinguaManager.T("LblComitato"), 20);
                var hostComitato = CreaInputDialog(datiGara.ComitatoProvinciale, 20, out var txtComitato);

                var lblOrg = CreaLabelDialog(LinguaManager.T("LblOrg"), 60);
                var hostOrg = CreaInputDialog(datiGara.Organizzazione, 60, out var txtOrg);

                var lblDenom = CreaLabelDialog(LinguaManager.T("LblDenom"), 100);
                var hostDenom = CreaInputDialog(datiGara.Denominazione, 100, out var txtDenom);

                var lblLoc = CreaLabelDialog(LinguaManager.T("LblLoc"), 140);
                var hostLoc = CreaInputDialog(datiGara.Localita, 140, out var txtLoc);

                var lblData = CreaLabelDialog(LinguaManager.T("LblData"), 180);
                var hostData = CreaInputDialog(datiGara.DataGara, 180, out var txtData);

                var lblSport = CreaLabelDialog(LinguaManager.T("LblSport"), 220);
                var hostSport = CreaInputDialog(datiGara.Sport, 220, out var txtSport);

                var btnSalva = new RoundedButton
                {
                    Text = LinguaManager.T("BtnSalvaDati"),
                    DialogResult = DialogResult.OK,
                    Left = 205,
                    Top = 272,
                    Width = 135,
                    Height = 38,
                    CornerRadius = 9,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                };
                btnSalva.ApplyScheme(ThemeManager.BtnStart);

                var btnAnnulla = new RoundedButton
                {
                    Text = LinguaManager.T("BtnAnnulla"),
                    DialogResult = DialogResult.Cancel,
                    Left = 348,
                    Top = 272,
                    Width = 122,
                    Height = 38,
                    CornerRadius = 9,
                    Font = new Font("Segoe UI", 9F)
                };
                btnAnnulla.ApplyScheme(ThemeManager.BtnIconNeutral);

                dlg.Controls.AddRange(new Control[]
                {
                    lblComitato, hostComitato,
                    lblOrg, hostOrg,
                    lblDenom, hostDenom,
                    lblLoc, hostLoc,
                    lblData, hostData,
                    lblSport, hostSport,
                    btnSalva, btnAnnulla
                });

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    datiGara.ComitatoProvinciale = txtComitato.Text.Trim();
                    datiGara.Organizzazione = txtOrg.Text.Trim();
                    datiGara.Denominazione = txtDenom.Text.Trim();
                    datiGara.Localita = txtLoc.Text.Trim();
                    datiGara.DataGara = txtData.Text.Trim();
                    datiGara.Sport = txtSport.Text.Trim();
                    SportCorrente = datiGara.Sport;

                    SalvaSessioneTemporanea();
                    MessageBox.Show(
                        LinguaManager.T("MsgDatiSalvati"),
                        LinguaManager.T("TitleDatiSalvati"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }

                /// <summary>
                /// Provides real-time rider lookup feedback beneath the main bib input box as the operator types.
                /// </summary>
                private void TxtInserisciNumero_TextChanged(object? sender, EventArgs e)
        {
            string input = txtInserisciNumero.Text.Trim();
            if (int.TryParse(input, out int pettorale) && pettorale > 0)
            {
                mappaCiclisti.TryGetValue(pettorale, out var ciclista);
                if (ciclista != null && !ciclista.NonInElenco)
                {
                    hostInserisciNumero.SyncColors(ThemeManager.InputBack, ThemeManager.TextPrimary, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
                    lblNomeAnteprima.ForeColor = ThemeManager.SuccessText;
                    lblNomeAnteprima.Text = $"{ciclista.Nome} ({ciclista.Categoria})";
                }
                else
                {
                    hostInserisciNumero.SyncColors(ThemeManager.InputBack, ThemeManager.WarningText, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
                    lblNomeAnteprima.ForeColor = ThemeManager.WarningText;
                    lblNomeAnteprima.Text = LinguaManager.T("LblSconosciutoAnteprima");
                }
            }
            else
            {
                hostInserisciNumero.SyncColors(ThemeManager.InputBack, ThemeManager.TextPrimary, ThemeManager.InputBorder, ThemeManager.InputFocusBorder);
                lblNomeAnteprima.Text = "";
            }
        }


        /// <summary>
        /// Opens a file picker to import an athlete start list (.xlsx, .csv, or .txt), merges
        /// loaded rider details with any existing bib records, and reports duplicate bibs if found.
        /// </summary>
        private void BtnCaricaDati_Click(object? sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog
            {
                Filter = LinguaManager.T("FilterAnagrafica"),
                Title = LinguaManager.T("TitleSelezionaAnagrafica")
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var atletiCaricati = ExcelHelper.LeggiFileAtleti(ofd.FileName);

                        if (atletiCaricati == null || atletiCaricati.Count == 0)
                        {
                            MessageBox.Show(
                                LinguaManager.T("MsgNessunAtleta"),
                                LinguaManager.T("TitleAttenzioneCaricamento"),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            return;
                        }

                        RicostruisciMappaCiclisti();
                        foreach (var atletaNuovo in atletiCaricati)
                        {
                            if (mappaCiclisti.TryGetValue(atletaNuovo.Pettorale, out var ciclistaEsistente))
                            {
                                ciclistaEsistente.Nome = atletaNuovo.Nome;
                                ciclistaEsistente.Categoria = atletaNuovo.Categoria;
                                ciclistaEsistente.NomeSocieta = atletaNuovo.NomeSocieta;
                                ciclistaEsistente.Ente = atletaNuovo.Ente;
                                ciclistaEsistente.NonInElenco = false;
                            }
                            else
                            {
                                listaCiclisti.Add(atletaNuovo);
                                mappaCiclisti[atletaNuovo.Pettorale] = atletaNuovo;
                            }
                        }

                        PopolaMenuCategorie();
                        RicalcolaEAggiorna();
                        SalvaSessioneTemporanea();

                        string msgOk = string.Format(LinguaManager.T("MsgImportAnagraficaOk"), atletiCaricati.Count, cmbFiltroCategoria.Items.Count - 1);
                        if (ExcelColumnMapper.UltimiPettoraliDuplicati.Count > 0)
                        {
                            string dupList = string.Join(", ", ExcelColumnMapper.UltimiPettoraliDuplicati);
                            msgOk += $"\n\n[!] Pettorali duplicati ignorati nel file: {dupList}";
                        }

                        MessageBox.Show(
                            msgOk,
                            LinguaManager.T("TitleEsitoCaricamento"),
                            MessageBoxButtons.OK,
                            ExcelColumnMapper.UltimiPettoraliDuplicati.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            string.Format(LinguaManager.T("MsgErroreAnagrafica"), ex.Message),
                            LinguaManager.T("TitleErroreImportazione"),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Rebuilds the category filter dropdown with "ALL" followed by every distinct category
        /// present among loaded or active participants.
        /// </summary>
        private void PopolaMenuCategorie()
        {
            cmbFiltroCategoria.Items.Clear();
            cmbFiltroCategoria.Items.Add(LinguaManager.T("FiltroTutte"));
            var categorie = listaCiclisti
                .Select(c => c.Categoria.Trim())
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(c => c, StringComparer.OrdinalIgnoreCase);

            foreach (var cat in categorie)
            {
                cmbFiltroCategoria.Items.Add(cat);
            }
            cmbFiltroCategoria.SelectedIndex = 0;
        }

        /// <summary>
        /// Prompts for confirmation and clears all recorded lap crossings, unknown riders,
        /// stopwatch time, and the temporary recovery file while keeping loaded start-list riders.
        /// </summary>
        private void BtnAzzeraDati_Click(object? sender, EventArgs e)
        {
            var res = MessageBox.Show(
                LinguaManager.T("MsgConfermaAzzeraDati"),
                LinguaManager.T("TitleConfermaAzzeraDati"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (res == DialogResult.Yes)
            {
                timerGara.Stop();
                timerInEsecuzione = false;
                tempoTrascorso = TimeSpan.Zero;
                lblCronometro.Text = "00:00:00";
                AggiornaColoreCronometro();
                btnAzzeraTimer.Visible = false;

                listaPassaggi.Clear();
                listaCiclisti.RemoveAll(c => c.NonInElenco);
                foreach (var c in listaCiclisti) c.Passaggi.Clear();

                lblNomeAnteprima.Text = "";

                RicalcolaEAggiorna();
                EliminaSessioneTemporanea();
            }
        }

        /// <summary>
        /// Opens a SaveFileDialog and exports the current race standings to a styled .xlsx workbook.
        /// </summary>
        private void BtnEsportaExcel_Click(object? sender, EventArgs e)
        {
            if (!listaCiclisti.Any(c => c.Giri > 0))
            {
                MessageBox.Show(LinguaManager.T("MsgNessunDatoExport"), LinguaManager.T("TitleAttenzione"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string defaultName = $"{LinguaManager.T("ExportBaseFileName")}.xlsx";
            using (var sfd = new SaveFileDialog { Filter = LinguaManager.T("FilterExcelSave"), FileName = defaultName })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportHelper.EsportaExcelCsv(sfd.FileName, listaCiclisti, datiGara);
                        // Temporary recovery backup is retained until Reset Data is explicitly triggered
                        MessageBox.Show(LinguaManager.T("MsgExcelOk"), LinguaManager.T("TitleExcelOk"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(LinguaManager.T("MsgErroreExcel") + ex.Message, LinguaManager.T("TitleErroreExport"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Opens a SaveFileDialog and exports the official A4 PDF classification report.
        /// </summary>
        private void BtnEsportaPdf_Click(object? sender, EventArgs e)
        {
            if (!listaCiclisti.Any(c => c.Giri > 0))
            {
                MessageBox.Show(LinguaManager.T("MsgNessunDatoExport"), LinguaManager.T("TitleAttenzione"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string defaultName = $"{LinguaManager.T("ExportBaseFileName")}.pdf";
            using (var sfd = new SaveFileDialog { Filter = LinguaManager.T("FilterPdfSave"), FileName = defaultName })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportHelper.EsportaPdf(sfd.FileName, listaCiclisti, datiGara);
                        // Temporary recovery backup is retained until Reset Data is explicitly triggered
                        MessageBox.Show(LinguaManager.T("MsgPdfOk"), LinguaManager.T("TitlePdfOk"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(LinguaManager.T("MsgErrorePdf") + ex.Message, LinguaManager.T("TitleErroreExport"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
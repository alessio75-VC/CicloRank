// ============================================================================
// File: ExportHelper.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Coordinates official race classification exports to Excel (.xlsx) and
//   printable A4 PDF documents (.pdf). Renders a self-contained HTML5/CSS3
//   report and converts it via headless Chromium (Edge/Chrome/Brave), with
//   automatic fallback to printable .html if no compatible browser is found.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace CicloRank
{
    /// <summary>
    /// Handles exporting official race classifications to Excel (.xlsx) and printable
    /// A4 PDF reports (.pdf) using headless Chromium browser rendering.
    /// </summary>
    public static class ExportHelper
    {
        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

        /// <summary>
        /// Exports race standings to an Excel (.xlsx) workbook via <see cref="ExcelHelper.EsportaXlsx"/>.
        /// </summary>
        public static void EsportaExcelCsv(string filePath, List<Ciclista> tuttiAtleti, DatiGara infoGara)
        {
            ExcelHelper.EsportaXlsx(filePath, tuttiAtleti, infoGara);
        }

        /// <summary>
        /// Generates a print-ready A4 PDF classification report by rendering a temporary HTML5/CSS3
        /// document through a headless Chromium-based browser (Edge, Chrome, or Brave).
        /// Falls back to saving a printable .html file if no supported browser executable is found.
        /// </summary>
        /// <param name="filePath">Destination .pdf file path.</param>
        /// <param name="tuttiAtleti">Complete list of participants.</param>
        /// <param name="infoGara">Event header metadata.</param>
        public static void EsportaPdf(string filePath, List<Ciclista> tuttiAtleti, DatiGara infoGara)
        {
            string tempHtmlPath = Path.Combine(Path.GetTempPath(), $"classifica_temp_{Guid.NewGuid():N}.html");
            string tempProfileDir = Path.Combine(Path.GetTempPath(), $"ciclorank_pdf_profile_{Guid.NewGuid():N}");

            try
            {
                string htmlContent = GeneraHtmlDocumento(tuttiAtleti, infoGara);
                File.WriteAllText(tempHtmlPath, htmlContent, Utf8NoBom);

                string browserPath = TrovaEseguibileBrowser();
                if (string.IsNullOrEmpty(browserPath))
                {
                    string fallbackHtml = Path.ChangeExtension(filePath, ".html");
                    File.Copy(tempHtmlPath, fallbackHtml, true);
                    throw new Exception(LinguaManager.T("MsgFallbackHtml"));
                }

                string fileUri = new Uri(tempHtmlPath).AbsoluteUri;

                if (File.Exists(filePath)) File.Delete(filePath);
                Directory.CreateDirectory(tempProfileDir);

                // Launch headless Chromium browser with isolated temporary profile to render HTML cleanly to A4 PDF.
                var psi = new ProcessStartInfo
                {
                    FileName = browserPath,
                    Arguments = $"--headless=new --disable-gpu --no-sandbox --allow-file-access-from-files --user-data-dir=\"{tempProfileDir}\" --no-pdf-header-footer --print-to-pdf=\"{filePath}\" \"{fileUri}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var proc = Process.Start(psi))
                {
                    proc?.WaitForExit(20000);
                }

                // Retry with legacy --headless flag for older Chromium/Edge builds if --headless=new produced no output.
                if (!File.Exists(filePath))
                {
                    psi.Arguments = $"--headless --disable-gpu --no-sandbox --allow-file-access-from-files --user-data-dir=\"{tempProfileDir}\" --no-pdf-header-footer --print-to-pdf=\"{filePath}\" \"{fileUri}\"";
                    using (var proc2 = Process.Start(psi))
                    {
                        proc2?.WaitForExit(20000);
                    }
                }
            }
            finally
            {
                try { if (File.Exists(tempHtmlPath)) File.Delete(tempHtmlPath); } catch { }
                try { if (Directory.Exists(tempProfileDir)) Directory.Delete(tempProfileDir, true); } catch { }
            }
        }

        /// <summary>
        /// Searches standard Windows installation directories for Microsoft Edge, Google Chrome,
        /// or Brave Browser executables to perform headless HTML-to-PDF conversion.
        /// </summary>
        /// <returns>The full executable path if found; otherwise, <see cref="string.Empty"/>.</returns>
        private static string TrovaEseguibileBrowser()
        {
            string[] candidates = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"BraveSoftware\Brave-Browser\Application\brave.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"BraveSoftware\Brave-Browser\Application\brave.exe")
            };

            return candidates.FirstOrDefault(File.Exists) ?? string.Empty;
        }

        /// <summary>
        /// Determines whether a metadata string contains meaningful content rather than empty
        /// placeholders such as "-", "N/D", "N/A", or "Nessuna".
        /// </summary>
        private static bool IsValoreSignificativo(string? val)
        {
            if (string.IsNullOrWhiteSpace(val)) return false;
            string t = val.Trim();
            return t != "-" &&
                   t != "\u2014" &&
                   t != "\u2013" &&
                   !t.Equals("N/D", StringComparison.OrdinalIgnoreCase) &&
                   !t.Equals("N.D.", StringComparison.OrdinalIgnoreCase) &&
                   !t.Equals("N/A", StringComparison.OrdinalIgnoreCase) &&
                   !t.Equals("Nessuna", StringComparison.OrdinalIgnoreCase) &&
                   !t.Equals("Nessuno", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Builds a self-contained HTML5/CSS3 A4 portrait document containing the race header,
        /// event metadata, and either category-grouped tables or a single overall standings table.
        /// </summary>
        private static string GeneraHtmlDocumento(List<Ciclista> tuttiAtleti, DatiGara infoGara)
        {
            infoGara ??= new DatiGara();

            var classificaAssoluta = tuttiAtleti
                .Where(c => c.Giri > 0)
                .OrderByDescending(c => c.Giri)
                .ThenBy(c => c.TempoTotale)
                .ToList();

            var posizioniAssolute = new Dictionary<Ciclista, int>(classificaAssoluta.Count);
            for (int i = 0; i < classificaAssoluta.Count; i++)
            {
                posizioniAssolute[classificaAssoluta[i]] = i + 1;
            }

            bool hasCat = classificaAssoluta.Any(c => IsValoreSignificativo(c.Categoria));
            bool hasSoc = classificaAssoluta.Any(c => IsValoreSignificativo(c.NomeSocieta));
            bool hasEnte = classificaAssoluta.Any(c => !c.NonInElenco && IsValoreSignificativo(c.Ente));

            string titoloDinamico = LinguaManager.CostruisciTitoloClassificaSport(infoGara.Sport);
            string htmlTitle = WebUtility.HtmlEncode(titoloDinamico);
            string headerFed = WebUtility.HtmlEncode(titoloDinamico);
            string lblCommittee = WebUtility.HtmlEncode(LinguaManager.T("RptCommittee"));
            string lblOrg = WebUtility.HtmlEncode(LinguaManager.T("RptOrg"));
            string lblDenom = WebUtility.HtmlEncode(LinguaManager.T("RptDenom"));
            string lblLoc = WebUtility.HtmlEncode(LinguaManager.T("RptLoc"));
            string lblDate = WebUtility.HtmlEncode(LinguaManager.T("RptDate"));
            string lblSport = WebUtility.HtmlEncode(LinguaManager.T("RptSport"));
            string mainTitle = WebUtility.HtmlEncode(LinguaManager.T("RptMainTitle"));
            string lblCategory = WebUtility.HtmlEncode(LinguaManager.T("RptCategory"));
            string colAss = WebUtility.HtmlEncode(LinguaManager.T("RptColAss"));
            string colPos = WebUtility.HtmlEncode(LinguaManager.T("RptColPos"));
            string colDor = WebUtility.HtmlEncode(LinguaManager.T("RptColDor"));
            string colName = WebUtility.HtmlEncode(LinguaManager.T("RptColName"));
            string colCat = WebUtility.HtmlEncode(LinguaManager.T("RptColCat"));
            string colTeam = WebUtility.HtmlEncode(LinguaManager.T("RptColTeam"));
            string colEnte = WebUtility.HtmlEncode(LinguaManager.T("RptColEnte"));
            string colTime = WebUtility.HtmlEncode(LinguaManager.T("RptColTime"));
            string colLap = WebUtility.HtmlEncode(LinguaManager.T("RptColLap"));
            string unknownRider = WebUtility.HtmlEncode(LinguaManager.T("SconosciutoNome"));

            var sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html><head><meta charset='utf-8'>");
            sb.Append($"<title>{htmlTitle}</title>");
            sb.Append("<style>");
            sb.Append("@page { size: A4 portrait; margin: 12mm 14mm 12mm 14mm; }");
            sb.Append("body { font-family: 'Segoe UI', 'Microsoft YaHei', 'Yu Gothic', 'Meiryo', Arial, sans-serif; font-size: 11px; color: #111; margin: 0; padding: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }");
            sb.Append(".header-box { border-bottom: 1.5px solid #222; padding-bottom: 8px; margin-bottom: 10px; }");
            sb.Append(".header-title { font-size: 16px; font-weight: bold; letter-spacing: 0.2px; }");
            sb.Append(".header-sub { font-size: 12px; font-weight: bold; margin-top: 2px; color: #333; }");
            sb.Append(".header-meta { margin-top: 6px; font-size: 11px; line-height: 1.45; }");
            sb.Append(".main-title { text-align: center; font-size: 15px; font-weight: bold; margin: 12px 0 8px 0; border-bottom: 1px solid #333; padding-bottom: 4px; }");
            sb.Append(".cat-header { font-size: 12px; font-weight: bold; margin-top: 12px; margin-bottom: 4px; border-bottom: 1px solid #888; padding-bottom: 2px; page-break-after: avoid; }");
            sb.Append("table.results { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 10px; }");
            sb.Append("table.results th { text-align: left; font-weight: bold; border-bottom: 1.5px solid #333; padding: 4px 3px; white-space: nowrap; }");
            sb.Append("table.results td { padding: 3px 3px; border-bottom: 0.5px solid #e5e5e5; vertical-align: middle; }");
            sb.Append("table.results tr { page-break-inside: avoid; }");
            sb.Append(".num { text-align: center; }");
            sb.Append("</style></head><body>");

            sb.Append("<div class='header-box'>");
            sb.Append($"<div class='header-title'>{headerFed}</div>");
            if (!string.IsNullOrWhiteSpace(infoGara.ComitatoProvinciale))
            {
                sb.Append($"<div class='header-sub'>{lblCommittee} {WebUtility.HtmlEncode(infoGara.ComitatoProvinciale.Trim())}</div>");
            }

            var metaRighe = new List<string>();
            if (!string.IsNullOrWhiteSpace(infoGara.Organizzazione))
                metaRighe.Add($"{lblOrg} <strong>{WebUtility.HtmlEncode(infoGara.Organizzazione.Trim())}</strong>");
            if (!string.IsNullOrWhiteSpace(infoGara.Denominazione))
                metaRighe.Add($"{lblDenom} <strong>{WebUtility.HtmlEncode(infoGara.Denominazione.Trim())}</strong>");

            var locDataParti = new List<string>();
            if (!string.IsNullOrWhiteSpace(infoGara.Localita))
                locDataParti.Add($"{lblLoc} <strong>{WebUtility.HtmlEncode(infoGara.Localita.Trim())}</strong>");
            if (!string.IsNullOrWhiteSpace(infoGara.DataGara))
                locDataParti.Add($"{lblDate} <strong>{WebUtility.HtmlEncode(infoGara.DataGara.Trim())}</strong>");
            if (locDataParti.Count > 0)
                metaRighe.Add(string.Join(" &nbsp;&nbsp;&nbsp; ", locDataParti));

            if (metaRighe.Count > 0)
            {
                sb.Append("<div class='header-meta'>");
                sb.Append(string.Join("<br/>", metaRighe));
                sb.Append("</div>");
            }
            sb.Append("</div>");

            // Render separate category-grouped tables when categories are populated; otherwise render one overall table.
            if (hasCat)
            {
                sb.Append($"<div class='main-title'>{mainTitle}</div>");

                var gruppiCategoria = classificaAssoluta
                    .GroupBy(c => c.Categoria)
                    .OrderBy(g => g.Key);

                foreach (var gruppo in gruppiCategoria)
                {
                    sb.Append($"<div class='cat-header'>{lblCategory} &nbsp; {WebUtility.HtmlEncode(gruppo.Key)}</div>");
                    ScriviIntestazioneTabella(sb, true, hasSoc, hasEnte, colAss, colPos, colDor, colName, colCat, colTeam, colEnte, colTime);

                    int posCat = 1;
                    foreach (var c in gruppo)
                    {
                        int posAss = posizioniAssolute[c];
                        ScriviRigaAtleta(sb, c, posAss, posCat, true, hasSoc, hasEnte, unknownRider, colLap);
                        posCat++;
                    }
                    sb.Append("</tbody></table>");
                }
            }
            else
            {
                ScriviIntestazioneTabella(sb, false, hasSoc, hasEnte, colAss, colPos, colDor, colName, colCat, colTeam, colEnte, colTime);
                int pos = 1;
                foreach (var c in classificaAssoluta)
                {
                    ScriviRigaAtleta(sb, c, pos, pos, false, hasSoc, hasEnte, unknownRider, colLap);
                    pos++;
                }
                sb.Append("</tbody></table>");
            }

            sb.Append("</body></html>");
            return sb.ToString();
        }

        /// <summary>
        /// Appends a localized HTML table header (&lt;thead&gt;) to the report buffer, dynamically
        /// including or omitting Category, Team/Club, and Federation columns based on data presence.
        /// </summary>
        private static void ScriviIntestazioneTabella(
            StringBuilder sb,
            bool showCategory,
            bool showTeam,
            bool showEnte,
            string colAss,
            string colPos,
            string colDor,
            string colName,
            string colCat,
            string colTeam,
            string colEnte,
            string colTime)
        {
            sb.Append("<table class='results'><thead><tr>");
            if (showCategory)
            {
                sb.Append($"<th style='width:6%;' class='num'>{colAss}</th>");
                sb.Append($"<th style='width:6%;' class='num'>{colPos}</th>");
            }
            else
            {
                sb.Append($"<th style='width:7%;' class='num'>{colPos}</th>");
            }
            sb.Append($"<th style='width:7%;' class='num'>{colDor}</th>");
            sb.Append($"<th>{colName}</th>");
            if (showCategory)
                sb.Append($"<th style='width:10%;'>{colCat}</th>");
            if (showTeam)
                sb.Append($"<th style='width:26%;'>{colTeam}</th>");
            if (showEnte)
                sb.Append($"<th style='width:9%;'>{colEnte}</th>");
            sb.Append($"<th style='width:16%;'>{colTime}</th>");
            sb.Append("</tr></thead><tbody>");
        }

        /// <summary>
        /// Appends a single ranked athlete row (&lt;tr&gt;) to the HTML standings table, formatting
        /// total elapsed time alongside completed lap count.
        /// </summary>
        private static void ScriviRigaAtleta(
            StringBuilder sb,
            Ciclista c,
            int posAss,
            int posCat,
            bool showCategory,
            bool showTeam,
            bool showEnte,
            string unknownRider,
            string colLap)
        {
            string ente = c.NonInElenco ? "" : WebUtility.HtmlEncode(c.Ente);
            string nomeAtleta = c.NonInElenco ? unknownRider : WebUtility.HtmlEncode(c.Nome);
            string tempoFormattato = $"{c.TempoTotale:hh\\:mm\\:ss} &nbsp;({c.Giri}\u00B0 {colLap})";

            sb.Append("<tr>");
            if (showCategory)
            {
                sb.Append($"<td class='num'>{posAss}</td>");
                sb.Append($"<td class='num'><strong>{posCat}</strong></td>");
            }
            else
            {
                sb.Append($"<td class='num'><strong>{posAss}</strong></td>");
            }
            sb.Append($"<td class='num'>{c.Pettorale}</td>");
            sb.Append($"<td>{nomeAtleta}</td>");
            if (showCategory)
                sb.Append($"<td>{WebUtility.HtmlEncode(c.Categoria)}</td>");
            if (showTeam)
                sb.Append($"<td>{WebUtility.HtmlEncode(c.NomeSocieta)}</td>");
            if (showEnte)
                sb.Append($"<td>{ente}</td>");
            sb.Append($"<td>{tempoFormattato}</td>");
            sb.Append("</tr>");
        }
    }
}
// ============================================================================
// File: ExcelHelper.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Handles start-list import delegation and native, dependency-free generation
//   of styled OpenXML (.xlsx) spreadsheets containing race header metadata,
//   overall and category rankings, and individual lap split columns.
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;

namespace CicloRank
{
    /// <summary>
    /// Provides helper methods for importing athlete start lists and generating native,
    /// styled OpenXML (.xlsx) classification spreadsheets without external dependencies.
    /// </summary>
    public static class ExcelHelper
    {
        // OpenXML SpreadsheetML main schema namespace URI.
        private const string MainSpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        // UTF-8 encoding without Byte Order Mark (BOM) required for valid OpenXML ZIP parts.
        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

        /// <summary>
        /// Validates the target start-list file path and delegates parsing to
        /// <see cref="ExcelColumnMapper.LeggiCiclistiDaFile"/>, enriching any error with the filename.
        /// </summary>
        /// <param name="filePath">Full path to the start-list file.</param>
        /// <returns>A list of imported <see cref="Ciclista"/> objects, or an empty list if missing.</returns>
        public static List<Ciclista> LeggiFileAtleti(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return new List<Ciclista>();
            }

            try
            {
                return ExcelColumnMapper.LeggiCiclistiDaFile(filePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"'{Path.GetFileName(filePath)}': {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Exports overall and category-grouped race standings to a styled .xlsx workbook.
        /// Builds the complete OpenXML ZIP archive structure ([Content_Types].xml, relationships,
        /// workbook, styles, and worksheet XML) with auto-fitted column widths and lap splits.
        /// </summary>
        /// <param name="filePath">Destination .xlsx file path.</param>
        /// <param name="tuttiAtleti">Complete list of race participants.</param>
        /// <param name="infoGara">Event header metadata to display at the top of the sheet.</param>
        public static void EsportaXlsx(string filePath, List<Ciclista> tuttiAtleti, DatiGara infoGara)
        {
            // Rank all active riders (at least 1 completed lap) by descending laps, then ascending total time.
            var classificaAssoluta = tuttiAtleti
                .Where(c => c.Giri > 0)
                .OrderByDescending(c => c.Giri)
                .ThenBy(c => c.TempoTotale)
                .ToList();

            var posizioniAssolute = new Dictionary<Ciclista, int>(classificaAssoluta.Count);
            int maxGiri = 0;
            for (int i = 0; i < classificaAssoluta.Count; i++)
            {
                var c = classificaAssoluta[i];
                posizioniAssolute[c] = i + 1;
                if (c.Giri > maxGiri) maxGiri = c.Giri;
            }

            string lapLabel = LinguaManager.T("RptColLap");
            var headers = new List<string>(8 + maxGiri)
            {
                LinguaManager.T("RptColAss"),
                LinguaManager.T("RptColPos"),
                LinguaManager.T("RptColDor"),
                LinguaManager.T("RptColName"),
                LinguaManager.T("RptColCat"),
                LinguaManager.T("RptColTeam"),
                LinguaManager.T("RptColEnte")
            };
            for (int g = 1; g <= maxGiri; g++)
            {
                headers.Add($"{lapLabel} {g}");
            }
            headers.Add(LinguaManager.T("RptColTotalTime"));

            var colWidths = new int[headers.Count];
            for (int c = 0; c < headers.Count; c++)
            {
                int minWidth = c switch { 3 => 28, 5 => 28, _ => 14 };
                colWidths[c] = Math.Max(minWidth, headers[c].Length + 4);
            }

            var dataRows = new List<string[]>(classificaAssoluta.Count);

            // Group classified athletes alphabetically by category and compute per-category rank positions.
            var gruppiCategoria = classificaAssoluta
                .GroupBy(c => c.Categoria)
                .OrderBy(g => g.Key);

            foreach (var gruppo in gruppiCategoria)
            {
                int posCat = 1;
                foreach (var c in gruppo)
                {
                    int posAss = posizioniAssolute[c];
                    var row = new string[headers.Count];
                    row[0] = posAss.ToString();
                    row[1] = posCat.ToString();
                    row[2] = c.Pettorale.ToString();
                    row[3] = c.NonInElenco ? LinguaManager.T("SconosciutoNome") : c.Nome;
                    row[4] = c.Categoria;
                    row[5] = c.NomeSocieta;
                    row[6] = c.NonInElenco ? "" : c.Ente;

                    for (int g = 0; g < maxGiri; g++)
                    {
                        row[7 + g] = g < c.Passaggi.Count
                            ? c.Passaggi[g].ToString(@"hh\:mm\:ss")
                            : "";
                    }

                    row[headers.Count - 1] = c.TempoTotale.ToString(@"hh\:mm\:ss");

                    for (int col = 0; col < row.Length; col++)
                    {
                        int needed = row[col].Length + 3;
                        if (needed > colWidths[col])
                        {
                            colWidths[col] = Math.Min(needed, 45);
                        }
                    }

                    dataRows.Add(row);
                    posCat++;
                }
            }

            // Assemble a valid OpenXML (.xlsx) ZIP archive containing all required XML parts and relationships.
            using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create);

            WriteZipEntry(zip, "[Content_Types].xml", GetContentTypesXml());
            WriteZipEntry(zip, "_rels/.rels", GetRootRelsXml());
            WriteZipEntry(zip, "xl/_rels/workbook.xml.rels", GetWorkbookRelsXml());
            WriteZipEntry(zip, "xl/workbook.xml", GetWorkbookXml());
            WriteZipEntry(zip, "xl/styles.xml", GetStylesXml());
            WriteZipEntry(zip, "xl/worksheets/sheet1.xml", GetSheetXml(infoGara, headers, dataRows, colWidths));
        }

        /// <summary>
        /// Creates an entry inside the target .xlsx <see cref="ZipArchive"/> and writes UTF-8 XML content.
        /// </summary>
        private static void WriteZipEntry(ZipArchive zip, string entryName, string content)
        {
            var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
            using var writer = new StreamWriter(entry.Open(), Utf8NoBom);
            writer.Write(content);
        }

        /// <summary>
        /// Generates the [Content_Types].xml OpenXML manifest declaring workbook, worksheet, and style MIME types.
        /// </summary>
        private static string GetContentTypesXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                   "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                   "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                   "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
                   "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n" +
                   "  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n" +
                   "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n" +
                   "</Types>";
        }

        /// <summary>
        /// Generates the root _rels/.rels XML relationship file pointing to xl/workbook.xml.
        /// </summary>
        private static string GetRootRelsXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                   "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n" +
                   "</Relationships>";
        }

        /// <summary>
        /// Generates xl/_rels/workbook.xml.rels linking the workbook to sheet1.xml and styles.xml.
        /// </summary>
        private static string GetWorkbookRelsXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                   "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n" +
                   "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n" +
                   "</Relationships>";
        }

        /// <summary>
        /// Generates xl/workbook.xml defining the single "CicloRank" worksheet.
        /// </summary>
        private static string GetWorkbookXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
                   "  <sheets>\n" +
                   "    <sheet name=\"CicloRank\" sheetId=\"1\" r:id=\"rId1\"/>\n" +
                   "  </sheets>\n" +
                   "</workbook>";
        }

        /// <summary>
        /// Generates xl/styles.xml defining Segoe UI fonts, emerald green header fill (#2E7D4F),
        /// zebra-striped alternating row shading (#F2F4F7), thin cell borders, and alignment styles.
        /// </summary>
        private static string GetStylesXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                   "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
                   "  <fonts count=\"3\">\n" +
                   "    <font><sz val=\"11\"/><name val=\"Segoe UI\"/></font>\n" +
                   "    <font><b/><sz val=\"11\"/><color rgb=\"FFFFFFFF\"/><name val=\"Segoe UI\"/></font>\n" +
                   "    <font><b/><sz val=\"11\"/><color rgb=\"FF000000\"/><name val=\"Segoe UI\"/></font>\n" +
                   "  </fonts>\n" +
                   "  <fills count=\"4\">\n" +
                   "    <fill><patternFill patternType=\"none\"/></fill>\n" +
                   "    <fill><patternFill patternType=\"gray125\"/></fill>\n" +
                   "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF2E7D4F\"/><bgColor indexed=\"64\"/></patternFill></fill>\n" +
                   "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF2F4F7\"/><bgColor indexed=\"64\"/></patternFill></fill>\n" +
                   "  </fills>\n" +
                   "  <borders count=\"2\">\n" +
                   "    <border><left/><right/><top/><bottom/></border>\n" +
                   "    <border>\n" +
                   "      <left style=\"thin\"><color rgb=\"FFD9D9D9\"/></left>\n" +
                   "      <right style=\"thin\"><color rgb=\"FFD9D9D9\"/></right>\n" +
                   "      <top style=\"thin\"><color rgb=\"FFD9D9D9\"/></top>\n" +
                   "      <bottom style=\"thin\"><color rgb=\"FFD9D9D9\"/></bottom>\n" +
                   "    </border>\n" +
                   "  </borders>\n" +
                   "  <cellStyleXfs count=\"1\">\n" +
                   "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n" +
                   "  </cellStyleXfs>\n" +
                   "  <cellXfs count=\"8\">\n" +
                   "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n" +
                   "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\n" +
                   "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                   "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\n" +
                   "    <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                   "    <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\n" +
                   "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                   "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\n" +
                   "  </cellXfs>\n" +
                   "</styleSheet>";
        }

        /// <summary>
        /// Generates xl/worksheets/sheet1.xml containing column width definitions, race header
        /// metadata rows, styled column headers, and zebra-striped participant result rows.
        /// </summary>
        private static string GetSheetXml(DatiGara info, List<string> headers, List<string[]> rows, int[] colWidths)
        {
            int estimatedCapacity = 4096 + (rows.Count + 2) * headers.Count * 95;
            var sb = new StringBuilder(estimatedCapacity);

            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
            sb.Append($"<worksheet xmlns=\"{MainSpreadsheetNs}\">\n");

            sb.Append("<cols>\n");
            for (int i = 0; i < headers.Count; i++)
            {
                int colNum = i + 1;
                sb.Append($"  <col min=\"{colNum}\" max=\"{colNum}\" width=\"{colWidths[i]}\" customWidth=\"1\"/>\n");
            }
            sb.Append("</cols>\n");

            sb.Append("<sheetData>\n");

            int rIdx = 1;
            void AddInfoRow(string label, string val)
            {
                sb.Append($"  <row r=\"{rIdx}\">\n");
                sb.Append($"    <c r=\"A{rIdx}\" t=\"inlineStr\" s=\"4\"><is><t xml:space=\"preserve\">{EscapeXml(label)}</t></is></c>\n");
                sb.Append($"    <c r=\"B{rIdx}\" t=\"inlineStr\" s=\"2\"><is><t xml:space=\"preserve\">{EscapeXml(val)}</t></is></c>\n");
                sb.Append("  </row>\n");
                rIdx++;
            }

            AddInfoRow(LinguaManager.T("RptCommittee"), info.ComitatoProvinciale);
            AddInfoRow(LinguaManager.T("RptOrg"), info.Organizzazione);
            AddInfoRow(LinguaManager.T("RptDenom"), info.Denominazione);
            AddInfoRow(LinguaManager.T("RptLoc"), info.Localita);
            AddInfoRow(LinguaManager.T("RptDate"), info.DataGara);
            if (!string.IsNullOrWhiteSpace(info.Sport))
            {
                AddInfoRow(LinguaManager.T("RptSport"), LinguaManager.FormattaNomeSport(info.Sport));
            }
            rIdx++;

            var colNames = new string[headers.Count];
            for (int c = 0; c < headers.Count; c++)
            {
                colNames[c] = GetColumnName(c + 1);
            }

            sb.Append($"  <row r=\"{rIdx}\">\n");
            for (int c = 0; c < headers.Count; c++)
            {
                sb.Append($"    <c r=\"{colNames[c]}{rIdx}\" t=\"inlineStr\" s=\"1\"><is><t xml:space=\"preserve\">{EscapeXml(headers[c])}</t></is></c>\n");
            }
            sb.Append("  </row>\n");
            rIdx++;

            for (int i = 0; i < rows.Count; i++)
            {
                bool isZebra = (i % 2 == 1);
                string[] currentRow = rows[i];
                sb.Append($"  <row r=\"{rIdx}\">\n");
                for (int c = 0; c < currentRow.Length; c++)
                {
                    bool isLeftAlign = (c == 3 || c == 4 || c == 5);
                    int styleId = isZebra ? (isLeftAlign ? 6 : 7) : (isLeftAlign ? 2 : 3);
                    sb.Append($"    <c r=\"{colNames[c]}{rIdx}\" t=\"inlineStr\" s=\"{styleId}\"><is><t xml:space=\"preserve\">{EscapeXml(currentRow[c])}</t></is></c>\n");
                }
                sb.Append("  </row>\n");
                rIdx++;
            }

            sb.Append("</sheetData>\n");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        /// <summary>
        /// Converts a 1-based column number into an Excel alphabetical column identifier (e.g., 1 -> "A", 27 -> "AA").
        /// </summary>
        private static string GetColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = (char)(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        /// <summary>
        /// Escapes XML special characters (&amp;, &lt;, &gt;, &quot;, &apos;) to guarantee well-formed OpenXML strings.
        /// </summary>
        private static string EscapeXml(string val)
        {
            if (string.IsNullOrEmpty(val)) return "";
            return SecurityElement.Escape(val) ?? "";
        }
    }
}
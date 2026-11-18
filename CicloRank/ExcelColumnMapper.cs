// ============================================================================
// File: ExcelColumnMapper.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Provides intelligent, zero-dependency start-list parsing and multilingual
//   column mapping across 11 languages (IT, EN, ES, FR, DE, PT, NL, PL, ZH, RU, JA).
//   Reads .xlsx OpenXML archives natively via ZipArchive + XDocument as well as
//   .csv and .txt files with automatic character encoding and delimiter detection.
// ============================================================================

#nullable disable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace CicloRank
{
    /// <summary>
    /// Intermediate data transfer object holding raw string values extracted from a single
    /// row of an imported start-list file (.xlsx, .csv, or .txt) prior to validation.
    /// </summary>
    public class RigaPartecipanteImportata
    {
        /// <summary>Raw bib or race plate number string from the source file.</summary>
        public string Pettorale { get; set; }
        /// <summary>Raw first/given name (or full name when surname is not separate).</summary>
        public string Nome { get; set; }
        /// <summary>Raw family name / surname when provided in a dedicated column.</summary>
        public string Cognome { get; set; }
        /// <summary>Combined display name constructed from surname and given name.</summary>
        public string NomeCompleto { get; set; }
        /// <summary>Raw age group or racing category string.</summary>
        public string Categoria { get; set; }
        /// <summary>Raw club, team, or sports association name.</summary>
        public string Societa { get; set; }
        /// <summary>Raw federation or governing body name.</summary>
        public string Ente { get; set; }
        /// <summary>Raw birth year string if present in the source table.</summary>
        public string Anno { get; set; }
        /// <summary>Raw license or membership card number if present in the source table.</summary>
        public string Tessera { get; set; }
    }

    /// <summary>
    /// Multilingual start-list reader and intelligent column mapper supporting 11 languages.
    /// Parses .xlsx workbooks directly via OpenXML ZIP inspection (requiring no Office or
    /// third-party NuGet packages) as well as delimited .csv/.txt files with automatic
    /// encoding and separator detection.
    /// </summary>
    public static class ExcelColumnMapper
    {
        // Normalized column header aliases for bib/plate/race numbers across all 11 supported languages.
        private static readonly HashSet<string> AliasPettorale = BuildSet(new[]
        {
            // Italian
            "Pettorale", "Numero di gara", "Numero pettorale", "Dorsale", "Numero", "Pett.", "Pett", "N. gara", "N\u00B0 gara", "Num. gara", "Num", "N.", "N\u00B0",
            // English
            "Bib", "Bib number", "Race number", "Start number", "Plate", "Plate number", "Number", "Bib no", "Bib #", "Race no", "Start no",
            // Spanish
            "Dorsal", "N\u00FAmero de dorsal", "N\u00FAmero de corredor", "N\u00FAmero de participaci\u00F3n", "N.\u00BA carrera", "N\u00B0 carrera", "No. carrera", "N. carrera", "Dor", "Dors.", "N\u00FAmero", "Numero",
            // French
            "Dossard", "Num\u00E9ro de dossard", "Num\u00E9ro de course", "Doss.", "Doss", "Num\u00E9ro", "N\u00B0 dossard",
            // German
            "Startnummer", "Wettkampfnummer", "Nummer", "R\u00FCckennummer", "Start-Nr", "Startnr", "Nr.", "Nr",
            // Portuguese
            "N\u00FAmero de peito", "N\u00FAmero de inscri\u00E7\u00E3o", "Peitoral", "N.\u00BA peito", "N.\u00BA inscri\u00E7\u00E3o",
            // Dutch
            "Borstnummer", "Wedstrijdnummer", "Rugnummer", "Stuurbordnummer", "Rugnr.", "Rugnr",
            // Polish
            "Numer startowy", "Numer zawodnika", "Numer", "Nr startowy",
            // Chinese
            "\u53F7\u7801\u5E03", "\u53C2\u8D5B\u53F7\u7801", "\u6BD4\u8D5B\u53F7\u7801", "\u9009\u624B\u7F16\u53F7", "\u7F16\u53F7", "\u53F7\u7801", "\u8F66\u53F7",
            // Russian
            "\u0421\u0442\u0430\u0440\u0442\u043E\u0432\u044B\u0439 \u2116", "\u0421\u0442\u0430\u0440\u0442\u043E\u0432\u044B\u0439 \u043D\u043E\u043C\u0435\u0440", "\u041D\u0430\u0433\u0440\u0443\u0434\u043D\u044B\u0439 \u043D\u043E\u043C\u0435\u0440", "\u041D\u043E\u043C\u0435\u0440 \u0443\u0447\u0430\u0441\u0442\u043D\u0438\u043A\u0430", "\u0411\u0438\u0431", "\u041D\u043E\u043C\u0435\u0440", "\u041D\u043E\u043C.", "\u041D\u043E\u043C", "\u2116",
            // Japanese
            "\u30BC\u30C3\u30B1\u30F3", "\u30BC\u30C3\u30B1\u30F3\u756A\u53F7", "\u30CA\u30F3\u30D0\u30FC", "\u30CA\u30F3\u30D0\u30FC\u30AB\u30FC\u30C9", "\u30D3\u30D6\u30B9", "\u7AF6\u6280\u8005\u756A\u53F7", "\u51FA\u8D70\u756A\u53F7", "\u756A\u53F7"
        });

        // Normalized column header aliases for athlete first name or full name across all 11 languages.
        private static readonly HashSet<string> AliasNome = BuildSet(new[]
        {
            // Italian
            "Nome", "Nome proprio", "Nome atleta", "Nome completo", "Nome e Cognome", "Cognome e Nome", "Cognome Nome", "Nome Cognome",
            "Nominativo", "Atleta", "Partecipante", "Concorrente", "Ciclista", "Corridore", "Iscritto", "Anagrafica",
            // English
            "First name", "Given name", "Forename", "Name", "Full name", "First and last name", "Last and first name",
            "Athlete", "Athlete name", "Rider", "Rider name", "Cyclist", "Racer", "Runner", "Competitor", "Participant", "Entrant", "Starter",
            // Spanish
            "Nombre", "Nombre de pila", "Nombre del atleta", "Nombre del corredor", "Nombre completo", "Nombre y apellidos", "Apellidos y nombre",
            "Corredor", "Ciclista", "Deportista", "Participante", "Inscrito",
            // French
            "Pr\u00E9nom", "Nom", "Nom complet", "Nom et pr\u00E9nom", "Pr\u00E9nom et nom", "Nom pr\u00E9nom",
            "Athl\u00E8te", "Coureur", "Cycliste", "Participant", "Engag\u00E9", "Partant",
            // German
            "Vorname", "Rufname", "Name", "Vollst\u00E4ndiger Name", "Vor- und Nachname", "Vorname Nachname", "Nachname Vorname",
            "Athlet", "Teilnehmer", "Fahrer", "Radfahrer", "Rennfahrer", "L\u00E4ufer", "Starter",
            // Portuguese
            "Primeiro nome", "Nome pr\u00F3prio", "Nome do atleta", "Nome completo", "Nome e apelido", "Nome e sobrenome",
            "Ciclista", "Corredor", "Atleta", "Participante", "Inscrito",
            // Dutch
            "Voornaam", "Roepnaam", "Naam", "Volledige naam", "Voor- en achternaam", "Voornaam achternaam",
            "Atleet", "Deelnemer", "Renner", "Wielrenner",
            // Polish
            "Imi\u0119", "Pierwsze imi\u0119", "Imi\u0119 i nazwisko", "Nazwisko i imi\u0119",
            "Zawodnik", "Kolarz", "Uczestnik",
            // Chinese
            "\u540D\u5B57", "\u540D", "\u59D3\u540D", "\u9009\u624B\u59D3\u540D", "\u8FD0\u52A8\u5458\u59D3\u540D", "\u53C2\u8D5B\u8005\u59D3\u540D",
            "\u8FD0\u52A8\u5458", "\u53C2\u8D5B\u8005", "\u9009\u624B", "\u9A91\u624B", "\u8F66\u624B", "\u81EA\u884C\u8F66\u624B",
            // Russian
            "\u0424\u0430\u043C\u0438\u043B\u0438\u044F \u0418\u043C\u044F", "\u0418\u043C\u044F \u0424\u0430\u043C\u0438\u043B\u0438\u044F",
            "\u0424\u0430\u043C\u0438\u043B\u0438\u044F \u0438 \u0438\u043C\u044F", "\u0418\u043C\u044F \u0438 \u0444\u0430\u043C\u0438\u043B\u0438\u044F",
            "\u0418\u043C\u044F", "\u041B\u0438\u0447\u043D\u043E\u0435 \u0438\u043C\u044F", "\u0424\u0418\u041E", "\u0424.\u0418.\u041E.",
            "\u0423\u0447\u0430\u0441\u0442\u043D\u0438\u043A", "\u0421\u043F\u043E\u0440\u0442\u0441\u043C\u0435\u043D", "\u0413\u043E\u043D\u0449\u0438\u043A", "\u0412\u0435\u043B\u043E\u0441\u0438\u043F\u0435\u0434\u0438\u0441\u0442",
            // Japanese
            "\u4E0B\u306E\u540D\u524D", "\u30D5\u30A1\u30FC\u30B9\u30C8\u30CD\u30FC\u30E0", "\u540D\u524D", "\u6C0F\u540D",
            "\u9078\u624B\u540D", "\u7AF6\u6280\u8005\u540D", "\u30E9\u30A4\u30C0\u30FC", "\u9078\u624B", "\u53C2\u52A0\u8005"
        });

        // Normalized column header aliases for athlete last name / surname across all 11 languages.
        private static readonly HashSet<string> AliasCognome = BuildSet(new[]
        {
            "Cognome", "Last name", "Surname", "Family name", "Apellido", "Apellidos",
            "Nom de famille", "Nachname", "Familienname", "Sobrenome", "Apelido",
            "Achternaam", "Familienaam", "Nazwisko", "\u0424\u0430\u043C\u0438\u043B\u0438\u044F", "\u59D3", "\u82D7\u5B57"
        });

        // Normalized column header aliases for racing category / age bracket across all 11 languages.
        private static readonly HashSet<string> AliasCategoria = BuildSet(new[]
        {
            // Italian
            "Categoria", "Fascia d'et\u00E0", "Fascia d et\u00E0", "Fascia eta", "Fascia", "Gruppo d'et\u00E0", "Gruppo d et\u00E0", "Gruppo eta", "Classe", "Divisione", "Cat.", "Cat", "Categ",
            // English
            "Category", "Age group", "AG", "Division", "Class", "Bracket",
            // Spanish
            "Categor\u00EDa", "Grupo de edad", "Divisi\u00F3n", "Rama",
            // French
            "Cat\u00E9gorie", "Cat\u00E9gorie d'\u00E2ge", "Groupe d'\u00E2ge",
            // German
            "Altersklasse", "AK", "Klasse", "Wertungsklasse", "Kategorie",
            // Portuguese
            "Escal\u00E3o", "Faixa et\u00E1ria", "Divis\u00E3o",
            // Dutch
            "Categorie", "Leeftijdscategorie", "Leeftijdsgroep", "Divisie",
            // Polish
            "Kategoria", "Kategoria wiekowa", "Grupa wiekowa", "Klasa", "Dywizja", "Kat.", "Kat",
            // Chinese
            "\u7EC4\u522B", "\u5E74\u9F84\u7EC4", "\u7EA7\u522B", "\u5206\u7C7B", "\u7C7B\u522B",
            // Russian
            "\u041A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u044F", "\u0412\u043E\u0437\u0440\u0430\u0441\u0442\u043D\u0430\u044F \u0433\u0440\u0443\u043F\u043F\u0430", "\u0412\u043E\u0437\u0440\u0430\u0441\u0442\u043D\u0430\u044F \u043A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u044F", "\u041A\u043B\u0430\u0441\u0441", "\u0414\u0438\u0432\u0438\u0437\u0438\u043E\u043D", "\u041A\u0430\u0442", "\u041A\u0430\u0442.",
            // Japanese
            "\u30AB\u30C6\u30B4\u30EA\u30FC", "\u30AB\u30C6\u30B4\u30EA", "\u90E8\u9580", "\u5E74\u4EE3\u5225", "\u30AF\u30E9\u30B9", "\u7A2E\u5225"
        });

        // Normalized column header aliases for sports club / team name across all 11 languages.
        private static readonly HashSet<string> AliasSocieta = BuildSet(new[]
        {
            // Italian
            "Societ\u00E0", "Societa", "Nome societ\u00E0", "Nome societa", "Denominazione societ\u00E0", "Squadra", "Club", "Associazione sportiva", "ASD", "A.S.D.", "SSD", "S.S.D.", "Team", "Gruppo sportivo", "GS", "G.S.",
            // English
            "Athletic club", "Running club", "Cycling club", "Club name", "Team name", "Squad", "Sports club",
            // Spanish
            "Equipo", "Sociedad deportiva", "Agrupaci\u00F3n", "Pe\u00F1a", "Nombre del club", "Club deportivo",
            // French
            "\u00C9quipe", "Nom du club", "Association sportive", "Soci\u00E9t\u00E9 sportive", "Club sportif",
            // German
            "Verein", "Vereinsname", "Sportverein", "Mannschaft", "Startgemeinschaft", "SG", "Klub",
            // Portuguese
            "Clube", "Nome do clube", "Equipa", "Equipe", "Associa\u00E7\u00E3o desportiva", "Associa\u00E7\u00E3o esportiva", "Agremia\u00E7\u00E3o", "Time",
            // Dutch
            "Vereniging", "Sportvereniging", "Ploeg",
            // Polish
            "Klub", "Nazwa klubu", "Klub sportowy", "Dru\u017Cyna", "Zesp\u00F3\u0142", "Stowarzyszenie",
            // Chinese
            "\u4FF1\u4E50\u90E8", "\u56E2\u961F", "\u8DD1\u56E2", "\u534F\u4F1A", "\u961F\u4F0D", "\u8F66\u961F", "\u5355\u4F4D",
            // Russian
            "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u043A\u043B\u0443\u0431\u0430", "\u041A\u043B\u0443\u0431", "\u0421\u043F\u043E\u0440\u0442\u0438\u0432\u043D\u044B\u0439 \u043A\u043B\u0443\u0431", "\u041A\u043E\u043C\u0430\u043D\u0434\u0430", "\u0421\u043F\u043E\u0440\u0442\u0438\u0432\u043D\u043E\u0435 \u043E\u0431\u0449\u0435\u0441\u0442\u0432\u043E", "\u041A\u043E\u043C\u0430\u043D\u0434\u0430 / \u041A\u043B\u0443\u0431", "\u041E\u0431\u0449\u0435\u0441\u0442\u0432\u043E",
            // Japanese
            "\u6240\u5C5E", "\u30AF\u30E9\u30D6", "\u30C1\u30FC\u30E0", "\u30C1\u30FC\u30E0\u540D", "\u56E3\u4F53", "\u540C\u597D\u4F1A", "\u6240\u5C5E\u30C1\u30FC\u30E0"
        });

        // Normalized column header aliases for sports federation / governing body across all 11 languages.
        private static readonly HashSet<string> AliasEnte = BuildSet(new[]
        {
            // Italian
            "Ente", "Federazione", "Ente di promozione sportiva", "EPS", "E.P.S.", "Organismo", "Affiliazione", "Federazione sportiva",
            // English
            "Federation", "Governing body", "Sanctioning body", "Organization", "Organisation", "Affiliation", "Association", "Body",
            // Spanish
            "Federaci\u00F3n", "Entidad", "Asociaci\u00F3n", "Ente deportivo", "Afiliaci\u00F3n",
            // French
            "F\u00E9d\u00E9ration", "Instance", "Ligue", "Organisme",
            // German
            "Verband", "Sportverband", "Dachverband", "Fachverband",
            // Portuguese
            "Federa\u00E7\u00E3o", "\u00D3rg\u00E3o", "Confedera\u00E7\u00E3o", "Associa\u00E7\u00E3o", "Entidade", "Filia\u00E7\u00E3o",
            // Dutch
            "Bond", "Sportbond", "Federatie", "Organisatie", "Overkoepelend orgaan",
            // Polish
            "Zwi\u0105zek", "Federacja", "Zwi\u0105zek sportowy", "Organizacja", "Podmiot",
            // Chinese
            "\u8054\u5408\u4F1A", "\u7BA1\u7406\u673A\u6784", "\u7EC4\u7EC7\u673A\u6784", "\u534F\u4F1A\u673A\u6784", "\u603B\u4F1A",
            // Russian
            "\u0424\u0435\u0434\u0435\u0440\u0430\u0446\u0438\u044F", "\u0421\u043F\u043E\u0440\u0442\u0438\u0432\u043D\u0430\u044F \u0444\u0435\u0434\u0435\u0440\u0430\u0446\u0438\u044F", "\u041E\u0440\u0433\u0430\u043D\u0438\u0437\u0430\u0446\u0438\u044F", "\u0412\u0435\u0434\u043E\u043C\u0441\u0442\u0432\u043E", "\u0424\u0435\u0434.", "\u0424\u0435\u0434",
            // Japanese
            "\u9023\u76DF", "\u5354\u4F1A", "\u7AF6\u6280\u56E3\u4F53", "\u7D71\u62EC\u56E3\u4F53", "\u52A0\u76DF\u56E3\u4F53"
        });

        // Substring fallback tokens used to recognize custom composite athlete-name headers.
        private static readonly string[] TokenNomeAtleta = new[]
        {
            "nome", "cognome", "nominativ", "atlet", "ciclist", "corridor", "partecip", "concorrent", "iscrit",
            "name", "surname", "rider", "cyclist", "racer", "competitor", "participant", "entrant", "starter",
            "nombre", "apellido", "corredor", "deportist",
            "prenom", "nom", "coureur", "engage", "partant",
            "vorname", "nachname", "fahrer", "teilnehmer",
            "apelido", "sobrenome",
            "naam", "achternaam", "renner", "deelnemer",
            "imie", "nazwisko", "zawodnik", "kolarz", "uczestnik",
            "\u0438\u043C\u044F", "\u0444\u0430\u043C\u0438\u043B", "\u0444\u0438\u043E", "\u0443\u0447\u0430\u0441\u0442\u043D\u0438\u043A", "\u0441\u043F\u043E\u0440\u0442\u0441\u043C\u0435\u043D", "\u0433\u043E\u043D\u0449\u0438\u043A", "\u0432\u0435\u043B\u043E\u0441\u0438\u043F\u0435\u0434\u0438\u0441\u0442",
            "\u59D3\u540D", "\u540D\u5B57", "\u9009\u624B", "\u8FD0\u52A8\u5458", "\u53C2\u8D5B", "\u9A91\u624B",
            "\u6C0F\u540D", "\u540D\u524D", "\u9078\u624B", "\u7AF6\u6280\u8005", "\u30E9\u30A4\u30C0\u30FC", "\u53C2\u52A0\u8005"
        };

        /// <summary>
        /// Builds a case-insensitive lookup set containing both accented and accent-stripped
        /// normalized forms of every header alias.
        /// </summary>
        /// <param name="items">Raw header alias strings.</param>
        /// <returns>A normalized <see cref="HashSet{T}"/> for fast O(1) header matching.</returns>
        private static HashSet<string> BuildSet(IEnumerable<string> items)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in items)
            {
                string n1 = NormalizzaHeader(item, false);
                string n2 = NormalizzaHeader(item, true);
                if (!string.IsNullOrEmpty(n1)) set.Add(n1);
                if (!string.IsNullOrEmpty(n2)) set.Add(n2);
            }
            return set;
        }

        /// <summary>
        /// Normalizes a column header or worksheet name by trimming, lowercasing, removing
        /// ordinal symbols, optionally stripping diacritics via Unicode FormD decomposition,
        /// and collapsing punctuation into single spaces.
        /// </summary>
        /// <param name="testo">The raw header text to normalize.</param>
        /// <param name="rimuoviAccenti">True to strip diacritical marks (accents); false to preserve them.</param>
        /// <returns>A clean, normalized lowercase string suitable for dictionary lookup.</returns>
        public static string NormalizzaHeader(string testo, bool rimuoviAccenti)
        {
            if (string.IsNullOrWhiteSpace(testo)) return string.Empty;
            string s = testo.Trim().ToLowerInvariant()
                .Replace('\u00BA', ' ')
                .Replace('\u00B0', ' ');

            if (rimuoviAccenti)
            {
                string formD = s.Normalize(NormalizationForm.FormD);
                var sbAcc = new StringBuilder(formD.Length);
                foreach (char ch in formD)
                {
                    if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                        sbAcc.Append(ch);
                }
                s = sbAcc.ToString().Normalize(NormalizationForm.FormC);
            }

            var sb = new StringBuilder(s.Length);
            bool prevSpace = false;
            foreach (char c in s)
            {
                char ch = c;
                if (ch == '\u0142') ch = 'l';
                if (char.IsLetterOrDigit(ch))
                {
                    sb.Append(ch);
                    prevSpace = false;
                }
                else if (!prevSpace && sb.Length > 0)
                {
                    sb.Append(' ');
                    prevSpace = true;
                }
            }
            return sb.ToString().Trim();
        }

        /// <summary>
        /// Maps a raw spreadsheet column header to one of the canonical internal field names:
        /// "Societa", "Pettorale", "Cognome", "Nome", "Categoria", or "Ente".
        /// </summary>
        /// <param name="header">The raw column header text from the spreadsheet.</param>
        /// <returns>The canonical field identifier, or <c>null</c> if the column is unrecognized.</returns>
        public static string MappaIntestazioneColonna(string header)
        {
            if (string.IsNullOrWhiteSpace(header)) return null;
            string n1 = NormalizzaHeader(header, false);
            string n2 = NormalizzaHeader(header, true);
            if (string.IsNullOrEmpty(n1)) return null;

            // Check Societa (Club/Team) first so "Club Name" headers are not mistaken for an athlete's personal name
            if (AliasSocieta.Contains(n1) || AliasSocieta.Contains(n2)) return "Societa";
            if (AliasPettorale.Contains(n1) || AliasPettorale.Contains(n2)) return "Pettorale";
            if (AliasCognome.Contains(n1) || AliasCognome.Contains(n2)) return "Cognome";
            if (AliasNome.Contains(n1) || AliasNome.Contains(n2)) return "Nome";
            if (AliasCategoria.Contains(n1) || AliasCategoria.Contains(n2)) return "Categoria";
            if (AliasEnte.Contains(n1) || AliasEnte.Contains(n2)) return "Ente";

            string[] parti = header.Split(new[] { '/', '|', '\\', '(', ')', '[', ']', '-', ':', '&', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parti.Length > 1)
            {
                foreach (var p in parti)
                {
                    string mappedPart = MappaIntestazioneColonna(p);
                    if (mappedPart != null) return mappedPart;
                }
            }

            // Fallback heuristic for custom or composite headers in any language (e.g., "Athlete Name", "Racer Name")
            bool isNonAthleteField =
                n2.Contains("soc") || n2.Contains("club") || n2.Contains("team") || n2.Contains("squadr") ||
                n2.Contains("verein") || n2.Contains("\u043A\u043B\u0443\u0431") || n2.Contains("\u043A\u043E\u043C\u0430\u043D\u0434") ||
                n2.Contains("\u4FF1\u4E50\u90E8") || n2.Contains("\u8F66\u961F") || n2.Contains("\u30AF\u30E9\u30D6") || n2.Contains("\u30C1\u30FC\u30E0") ||
                n2.Contains("gara") || n2.Contains("race") || n2.Contains("comitat") || n2.Contains("\u043A\u043E\u043C\u0438\u0442\u0435\u0442") ||
                n2.Contains("codic") || n2.Contains("fisc") || n2.Contains("\u043D\u0430\u043B\u043E\u0433");

            if (!isNonAthleteField)
            {
                foreach (var tok in TokenNomeAtleta)
                {
                    if (n1.Contains(tok) || n2.Contains(tok))
                        return "Nome";
                }
            }

            return null;
        }

        /// <summary>
        /// Determines whether a worksheet name corresponds to a non-participant sheet
        /// (such as club lookup tables, standings, results, category summaries, or print layouts)
        /// so it can be excluded from athlete roster extraction.
        /// </summary>
        /// <param name="foglioNorm">Normalized worksheet name.</param>
        /// <returns><c>true</c> if the sheet should be skipped; otherwise, <c>false</c>.</returns>
        private static bool IsFoglioNonAtleti(string foglioNorm)
        {
            if (string.IsNullOrEmpty(foglioNorm)) return false;

            // Club / team / squad sheets across all 11 supported languages
            if (foglioNorm == "societa" || foglioNorm == "societas" ||
                foglioNorm == "club" || foglioNorm == "clubs" || foglioNorm == "clubes" ||
                foglioNorm == "squadre" || foglioNorm == "teams" || foglioNorm == "equipes" ||
                foglioNorm == "vereine" || foglioNorm == "kluby" ||
                foglioNorm == "\u043A\u043B\u0443\u0431\u044B" || foglioNorm == "\u043A\u043E\u043C\u0430\u043D\u0434\u044B" ||
                foglioNorm == "\u4FF1\u4E50\u90E8" || foglioNorm == "\u8F66\u961F" ||
                foglioNorm == "\u30AF\u30E9\u30D6" || foglioNorm == "\u30C1\u30FC\u30E0")
            {
                return true;
            }

            // Standings, finish order, category, print layout, configuration, and summary sheets to exclude
            string[] prefissiEsclusi = new[]
            {
                "config", "konfig", "impresi", "stampa", "print", "imprim", "druck", "druk", "\u043F\u0435\u0447\u0430\u0442", "\u6253\u5370", "\u5370\u5237",
                "resumen", "summary", "riepilogo", "\u0441\u0432\u043E\u0434\u043A", "\u6C47\u603B", "\u6982\u8981",
                "clasif", "classific", "classement", "wertung", "rangliste", "klassement", "klasyfikac",
                "\u043A\u043B\u0430\u0441\u0441\u0438\u0444", "\u0437\u0430\u0447\u0435\u0442", "\u0437\u0430\u0447\u0451\u0442", "\u043F\u0440\u043E\u0442\u043E\u043A\u043E\u043B",
                "\u6392\u540D", "\u6210\u7EE9", "\u9806\u4F4D", "\u7D50\u679C",
                "categor", "kategor", "\u043A\u0430\u0442\u0435\u0433\u043E\u0440", "\u7EC4\u522B", "\u30AB\u30C6\u30B4\u30EA",
                "arriv", "llegad", "arrivee", "ankunft", "chegad", "aankomst", "przyjazd", "\u0444\u0438\u043D\u0438\u0448", "\u7EC8\u70B9", "\u30D5\u30A3\u30CB\u30C3\u30B7\u30E5"
            };

            if (foglioNorm == "class" || foglioNorm == "cl soc" || foglioNorm.StartsWith("klas ")) return true;
            foreach (var p in prefissiEsclusi)
            {
                if (foglioNorm.Contains(p)) return true;
            }
            return false;
        }

        /// <summary>
        /// Checks whether a worksheet name explicitly indicates a participant/start list
        /// (e.g., "Athletes", "StartList", "Iscritti"), granting it priority during sheet selection.
        /// </summary>
        /// <param name="foglioNorm">Normalized worksheet name.</param>
        /// <returns><c>true</c> if the worksheet name matches high-priority athlete keywords.</returns>
        private static bool IsFoglioAtletiPrioritario(string foglioNorm)
        {
            string[] kwAtleti = new[]
            {
                "atlet", "athlet", "partecip", "particip", "iscrit", "inscrit", "ciclist", "cyclist",
                "elenco", "anagraf", "start", "rider", "runner", "racer", "corredor", "coureur",
                "engage", "partant", "fahrer", "teilnehmer", "renner", "deelnemer", "zawodnik", "zawodnic", "kolarz", "atletas",
                "\u0441\u043F\u043E\u0440\u0442\u0441\u043C\u0435\u043D", "\u0433\u043E\u043D\u0449\u0438\u043A", "\u0443\u0447\u0430\u0441\u0442\u043D\u0438\u043A", "\u0441\u0442\u0430\u0440\u0442\u043E\u0432",
                "\u9009\u624B", "\u8FD0\u52A8\u5458", "\u53C2\u8D5B", "\u540D\u5355",
                "\u9078\u624B", "\u7AF6\u6280\u8005", "\u30E9\u30A4\u30C0\u30FC", "\u30A8\u30F3\u30C8\u30EA\u30FC", "\u53C2\u52A0\u8005"
            };
            foreach (var k in kwAtleti)
            {
                if (foglioNorm.Contains(k)) return true;
            }
            return false;
        }

        /// <summary>
        /// Reads raw participant rows from a .csv, .txt, or multi-sheet .xlsx file.
        /// For .xlsx files, evaluates all candidate worksheets using a heuristic quality score
        /// and resolves relational club IDs against secondary club sheets when present.
        /// </summary>
        /// <param name="filePath">Full path to the start-list file.</param>
        /// <returns>A list of raw imported participant rows.</returns>
        public static List<RigaPartecipanteImportata> LeggiPartecipantiDaFile(string filePath)
        {
            var risultato = new List<RigaPartecipanteImportata>();
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return risultato;

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".csv" || ext == ".txt")
                return EstraiPartecipantiDaTabella(LeggiRigheCsv(filePath), null);

            if (ext == ".xlsx")
            {
                var fogli = LeggiTuttiFogliXlsx(filePath);
                if (fogli.Count == 0) return risultato;

                var lookupClub = CostruisciLookupClubDaFogli(fogli);
                List<RigaPartecipanteImportata> migliore = risultato;
                int migliorPunteggio = -1;

                foreach (var kv in fogli)
                {
                    string foglioNorm = NormalizzaHeader(kv.Key, true);
                    if (IsFoglioNonAtleti(foglioNorm))
                        continue;

                    var candidati = EstraiPartecipantiDaTabella(kv.Value, lookupClub);
                    if (candidati.Count == 0) continue;

                    int score = candidati.Count * 10;
                    foreach (var p in candidati)
                    {
                        if (!string.IsNullOrWhiteSpace(p.Societa)) score += 4;
                        if (!string.IsNullOrWhiteSpace(p.Categoria)) score += 3;
                        if (!string.IsNullOrWhiteSpace(p.Ente)) score += 2;
                    }
                    if (IsFoglioAtletiPrioritario(foglioNorm))
                        score += 100000;

                    if (score > migliorPunteggio)
                    {
                        migliorPunteggio = score;
                        migliore = candidati;
                    }
                }
                return migliore;
            }
            return risultato;
        }

        /// <summary>
        /// Scans secondary club/team worksheets inside an .xlsx workbook to build a lookup
        /// dictionary mapping numeric club IDs to full club names.
        /// </summary>
        /// <param name="fogli">Dictionary of worksheet names to 2D string cell grids.</param>
        /// <returns>A case-insensitive dictionary mapping club ID strings to club names.</returns>
        private static Dictionary<string, string> CostruisciLookupClubDaFogli(Dictionary<string, List<List<string>>> fogli)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in fogli)
            {
                string sheetNorm = NormalizzaHeader(kv.Key, true);
                bool isClubSheet =
                    sheetNorm.Contains("club") || sheetNorm.Contains("societ") || sheetNorm.Contains("team") ||
                    sheetNorm.Contains("squadr") || sheetNorm.Contains("verein") || sheetNorm.Contains("klub") ||
                    sheetNorm.Contains("\u043A\u043B\u0443\u0431") || sheetNorm.Contains("\u043A\u043E\u043C\u0430\u043D\u0434") ||
                    sheetNorm.Contains("\u4FF1\u4E50\u90E8") || sheetNorm.Contains("\u30AF\u30E9\u30D6");

                if (!isClubSheet) continue;

                var rows = kv.Value;
                if (rows.Count < 2) continue;

                int colId = -1, colName = -1;
                var hdr = rows[0];
                for (int c = 0; c < hdr.Count; c++)
                {
                    string h = NormalizzaHeader(hdr[c], true);
                    if (colId < 0 && (
                        h == "numero" || h == "number" || h == "numer" || h == "id" ||
                        h == "codice" || h == "codigo" || h == "nr klubu" ||
                        h == "n club" || h == "n clube" || h == "n soc" ||
                        h == "verenigingsnr" || h == "\u756A\u53F7" || h == "\u7F16\u53F7" ||
                        h == "\u043D\u043E\u043C\u0435\u0440" || h == "\u043A\u043E\u0434 \u043A\u043B\u0443\u0431\u0430"))
                    {
                        colId = c;
                    }
                    if (MappaIntestazioneColonna(hdr[c]) == "Societa")
                        colName = c;
                }

                if (colId >= 0 && colName >= 0)
                {
                    for (int r = 1; r < rows.Count; r++)
                    {
                        if (colId < rows[r].Count && colName < rows[r].Count)
                        {
                            string idVal = (rows[r][colId] ?? "").Trim();
                            string nameVal = (rows[r][colName] ?? "").Trim().Trim('"').Trim();
                            if (!string.IsNullOrEmpty(idVal) && !string.IsNullOrEmpty(nameVal) && !map.ContainsKey(idVal))
                                map[idVal] = nameVal;
                        }
                    }
                }
            }
            return map;
        }

        /// <summary>
        /// Scans the top 15 rows of a parsed worksheet to automatically detect the header row,
        /// maps columns to canonical fields, infers an unmapped text column as athlete name if
        /// needed, resolves club IDs, and extracts structured participant records.
        /// </summary>
        /// <param name="rows">2D list of cell strings representing the worksheet.</param>
        /// <param name="lookupClub">Optional club ID-to-name lookup map from secondary sheets.</param>
        /// <returns>A list of extracted <see cref="RigaPartecipanteImportata"/> records.</returns>
        private static List<RigaPartecipanteImportata> EstraiPartecipantiDaTabella(List<List<string>> rows, Dictionary<string, string> lookupClub)
        {
            var list = new List<RigaPartecipanteImportata>();
            if (rows == null || rows.Count < 2) return list;

            int bestHeaderRow = -1, bestScore = -1, colClubId = -1, colAnno = -1, colTessera = -1;
            Dictionary<string, int> bestMap = null;

            // Scan up to the first 15 rows to locate the header row with the highest mapped-column score.
            int maxScan = Math.Min(15, rows.Count);
            for (int r = 0; r < maxScan; r++)
            {
                var currentMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int tempClubId = -1, tempAnno = -1, tempTess = -1;
                var row = rows[r];
                for (int c = 0; c < row.Count; c++)
                {
                    string cellText = (row[c] ?? "").Trim();
                    if (string.IsNullOrEmpty(cellText)) continue;

                    string hNorm = NormalizzaHeader(cellText, true);
                    if (hNorm == "n soc" || hNorm == "no soc" || hNorm == "num soc" ||
                        hNorm == "n club" || hNorm == "no club" || hNorm == "num club" ||
                        hNorm == "n clube" || hNorm == "no clube" || hNorm == "num clube" ||
                        hNorm == "nr klubu" || hNorm == "numer klubu" || hNorm == "verenigingsnr" ||
                        hNorm == "id club" || hNorm == "codice societa" ||
                        hNorm == "\u30C1\u30FC\u30E0\u756A\u53F7" || hNorm == "\u4FF1\u4E50\u90E8\u7F16\u53F7" ||
                        hNorm == "\u043A\u043E\u0434 \u043A\u043B\u0443\u0431\u0430" ||
                        hNorm == "\u043D\u043E\u043C\u0435\u0440 \u043A\u043B\u0443\u0431\u0430")
                    {
                        tempClubId = c;
                    }
                    if (hNorm == "anno" || hNorm == "ano" || hNorm == "year" || hNorm == "annee" || hNorm == "jahr" ||
                        hNorm == "jaar" || hNorm == "rok ur" || hNorm == "rok urodzenia" ||
                        hNorm == "\u751F\u5E74" || hNorm == "\u51FA\u751F\u5E74\u4EFD" ||
                        hNorm.StartsWith("\u0433\u043E\u0434"))
                    {
                        tempAnno = c;
                    }
                    if (hNorm.Contains("licenc") || hNorm.Contains("tesser") || hNorm.Contains("licens") ||
                        hNorm.Contains("licent") || hNorm.Contains("\u30E9\u30A4\u30BB\u30F3\u30B9") ||
                        hNorm.Contains("\u6267\u7167") || hNorm.Contains("\u043B\u0438\u0446\u0435\u043D\u0437"))
                    {
                        tempTess = c;
                    }

                    string field = MappaIntestazioneColonna(cellText);
                    if (field != null && !currentMap.ContainsKey(field))
                        currentMap[field] = c;
                }

                bool hasAnyNameSynonym = currentMap.ContainsKey("Nome") || currentMap.ContainsKey("Cognome");
                bool hasBibAndOther = currentMap.ContainsKey("Pettorale") && currentMap.Count >= 2;

                if (hasAnyNameSynonym || hasBibAndOther)
                {
                    int rowScore = currentMap.Count * 10 + (hasAnyNameSynonym ? 20 : 0) + (currentMap.ContainsKey("Pettorale") ? 10 : 0);
                    if (rowScore > bestScore)
                    {
                        bestScore = rowScore;
                        bestHeaderRow = r;
                        bestMap = currentMap;
                        colClubId = tempClubId;
                        colAnno = tempAnno;
                        colTessera = tempTess;
                    }
                }
            }

            if (bestHeaderRow < 0 || bestMap == null) return list;

            int idxPett = bestMap.ContainsKey("Pettorale") ? bestMap["Pettorale"] : -1;
            int idxNome = bestMap.ContainsKey("Nome") ? bestMap["Nome"] : -1;
            int idxCogn = bestMap.ContainsKey("Cognome") ? bestMap["Cognome"] : -1;
            int idxCat  = bestMap.ContainsKey("Categoria") ? bestMap["Categoria"] : -1;
            int idxSoc  = bestMap.ContainsKey("Societa") ? bestMap["Societa"] : -1;
            int idxEnte = bestMap.ContainsKey("Ente") ? bestMap["Ente"] : -1;

            // If no column header matches Name/Surname synonyms but a bib column exists,
            // automatically pick the first unmapped text column (e.g., custom or blank header)
            // Fallback column inference: when a bib column exists but no explicit name header was matched,
                        // inspect the first 10 data rows and select the first unmapped column containing text values.
            if (idxNome < 0 && idxCogn < 0 && idxPett >= 0)
            {
                var usedCols = new HashSet<int>(bestMap.Values);
                if (colClubId >= 0) usedCols.Add(colClubId);
                if (colAnno >= 0) usedCols.Add(colAnno);
                if (colTessera >= 0) usedCols.Add(colTessera);

                int maxCols = rows.Skip(bestHeaderRow + 1).Take(10).Select(r => r.Count).DefaultIfEmpty(0).Max();
                for (int c = 0; c < maxCols; c++)
                {
                    if (usedCols.Contains(c)) continue;
                    int textCount = 0;
                    foreach (var sampleRow in rows.Skip(bestHeaderRow + 1).Take(10))
                    {
                        string v = GetCell(sampleRow, c);
                        int dummy;
                        if (v.Length >= 3 && v.Any(char.IsLetter) && !int.TryParse(v, out dummy))
                            textCount++;
                    }
                    if (textCount >= 2)
                    {
                        idxNome = c;
                        break;
                    }
                }
            }

            if (idxNome < 0 && idxCogn < 0) return list;

            for (int r = bestHeaderRow + 1; r < rows.Count; r++)
            {
                var row = rows[r];
                string pett = GetCell(row, idxPett);
                string nome = GetCell(row, idxNome);
                string cogn = GetCell(row, idxCogn);
                string cat  = GetCell(row, idxCat);
                string soc  = GetCell(row, idxSoc);
                string ente = GetCell(row, idxEnte);
                string anno = GetCell(row, colAnno);
                string tess = GetCell(row, colTessera);

                string full = !string.IsNullOrEmpty(cogn)
                    ? (string.IsNullOrEmpty(nome) ? cogn : (cogn + " " + nome).Trim())
                    : nome;

                // Resolve relational club ID against secondary club worksheet lookup table when club name is missing or "0".
                if ((string.IsNullOrWhiteSpace(soc) || soc == "0") && colClubId >= 0 && lookupClub != null && lookupClub.Count > 0)
                {
                    string cid = GetCell(row, colClubId);
                    if (!string.IsNullOrEmpty(cid) && lookupClub.ContainsKey(cid))
                        soc = lookupClub[cid];
                }

                // Strip trailing ".0" artifact produced when Excel stores integer bib numbers as floating-point values.
                if (pett.EndsWith(".0", StringComparison.Ordinal))
                    pett = pett.Substring(0, pett.Length - 2);

                if (string.IsNullOrWhiteSpace(full))
                    continue;

                list.Add(new RigaPartecipanteImportata
                {
                    Pettorale = pett,
                    Nome = nome,
                    Cognome = cogn,
                    NomeCompleto = full,
                    Categoria = cat,
                    Societa = soc,
                    Ente = ente,
                    Anno = anno,
                    Tessera = tess
                });
            }
            return list;
        }

        /// <summary>
        /// Safely retrieves and trims a cell string from a row by column index.
        /// </summary>
        private static string GetCell(List<string> row, int col)
        {
            if (col < 0 || col >= row.Count) return string.Empty;
            return (row[col] ?? string.Empty).Trim();
        }

        /// <summary>
        /// Reads a text/CSV file and automatically detects its character encoding by inspecting
        /// UTF-8 / UTF-16 Byte Order Marks, validating strict UTF-8 decoding, and falling back
        /// to Windows-1252 / Latin-1 for legacy European spreadsheets.
        /// </summary>
        /// <param name="filePath">Path to the CSV or text file.</param>
        /// <returns>The decoded text content of the file.</returns>
        private static string LeggiTestoCsvConRilevamentoEncoding(string filePath)
        {
            byte[] bytes = File.ReadAllBytes(filePath);
            if (bytes.Length == 0) return string.Empty;

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);

            try
            {
                var utf8Strict = new UTF8Encoding(false, true);
                return utf8Strict.GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                try
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                    return Encoding.GetEncoding(1252).GetString(bytes);
                }
                catch
                {
                    return Encoding.Latin1.GetString(bytes);
                }
            }
        }

        /// <summary>
        /// Splits a CSV or delimited text file into rows and columns, automatically detecting
        /// whether semicolon (;), comma (,), or tab (\t) is the primary delimiter.
        /// </summary>
        /// <param name="filePath">Path to the CSV or text file.</param>
        /// <returns>A 2D list of trimmed cell strings.</returns>
        private static List<List<string>> LeggiRigheCsv(string filePath)
        {
            var result = new List<List<string>>();
            string content = LeggiTestoCsvConRilevamentoEncoding(filePath);
            if (string.IsNullOrWhiteSpace(content)) return result;

            var lines = content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            if (lines.Length == 0) return result;

            char sep = ';';
            if (lines[0].Count(c => c == ',') > lines[0].Count(c => c == ';')) sep = ',';
            else if (lines[0].Count(c => c == '\t') > lines[0].Count(c => c == ';')) sep = '\t';

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                result.Add(line.Split(sep).Select(x => x.Trim().Trim('"')).ToList());
            }
            return result;
        }

        /// <summary>
        /// Parses all worksheets of an .xlsx OpenXML workbook directly using <see cref="ZipArchive"/>
        /// and <see cref="XDocument"/>, resolving shared strings and cell coordinate references
        /// without external spreadsheet libraries.
        /// </summary>
        /// <param name="filePath">Path to the .xlsx workbook.</param>
        /// <returns>A dictionary mapping each worksheet name to its 2D table of cell strings.</returns>
        private static Dictionary<string, List<List<string>>> LeggiTuttiFogliXlsx(string filePath)
        {
            var sheets = new Dictionary<string, List<List<string>>>(StringComparer.OrdinalIgnoreCase);

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                // Step 1: Load the shared string table (xl/sharedStrings.xml) used by Excel to deduplicate cell text.
                var sharedStrings = new List<string>();
                var ssEntry = zip.GetEntry("xl/sharedStrings.xml");
                if (ssEntry != null)
                {
                    using (var s = ssEntry.Open())
                    {
                        var doc = XDocument.Load(s);
                        XNamespace ns = doc.Root.Name.Namespace;
                        foreach (var si in doc.Descendants(ns + "si"))
                            sharedStrings.Add(string.Concat(si.Descendants(ns + "t").Select(t => t.Value)));
                    }
                }

                // Step 2: Map relationship IDs (rId*) to internal worksheet XML paths via workbook.xml.rels.
                var relMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry != null)
                {
                    using (var s = relsEntry.Open())
                    {
                        var doc = XDocument.Load(s);
                        foreach (var rel in doc.Root.Elements())
                        {
                            var idAttr = rel.Attribute("Id");
                            var targetAttr = rel.Attribute("Target");
                            if (idAttr != null && targetAttr != null)
                            {
                                string target = targetAttr.Value.TrimStart('/');
                                if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                                    target = "xl/" + target;
                                relMap[idAttr.Value] = target;
                            }
                        }
                    }
                }

                var wbEntry = zip.GetEntry("xl/workbook.xml");
                if (wbEntry == null) return sheets;

                // Step 3: Enumerate worksheets in tab order from xl/workbook.xml.
                var orderedSheets = new List<KeyValuePair<string, string>>();
                using (var s = wbEntry.Open())
                {
                    var doc = XDocument.Load(s);
                    XNamespace ns = doc.Root.Name.Namespace;
                    XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    foreach (var sh in doc.Descendants(ns + "sheet"))
                    {
                        var nameAttr = sh.Attribute("name");
                        var ridAttr = sh.Attribute(rNs + "id");
                        if (nameAttr != null && ridAttr != null && relMap.ContainsKey(ridAttr.Value))
                            orderedSheets.Add(new KeyValuePair<string, string>(nameAttr.Value, relMap[ridAttr.Value]));
                    }
                }

                // Step 4: Parse each worksheet XML (xl/worksheets/sheet*.xml) cell by cell into a 2D string table.
                foreach (var sh in orderedSheets)
                {
                    var wsEntry = zip.GetEntry(sh.Value);
                    if (wsEntry == null) continue;

                    var table = new List<List<string>>();
                    using (var s = wsEntry.Open())
                    {
                        var doc = XDocument.Load(s);
                        XNamespace ns = doc.Root.Name.Namespace;
                        foreach (var rowEl in doc.Descendants(ns + "row"))
                        {
                            var rowList = new List<string>();
                            foreach (var cEl in rowEl.Elements(ns + "c"))
                            {
                                var refAttr = cEl.Attribute("r");
                                int colIndex = refAttr != null ? ColumnLettersToIndex(refAttr.Value) : rowList.Count;
                                while (rowList.Count <= colIndex) rowList.Add(string.Empty);

                                string cellType = cEl.Attribute("t") != null ? cEl.Attribute("t").Value : "";
                                string val = string.Empty;

                                if (cellType == "s")
                                {
                                    var vEl = cEl.Element(ns + "v");
                                    int idx;
                                    if (vEl != null && int.TryParse(vEl.Value, out idx) && idx >= 0 && idx < sharedStrings.Count)
                                        val = sharedStrings[idx];
                                }
                                else if (cellType == "inlineStr")
                                {
                                    val = string.Concat(cEl.Descendants(ns + "t").Select(t => t.Value));
                                }
                                else
                                {
                                    var vEl = cEl.Element(ns + "v");
                                    if (vEl != null) val = vEl.Value;
                                }
                                rowList[colIndex] = val ?? string.Empty;
                            }
                            table.Add(rowList);
                        }
                    }
                    sheets[sh.Key] = table;
                }
            }
            return sheets;
        }

        /// <summary>
        /// Converts an Excel cell coordinate's column letters (e.g., "A", "Z", "AA") into a
        /// zero-based column index.
        /// </summary>
        /// <param name="cellRef">Excel cell reference string (e.g., "C12").</param>
        /// <returns>Zero-based column index.</returns>
        private static int ColumnLettersToIndex(string cellRef)
        {
            int col = 0;
            foreach (char ch in cellRef)
            {
                if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
                else break;
            }
            return Math.Max(0, col - 1);
        }



        /// <summary>
        /// Gets the list of duplicate bib numbers encountered and skipped during the most
        /// recent call to <see cref="LeggiCiclistiDaFile"/>.
        /// </summary>
        public static List<int> UltimiPettoraliDuplicati { get; private set; } = new List<int>();

        /// <summary>
        /// High-level entry point that imports a start list file and converts raw rows into
        /// deduplicated <see cref="Ciclista"/> domain entities. Automatically assigns sequential
        /// bib numbers when omitted, skips duplicate bibs, and normalizes category casing.
        /// </summary>
        /// <param name="filePath">Full path to the .xlsx, .csv, or .txt start list.</param>
        /// <returns>A validated list of <see cref="Ciclista"/> participants ready for race timing.</returns>
        public static List<Ciclista> LeggiCiclistiDaFile(string filePath)
        {
            UltimiPettoraliDuplicati = new List<int>();
            var righe = LeggiPartecipantiDaFile(filePath);
            var ciclisti = new List<Ciclista>(righe.Count);
            var pettoraliUsati = new HashSet<int>();
            var pettoraliEspliciti = new HashSet<int>();
            var categorieCanonical = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int progressivo = 1;

            foreach (var r in righe)
            {
                int pettEsplicito;
                if (int.TryParse((r.Pettorale ?? "").Trim(), out pettEsplicito) && pettEsplicito > 0)
                    pettoraliEspliciti.Add(pettEsplicito);
            }

            foreach (var r in righe)
            {
                string nome = !string.IsNullOrWhiteSpace(r.NomeCompleto)
                    ? r.NomeCompleto.Trim()
                    : ((r.Cognome ?? "") + " " + (r.Nome ?? "")).Trim();

                if (string.IsNullOrWhiteSpace(nome))
                    continue;

                int pettorale;
                if (int.TryParse((r.Pettorale ?? "").Trim(), out pettorale) && pettorale > 0)
                {
                    if (pettoraliUsati.Contains(pettorale))
                    {
                        if (!UltimiPettoraliDuplicati.Contains(pettorale))
                            UltimiPettoraliDuplicati.Add(pettorale);
                        continue;
                    }
                }
                else
                {
                    // Auto-assign the lowest available sequential bib number that does not collide with explicit or used bibs.
                    while (pettoraliUsati.Contains(progressivo) || pettoraliEspliciti.Contains(progressivo))
                        progressivo++;
                    pettorale = progressivo;
                }

                pettoraliUsati.Add(pettorale);
                if (pettorale >= progressivo)
                    progressivo = pettorale + 1;

                string catRaw = string.IsNullOrWhiteSpace(r.Categoria) ? "N.D." : r.Categoria.Trim();
                if (!categorieCanonical.TryGetValue(catRaw, out string catNorm))
                {
                    categorieCanonical[catRaw] = catRaw;
                    catNorm = catRaw;
                }

                ciclisti.Add(new Ciclista
                {
                    Pettorale = pettorale,
                    Nome = nome,
                    Categoria = catNorm,
                    NomeSocieta = (r.Societa ?? "").Trim(),
                    Ente = (r.Ente ?? "").Trim(),
                    NonInElenco = false
                });
            }

            return ciclisti;
        }
    }
}
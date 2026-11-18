// ============================================================================
// File: Ciclista.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Defines the core domain model representing a single race participant
//   (cyclist or athlete). Stores registration metadata loaded from start lists
//   as well as live cumulative lap timestamps recorded during a race.
// ============================================================================

using System;
using System.Collections.Generic;

namespace CicloRank
{
    /// <summary>
    /// Represents an individual race participant (cyclist or athlete), storing
    /// biographical/club details and chronological lap timestamps.
    /// </summary>
    public class Ciclista
    {
        /// <summary>
        /// Gets or sets the unique race bib (plate) number assigned to the participant.
        /// </summary>
        public int Pettorale { get; set; }
        /// <summary>
        /// Gets or sets the full name of the participant.
        /// </summary>
        public string Nome { get; set; } = "";
        /// <summary>
        /// Gets or sets the competitive age group or racing category (e.g., Elite, Master, Junior).
        /// </summary>
        public string Categoria { get; set; } = "";
        /// <summary>
        /// Gets or sets the name of the sports club, team, or association the participant represents.
        /// </summary>
        public string NomeSocieta { get; set; } = "";
        /// <summary>
        /// Gets or sets the governing federation or sports promotion body (e.g., FCI, ACSI, UISP).
        /// Defaults to an empty string to prevent unintended federation assignments.
        /// </summary>
        public string Ente { get; set; } = ""; // Empty by default to prevent unintended federation assignments
        /// <summary>
        /// Gets or sets a value indicating whether this bib number was entered during the race
        /// without being present in the pre-loaded start list ("unknown rider").
        /// </summary>
        public bool NonInElenco { get; set; } = false;
        /// <summary>
        /// Gets or sets the chronological list of cumulative elapsed race timestamps
        /// recorded each time this participant crosses the finish/timing line.
        /// </summary>
        public List<TimeSpan> Passaggi { get; set; } = new List<TimeSpan>();

        /// <summary>
        /// Gets the total number of completed laps, derived from the count of recorded crossings.
        /// </summary>
        public int Giri => Passaggi.Count;
        /// <summary>
        /// Gets the cumulative elapsed race time at the participant's most recent lap crossing,
        /// or <see cref="TimeSpan.Zero"/> if no laps have been recorded yet.
        /// </summary>
        public TimeSpan TempoTotale => Passaggi.Count > 0 ? Passaggi[Passaggi.Count - 1] : TimeSpan.Zero;
    }
}
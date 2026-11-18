// ============================================================================
// File: DatiGara.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Defines the race header metadata model used to store event details such
//   as governing federation/committee, organizer, race title, venue location,
//   event date, and sport discipline for official Excel and PDF exports.
// ============================================================================

using System;

namespace CicloRank
{
    /// <summary>
    /// Stores event header metadata used to populate official classification reports
    /// (Excel/PDF exports) and saved race session files.
    /// </summary>
    public class DatiGara
    {
        /// <summary>
        /// Gets or sets the provincial/regional committee or governing federation overseeing the race.
        /// </summary>
        public string ComitatoProvinciale { get; set; } = "";
        /// <summary>
        /// Gets or sets the name of the organizing club, committee, or event host.
        /// </summary>
        public string Organizzazione { get; set; } = "";
        /// <summary>
        /// Gets or sets the official title or name of the sporting event.
        /// </summary>
        public string Denominazione { get; set; } = "";
        /// <summary>
        /// Gets or sets the town, city, or venue where the race takes place.
        /// </summary>
        public string Localita { get; set; } = "";
        /// <summary>
        /// Gets or sets the event date formatted as "dd/MM/yyyy" (defaults to today's date).
        /// </summary>
        public string DataGara { get; set; } = DateTime.Now.ToString("dd/MM/yyyy");
        /// <summary>
        /// Gets or sets the sport or discipline name (e.g., Cycling, MTB, Running)
        /// used to dynamically customize classification headers.
        /// </summary>
        public string Sport { get; set; } = "";
    }
}
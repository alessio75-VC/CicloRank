// ============================================================================
// File: SessioneGara.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Defines the serializable race session container used both for automatic
//   crash-recovery snapshots (sessione_temp.json) and manual session export/import.
// ============================================================================

using System.Collections.Generic;

namespace CicloRank
{
    /// <summary>
    /// Serializable data transfer object (DTO) representing a complete race snapshot.
    /// Used for automatic crash recovery (sessione_temp.json) as well as manual
    /// export and import of race sessions in JSON format.
    /// </summary>
    public class SessioneGara
    {
        /// <summary>
        /// Gets or sets the race header information (federation, organizer, event name, date, sport).
        /// </summary>
        public DatiGara DatiGara { get; set; } = new DatiGara();
        /// <summary>
        /// Gets or sets the full roster of registered and dynamically discovered participants.
        /// </summary>
        public List<Ciclista> ListaCiclisti { get; set; } = new List<Ciclista>();
        /// <summary>
        /// Gets or sets the chronological log of all recorded lap crossings.
        /// </summary>
        public List<Passaggio> ListaPassaggi { get; set; } = new List<Passaggio>();
        /// <summary>
        /// Gets or sets the total elapsed stopwatch duration in seconds when the session was saved.
        /// </summary>
        public double TempoTrascorsoSecondi { get; set; }
    }
}
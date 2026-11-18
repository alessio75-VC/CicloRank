// ============================================================================
// File: Passaggio.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Defines the atomic timing event model representing a single finish-line
//   crossing (bib number + elapsed race timestamp).
// ============================================================================

using System;

namespace CicloRank
{
    /// <summary>
    /// Represents a single recorded finish-line crossing event, pairing a participant's
    /// bib number with the exact elapsed race time at the moment of transit.
    /// </summary>
    public class Passaggio
    {
        /// <summary>
        /// Gets or sets the bib number recorded at the timing point.
        /// </summary>
        public int Pettorale { get; set; }
        /// <summary>
        /// Gets or sets the cumulative elapsed race time when the crossing was recorded.
        /// </summary>
        public TimeSpan Tempo { get; set; }

        /// <summary>
        /// Initializes a new empty instance of <see cref="Passaggio"/> for JSON deserialization.
        /// </summary>
        public Passaggio() { }

        /// <summary>
        /// Initializes a new instance of <see cref="Passaggio"/> with a bib number and elapsed timestamp.
        /// </summary>
        /// <param name="pettorale">The rider's race bib number.</param>
        /// <param name="tempo">The cumulative elapsed time from race start.</param>
        public Passaggio(int pettorale, TimeSpan tempo)
        {
            Pettorale = pettorale;
            Tempo = tempo;
        }
    }
}
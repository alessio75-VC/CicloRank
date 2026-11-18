// ============================================================================
// File: Program.cs
// Project: CicloRank - Race Management & Manual Lap Timing System
// Description:
//   Application entry point for the CicloRank Windows Forms desktop app.
//   Configures OS visual styles, text rendering compatibility, and starts the
//   main UI event loop on a Single-Threaded Apartment (STA) thread.
// ============================================================================

using System;
using System.Windows.Forms;

namespace CicloRank
{
    /// <summary>
    /// Provides the main entry point for the CicloRank Windows Forms desktop application.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Application entry point. Enables modern Windows visual styles, configures GDI+
        /// text rendering defaults, and launches <see cref="MainForm"/> on an STA thread.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
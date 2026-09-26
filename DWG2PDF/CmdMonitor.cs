using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
namespace DWG2PDF
{
    public static class CmdMonitor
    {
        public static bool WaitForMirror = false;

        public static void Init()
        {
            Application.DocumentManager.DocumentCreated += (s, e) =>
            {
                e.Document.CommandEnded += CommandEnded;
            };

            // Add cho document hiện tại
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.CommandEnded += CommandEnded;
            }
        }

        private static void CommandEnded(object sender, CommandEventArgs e)
        {
            if (e.GlobalCommandName.Equals("MIRROR", StringComparison.OrdinalIgnoreCase))
            {
                WaitForMirror = false;
            }
        }
    }
}


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DWG2PDF
{
    internal static class NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}

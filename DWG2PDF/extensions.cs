using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DWG2PDF
{
    class extensions
    {
        public static string extension(string plotname)
        {
            string result = "";
            if(plotname.Contains("PDF"))
            {
                result = ".pdf";
            }
            if (plotname.Contains("XPS"))
            {
                result = ".xps";
            }
            if (plotname.Contains("DWF"))
            {
                result = ".dwf";
            }
            if (plotname.Contains("JPG"))
            {
                result = ".jpg";
            }
            if (plotname.Contains("PNG"))
            {
                result = ".png";
            }
            return result;
        }
    }
}

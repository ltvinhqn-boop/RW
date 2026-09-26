using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DWG2PDF
{
    class CHECKTIME
    {
        public static bool CHECKLICENSE(Double time)
        {
            bool CHECK = false;
            DateTime networkTime = NtpTime.GetNetworkTime();
            string dateTimeString = networkTime.ToString("yyyy MM dd").Replace(" ", "");
            if(time-Double.Parse(dateTimeString)>0)
            {
                CHECK = true;
            }
            return CHECK;
        }
    }
}

using System;
using System.Net.Sockets;

using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace DWG2PDF
{
    public static class NtpTime
    {
        private static readonly string NtpServer = "pool.ntp.org";

        public static DateTime GetNetworkTime()
        {
            try
            {
                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
                {
                    socket.Connect(NtpServer, 123); // Port 123 is used for NTP

                    var data = new byte[48]; // NTP message size is 48 bytes
                    data[0] = 0x1B; // Set the first byte as the NTP request type

                    socket.Send(data);
                    socket.Receive(data);

                    // Get the timestamp from the NTP response
                    ulong timestamp = (ulong)data[40] << 24 | (ulong)data[41] << 16 | (ulong)data[42] << 8 | (ulong)data[43];

                    // Convert the timestamp to a DateTime object
                    ulong unixEpoch = 2208988800UL; // Seconds between 1900-01-01 and 1970-01-01
                    ulong unixTime = timestamp - unixEpoch;
                    DateTime networkDateTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(unixTime);

                    return networkDateTime.ToLocalTime();
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions
               // Console.WriteLine("Error: " + ex.Message);
                return DateTime.Now;
            }
        }
    }
}

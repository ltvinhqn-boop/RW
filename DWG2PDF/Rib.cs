using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HBTools
{
    internal class Rib
    {
    }


public class RibRegion
    {
        public int OL_Min { get; set; }
        public int OL_Max { get; set; }
        public int OW_Min { get; set; }
        public int OW_Max { get; set; }
        public int RibX { get; set; }
        public int RibY { get; set; }

        public bool Contains(int ol, int ow)
        {
            return ol >= OL_Min && ol < OL_Max && ow >= OW_Min && ow < OW_Max;
        }
    }

    public class RibLookup
    {
        private List<RibRegion> regions = new List<RibRegion>
    {
        //new RibRegion { OL_Min = 550, OL_Max = 640, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
        //new RibRegion { OL_Min = 640, OL_Max = 800, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
        //new RibRegion { OL_Min = 800, OL_Max = 810, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
        //new RibRegion { OL_Min = 810, OL_Max = 1110, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
        //new RibRegion { OL_Min = 1110, OL_Max = 1240, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
        //new RibRegion { OL_Min = 1240, OL_Max = 1350, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },

        //new RibRegion { OL_Min = 810, OL_Max = 1110, OW_Min = 800, OW_Max = 1110, RibX = 2, RibY = 0 },
        //new RibRegion { OL_Min = 1110, OL_Max = 1240, OW_Min = 800, OW_Max = 1110, RibX = 2, RibY = 0 },
        //new RibRegion { OL_Min = 1240, OL_Max = 1850, OW_Min = 800, OW_Max = 1110, RibX = 2, RibY = 0 },

        //new RibRegion { OL_Min = 1110, OL_Max = 1240, OW_Min = 1110, OW_Max = 1240, RibX = 2, RibY = 0 },
        //new RibRegion { OL_Min = 1240, OL_Max = 1850, OW_Min = 1110, OW_Max = 1240, RibX = 2, RibY = 0 },

        //new RibRegion { OL_Min = 1240, OL_Max = 1850, OW_Min = 1240, OW_Max = 1850, RibX = 4, RibY = 2 },
        //new RibRegion { OL_Min = 1850, OL_Max = 2200, OW_Min = 1240, OW_Max = 1850, RibX = 4, RibY = 2 },
        //new RibRegion { OL_Min = 2200, OL_Max = 3000, OW_Min = 1240, OW_Max = 1850, RibX = 4, RibY = 2 },
        //new RibRegion { OL_Min = 3000, OL_Max = 3365, OW_Min = 1240, OW_Max = 1850, RibX = 4, RibY = 4 },

        //new RibRegion { OL_Min = 1850, OL_Max = 2200, OW_Min = 1850, OW_Max = 2200, RibX = 6, RibY = 0 },
        //new RibRegion { OL_Min = 2200, OL_Max = 3000, OW_Min = 1850, OW_Max = 2200, RibX = 6, RibY = 2 },
        //new RibRegion { OL_Min = 3000, OL_Max = 3365, OW_Min = 1850, OW_Max = 2200, RibX = 6, RibY = 4 },

        // ✅ Đã sửa vùng này:
        //new RibRegion { OL_Min = 3000, OL_Max = 3365, OW_Min = 1110, OW_Max = 1240, RibX = 4, RibY = 2 },
    // SS1 - SS3 và Q6 vùng xanh lá
    //new RibRegion { OL_Min = 550, OL_Max = 640, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
    //new RibRegion { OL_Min = 640, OL_Max = 800, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
    //new RibRegion { OL_Min = 800, OL_Max = 810, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
    //new RibRegion { OL_Min = 810, OL_Max = 1110, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },
    //new RibRegion { OL_Min = 1110, OL_Max = 1240, OW_Min = 550, OW_Max = 800, RibX = 0, RibY = 0 },

    //// S1 và S2 vùng xanh dương
    //new RibRegion { OL_Min = 810, OL_Max = 1240, OW_Min = 800, OW_Max = 1110, RibX = 2, RibY = 0 },
    //new RibRegion { OL_Min = 810, OL_Max = 1240, OW_Min = 1110, OW_Max = 1240, RibX = 2, RibY = 0 },
    //new RibRegion { OL_Min = 1240, OL_Max = 1350, OW_Min = 550, OW_Max = 1240, RibX = 0, RibY = 0 },
    //new RibRegion { OL_Min = 1350, OL_Max = 1850, OW_Min = 550, OW_Max = 1240, RibX = 2, RibY = 0 },
    // Vùng xanh nhạt (RIB X: 0, Y: 0)
    new RibRegion { OL_Min = 550, OL_Max = 1349, OW_Min = 550, OW_Max = 809, RibX = 0, RibY = 0 },
    // Vùng vàng (RIB X: 2, Y: 2)
    new RibRegion { OL_Min = 550, OL_Max = 1849, OW_Min = 1240, OW_Max = 1849, RibX = 2, RibY = 2 },
    new RibRegion { OL_Min = 1350, OL_Max = 2999, OW_Min = 810, OW_Max = 1239, RibX = 2, RibY = 2 },
    new RibRegion { OL_Min = 2200, OL_Max = 3364, OW_Min = 810, OW_Max = 1109, RibX = 2, RibY = 2 },
    // Vùng cam nhạt (RIB X: 4, Y: 2)
    new RibRegion { OL_Min = 1850, OL_Max = 3364, OW_Min = 1240, OW_Max = 1849, RibX = 4, RibY = 2 },
    new RibRegion { OL_Min = 3000, OL_Max = 3364, OW_Min = 1100, OW_Max = 1239, RibX = 4, RibY = 2 },
    // Vùng cam đậm (RIB X: 6, Y: 2)
    new RibRegion { OL_Min = 810, OL_Max = 3364, OW_Min = 1850, OW_Max = 2200, RibX = 6, RibY = 2 },

    // Vùng xanh dương đậm (RIB X: 6, Y: 0)
    new RibRegion { OL_Min = 550, OL_Max = 809, OW_Min = 1850, OW_Max = 2200, RibX = 6, RibY = 0 },

    //// Vùng vàng nhạt (RIB X: 4, Y: 2)
    //new RibRegion { OL_Min = 3000, OL_Max = 3365, OW_Min = 550, OW_Max = 1350, RibX = 4, RibY = 2 },

    // Vùng dương nhạt (RIB X: 2, Y: 0)
    new RibRegion { OL_Min = 550, OL_Max = 1349, OW_Min = 810, OW_Max = 1239, RibX = 2, RibY = 0 },
    new RibRegion { OL_Min = 1350, OL_Max = 2199, OW_Min = 550, OW_Max = 809, RibX = 2, RibY = 0 },
    // Vùng xám (RIB X: 4, Y: 4)
    new RibRegion { OL_Min = 3365, OL_Max = 5000, OW_Min = 550, OW_Max = 1849, RibX = 4, RibY = 4 },

    // Vùng vàng đậm (RIB X: 6, Y: 4)
    new RibRegion { OL_Min = 3365, OL_Max = 5000, OW_Min = 1850, OW_Max = 2200, RibX = 6, RibY = 4 },
    };

        public (int RibX, int RibY)? GetRibCount(int ol, int ow)
        {
            foreach (var region in regions)
            {
                if (region.Contains(ol, ow))
                    return (region.RibX, region.RibY);
            }
            return null;
        }
    }

}

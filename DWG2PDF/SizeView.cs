using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Geometry;

namespace HBTools
{
    internal class SizeView
    {
    }
    public class SizeView1
    {
        public double Size_L { get; set; }
        public double Size_R { get; set; }    
        public double Size_H { get; set; }

      //  public Point3d Position { get; set; }


        public SizeView1() { }
        public SizeView1(double size_L, double size_R, double size_H/*, Point3d position*/)
        {
            Size_L = size_L;
            Size_R = size_R;
            Size_H = size_H;
          //  Position = position;
        }
    }
}

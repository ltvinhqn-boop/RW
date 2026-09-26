using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DWG2PDF
{
    internal class DimZam
    {
    }
    public class ZamPara
    {
        public double LengthX { get; set; }
        public double HeightY { get; set; }
        public int Quanlity { get; set; }
    }
    public class Zamparameters
    {
        public static ZamPara GetDimZam(double lengthX, double heightY, int quanlity)
        {
            return new ZamPara
            {
                LengthX = lengthX,
                HeightY = heightY,
                Quanlity = quanlity
            };
        }
    }
}
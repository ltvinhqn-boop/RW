using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using System.Linq;
namespace DWG2PDF
{
   

    public class Extents2dComparer : IEqualityComparer<Extents2d>
    {
        public bool Equals(Extents2d x, Extents2d y)
        {
            return x.MinPoint.IsEqualTo(y.MinPoint) && x.MaxPoint.IsEqualTo(y.MaxPoint);
        }

        public int GetHashCode(Extents2d obj)
        {
            return obj.MinPoint.GetHashCode() ^ obj.MaxPoint.GetHashCode();
        }
    }

    public class Extents2dFilter
    {
        public static List<Extents2d> RemoveDuplicates(List<Extents2d> extentsList)
        {
            var distinctExtents = new HashSet<Extents2d>(extentsList, new Extents2dComparer());
            return distinctExtents.ToList();
        }
    }
}

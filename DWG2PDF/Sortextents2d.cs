using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace DWG2PDF
{
           public static class Sortextents2d
        {
            public static void SortExtentsByPosition(List<Extents2d> extentsList, SortDirection sortDirection)
            {
                extentsList.Sort((extents1, extents2) =>
                {
                    Point2d point1 = extents1.MinPoint;
                    Point2d point2 = extents2.MinPoint;

                    if (sortDirection == SortDirection.TopToBottomLeftToRight)
                    {
                        if (point1.Y != point2.Y)
                        {
                            return point2.Y.CompareTo(point1.Y); // Sắp xếp từ trên xuống dưới
                        }
                        else
                        {
                            return point1.X.CompareTo(point2.X); // Sắp xếp từ trái qua phải
                        }
                    }
                    else if (sortDirection == SortDirection.BottomToTopLeftToRight)
                    {
                        if (point1.Y != point2.Y)
                        {
                            return point1.Y.CompareTo(point2.Y); // Sắp xếp từ dưới lên trên
                        }
                        else
                        {
                            return point1.X.CompareTo(point2.X); // Sắp xếp từ trái qua phải
                        }
                    }                   
                 
                    else if (sortDirection == SortDirection.RightToLeftTopToBottom)
                    {
                        if (point1.Y != point2.Y)
                        {
                            return point2.Y.CompareTo(point1.Y); // Sắp xếp từ trên xuống dưới
                        }
                        else
                        {
                            return point2.X.CompareTo(point1.X); // Sắp xếp từ phải qua trái
                        }
                    }
                    else if (sortDirection == SortDirection.RightToLeftBottomToTop)
                    {
                        if (point1.Y != point2.Y)
                        {
                            return point1.Y.CompareTo(point2.Y); // Sắp xếp từ dưới lên trên
                        }
                        else
                        {
                            return point2.X.CompareTo(point1.X); // Sắp xếp từ phải qua trái
                        }
                    }

                    return 0;
                });
            }
        }

        public enum SortDirection
        {
            TopToBottomLeftToRight,
            BottomToTopLeftToRight,                      
            RightToLeftTopToBottom,
            RightToLeftBottomToTop
        }
    
}

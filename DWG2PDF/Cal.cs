using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.Internal;
using Microsoft.Office.Interop.Excel;
using Google.OrTools.Sat;

namespace DWG2PDF
{
    public class SheetArrangement
    {
        // Constructor with a body and proper return type (fix for CS1520 and CS0501)  
        public double WidthZam { get; set; }
        public double LengthZam { get; set; }

        public SheetArrangement(double lengthZam, double widthZam)
        {
            LengthZam = lengthZam;
            WidthZam = widthZam;
        }


        public class SheetSplitResult
        {
            public int TotalSheets { get; set; }
            public int SplitX { get; set; }
            public int SplitY { get; set; }
            public double SheetWidth { get; set; }
            public double SheetHeight { get; set; }
            public string SheetSizeUsed { get; set; }
            public double GapX { get; set; }
            public double GapY { get; set; }
            public double TotalAreaUsed { get; set; }
            public double AreaNesting { get; set; }
        }

        public static SheetSplitResult ArrangeSheets(List<SheetArrangement> listZam)
        {
            listZam = listZam.OrderByDescending(s => s.LengthZam).ThenByDescending(s => s.WidthZam).ToList();
            // Existing implementation of ArrangeSheets  
            var sizes = new List<(double width, double height, string name)>
               {
                   (914, 2438, "914x2438"),
                   (1219, 2438, "1219x2438")
               };

            const double gapBetweenParts = 5.0;
            const double extraGapLongParts = 55.0;

            var parts = listZam.Select(s => (s.WidthZam, s.LengthZam)).ToList();

            bool FitsAnySheet(double w, double h)
            {
                double adjW = w > 1600 ? w : w + (h > 1600 ? extraGapLongParts : 0);
                double adjH = h > 1600 ? h : h + (w > 1600 ? extraGapLongParts : 0);

                return sizes.Any(sheet =>
                    (adjW <= sheet.width && adjH <= sheet.height) ||
                    (adjH <= sheet.width && adjW <= sheet.height));
            }

            for (int i = 0; i < parts.Count;)
            {
                var (w, h) = parts[i];

                if (FitsAnySheet(w, h))
                {
                    i++;
                    continue;
                }

                bool canSplit = false;

                for (int n = 2; n <= 10; n++)
                {
                    if (h >= w)
                    {
                        double splitH = h / n;
                        if (FitsAnySheet(w, splitH))
                        {
                            parts.RemoveAt(i);
                            for (int j = 0; j < n; j++)
                                parts.Insert(i + j, (w, splitH));
                            canSplit = true;
                            break;
                        }
                    }
                    else
                    {
                        double splitW = w / n;
                        if (FitsAnySheet(splitW, h))
                        {
                            parts.RemoveAt(i);
                            for (int j = 0; j < n; j++)
                                parts.Insert(i + j, (splitW, h));
                            canSplit = true;
                            break;
                        }
                    }
                }

                if (!canSplit)
                    throw new Exception($"Cannot split part ({w}x{h}) to fit the sheet.");
            }

            SheetSplitResult bestResult = null;
            parts = parts.Select(p => p.LengthZam >= p.WidthZam ? p : (p.WidthZam, p.LengthZam)).ToList();
            if (parts.Count > 0)
            {
                foreach (var sheet in sizes)
                {
                    double sw = sheet.height;
                    double sh = sheet.width;

                    int totalParts = parts.Count;

                    double areaUsed = 0;
                    int totalSheets = int.MaxValue;
                    double gapY = 0;

                    foreach (var orientation in parts)
                    {
                        double pw;
                        double ph;
                        if (orientation.LengthZam < sh)
                        {
                            ph = orientation.LengthZam;
                            pw = orientation.WidthZam;
                        }
                        else
                        {
                            pw = orientation.LengthZam;
                            ph = orientation.WidthZam;
                        }
                        if (pw >= 1600 || (totalParts > 1 && pw < (sw - gapBetweenParts) / 2))
                        {
                            gapY = extraGapLongParts;
                        }
                        else
                        {
                            gapY = 0;
                        }

                        int partsPerRow = (int)Math.Floor((sh - gapY) / (ph));
                        int partsPerCol = (int)Math.Floor((sw) / (pw));
                        if ((ph * partsPerRow + gapBetweenParts * (partsPerRow - 1) + gapY > sh))
                        {
                            partsPerRow = partsPerRow - 1;
                        }
                        if (partsPerRow <= 0 || partsPerCol <= 0)
                            continue;

                        int partsPerSheet = partsPerRow * partsPerCol;
                        if (partsPerSheet == 0) continue;

                        int sheets = (int)Math.Ceiling((double)totalParts / partsPerSheet);
                        double usedArea = sheets * sw * sh;

                        if (bestResult == null ||
                            usedArea < bestResult.TotalAreaUsed || (usedArea == bestResult.TotalAreaUsed && sheets < bestResult.TotalSheets))
                        {
                            bestResult = new SheetSplitResult
                            {
                                TotalSheets = sheets,
                                SplitX = partsPerRow,
                                SplitY = partsPerCol,
                                SheetWidth = sw,
                                SheetHeight = sh,
                                SheetSizeUsed = sheet.name,
                                GapX = (partsPerRow - 1) * gapBetweenParts,
                                GapY = (partsPerCol - 1) * gapBetweenParts + gapY,
                                TotalAreaUsed = usedArea,
                                AreaNesting = partsPerSheet * sheets * ((pw * (ph + gapY + gapBetweenParts * (partsPerRow - 1))))
                            };
                        }
                    }
                }
            }

            if (bestResult == null)
                throw new Exception("Cannot arrange parts into any sheet size, even after splitting.");

            return bestResult;
        }
    }
}

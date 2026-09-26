using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.OrTools.Sat;

namespace DWG2PDF
{
    internal class Binparking
    {
    }
    public class Item
    {
        public double Length { get; set; }
        public double Width { get; set; }

        public int IntLength => (int)Math.Ceiling(Length);
        public int IntWidth => (int)Math.Ceiling(Width);
    }

    public class Bin
    {
        public int Id { get; set; } // unique identifier
        public int Length { get; set; }
        public int Width { get; set; }
    }



public static class MultiBinPacking
    {
        public static Dictionary<int, int> Optimize(
            List<Item> items,
            List<Bin> bins,
            int gap)
        {
            int nItems = items.Count;
            int nBinTypes = bins.Count;
            int maxBins = nItems; // worst case

            var model = new CpModel();

            // Biến xác định bin type & bin index mỗi item dùng
            IntVar[] binType = new IntVar[nItems];
            IntVar[] binIndex = new IntVar[nItems];

            // Biến vị trí & xoay
            IntVar[] x = new IntVar[nItems];
            IntVar[] y = new IntVar[nItems];
            BoolVar[] rotated = new BoolVar[nItems];

            for (int i = 0; i < nItems; i++)
            {
                binType[i] = model.NewIntVar(0, nBinTypes - 1, $"binType_{i}");
                binIndex[i] = model.NewIntVar(0, maxBins - 1, $"binIndex_{i}");
                x[i] = model.NewIntVar(0, 5000, $"x_{i}"); // giới hạn tạm
                y[i] = model.NewIntVar(0, 5000, $"y_{i}");
                rotated[i] = model.NewBoolVar($"rotated_{i}");
            }

            // Tạo biến kích thước thực (phụ thuộc xoay)
            IntVar[] w = new IntVar[nItems];
            IntVar[] h = new IntVar[nItems];

            for (int i = 0; i < nItems; i++)
            {
                var item = items[i];
                int len = item.IntLength + gap;
                int wid = item.IntWidth + gap;
                w[i] = model.NewIntVar(0, 5000, $"w_{i}");
                h[i] = model.NewIntVar(0, 5000, $"h_{i}");

                model.Add(w[i] == len).OnlyEnforceIf(rotated[i].Not());
                model.Add(w[i] == wid).OnlyEnforceIf(rotated[i]);

                model.Add(h[i] == wid).OnlyEnforceIf(rotated[i].Not());
                model.Add(h[i] == len).OnlyEnforceIf(rotated[i]);
            }

            // Ràng buộc không vượt quá kích thước bin
            for (int i = 0; i < nItems; i++)
            {
                for (int j = 0; j < nBinTypes; j++)
                {
                    var bin = bins[j];
                    var isThisType = model.NewBoolVar($"isType_{i}_{j}");
                    model.Add(binType[i] == j).OnlyEnforceIf(isThisType);
                    model.Add(x[i] + w[i] <= bin.Length).OnlyEnforceIf(isThisType);
                    model.Add(y[i] + h[i] <= bin.Width).OnlyEnforceIf(isThisType);
                }
            }

            // Ràng buộc không chồng nếu cùng binIndex + binType
            for (int i = 0; i < nItems; i++)
            {
                for (int j = i + 1; j < nItems; j++)
                {
                    var sameBin = model.NewBoolVar($"sameBin_{i}_{j}");
                    model.Add(binType[i] == binType[j]).OnlyEnforceIf(sameBin);
                    model.Add(binIndex[i] == binIndex[j]).OnlyEnforceIf(sameBin);

                    var noOverlap = new List<ILiteral>
                {
                    model.NewBoolVar($"left_{i}_{j}"),
                    model.NewBoolVar($"right_{i}_{j}"),
                    model.NewBoolVar($"above_{i}_{j}"),
                    model.NewBoolVar($"below_{i}_{j}")
                };

                    model.Add(x[i] + w[i] <= x[j]).OnlyEnforceIf(noOverlap[0]);
                    model.Add(x[j] + w[j] <= x[i]).OnlyEnforceIf(noOverlap[1]);
                    model.Add(y[i] + h[i] <= y[j]).OnlyEnforceIf(noOverlap[2]);
                    model.Add(y[j] + h[j] <= y[i]).OnlyEnforceIf(noOverlap[3]);

                    model.AddBoolOr(noOverlap).OnlyEnforceIf(sameBin);
                }
            }

            // Mục tiêu: giảm tổng số bin đã dùng trên từng loại
            List<BoolVar> binUsed = new List<BoolVar>();
            for (int bType = 0; bType < nBinTypes; bType++)
            {
                for (int bIdx = 0; bIdx < maxBins; bIdx++)
                {
                    var used = model.NewBoolVar($"used_{bType}_{bIdx}");
                    binUsed.Add(used);

                    for (int i = 0; i < nItems; i++)
                    {
                        var assigned = model.NewBoolVar($"assign_{i}_{bType}_{bIdx}");
                        model.Add(binType[i] == bType).OnlyEnforceIf(assigned);
                        model.Add(binIndex[i] == bIdx).OnlyEnforceIf(assigned);
                        model.AddImplication(assigned, used);
                    }
                }
            }

            model.Minimize(LinearExpr.Sum(binUsed));

            // Solve
            CpSolver solver = new CpSolver();
            solver.StringParameters = "max_time_in_seconds:30.0";

            var status = solver.Solve(model);
            var result = new Dictionary<int, int>(); // binType => used count

            if (status == CpSolverStatus.Optimal || status == CpSolverStatus.Feasible)
            {
                for (int b = 0; b < nBinTypes; b++)
                    result[bins[b].Id] = 0;

                var usedBins = new HashSet<(int binType, int binIndex)>();

                for (int i = 0; i < nItems; i++)
                {
                    int bt = (int)solver.Value(binType[i]);
                    int bi = (int)solver.Value(binIndex[i]);
                    var key = (bt, bi);
                    if (!usedBins.Contains(key))
                    {
                        usedBins.Add(key);
                        result[bins[bt].Id]++;
                    }
                }
            }

            return result;
        }
    }

}

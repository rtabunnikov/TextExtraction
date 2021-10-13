using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.Spreadsheet;

namespace TextExtraction {
    class Program {
        static void Main(string[] args) {
            Stopwatch sw = new Stopwatch();
            using (Workbook workbook = new Workbook()) {
                sw.Start();
                workbook.LoadDocument("sample.xlsx");
                sw.Stop();
                var loadTime = sw.Elapsed;
                sw.Restart();

                // Uncomment if you need unique strings
                //var query = GetCellDisplayText(workbook)
                //    .Union(GetChartTitles(workbook))
                //    .Union(GetShapeText(workbook));

                // Comment if you don't need unique strings
                var query = GetCellDisplayText(workbook)
                    .Concat(GetChartTitles(workbook))
                    .Concat(GetShapeText(workbook));

                foreach (string str in query)
                    Console.WriteLine(str);

                sw.Stop();
                var extractTime = sw.Elapsed;
                Console.WriteLine($"Load {loadTime}");
                Console.WriteLine($"Extract {extractTime}");
                Console.ReadLine();
            }
        }

        static IEnumerable<string> GetCellTextOnly(Workbook workbook) =>
            workbook.Worksheets.SelectMany(x => x.GetExistingCells()
                .Where(c => c.Value.IsText)
                .Select(c => c.Value.TextValue));

        static IEnumerable<string> GetCellDisplayText(Workbook workbook) =>
            workbook.Worksheets.SelectMany(x => x.GetExistingCells().Select(c => c.DisplayText));

        static IEnumerable<string> GetShapeText(Workbook workbook) =>
             workbook.Worksheets.SelectMany(x => x.Shapes
                .Flatten()
                .Where(s => s.ShapeType == ShapeType.Shape && s.ShapeText.HasText)
                .Select(s => s.ShapeText.Characters().Text));

        static IEnumerable<string> GetChartTitles(Workbook workbook) =>
            workbook.Worksheets.SelectMany(x => x.Charts.Select(c => c.Title.PlainText));
    }
}

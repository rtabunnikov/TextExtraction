using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.Spreadsheet;

namespace TextExtraction {
    class Program {
        static void Main(string[] args) {
            using (Workbook workbook = new Workbook()) {
                workbook.LoadDocument("sample.xlsx");
                var query = GetCellDisplayText(workbook)
                    .Union(GetChartTitles(workbook))
                    .Union(GetShapeText(workbook));
                foreach (string str in query)
                    Console.WriteLine(str);
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
                .Where(s => s.ShapeType == ShapeType.Shape && s.ShapeText.HasText)
                .Select(s => s.ShapeText.Characters().Text));

        static IEnumerable<string> GetChartTitles(Workbook workbook) =>
            workbook.Worksheets.SelectMany(x => x.Charts.Select(c => c.Title.PlainText));
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Linq;

namespace exceltk
{
    public static class ExcelExtension
    {
        public class MarkDownTable
        {
            public string Name{get;set;}
            public string Value{get;set;} 
        }

        public static MarkDownTable ToMd(this string xls, string sheet)
        {
            var workBook = new XLWorkbook(xls);
            var worksheet = workBook.Worksheet(sheet);

            if (worksheet == null) return null;

            return new MarkDownTable
            {
                Name = worksheet.Name,
                Value = worksheet.ToMd()
            };
        }
        public static IEnumerable<MarkDownTable> ToMd(this string xls)
        {
            var workBook = new XLWorkbook(xls);
            var worksheets = workBook.Worksheets;
            foreach (var sheet in worksheets)
            {
                var  table = new MarkDownTable
                {
                    Name = sheet.Name,
                    Value = sheet.ToMd()
                };

                yield return table;
            }
        }

        private static string ToMd(this IXLWorksheet worksheet)
        {
            try
            {
                var excelRows = worksheet.RowsUsed().Select((row, index) => new { Row = row, Index = index });

                var mdRows =
                from xrow in excelRows
                let xr = xrow.Row
                let i = xrow.Index
                let mr =
                    from xc in xr.Cells()
                    let v = xc.HasHyperlink
                        ? string.Format("[{0}]({1})", xc.Value, xc.Hyperlink.ExternalAddress.AbsoluteUri)
                        : xc.Value.ToString().Replace(" ", "&nbsp;").Replace("`","<code>").Replace("\r\n","<br/>")
                    select string.Format("{0}|", v)
                let mrr = i == 0
                    ? string.Format("|{0}\r\n|{1}\r\n", mr.Aggregate((a, b) => a + b), mr.Aggregate("", (s, a) => s + ":--:|"))
                    : string.Format("|{0}\r\n", mr.Aggregate((a, b) => a + b))
                select mrr;

                var mt = mdRows.Aggregate((a, b) => a + b);
                return mt.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return "";
            }
        }
    }
}

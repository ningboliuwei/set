using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Linq;
using System.IO;
using System.Data;
using Excel;
using Excel.Core;

namespace exceltk
{
    public static class ExcelExtension
    {
        public class MarkDownTable
        {
            public string Name{get;set;}
            public string Value{get;set;} 
        }

        public static MarkDownTable ToMd(this string xlsx, string sheet)
        {
            var ext = Path.GetExtension(xlsx);
            if (ext == ".xls")
            {
                return xlsx.XlsToMd(sheet);
            }
            else if (ext == "xlsx")
            {
                return xlsx.XlsxToMd(sheet);
            }
            else
            {
                throw new ArgumentException("Not Support Formart");
            }
        }

        public static IEnumerable<MarkDownTable> ToMd(this string xlsx)
        {
            var ext = Path.GetExtension(xlsx);
            if (ext == ".xls")
            {
                return xlsx.XlsToMd();
            }
            else if (ext == ".xlsx")
            {
                return xlsx.XlsxToMd();
            }
            else
            {
                throw new ArgumentException("Not Support Formart");
            }
        }

        public static MarkDownTable XlsxToMd(this string xlsx, string sheet)
        {
            var workBook = new XLWorkbook(xlsx);
            var worksheet = workBook.Worksheet(sheet);

            if (worksheet == null) return null;

            return new MarkDownTable
            {
                Name = worksheet.Name,
                Value = worksheet.ToMd()
            };
        }

        public static IEnumerable<MarkDownTable> XlsxToMd(this string xlsx)
        {
            var workBook = new XLWorkbook(xlsx);
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

        public static MarkDownTable XlsToMd(this string xls, string sheet)
        {
            FileStream stream = File.Open(xls, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            DataSet dataSet = excelReader.AsDataSet();
            DataTable dataTable = dataSet.Tables[sheet];

            var table = new MarkDownTable
            {
                Name = dataTable.TableName,
                Value = dataTable.ToMd()
            };

            excelReader.Close();

            return table;
        }

        public static IEnumerable<MarkDownTable> XlsToMd(this string xls)
        {
            FileStream stream = File.Open(xls, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            DataSet dataSet = excelReader.AsDataSet();

            foreach (DataTable dataTable in dataSet.Tables)
            {
                var table = new MarkDownTable
                {
                    Name = dataTable.TableName,
                    Value = dataTable.ToMd()
                };

                yield return table;
            }

            excelReader.Close();
        }

        private static string ToMd(this DataTable table)
        {
            var sb = new StringBuilder();

            int i = 0;
            foreach (DataRow row in table.Rows)
            {
                sb.Append("|");
                foreach (var cell in row.ItemArray)
                {
                    string value = "";
                    XlsCell xlsCell = cell as XlsCell;
                    if (xlsCell != null)
                    {
                        value = xlsCell.MarkDownText;
                    }
                    else
                    {
                        value = cell.ToString();
                    }

                    //var value = cell.ToString().Replace(" ", "&nbsp;").Replace("\r\n","<br/>").Replace("\n","<br/>").Replace("|","");
                    sb.Append(value).Append("|");
                }
                sb.Append("\r\n");
                if (i == 0)
                {
                    sb.Append("|");
                    foreach (DataColumn col in table.Columns)
                    {
                        sb.Append(":--|");
                    }
                    sb.Append("\r\n");
                }
                i++;
            }
            return sb.ToString();
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
                        : xc.Value.ToString().Replace(" ", "&nbsp;").Replace("\r\n","<br/>")
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

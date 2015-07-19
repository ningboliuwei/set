﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Data;
using System.Globalization;
using System.IO;

namespace ExcelToolKit
{
    public static class Extension
    {
        public class MarkDownTable
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }

        public static MarkDownTable ToMd(this string xlsx, string sheet)
        {
            var ext = Path.GetExtension(xlsx);
            return xlsx.XlsToMd(sheet);
        }

        public static IEnumerable<MarkDownTable> ToMd(this string xlsx)
        {
            var ext = Path.GetExtension(xlsx);
            return xlsx.XlsToMd();
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
            IExcelDataReader excelReader = null;
            if (Path.GetExtension(xls) == ".xls")
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (Path.GetExtension(xls) == ".xlsx")
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            else
            {
                throw new ArgumentException("Not Support Format: ");
            }
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

        public static bool IsSingleByteEncoding(this Encoding encoding)
        {
            return encoding.IsSingleByte;
        }
        public static double Int64BitsToDouble(this long value)
        {
            return BitConverter.ToDouble(BitConverter.GetBytes(value), 0);
        }

        private static Regex re = new Regex("_x([0-9A-F]{4,4})_");

        public static string ConvertEscapeChars(this string input)
        {
            return re.Replace(input, m => (((char)UInt32.Parse(m.Groups[1].Value, NumberStyles.HexNumber))).ToString());
        }
        public static object ConvertFromOATime(this double value)
        {
            if ((value >= 0.0) && (value < 60.0))
            {
                value++;
            }
            return DateTime.FromOADate(value);
        }

        public static void FixDataTypes(this DataSet dataset)
        {
            var tables = new List<DataTable>(dataset.Tables.Count);
            bool convert = false;
            foreach (DataTable table in dataset.Tables)
            {

                if (table.Rows.Count == 0)
                {
                    tables.Add(table);
                    continue;
                }
                DataTable newTable = null;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    Type type = null;
                    foreach (DataRow row in table.Rows)
                    {
                        if (row.IsNull(i))
                            continue;
                        var curType = row[i].GetType();
                        if (curType != type)
                        {
                            if (type == null)
                                type = curType;
                            else
                            {
                                type = null;
                                break;
                            }
                        }
                    }
                    if (type != null)
                    {
                        convert = true;
                        if (newTable == null)
                            newTable = table.Clone();
                        newTable.Columns[i].DataType = type;

                    }
                }
                if (newTable != null)
                {
                    newTable.BeginLoadData();
                    foreach (DataRow row in table.Rows)
                    {
                        newTable.ImportRow(row);
                    }

                    newTable.EndLoadData();
                    tables.Add(newTable);

                }
                else tables.Add(table);
            }
            if (convert)
            {
                dataset.Tables.Clear();
                dataset.Tables.AddRange(tables.ToArray());
            }
        }
        public static void AddColumnHandleDuplicate(this DataTable table, string columnName)
        {
            //if a colum  already exists with the name append _i to the duplicates
            var adjustedColumnName = columnName;
            var column = table.Columns[columnName];
            var i = 1;
            while (column != null)
            {
                adjustedColumnName = string.Format("{0}_{1}", columnName, i);
                column = table.Columns[adjustedColumnName];
                i++;
            }

            table.Columns.Add(adjustedColumnName, typeof(Object));
        }
        public static int[] ReferenceToColumnAndRow(this string reference)
        {
            //split the string into row and column parts


            Regex matchLettersNumbers = new Regex("([a-zA-Z]*)([0-9]*)");
            string column = matchLettersNumbers.Match(reference).Groups[1].Value.ToUpper();
            string rowString = matchLettersNumbers.Match(reference).Groups[2].Value;

            //.net 3.5 or 4.5 we could do this awesomeness
            //return reference.Aggregate(0, (s,c)=>{s*26+c-'A'+1});
            //but we are trying to retain 2.0 support so do it a longer way
            //this is basically base 26 arithmetic
            int columnValue = 0;
            int pow = 1;

            //reverse through the string
            for (int i = column.Length - 1; i >= 0; i--)
            {
                int pos = column[i] - 'A' + 1;
                columnValue += pow * pos;
                pow *= 26;
            }

            return new int[2] { int.Parse(rowString), columnValue };
        }
    }
}

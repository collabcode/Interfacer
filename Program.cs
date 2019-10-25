using System;
using System.IO;
using CsvHelper;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Data.Common;
using System.Data;

namespace Interfacer
{
    class Program
    {
        public static List<string> ReadInCSV(string absolutePath)
        {
            List<string> result = new List<string>();
            string value;
            using (TextReader fileReader = File.OpenText(absolutePath))
            {
                var csv = new CsvReader(fileReader);
                csv.Configuration.HasHeaderRecord = false;

                var itr = 0;

                while (csv.Read())
                {
                    for (int i = 0; csv.TryGetField<string>(i, out value); i++)
                    {
                        result.Add(value);

                    }
                    result.Add("-|-");
                    itr++;
                }
            }
            return result;
        }
        public static string listToJSON(List<string> l)
        {
            List<string> headers = new List<string>();
            List<string> fieldDataType = new List<string>();

            string jsonStr = "[";
            int numColumns = 1; //considering line break as an item in the row
            int numRows = 0;
            foreach (var item in l)
            {
                if (item == "-|-")
                    break;
                headers.Add(item);
                numColumns++;
            }
            if (l.Count % numColumns == 0)
            {
                numRows = l.Count / numColumns; //first row is header row

                for (int i = 0; i < numColumns; i++)
                {
                    for (int j = 0; j < numRows; j++)
                    {
                        var val = l[i * j];
                        var type = "string";
                        if (Int32.TryParse(val, out int numValue))
                        {
                            type = "int";
                        }
                        else
                        {
                            if (Double.TryParse(val, out Double doubleValue))
                            {
                                type = "double";
                            }
                            else
                            {
                                if (DateTime.TryParse(val, out DateTime datetimeValue))
                                {
                                    type = "date";
                                }
                                else
                                {
                                    type = "string";
                                }
                            }
                        }

                    }
                }
            }
            else
            {
                Console.WriteLine("Item count mismatch");
            }
            jsonStr = jsonStr.Replace("\",}", "\"}").Replace("{},{", "{").Replace("},]", "}]");
            return jsonStr;
        }

        private static DataTable GetDataTableFromExcel(string path, bool hasHeader = true, int sheetNum = 0)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets[sheetNum];
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column { 0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            DataTable dtContent = GetDataTableFromExcel("temp//Definition.xlsx", true, 0);
            foreach (var row in dtContent.Rows)
            {
                Console.WriteLine(row);
            }

            //var result = ReadInCSV("temp//test.csv");
            //var jsonStr = listToJSON(result);
            //Console.WriteLine(jsonStr);

        }
    }
}

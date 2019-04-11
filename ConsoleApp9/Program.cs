using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
namespace ConsoleApp9
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            string con = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Santhosh\tifFileDetails.xlsx; Extended Properties = 'Excel 12.0 Xml;HDR=YES;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {

                OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "select * from [Sheet1$]";
                    comm.Connection = connection;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);

                    }

                }
            }

            if (dt.Rows.Count > 0)
            {
                DataTable workTable = new DataTable("TIFfiles");

                workTable.Columns.Add("RequestId", typeof(String));
                workTable.Columns.Add("RecordId", typeof(String));
                workTable.Columns.Add("PageCount", typeof(int));
                // List<int> vs = new List<int>();
                foreach (DataRow item in dt.AsEnumerable())
                {
                    string[] files = Directory.GetFiles(item[1].ToString());
                    foreach (var file in files)
                    {
                        if (file.EndsWith(".tif"))
                        {
                            Stream imageStreamSource = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read);
                            TiffBitmapDecoder decoder = new TiffBitmapDecoder(imageStreamSource, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
                            // vs.Add(decoder.Frames.Count);
                            DataRow newCustomersRow = workTable.NewRow();

                            newCustomersRow["RequestId"] = item[0];
                            newCustomersRow["RecordId"] = Path.GetFileName(file);
                            newCustomersRow["PageCount"] = decoder.Frames.Count;
                            workTable.Rows.Add(newCustomersRow);
                            imageStreamSource.Close();
                        }
                    }

                }

                StringBuilder sb = new StringBuilder();
                //adding header
                sb.Append("RequestId,RecordId,PageCount");
                sb.AppendLine();
                foreach (DataRow dr in workTable.Rows)
                {
                    foreach (DataColumn dc in workTable.Columns)
                        sb.Append(FormatCSV(dr[dc.ColumnName].ToString()) + ",");
                    sb.Remove(sb.Length - 1, 1);
                    sb.AppendLine();
                }
                File.WriteAllText("C:\\Santhosh\\tfiresult.csv", sb.ToString());
            }
            else
            {
                Console.WriteLine("No records found");
            }
            //DataSet dsResult = new DataSet();
            //dsResult.Tables.Add(workTable);
            // ExportDataSet(dsResult, "C:\\Santhosh\\tfiresult.xlsx");

        }

        public static string FormatCSV(string input)
        {
            try
            {
                if (input == null)
                    return string.Empty;

                bool containsQuote = false;
                bool containsComma = false;
                int len = input.Length;
                for (int i = 0; i < len && (containsComma == false || containsQuote == false); i++)
                {
                    char ch = input[i];
                    if (ch == '"')
                        containsQuote = true;
                    else if (ch == ',')
                        containsComma = true;
                }

                if (containsQuote && containsComma)
                    input = input.Replace("\"", "\"\"");

                if (containsComma)
                    return "\"" + input + "\"";
                else
                    return input;
            }
            catch
            {
                throw;
            }
        }


        //public static void ExportDataSet(DataSet ds, string destination)
        //{
        //    using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        //    {
        //        var workbookPart = workbook.AddWorkbookPart();

        //        workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

        //        workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

        //        foreach (System.Data.DataTable table in ds.Tables)
        //        {

        //            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
        //            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
        //            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

        //            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
        //            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

        //            uint sheetId = 1;
        //            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
        //            {
        //                sheetId =
        //                    sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        //            }

        //            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
        //            sheets.Append(sheet);

        //            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

        //            List<String> columns = new List<string>();
        //            foreach (System.Data.DataColumn column in table.Columns)
        //            {
        //                columns.Add(column.ColumnName);

        //                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
        //                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
        //                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
        //                headerRow.AppendChild(cell);
        //            }


        //            sheetData.AppendChild(headerRow);

        //            foreach (System.Data.DataRow dsrow in table.Rows)
        //            {
        //                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
        //                foreach (String col in columns)
        //                {
        //                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
        //                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
        //                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
        //                    newRow.AppendChild(cell);
        //                }

        //                sheetData.AppendChild(newRow);
        //            }

        //        }
        //    }
        //}
    }
}

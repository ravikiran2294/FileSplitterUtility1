using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Net.Http.Headers;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileSplitterUtility1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter file path: ");
            string filePath = Console.ReadLine();
			Console.WriteLine("Enter variable 1: ");
			string value1 = Console.ReadLine();
			Console.WriteLine("Enter date: ");
			string date = Console.ReadLine();
			//var excelData = ReadExcelDataUsingOledb(filePath);
			var excelData = ReadExcelFileUsingClosedXml(filePath);
            var dataSet = SplitDataAndWriteToFile(excelData);
            
            string outFilePath = WriteDataSetToExcel(dataSet, filePath, value1, date);
            Console.WriteLine($"Output files generated at : {outFilePath}");
            Console.WriteLine("Press any key to close this window..");
            Console.ReadLine();
        }

        static DataTable ReadExcelDataUsingOledb(string filePath)
        {
            DataTable dt = new DataTable();
            using (OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0; HDR=YES;IMEX=1;';"))
            {
                con.Open();
                try
                {
                    List<DataRow> sheetNames = con.GetSchema("Tables").AsEnumerable().ToList<DataRow>();
                    foreach (var sheetName in sheetNames)
                    {
                        OleDbDataAdapter oda = new OleDbDataAdapter(" SELECT * FROM [" + sheetName["TABLE_NAME"] + "]", con);
                        oda.Fill(dt);
                    }
                    con.Close();
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message + Environment.NewLine + ex.StackTrace);
                }
            }
            return dt;
        }

        static DataSet SplitDataAndWriteToFile(DataTable data)
        {
            try
            {
                DataSet ds = new DataSet();
                //var distinctFirstColVals = data.AsDataView().OfType<DataRow>().Select(x => x[0].ToString()).Distinct().ToList();
                List<string> distVals = new List<string>();
                foreach (DataRow row in data.Rows)
                {
                    string value = row[0].ToString();
                    if (!distVals.Contains(value))
                    {
                        distVals.Add(value);
                    }
                }

                foreach (var item in distVals)
                {
                    var dt = data.Clone();
                    dt.TableName = item.ToString();
                    foreach (DataRow row in data.Rows)
                    {
                        string value = row[0].ToString();
                        if (item == value)
                        {
                            //dt.ImportRow(row);
                            dt.Rows.Add(row.ItemArray);
                        }
                    }
                    ds.Tables.Add(dt);
                }
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        static string WriteDataSetToExcel(DataSet ds, string filePath, string value1, string date)
        {
            filePath = filePath.Trim('\"');
            var outFileDir = Path.GetDirectoryName(filePath) + Path.DirectorySeparatorChar.ToString() + "outputFiles";
            try
            {
                if (Directory.Exists(outFileDir))
                {
                    Directory.Delete(outFileDir, false);
                }

                foreach (DataTable dataTable in ds.Tables)
                {
                    XLWorkbook wb = new XLWorkbook();
                    wb.AddWorksheet(dataTable, "sheet-"+dataTable.TableName);
                    string outputFilePath = string.Concat(outFileDir, Path.DirectorySeparatorChar, value1, "_", dataTable.TableName, "_", date, ".xlsx");
                    wb.SaveAs(outputFilePath);
                }
                return outFileDir;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }


        static DataTable ReadExcelFileUsingClosedXml(string filePath)
        {
            try
            {
                DataTable dt = new DataTable();
                using (XLWorkbook workBook = new XLWorkbook(filePath))
                {
                    IXLWorksheet workSheet1 = workBook.Worksheets.FirstOrDefault();
                    //foreach (IXLWorksheet workSheet in workBook.Worksheets)
                    //{ }
                    dt = new DataTable(workSheet1.Name);

                    // Read First Row of Excel Sheet to add Columns to DataTable
                    workSheet1.FirstRowUsed().CellsUsed().ToList()
                    .ForEach(x => { dt.Columns.Add(x.Value.ToString()); });

                    foreach (IXLRow row in workSheet1.RowsUsed().Skip(1))
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            dr[i] = row.Cell(i + 1).Value.ToString();
                        }
                        dt.Rows.Add(dr);
                    }

                }
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

    }
}
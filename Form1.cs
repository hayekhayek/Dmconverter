using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Style;


//https://itenium.be/blog/dotnet/create-xlsx-excel-with-epplus-csharp/
//https://www.programmersought.com/article/4453654835/
//https://www.xspdf.com/resolution/52117722.html
namespace dmconverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        public object Assert { get; private set; }

        private void button1_Click(object sender, EventArgs e)
        {

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "Feri";
                excelPackage.Workbook.Properties.Title = "Data Manager";
                excelPackage.Workbook.Properties.Created = DateTime.Now;
                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("FeriSpec");
                ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets.Add("Tabelle1");

                ExcelRange range = worksheet1.Cells["A1"];           // ws is the worksheet name
                //ws.Cells["A1"].Value = "Name";
                string sNamedRange = "FeriTable1";
                //worksheet1.Names.Add(sNamedRange, range);
                excelPackage.Workbook.Names.Add(sNamedRange, range);

                //Standards Cells
                worksheet.Cells["A1"].Value = "TableSpec";
                worksheet.Cells["A2"].Value = "CoreSpec";
                worksheet.Cells["A3"].Value = "TimeSpecs";
                worksheet.Cells["A4"].Value = "Info";
                worksheet.Cells["A5"].Value = "Info";
                worksheet.Cells["A6"].Value = "Info";
                worksheet.Cells["A7"].Value = "TimeFrame";

                worksheet.Cells["A9"].Value = "DataSpec";
                worksheet.Cells["A10"].Value = "DataSpec";
                worksheet.Cells["A11"].Value = "End";

                worksheet.Cells["B1"].Value = "FeriTable1";

                worksheet.Cells["B2"].Value = "Title=Unbenannt";
                worksheet.Cells["C2"].Value = "Type=Col";
                worksheet.Cells["D2"].Value = "UseSettings=Yes";
                worksheet.Cells["E2"].Value = "ExcelFormat=Yes";
                worksheet.Cells["F2"].Value = "DataForeColor=4194432";
                worksheet.Cells["G2"].Value = "DataBackColor=12632256";
                worksheet.Cells["H2"].Value = "TitleForeColor=12632256";
                worksheet.Cells["I2"].Value = "TitleBackColor=8404992";
                worksheet.Cells["J2"].Value = "IntegratedSpec=No";
                worksheet.Cells["K2"].Value = "RecreateView=Yes";
                worksheet.Cells["B4"].Value = "Type=Title";
                worksheet.Cells["B5"].Value = "Type=DataField";
                worksheet.Cells["B6"].Value = "Type=Transform";
                worksheet.Cells["B7"].Value = "Type=Customized";
                worksheet.Cells["C5"].Value = "DataField=Description";
                worksheet.Cells["C7"].Value = "TargetFrequency=Monthly";
                worksheet.Cells["D7"].Value = "PeriodsAscending=Yes";

                worksheet.Cells["A8"].Value = "DataSpecs";
                worksheet.Cells["B8"].Value = "Type";
                worksheet.Cells["C8"].Value = "Database";
                worksheet.Cells["D8"].Value = "Table";
                worksheet.Cells["E8"].Value = "Series";
                worksheet.Cells["F8"].Value = "Trans";
                worksheet.Cells["G8"].Value = "SeriesTitle";
                worksheet.Cells["H8"].Value = "SeriesSubtitle";
                worksheet.Cells["I8"].Value = "Aggregation";
                worksheet.Cells["J8"].Value = "Disaggregation";
                worksheet.Cells["K8"].Value = "UpdateControl";
                worksheet.Cells["L8"].Value = "ExportControl";
                worksheet.Cells["M8"].Value = "Precision";

                //DM_web Excel Reading 

                String readfile = "O:\\Mitarbeiter\\Henninger, Dirk\\Allgemein\\Studenten\\Jihad\\!DM_Converter\\dmconverter\\DM_Web.xlsx";

                FileInfo fileInfo = new FileInfo(readfile);
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet WS = package.Workbook.Worksheets.FirstOrDefault();
                int rows = worksheet.Dimension.Rows; // 20
                int columns = worksheet.Dimension.Columns; // 7

                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {

                        string content = worksheet.Cells[i, j].Value.ToString();
                        /* Do something ...*/
                    }
                    //worksheet.Cells.AutoFitColumns();

                    //Save your file
                    FileInfo fi = new FileInfo(@"O:\Mitarbeiter\Henninger, Dirk\Allgemein\Studenten\Jihad\!DM_Converter\dmconverter\bin\Debug\File.xlsx");
                    excelPackage.SaveAs(fi);
                }
            }
            //FileInfo fi_ = new FileInfo(@"D:\für Arbeit\DmConverter\dmconverter\bin\Debug\File.xlsx");
            //using (ExcelPackage excelPackage = new ExcelPackage(fi_))
            //{
            //    //Get a WorkSheet by index. Note that EPPlus indexes are base 1, not base 0!
            //    ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];
            //    //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
            //    ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["SomeWorksheet"];
            //    //If you don't know if a worksheet exists, you could use LINQ,
            //    //So it doesn't throw an exception, but return null in case it doesn't find it
            //    ExcelWorksheet anotherWorksheet =
            //    excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SomeWorksheet");
            //    //Get the content from cells A1 and B1 as string, in two different notations
            //    string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
            //    string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();
            //    //Save your file
            //    excelPackage.Save();
            //}



            //the path of the file
            //string filePath = "O:\\Mitarbeiter\\Henninger, Dirk\\Allgemein\\Studenten\\Jihad\\!DM_Converter\\dmconverter\\bin\\DebugFile.xlsx";
            
            ////create a fileinfo object of an excel file on the disk
            //FileInfo file = new FileInfo(filePath);
            ////create a new Excel package from the file
            //using (ExcelPackage excelPackage = new ExcelPackage(file))
            //{
            //    //create an instance of the the first sheet in the loaded file
            //    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];
            //    //add some data
            //    worksheet.Cells[1, 1].Value = "FeriTable1";
            //    worksheet.Cells[1, 2].Value = "CoreSpec";
            //    worksheet.Cells[1, 3].Value = "TimeSpecs";
            //    worksheet.Cells[1, 4].Value = "Info";
            //    worksheet.Cells[1, 5].Value = "Info";
            //    worksheet.Cells[1, 6].Value = "Info";
            //    worksheet.Cells[1, 6].Value = "TimeFrame";
            //    worksheet.Cells[1, 6].Value = "DataSpecs";
            //    worksheet.Cells[1, 6].Value = "DataSpec";
            //    worksheet.Cells[1, 6].Value = "DataSpec";
            //    worksheet.Cells[1, 6].Value = "End";

            //    //save the changes
            //    excelPackage.Save();
            //}



            //create a list to hold all the values
            List<string> excelData = new List<string>();
            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes("O:\\Mitarbeiter\\Henninger, Dirk\\Allgemein\\Studenten\\Jihad\\!DM_Converter\\dmconverter\\bin\\Debug\\File.xlsx");
            
            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                //loop all worksheets
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <=
                       worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                excelData.Add(worksheet.Cells[i, j].Value.ToString());
                            }
                        }
                    }
                }
            }

        }
    }
}
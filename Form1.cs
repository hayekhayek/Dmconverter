using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;


//https://itenium.be/blog/dotnet/create-xlsx-excel-with-epplus-csharp/
//https://www.programmersought.com/article/4453654835/
//https://www.xspdf.com/resolution/52117722.html
//https://stackoverflow.com/questions/42042662/c-sharp-trying-to-split-a-string-to-get-json-object-value/50362239
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
                //worksheet.Cells["C2"].Value = "Type=Col";
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
                ExcelWorksheet WS_1 = package.Workbook.Worksheets[1];   //Select sheet FeDaX_Spec
                ExcelWorksheet WS_2 = package.Workbook.Worksheets[2];   //Select sheet FeDaX_Table
                ExcelWorksheet WS_3 = package.Workbook.Worksheets[3];   //Select sheet für das Country als SeriesTitel && für SeriesSubtitle zu builden 


               //WS_2.Cells["C2"].Copy(worksheet.Cells["C2"]);
                worksheet.Cells["C2"].Value = "Type=" + WS_2.Cells["C2"].Value;

                string input = WS_1.Cells["C5"].Text;
                //input = input.Remove('');
                JArray a = JArray.Parse(json);


                FileInfo fi = new FileInfo(@"O:\Mitarbeiter\Henninger, Dirk\Allgemein\Studenten\Jihad\!DM_Converter\dmconverter\bin\Debug\File.xlsx");
                    excelPackage.SaveAs(fi);
                



                //    ExcelWorksheet sheet = package.Workbook.Worksheets.Add("MySheet");

                //    // One cell
                //    ExcelRange cellA2 = sheet.Cells["A2"];
                //    var alsoCellA2 = sheet.Cells[2, 1];
                //    Assert.That(cellA2.Address, Is.EqualTo("A2"));
                //    Assert.That(cellA2.Address, Is.EqualTo(alsoCellA2.Address));

                //    // Column from a cell
                //    // ExcelRange.Start is the top and left most cell
                //    Assert.That(cellA2.Start.Column, Is.EqualTo(1));
                //    // To really get the column: sheet.Column(1)

                //    // A range
                //    ExcelRange ranger = sheet.Cells["A2:C5"];
                //    var sameRanger = sheet.Cells[2, 1, 5, 3];
                //    Assert.That(ranger.Address, Is.EqualTo(sameRanger.Address));

                //    //sheet.Cells["A1,A4"] // Just A1 and A4
                //    //sheet.Cells["1:1"] // A row
                //    //sheet.Cells["A:B"] // Two columns

                //    // Linq
                //    var l = sheet.Cells["A1:A5"].Where(range => range.Comment != null);

                //    // Dimensions used
                //    Assert.That(sheet.Dimension, Is.Null);

                //    ranger.Value = "pushing";
                //    var usedDimensions = sheet.Dimension;
                //    Assert.That(usedDimensions.Address, Is.EqualTo(ranger.Address));

                //    // Offset: down 5 rows, right 10 columns
                //    var movedRanger = ranger.Offset(5, 10);
                //    Assert.That(movedRanger.Address, Is.EqualTo("K7:M10"));
                //    movedRanger.Value = "Moved";

                package.SaveAs(new FileInfo(@""));
                
            }
          
        }
    }
}
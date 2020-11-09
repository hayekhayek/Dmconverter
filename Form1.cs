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
using System.Text.Json;
using System.Text.Json.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;




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
        List<string> listFiles = new List<string>();

        private void convert_button(object sender, EventArgs e)


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
                worksheet.Cells["C7"].Value = "Type=" + WS_2.Cells["C2"].Value;


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
                if (listFiles.Count() == 0)
                {
                    MessageBox.Show("Bitte wählen Sie Datei/en!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    //for (int i = 0; i <= listFiles.Count(); i++)
                    List<string> seriesname = new List<string>();
                    
                    foreach (string listfile in listFiles)
                    {


                            String readfile = listfile;
                            FileInfo fileInfo = new FileInfo(readfile);
                            ExcelPackage package = new ExcelPackage(fileInfo);
                            ExcelWorksheet WS_1 = package.Workbook.Worksheets[1];   //Select sheet FeDaX_Spec
                            ExcelWorksheet WS_2 = package.Workbook.Worksheets[2];   //Select sheet FeDaX_Table
                            ExcelWorksheet WS_3 = package.Workbook.Worksheets[3];   //Select sheet für das Country als SeriesTitel && für SeriesSubtitle zu builden 
                                                                                    //hier ist das Type Col or Row

                            worksheet.Cells["C2"].Value = "Type=" + WS_2.Cells["C2"].Value;


                            string jsonString = WS_1.Cells["C5"].Text.ToString();
                            getData(jsonString);

                            //for (int n = 7; n <= WS_1.Cells.Count(); n++)

                            //int lastRow = WS_1.Cells[7, 3].Where(cell => !cell.Value.ToString().Equals("")).Last().End.Row;

                            for (int n = 7; n <= WS_1.Cells.Last().End.Row; n++)
                            {

                                if (WS_1.Cells.Text != null)
                                {
                                    string cellvalue = WS_1.Cells[n, 3].Value.ToString();
                                    getData(cellvalue);
                                    
                                }

                            }

                        
                        FileInfo fi = new FileInfo(@"D:\für Arbeit\DmConverter\dmconverter\bin\Debug\File.xlsx");
                        excelPackage.SaveAs(fi);

                        //MessageBox.Show("Die Konvertierung war erfolgreich!", "Fertig!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                    }
                    MessageBox.Show("Die Konvertierung war erfolgreich!", "Fertig!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                    
            }
        }

   

        public static void getData(string json)
        {

            Spec jsonc = JsonConvert.DeserializeObject<Spec>(json);

            string TargetFrequency = jsonc.TargetFrequency;
            string PeriodsAscending = jsonc.PeriodsAscending;
            string series = jsonc.SeriesName;
            string TableName = jsonc.TableName;
            
        }

        private void SetText(string text)
        {
            dataGridView1.Text = text;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in openFileDialog1.FileNames)
                {
                    listFiles.Add(file);
                    dataGridView1.Rows.Add(file);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.ExitThread();
        }
    }
}
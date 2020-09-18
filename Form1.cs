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
                excelPackage.Workbook.Properties.Author = "VDWWD";
                excelPackage.Workbook.Properties.Title = "Title of Document";
                excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
                excelPackage.Workbook.Properties.Created = DateTime.Now;
                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                //Add some text to cell A1
                worksheet.Cells["A1"].Value = "My first EPPlus spreadsheet!";
                //You could also use [line, column] notation:
                worksheet.Cells[1, 2].Value = "This is cell B1!";
                //Save your file
                FileInfo fi = new FileInfo(@"D:\für Arbeit\DmConverter\dmconverter\bin\Debug\File.xlsx");
                excelPackage.SaveAs(fi);
            }

            FileInfo fi_ = new FileInfo(@"D:\für Arbeit\DmConverter\dmconverter\bin\Debug\File.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(fi_))
            {
                //Get a WorkSheet by index. Note that EPPlus indexes are base 1, not base 0!
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];
                //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["SomeWorksheet"];
                //If you don't know if a worksheet exists, you could use LINQ,
                //So it doesn't throw an exception, but return null in case it doesn't find it
                ExcelWorksheet anotherWorksheet =
                excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SomeWorksheet");
                //Get the content from cells A1 and B1 as string, in two different notations
                string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
                string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();
                //Save your file
                excelPackage.Save();
            }



            //the path of the file
            string filePath = "D:\\für Arbeit\\DmConverter\\dmconverter\\bin\\Debug\\File.xlsx";
            
            //create a fileinfo object of an excel file on the disk
            FileInfo file = new FileInfo(filePath);
            //create a new Excel package from the file
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                //create an instance of the the first sheet in the loaded file
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];
                //add some data
                worksheet.Cells[4, 1].Value = "Added data in Cell A4";
                worksheet.Cells[4, 2].Value = "Added data in Cell B4";
                //save the changes
                excelPackage.Save();
            }



            //create a list to hold all the values
            List<string> excelData = new List<string>();
            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes("D:\\für Arbeit\\DmConverter\\dmconverter\\bin\\Debug\\File.xlsx");
            
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

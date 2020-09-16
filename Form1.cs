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
			using (var package = new ExcelPackage())
			{
				ExcelWorksheet sheet = package.Workbook.Worksheets.Add("MySheet");

				// Setting & getting values
				ExcelRange firstCell = sheet.Cells[1, 1];
				firstCell.Value = "will it work?";
				sheet.Cells["A2"].Formula = "CONCATENATE(A1,\" ... Of course it will!\")";
				//Assert.That(firstCell.Text, Is.EqualTo("will it work?"));

				// Numbers
				var moneyCell = sheet.Cells["A3"];
				moneyCell.Style.Numberformat.Format = "$#,##0.00";
				moneyCell.Value = 15.25M;

				// Easily write any Enumerable to a sheet
				// In this case: All Excel functions implemented by EPPlus
				var funcs = package.Workbook.FormulaParserManager.GetImplementedFunctions()
					.Select(x => new { FunctionName = x.Key, TypeName = x.Value.GetType().FullName });
				sheet.Cells["A4"].LoadFromCollection(funcs, true);

				// Styling cells
				var someCells = sheet.Cells["A1,A4:B4"];
				someCells.Style.Font.Bold = true;
				someCells.Style.Font.Color.SetColor(Color.Ivory);
				someCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
				someCells.Style.Fill.BackgroundColor.SetColor(Color.Navy);

				sheet.Cells.AutoFitColumns();
				package.SaveAs(new FileInfo(@"basicUsage.xslx"));
			}

		}
    }
}

using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Comparator;

namespace Comparator.FileOpenCount
{
	class FileOC
	{
		public bool CheckOpenFileOUR { get; set; }
		public bool CheckOpenFileKP { get; set; }
		public string Cell { get; set; }
		public string TextCells { get; private set; }
		public static string filenameOurSpecification { get;  set; }
		public static string filenameKPSpecification { get;  set; }
		List<string>  OurSpec = new List<string>();
		public static Worksheet worksheetOurSpecification { get;  set; }
		public static Workbook workbookOurSpecification { get;  set; }
		public static Worksheet worksheetKPSpecification { get; set; }
		public static Workbook workbookKPSpecification { get; set; }
		public static string filename1 { get; set; }
		public static string filename2 { get; set; }

		public void OpenFile()
		{
			try
			{
				Excel.Application ExcelApp = new Excel.Application();
				ExcelApp.Visible = true;
				//ExcelApp.WindowState = XlWindowState.xlMaximized;
				//ExcelApp.DisplayFullScreen = true;
				if (CheckOpenFileOUR)
				{
					workbookOurSpecification = ExcelApp.Workbooks.Open(filenameOurSpecification);
					worksheetOurSpecification = workbookOurSpecification.Worksheets[1];
					filename1 = filenameOurSpecification;

				}
				if (CheckOpenFileKP)
				{
					workbookKPSpecification = ExcelApp.Workbooks.Open(filenameKPSpecification);
					worksheetKPSpecification = workbookKPSpecification.Worksheets[1];
					filename2 = filenameKPSpecification;
				}
				//CheckOpenFileOUR = false;
				//CheckOpenFileKP = false;
				//Excel.Range range = (Excel.Range)worksheet.Cells[1, 1];
				//textBox.Text = range.Value2;
				//Cell = .Text;
				//textBox.Text = worksheet.Range[Cell].Value;

				//ExcelApp.Application.DisplayFullScreen=true;

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
				//Close();
			}
		}
		public void Compare()
		{
			try
			{
				MessageBox.Show(worksheetOurSpecification.Range[1, 1].Value);

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);

			}
			//TextCells = worksheetOurSpecification.Range[1].Value2;
			//MessageBox.Show(worksheetKPSpecification.Range[1,1].Value2);

		}

	}
}

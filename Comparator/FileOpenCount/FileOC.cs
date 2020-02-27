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
		public string Cell { get; set; }
		public string TextCells { get; private set; }
		public string filename { get; private set; }
		List<string>  OurSpec = new List<string>();
		public Worksheet worksheet { get; private set; }


		public void OpenFile()
		{
			try
			{
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Filter = "Excel files|*.xlsx|All files|*.*";
				openFileDialog.DefaultExt = "*.xlsx";
				openFileDialog.Title = "Choose your destiny";
				if (openFileDialog.ShowDialog() == true)
				{
					filename = openFileDialog.FileName;
				}
				Excel.Application ExcelApp = new Excel.Application();
				ExcelApp.Visible = true;
				//ExcelApp.WindowState = XlWindowState.xlMaximized;
				//ExcelApp.DisplayFullScreen = true;
				Workbook workbook;
				workbook = ExcelApp.Workbooks.Open(filename);
				worksheet = workbook.Worksheets[1];
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
			TextCells = worksheet.Range[Cell].Value;

		}

	}
}

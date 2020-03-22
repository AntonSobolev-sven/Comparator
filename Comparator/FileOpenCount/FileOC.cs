using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Comparator.FileOpenCount
{
	
	class FileOC
	{
		List<string> Order = new List<string>();
		public bool CheckOpenFileOUR { get; set; }
		public bool CheckOpenFileKP { get; set; }
		public string Cell { get; set; }
		public static string StartCellOur { get; set; }
		public static string StartCellProvider { get; set; }
		public string TextCells { get; private set; }
		public static string filenameOurSpecification { get; set; }
		public static string filenameKPSpecification { get; set; }
		public string OrderN { get; set; }
		public int RowN { get; set; }
		public static Worksheet worksheetOurSpecification { get; set; }
		public static Workbook workbookOurSpecification { get; set; }
		public static Worksheet worksheetKPSpecification { get; set; }
		public static Workbook workbookKPSpecification { get; set; }

		//MainWindow MainWindow = new MainWindow();


		public void OpenFile()
		{
			try
			{
				//создаем приложение экселя
				Excel.Application ExcelApp = new Excel.Application();
				ExcelApp.Visible = true;
				//ExcelApp.WindowState = XlWindowState.xlMaximized;
				//ExcelApp.DisplayFullScreen = true;
				//проверка на то, какая кнопка сработала - первая, или вторая
				if (CheckOpenFileOUR)
				{
					workbookOurSpecification = ExcelApp.Workbooks.Open(filenameOurSpecification);
					worksheetOurSpecification = workbookOurSpecification.Worksheets[1];

				}
				if (CheckOpenFileKP)
				{
					workbookKPSpecification = ExcelApp.Workbooks.Open(filenameKPSpecification);
					worksheetKPSpecification = workbookKPSpecification.Worksheets[1];
				}
				CheckOpenFileOUR = false;
				CheckOpenFileKP = false;
				//Excel.Range range = (Excel.Range)worksheet.Cells[1, 1];
				//textBox.Text = range.Value2;
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
				//MessageBox.Show(worksheetOurSpecification.Range[StartCellOur].Row.ToString());
				//MessageBox.Show(worksheetOurSpecification.Cells[worksheetOurSpecification.Range[StartCellOur].Row, worksheetOurSpecification.Range[StartCellOur].Column].Value2);
				//if (worksheetOurSpecification.Cells[worksheetOurSpecification.Range[StartCellOur].Row, worksheetOurSpecification.Range[StartCellOur].Column].Value2.Replace("-","") == worksheetKPSpecification.Cells[worksheetKPSpecification.Range[StartCellProvider].Row, worksheetOurSpecification.Range[StartCellProvider].Column].Value2)
				//{
				//	MessageBox.Show("Заебись-одно и то же");

				//}
				//else
				//{
				//	MessageBox.Show("Что-то сука не так");
				//}
				bool find = false;//переменная для выхода из проверки при обнаружении искомого элемента. Также используется при выводе элемента, которы не был найден
				int OurSpecificationRowsMax = worksheetOurSpecification.UsedRange.Rows.Count;
				int ProviderSpecificationRowsMax = worksheetKPSpecification.UsedRange.Rows.Count;
				for (int i = worksheetOurSpecification.Range[StartCellOur].Row; i < OurSpecificationRowsMax + worksheetOurSpecification.Range[StartCellOur].Row; i++)
				{
					for (int j = worksheetKPSpecification.Range[StartCellProvider].Row; j <= ProviderSpecificationRowsMax; j++)
					{
						//Добавить проверку на пустую ячейку
						//где-то здесь
						//
						//Проверка на сравнение. Сравниваем по ячейчкам, задавая ее адрес в формате [i,j] - где i - строка j - столбец. У нас i-строка, а так так мы идем по одному столбцу, то он не изменен
						// и получаем его из свойства Column диапазона, образованного начальной ячейкой.
 						if (worksheetOurSpecification.Cells[i, worksheetOurSpecification.Range[StartCellOur].Column].Value2.Replace("-", "") == worksheetKPSpecification.Cells[j, worksheetOurSpecification.Range[StartCellProvider].Column].Value2.Replace("-", ""))
						{
							//OrderN = worksheetOurSpecification.Cells[i, worksheetOurSpecification.Range[StartCellOur].Column].Value2.Replace("-", "");
							//RowN = i;
							MessageBox.Show("Нашел" + " " + worksheetOurSpecification.Cells[i, worksheetOurSpecification.Range[StartCellOur].Column].Value2 + " " +
								"=" + " " + worksheetKPSpecification.Cells[j, worksheetOurSpecification.Range[StartCellProvider].Column].Value2);
							find = true;
							break;
						}

					}
					if (!find)
					{
						//MainWindow.NotfoundList.Items.Add(new { OrderN = worksheetOurSpecification.Cells[i, worksheetOurSpecification.Range[StartCellOur].Column].Value2.Replace("-", ""), RowN = worksheetOurSpecification.Cells[i, worksheetOurSpecification.Range[StartCellOur].Column].Row }) ;
						MessageBox.Show("Не нашел" + " " + worksheetOurSpecification.Cells[i, worksheetOurSpecification.Range[StartCellOur].Column].Value2);
						Order.Add(worksheetOurSpecification.Cells[i, worksheetOurSpecification.Range[StartCellOur].Column].Value2);
					}
					find = false;

				}

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

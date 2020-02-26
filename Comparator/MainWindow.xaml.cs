using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Comparator.Properties;

namespace Comparator
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	/// 
	public delegate void TextChangedEventHandler(object sender, TextChangedEventArgs e);

	public partial class MainWindow : System.Windows.Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		public string Cell { get; private set; }
		public Worksheet worksheet { get; private set; }

		private void Open_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Filter = "Excel files|*.xlsx|All files|*.*";
				openFileDialog.DefaultExt = "*.xlsx";
				openFileDialog.Title = "Choose your destiny";
				string filename = "";
				if (openFileDialog.ShowDialog() == true)
				{
					filename = openFileDialog.FileName;
					//textBox.Text = filename;
				}
				//textBox.Clear();
				Excel.Application ExcelApp = new Excel.Application();
				ExcelApp.Visible = true;
				//ExcelApp.WindowState = XlWindowState.xlMaximized;

				//ExcelApp.DisplayFullScreen = true;

				Workbook workbook;
				//Worksheet worksheet;
				workbook = ExcelApp.Workbooks.Open(filename);
				worksheet = workbook.Worksheets[1];
				Excel.Range range = (Excel.Range)worksheet.Cells[1, 1];
				//Excel.Range range = worksheet.
				//textBox.Text = range.Value2;
				//string Cell = "";
				//Cell = textBox1.Text;
				//textBox.Text = worksheet.Range[Cell].Value;

				//ExcelApp.Application.DisplayFullScreen=true;

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
				//Close();
			}

		}

	}
}

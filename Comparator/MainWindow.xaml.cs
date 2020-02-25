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

namespace Comparator
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	public partial class MainWindow : System.Windows.Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void Open_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Filter = "Excel files|*.xlsx|All files|*.*";
				openFileDialog.DefaultExt = "*.xlsx";
				openFileDialog.Title= "Choose your destiny";
				string filename = "";
				if (openFileDialog.ShowDialog() == true)
				{
					filename = openFileDialog.FileName;
					textBox.Text = filename;
				}

				Excel.Application ExcelApp = new Excel.Application();
				ExcelApp.Visible = true;
				ExcelApp.WindowState = Excel.XlWindowState.xlMaximized;
				Excel.

				//ExcelApp.DisplayFullScreen = true;

				Excel.Workbook workbook;
				Excel.Worksheet worksheet;
				workbook = ExcelApp.Workbooks.Open(filename);
				worksheet = workbook.ActiveSheet();
				//ExcelApp.Application.DisplayFullScreen=true;

			}
			catch (Exception)
			{

			}
			
		}
	}
}

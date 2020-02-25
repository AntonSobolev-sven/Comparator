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
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Excel files|*.xlsx|All files|*.*";
			openFileDialog.DefaultExt = "*.xlsx";
			string filename = "";
			if (openFileDialog.ShowDialog() == true)
			{
				filename = openFileDialog.FileName;
				textBox.Text = filename;
			}

			_Application ExcelApp = new Excel.Application();
			ExcelApp.Visible = true;

			Workbook workbook = new Excel.Workbook();
			Worksheet worksheet = workbook.Open(filename);
		}
	}
}

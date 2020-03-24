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
using Comparator.FileOpenCount;

namespace Comparator
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	/// 

	public partial class MainWindow : System.Windows.Window
	{
		OpenFileDialog openFileDialog = new OpenFileDialog();

		public MainWindow()
		{
			InitializeComponent();
			openFileDialog.Filter = "Excel files|*.xlsx;*.xls|All files|*.*";
			openFileDialog.DefaultExt = "*.xlsx";
			openFileDialog.Title = "Choose your destiny";
		}

		List<FileOC> fileOCs = new List<FileOC>();

		FileOC OpenfileOURSpec = new FileOC();
		FileOC OpenfileProvSpec = new FileOC();
		FileOC CompareFiles = new FileOC();
		private void Open_Click(object sender, RoutedEventArgs e)
		{
			//FileOC fileOC = new FileOC();
			//FileNamePath.Text = OpenfileOURSpec.filename;
			if (openFileDialog.ShowDialog() == true)
			{
				FileNamePath.Text = openFileDialog.FileName;
				FileOC.filenameOurSpecification = FileNamePath.Text;
				OpenfileOURSpec.CheckOpenFileOUR = true;
				OpenfileOURSpec.OpenFile();
			}
			else
			{
				FileNamePath.Text = null;
			}

		}

		private void OpenKP_Click(object sender, RoutedEventArgs e)
		{
			//FileOC fileOC = new FileOC();
			//FileNamePathKP.Text = OpenfileOURSpec.filename;
			if (openFileDialog.ShowDialog() == true)
			{
				FileNamePathKP.Text = openFileDialog.FileName;
				FileOC.filenameKPSpecification = FileNamePathKP.Text;
				OpenfileProvSpec.CheckOpenFileKP = true;
				OpenfileProvSpec.OpenFile();
			}
			else
			{
				FileNamePathKP.Text = null;
			}
		}

		private void Go_Click(object sender, RoutedEventArgs e)
		{
			FileOC.StartCellOur = StartOurCell.Text;
			FileOC.StartCellProvider = StartProviderCell.Text;
			CompareFiles.Compare();
			//NotfoundList.Items.Add("sdfdfsgggggg");
			NotfoundList.Items.Add(CompareFiles.Order);

		}

	}
}

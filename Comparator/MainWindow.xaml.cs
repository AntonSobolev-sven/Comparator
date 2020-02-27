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
		public MainWindow()
		{
			InitializeComponent();
		}

		FileOC OpenfileOURSpec = new FileOC();
		FileOC OpenfileProvSpec = new FileOC();
		FileOC CompareFiles = new FileOC();
		private void Open_Click(object sender, RoutedEventArgs e)
		{
			//FileOC fileOC = new FileOC();
			OpenfileOURSpec.OpenFile();
			FileNamePath.Text = OpenfileOURSpec.filename;

		}

		private void OpenKP_Click(object sender, RoutedEventArgs e)
		{
			//FileOC fileOC = new FileOC();
			OpenfileProvSpec.OpenFile();
			FileNamePathKP.Text = OpenfileOURSpec.filename;
		}

		private void Go_Click(object sender, RoutedEventArgs e)
		{
			//FileOC fileOC = new FileOC();
			CompareFiles.Compare();
			CompareFiles.Cell = OurCell.Text;
		}
	}
}

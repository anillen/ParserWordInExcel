using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;
using System.Windows.Threading;
using Microsoft.Win32;
using System.Diagnostics;



namespace ParserWordInExcel
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	/// 



	public partial class MainWindow : Window
	{

		Object missingObj = System.Reflection.Missing.Value;
		Object trueObj = true;
		Object falseObj = false;
		double step = 1; //шаг по умолчанию для полоски
		string razdel = "";
		string path = "D:\\Resource"; //Путь по умолчанию


		public MainWindow()
		{
			InitializeComponent();
			UpdateContent();
		}
		

		string GetPathInFile()
		{
			//Возвращает путь из файла 'setings.txt'
			try
			{
				using (StreamReader reader = new StreamReader("settings.txt"))
				{
					path = reader.ReadToEnd();	
				}
				path = ReplacePath(path);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

			return path;
		}

		string ReplacePath(string str)
		{
			//Очищаем строку от лишних символов
			List<char> charsToRemove = new List<char>() { '\r', '\a','\v','\n','+','*','!'};
			foreach (char c in charsToRemove)
			{
				str = str.Replace(c.ToString(), String.Empty);
			}

			return str;
		}
		string ReplaceStr(string str)
		{
			//Очищаем строку от лишних символов
			str = str.Replace('\a'.ToString(), String.Empty);
			str = str.Replace('\r'.ToString(), String.Empty);
			str = str.Replace('\v'.ToString(), "\n\r");
			return str;
		}

		public string[] GetFilters()
		{
			string[] filters = new string[5];
			using (StreamReader reader = new StreamReader("filter.txt"))
			{
				string text = reader.ReadToEnd();
				if (text != "")
				{
					string tmp = "";
					for (int i = 0, j = 0; i < text.Length; i++)
					{
						if (text[i] == ';')
						{
							filters[j] = tmp;
							tmp = "";
							j++;
						}
						else
							tmp += text[i];

					}
				}
				else
				{
					System.Windows.MessageBox.Show("Не предвиденная ошибка пути к фильтру!");
				}
				return filters;
			}
		}
		void UpdateContent()
		{
			//Обновление содержимого окна
			progressBar_Main.Value = 0;
			progressBar_Main.Maximum = 100;
			path = GetPathInFile();

			btn_Convert.IsEnabled = true;
			lBox_FilesExport.Items.Clear();
			lBox_FilesImport.Items.Clear();
			tb_Razdel.Text = "";

			try
			{
				string[] patheth = Directory.GetFiles(path, "*.doc");
				if (patheth.Count() != 0)
				{
					for (int i = 0; i < patheth.Count(); i++)
					{
						FileInfo file = new FileInfo(patheth[i]);
						if (file.Name[0] != '~' && file.Name[1] != '$')
							lBox_FilesImport.Items.Add(file.Name);
					}
				}
				else
				{
					MessageBox.Show("Нет файлов в папке c ресурсами!");
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			
		}

		void ConvertFile(object obj)
		{
			//Создаем документ word и excel
			Word._Application application = new Word.Application();
			Word._Document document = new Word.Document();
			Excel.Application ObjExcel = new Excel.Application();
			// Определяем путь до файла элемента
			path = GetPathInFile();
			Object FilePath = path + "\\" + obj.ToString();


			try
			{
				Excel.Workbook ObjWorkBook;
				Excel.Worksheet ObjWorkSheet;


				ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
				ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

				document = application.Documents.Open(FilePath, ref missingObj, false, true);
				Word.Tables tables = document.Tables;

				string[] Filters = GetFilters();

				//Разбираем таблицу
				foreach (Word.Table t in tables)
				{
					int i = 2;
					foreach (Word.Row r in t.Rows)
					{
						if (r.Cells[2].Range.Text.Contains(Filters[0])  && r.Cells[3].Range.Text.Contains(Filters[1]) && r.Cells[4].Range.Text.Contains(Filters[2]) && r.Cells[5].Range.Text.Contains(Filters[3]) && r.Cells[6].Range.Text.Contains(Filters[4]) )
						{
							string tmp = r.Cells[5].Range.Text;
							string[] RowText = { (i-1).ToString(), " ", " ", razdel, ReplaceStr(r.Cells[2].Range.Text),ReplaceStr(r.Cells[3].Range.Text)," ",
								ReplaceStr(r.Cells[5].Range.Text), ReplaceStr(r.Cells[6].Range.Text),ReplaceStr(r.Cells[7].Range.Text)," "," " };

							ObjWorkSheet.Cells[1, 1] = "№п/п";
							ObjWorkSheet.Cells[1, 2] = "Объект";
							ObjWorkSheet.Cells[1, 3] = "Заказчик"; ;
							ObjWorkSheet.Cells[1, 4] = "Раздел проекта";
							ObjWorkSheet.Cells[1, 5] = "Наименование";
							ObjWorkSheet.Cells[1, 6] = "    ";
							ObjWorkSheet.Cells[1, 7] = "Номенклатура ТО";
							ObjWorkSheet.Cells[1, 8] = "Поставщик по проекту";
							ObjWorkSheet.Cells[1, 9] = "Ед. изм.";
							ObjWorkSheet.Cells[1, 10] = "Кол-во";
							ObjWorkSheet.Cells[1, 11] = "Цена с НДС, руб. ";
							ObjWorkSheet.Cells[1, 12] = "Сумма с НДС, руб.";


							ObjWorkSheet.Cells[i, 1] = RowText[0];
							ObjWorkSheet.Cells[i, 2] = RowText[1];
							ObjWorkSheet.Cells[i, 3] = RowText[2];
							ObjWorkSheet.Cells[i, 4] = RowText[3];
							ObjWorkSheet.Cells[i, 5] = RowText[4];
							ObjWorkSheet.Cells[i, 6] = RowText[5];
							ObjWorkSheet.Cells[i, 7] = RowText[6];
							ObjWorkSheet.Cells[i, 8] = RowText[7];
							ObjWorkSheet.Cells[i, 9] = RowText[8];
							ObjWorkSheet.Cells[i, 10] = RowText[9];
							ObjWorkSheet.Cells[i, 11] = RowText[10];
							ObjWorkSheet.Cells[i, 12] = RowText[11];
							i++;
						}
					}
					//Форматирование ячеек
					Excel.Range range1 = ObjWorkSheet.Range["A1", "L" + (i-1)];
					range1.Cells.Font.Size = 12;
					range1.Cells.Font.Name = "Times New Roman";
					range1.Cells.VerticalAlignment = -4160; //Выравнивание по верхнему краю
					range1.Cells.WrapText = false;
					range1.EntireColumn.AutoFit();
					range1.EntireRow.AutoFit();
					range1.Cells.NumberFormat = "@";
					Excel.Borders border = range1.Cells.Borders;
					border.LineStyle = Excel.XlLineStyle.xlContinuous;
					border.Weight = 2d;

					Excel.Range range2 = ObjWorkSheet.Range["E2", "F" + (i-1)];
					range2.Cells.WrapText = true;
					Excel.Range range3 = ObjWorkSheet.Range["J2", "L" + (i-1)];
					range3.Cells.NumberFormat = "0.00";
					
				}
				path = GetPathInFile();

				if (!Directory.Exists(path + "\\ExcelExport"))
				{
					Directory.CreateDirectory(path + "\\ExcelExport");
				}

				string fileName = obj.ToString();
				fileName = fileName.Remove((fileName.Length-4),4);
				path = path + "\\ExcelExport\\" + fileName + ".xls";

				path = ReplacePath(path);
				//Сохраняем наш excel файл
				ObjWorkBook.SaveAs(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive,
				Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

				ObjExcel.Quit();
				ObjExcel = null;
				document.Close(ref falseObj, ref missingObj, ref missingObj);
				application.Quit(ref missingObj, ref missingObj, ref missingObj);
				document = null;
				application = null;



				//Обновление progressBar из потока.
				progressBar_Main.Dispatcher.BeginInvoke(
		   System.Windows.Threading.DispatcherPriority.Normal
		   , new DispatcherOperationCallback(delegate
		   {
			   progressBar_Main.Value = progressBar_Main.Value + step;
			   if(progressBar_Main.Value>=100)
			   {
				   btn_Convert.IsEnabled = true;
				   btn_Update.IsEnabled = true;
			   }
			   return null;
		   }), null);


			}
			catch (Exception error)
			{
				if(document!=null)
				{
					document.Close(ref falseObj, ref missingObj, ref missingObj);
					application.Quit(ref missingObj, ref missingObj, ref missingObj);
					document = null;
					application = null;
				}
				if(ObjExcel!=null)
				{
					ObjExcel.Quit();
					ObjExcel = null;
				}
				MessageBox.Show(error.Message);
			}
		}


		void ConvertFiles()
		{
			if(lBox_FilesExport.Items.Count!=0)
			{
				// = (100 / lBox_FilesExport.Items.Count)+0.001;//Вычисляем шаг для ProgressBar
				//step = Math.Ceiling(step);
			}

			foreach (object obj in lBox_FilesExport.Items)
			{
				ConvertFile(obj);
			}
		}

		private void MenuItem_Click_Help(object sender, RoutedEventArgs e)
		{
			Process proc = Process.Start("notepad.exe", "ReadMe.txt");
		}

		private void LBox_FilesImport_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (lBox_FilesImport.SelectedItem == null)
			{
				MessageBox.Show("Выберите элемент");
				return;
			}
			lBox_FilesExport.Items.Add(lBox_FilesImport.SelectedItem);
			lBox_FilesImport.Items.RemoveAt(lBox_FilesImport.SelectedIndex);
		}

		private void Btn_Convert_Click(object sender, RoutedEventArgs e)
		{
			if(lBox_FilesExport.Items.Count==0)
			{
				MessageBox.Show("Выберите элементы для конвертации");
				return;
			}
			btn_Update.IsEnabled = false;
			progressBar_Main.Value = 0;
			progressBar_Main.Maximum = lBox_FilesExport.Items.Count;
			btn_Convert.IsEnabled = false;
			razdel = tb_Razdel.Text;
			Thread thread = new Thread(ConvertFiles);
			thread.Start();
		}

		private void LBox_FilesExport_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (lBox_FilesExport.SelectedItem == null)
			{
				MessageBox.Show("Выберите элемент");
				return;
			}
			lBox_FilesImport.Items.Add(lBox_FilesExport.SelectedItem);
			lBox_FilesExport.Items.RemoveAt(lBox_FilesExport.SelectedIndex);
		}

		private void MenuItem_Click_Options(object sender, RoutedEventArgs e)
		{
			Settings win = new Settings();
			win.Show();
		}

		private void Btn_Alin_Click(object sender, RoutedEventArgs e)
		{
			foreach(var ob in lBox_FilesImport.Items)
			{
				lBox_FilesExport.Items.Add(ob);
			}
			lBox_FilesImport.Items.Clear();
		}

		private void Btn_Alof_Click(object sender, RoutedEventArgs e)
		{
			foreach (var ob in lBox_FilesExport.Items)
			{
				lBox_FilesImport.Items.Add(ob);
			}
			lBox_FilesExport.Items.Clear();
		}

		private void Btn_Update_Click(object sender, RoutedEventArgs e)
		{
			UpdateContent();
		}

		private void MenuItem_Click_Files(object sender, RoutedEventArgs e)
		{
			path = GetPathInFile();
			if(path!="")
			{
				Process proc = Process.Start("explorer.exe", path);
			}
			else
			{
				MessageBox.Show("Выбран не верный путь!");
			}
		}
	}

}

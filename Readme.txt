git add *.pdf
git commit -m "Comments"
git pull
git push



https://www.youtube.com/watch?v=sbGJZPAwEyg

https://www.youtube.com/watch?v=cX3c57phPR8

Спинер
https://aliexpress.ru/item/1005001706750028.html

https://aliexpress.ru/item/32417857399.html
https://aliexpress.ru/item/32958232064.html
https://aliexpress.ru/item/33048773910.html
https://aliexpress.ru/item/32476516056.html

https://aliexpress.ru/item/580449904.html

https://www.youtube.com/watch?v=iMzA-RkqOc4


c#
https://www.grapecity.com/blogs/4-steps-to-transpose-and-invert-wpf-datagrid
http://www.oszone.net/16661/WPF-DataGrid
https://www.programmersought.com/article/57831471330/
https://www.youtube.com/watch?v=DPbanNt1Ss4

https://learn.microsoft.com/en-us/dotnet/api/system.security.cryptography.tripledescryptoserviceprovider?redirectedfrom=MSDN&view=net-6.0

https://www.codeproject.com/Articles/509824/Creating-a-NumericUpDown-control-from-scratch

//DataGrid Filtering
https://github.com/MishkinIN/ItemsFilter

for (int i = 0; i < textBox1.Text.Length; i++)
{
    if (Char.IsLetter(textBox1.Text, i) || char.IsSymbol(textBox1.Text, i) )
    {
        MessageBox.Show($"{textBox1.Text.Substring(0, i).Trim()}|");
        MessageBox.Show($"{textBox1.Text.Substring(i).Trim()}|");
        break;
    }
    if (char.IsPunctuation(textBox1.Text, i) )
    {
        if (textBox1.Text[i] == ',' || textBox1.Text[i] == '-')
        {
            continue;
        }
        MessageBox.Show(textBox1.Text.Substring(0, i));
        MessageBox.Show($"{textBox1.Text.Substring(i).Trim()}|");
        break;
    }
    
    SortedSet
    
    
    https://www.youtube.com/watch?v=1QmTkHUMwiA
    
    https://www.youtube.com/@vector-massage-school
    
    https://github.com/HangfireIO/Hangfire

Шифрование Go
https://www.golinuxcloud.com/golang-encrypt-decrypt/

https://gist.github.com/andrewloable/afa1c683701cec18c4530f6a91692e0b

=========================================================================================
using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace pdfSpecification {
	public partial class frmCreatePDFSpec : Form {

		public struct ElementData {
			public string Pos;
			public string ElName;
			public string ElType;
			public string CodeName;
			public string ElMaker;
			public string ElUnit;
			public string ElTotal;
			public string ElMass;
			public string Notes;
		};

		public struct GroupData {
			public string GroupName;
			public List<ElementData> elArray;
		};

		#region ClassData

		const float PTINMM = 2.83333333333f;
		const float MAX_TAB_HEIGHT = 223f;
		const int NUM_SHEET = 7;
		
		private string ExcelFileName;
		private string FontFileNameI;
		private string FontFileNameBI;
		private string PdfFileName;

		private bool pdfCreated;
		private bool excelOpened;

		private Document doc;
		private BaseFont bfArialI;
		private BaseFont bfArialBI;
		
		private PdfContentByte cb;
		private PdfTemplate templOT;
		private PdfTemplate templLeft;
		private PdfTemplate templTabHeader;

		//размеры колонок основной таблицы
		private float[] tabcol = { 20f, 135f, 59f, 40f, 40f, 20f, 20f, 25f, 35f };

		private List<GroupData> AllData;
		private List<int> SequensData;
		private int workGpNum;
		private int workElNum;
		private bool GpTitlePrinted;
		private int PageNum;
		private int GroupNum;
		private string Sequens;

		#endregion

		#region ExcelData
		//Переменные для работы с Excel
		private Excel.Application exapp;
		private Excel.Workbook wbSource;
		private Excel.Worksheet wsSource;
		//------------------------------
		#endregion

		/// <summary>
		/// Обработка загрузки формы
		/// </summary>
		private void frmCreatePDFSpec_Load(object sender, EventArgs e) {
			tbProjName.Text = Properties.Settings.Default.ProjectName;
		}

		/// <summary>
		/// Конструктор формы frmCreatePDFSpec
		/// </summary>
		public frmCreatePDFSpec() {
			InitializeComponent();
		}

		/// <summary>
		/// Процедура пуска генерации спецификации
		/// </summary>
		private void btnCreatePDFSpec_Click(object sender, EventArgs e) {
			ClearStatusLabel(0);
			Properties.Settings.Default.ProjectName = tbProjName.Text.Trim();
			Properties.Settings.Default.Save();
			PdfFileName = Application.StartupPath + "\\PDF\\1.pdf";
			if (!Conditions())
				return;
			Refresh();
			excelOpened = OpenExcelFile();
			if (!excelOpened)
				return;
			int workSheet = 1;
			while (workSheet < exapp.Worksheets.Count + 1) {
				ReadExcelPage(workSheet);
				workSheet++;
			}
			//Закрываем Excel файл
			if (excelOpened) {
				exapp.Quit();
				PrintMess(ExcelFileName + " прочитан и закрыт.", 2);
			}
			lbReportLog.Refresh();

			{
				//Выбор порядка вывода групп
				frmGroupSel gs = new frmGroupSel();
				for (int i = 0; i < AllData.Count; i++) {
					gs.AddSourceGroup(AllData[i].GroupName, i);
				}
				gs.ShowDialog();
				Sequens = gs.OutList;
				PageNum = gs.PageNum;
				GroupNum = gs.GroupNum;
				gs.OutList = "";
				gs.Dispose();
			}  
			if (!LoadSequens()) {
				AllData.Clear();
				SequensData.Clear();
				Sequens = "";
				return;
			}
			PrintMess("Выбраны групп: " + SequensData.Count.ToString(), 2);

			pdfCreated = CreatePDFDoc();
			templOT = CreateTemplOT();
			templLeft = CreateTemplLeft();
			templTabHeader = CreateTemplTabHeader();

			try {
				if (!DrawPageContent()) {
					throw new Exception("Не смог напечатать первую страницу файла!");
				}
				while (workGpNum < SequensData.Count) {
					doc.NewPage();
					if (!DrawPageContent())
						break;
				}
			}
			catch (Exception ex) {
				PrintMess(ex.Message, 2);
			}

			//Вывод в pdf
			if (pdfCreated) {
				doc.Close();
				PrintMess("Спецификация создана.", 2);
			}
			else
				PrintMess("Работа завершена с ошибками!", 2);

			//Сохраним результат на диске
			SaveFileDialog sd = new SaveFileDialog();
			sd.Filter = "Файл PDF (*.pdf) |*.pdf";
			DialogResult res = sd.ShowDialog();
			if (res != DialogResult.OK) {
				PrintMess("Спецификация не сохранена!", 2);
			}
			else {
				string NewPdfFileName = sd.FileName;
				try {
					File.Copy(PdfFileName, NewPdfFileName, true);
					PrintMess("Файл сохранён:", 2);
					PrintMess(NewPdfFileName, 2);
				}
				catch (Exception ex) {
					PrintMess("Не удалось сохранить файл!", 2);
					PrintMess(ex.Message, 2);
				}        
			}
		}

		//========================================================================

		#region ServiceFunc

		/// <summary>
		/// Собирает необходимые для генерации данные
		/// </summary>
		/// <returns>Возвращает true, если все условия выполнены</returns>
		private bool Conditions() {
			bool Result = false;
			if (tbProjName.Text == "") {
				PrintMess("Нет названия проекта!", 1);
				return Result;
			}
			//Выбираем файл для загрузки
			OpenFileDialog od = new OpenFileDialog();
			od.Filter = "Файл Excel (*.xls) |*.xls";
			DialogResult res = od.ShowDialog();
			if (res != DialogResult.OK) {
				PrintMess("Не выбран файл для загрузки!", 1);
				return Result;
			}
			ExcelFileName = od.FileName;
			
			FontFileNameI = Application.StartupPath + "\\Fonts\\Ariali.TTF";
			FontFileNameBI = Application.StartupPath + "\\Fonts\\Arialbi.ttf";
			if (!File.Exists(FontFileNameI)) {
				od.Filter = "Файл шрифта (*.TTF) |*.TTF";
				res = od.ShowDialog();
				if (res != DialogResult.OK) {
					PrintMess("Не выбран шрифт для формирования pdf!", 1);
				}
				else {
					FontFileNameI = od.FileName;
				}          
			}
			if (!File.Exists(FontFileNameBI)) {
				od.Filter = "Файл шрифта (*.TTF) |*.TTF";
				res = od.ShowDialog();
				if (res != DialogResult.OK) {
					PrintMess("Не выбран шрифт для формирования pdf!", 1);
				}
				else {
					FontFileNameBI = od.FileName;
					Result = true;
				}
			}
			else
				Result = true;
			//Инициализация
			workGpNum = 0;
			workElNum = 0;
			GpTitlePrinted = false;
			PageNum = 1;
			GroupNum = 1;
			return Result;
		}

		/// <summary>
		/// Печатает сообщение
		/// </summary>
		/// <param name="mess">Текст сообшения об ошибке</param>
		/// <param name="Direction">Направление вывода текста</param>
		private void PrintMess(string mess, int Direction) {
			switch (Direction) {
				case 0:
					lbReportLog.Items.Add(mess);
					ClearStatusLabel(1);
					statusLabel.Text = mess;
					Refresh();
					break;
				case 1:
					ClearStatusLabel(1);
					statusLabel.Text = mess;
					Refresh();
					break;
				case 2:
					lbReportLog.Items.Add(mess);
					lbReportLog.Refresh();
					break;
			}
		}

		/// <summary>
		/// функция очиски содержимого
		/// </summary>
		/// <param name="Direction">Элемент</param>
		private void ClearStatusLabel(int Direction) {
			switch (Direction) {
				case 0:
					statusLabel.Text = "";
					lbReportLog.Items.Clear();
					break;
				case 1:
					statusLabel.Text = "";
					break;
				case 2:
					lbReportLog.Items.Clear();
					break;
			}
		}

		/// <summary>
		/// Переводит миллиметры в дюймы
		/// </summary>
		/// <param name="inptval">Значение в миллиметрах</param>
		/// <returns>Значение в дюймах</returns>
		private float InMM(float inptval) {
			return inptval * PTINMM;
		}

		private bool LoadSequens() {
			bool Result = false;
			SequensData = new List<int>();
			Regex rgx = new Regex(@"|");
			try {
				if (rgx.Match(Sequens).Success) {
					string[] SeqNumArr = Sequens.Split('|');
					for (int i = 0; i < SeqNumArr.Length; i++) {
						try {
							SequensData.Add(Convert.ToInt32(SeqNumArr[i]));
						}
						catch (Exception ex) {
							PrintMess("Конвертация строки в целое. " + ex.Message, 2);
						}
					}
				}
				else {
					//Только одна группа
					int Seq = Convert.ToInt32(Sequens);
					SequensData.Add(Seq);
				}
				Result = true;
			}
			catch (Exception ex) {
				PrintMess(ex.Message, 2);
			}
			return Result;
		}

		#endregion

		//========================================================================

		#region PDF

		/// <summary>
		/// Создаём pdf файл
		/// </summary>
		/// <returns>При успешном завершении операций возвращает true</returns>
		private bool CreatePDFDoc() {
			bool Result = false;
			try {
				doc = new Document(PageSize.A3.Rotate(), 0, 0, 0, 0);
				PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(PdfFileName, FileMode.Create));
				doc.Open();
				cb = writer.DirectContent;
				if (CreateFont()) {
					Result = true;
				}
			}
			catch (Exception ex) {
				PrintMess("Ошибка при создании pdf файла. " + ex.Message, 0);
			}
			return Result;
		}

		/// <summary>
		/// Создаёт шрифт для нанесения записей в файле pdf
		/// </summary>
		/// <param name="FontName">Путь к шрифту</param>
		/// <returns>Возвращает true, при удачном создании шрифта</returns>
		private bool CreateFont() {
			bool Result = false;
			try {
				bfArialI = BaseFont.CreateFont(FontFileNameI, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
				bfArialBI = BaseFont.CreateFont(FontFileNameBI, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
				Result = true;
			}
			catch(Exception ex){
				PrintMess("Ошибка при создании шрифт для pdf. " + ex.Message, 0);
			}
			return Result;
		}

		/// <summary>
		/// Создаёт новую страницу в pdf документе
		/// </summary>
		/// <returns>true - при успешном выполнении</returns>
		private bool DrawPageContent() {
			bool Result = false;
			try {
				if (pdfCreated) {
					if (templOT != null) {
						cb.AddTemplate(templOT, InMM(420f - 5.7f - 185f), InMM(5.7f));
						cb.Stroke();
					}
					if (templLeft != null) {
						cb.AddTemplate(templLeft, 0, 1, -1, 0, InMM(20f), InMM(6f - 0.3f));
						cb.Stroke();
					}
					{
						PdfPTable tb = CreateDataTable();
						tb.WriteSelectedRows(0, -1, InMM(20.3f), InMM(254f), cb);
					}
					if (templTabHeader != null) {
						cb.AddTemplate(templTabHeader, InMM(20f), InMM(297f - 43f));
						cb.Stroke();
					}
					cb.SetLineWidth(InMM(0.6f));
					cb.Rectangle(InMM(20f), InMM(6f), InMM(420f - 26f), InMM(297f - 12f));
					cb.Stroke();

					cb.MoveTo(InMM(20f), InMM(31f));
					cb.LineTo(InMM(414.3f), InMM(31f));
					cb.Stroke();

					float tempX = 20f;
					for (int i = 0; i < tabcol.Length; i++) {
						tempX += tabcol[i];
						cb.MoveTo(InMM(tempX), InMM(31f));
						cb.LineTo(InMM(tempX), InMM(290.7f));
					}
					cb.Stroke();

					cb.SetFontAndSize(bfArialI, 12);
					cb.BeginText();
					cb.ShowTextAligned(Element.ALIGN_CENTER, PageNum.ToString(), InMM(420f - 11f), InMM(8f), 0);
					cb.EndText();
					PageNum++;

					Result = true;
				}
			}
			catch (Exception ex) {
				PrintMess("Ошибка создания новой страницы pdf файла. " + ex.Message, 2);
			}
			return Result;
		}

		/// <summary>
		/// Создаёт основную таблицу на странице
		/// </summary>
		/// <returns>При успешном завершении возвращает true</returns>
		private PdfPTable CreateDataTable() {
			PdfPTable tb = new PdfPTable(9);
			tb.TotalWidth = InMM(394f);
			tb.SetWidths(new int[] { 20, 135, 59, 40, 40, 20, 20, 25, 35});
			tb.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
			tb.DefaultCell.Padding = InMM(2f);
			tb.SkipFirstHeader = true;
			//Считывание данных из Excel и добавление ячеек
			if (!(AllData.Count > 0)) {
				PrintMess("Нет данных для отображения в спецификации!", 2);
				return null;
			}
			bool CanAddCell = true;
			while (CanAddCell) {
				CanAddCell = CreateTableRow(tb);
				if (tb.TotalHeight > InMM(MAX_TAB_HEIGHT)) {
					tb.DeleteLastRow();
					workElNum--;
					CanAddCell = false;
				}        
			}
			return tb;
		}

		/// <summary>
		/// Формирует строку основной таблицы
		/// </summary>
		/// <param name="tb">Целевая таблица</param>
		/// <returns>При успешном завершении возвращает true</returns>
		private bool CreateTableRow(PdfPTable tb) {
			bool Result = false;
			if (workElNum > AllData[SequensData[workGpNum]].elArray.Count-1) {
				workGpNum++;
				workElNum = 0;
				GpTitlePrinted = false;
				return Result;
			}
			if ((workElNum == 0) && (!GpTitlePrinted)){
				//печатаем заголовок группы
				tb.AddCell("");
				tb.AddCell(DrawTitleCell(GroupNum.ToString() + "." + AllData[SequensData[workGpNum]].GroupName, Element.ALIGN_CENTER));
				GroupNum++;
				for(int i=0;i<7;i++)
					tb.AddCell("");
				GpTitlePrinted = true;
				Result = true;
			}
			else {
				//Печатаем строку с елементом группы
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].Pos, Element.ALIGN_RIGHT));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].ElName, Element.ALIGN_LEFT));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].ElType, Element.ALIGN_LEFT));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].CodeName, Element.ALIGN_CENTER));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].ElMaker, Element.ALIGN_LEFT));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].ElUnit, Element.ALIGN_CENTER));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].ElTotal, Element.ALIGN_CENTER));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].ElMass, Element.ALIGN_CENTER));
				tb.AddCell(DrawCell(AllData[SequensData[workGpNum]].elArray[workElNum].Notes, Element.ALIGN_LEFT));
				workElNum++;
				Result = true;
			}
			return Result;
		}

		/// <summary>
		/// Создаёт ячейку стблицы с заголовком группы
		/// </summary>
		/// <param name="celltext">Текст в ячейке</param>
		/// <param name="horz_alig">Горизонтальное выравнивание в ячейке</param>
		/// <returns>При успешном завершении возвращает true</returns>
		private PdfPCell DrawTitleCell(string celltext, int horz_alig) {
			iTextSharp.text.Font fArialUn = new iTextSharp.text.Font(bfArialI, 12,iTextSharp.text.Font.UNDERLINE);
			PdfPCell cell = new PdfPCell(new Phrase(celltext, fArialUn));
			cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
			cell.HorizontalAlignment = horz_alig;
			cell.Padding = InMM(2f);
			return cell;
		}

		/// <summary>
		/// Создаёт ячейку таблицы
		/// </summary>
		/// <param name="celltext">Текст в ячейке</param>
		/// <param name="horz_alig">Горизонтальное выравнивание текста</param>
		/// <returns>При успешном завершении возвращает true</returns>
		private PdfPCell DrawCell(string celltext, int horz_alig) {
			iTextSharp.text.Font fArial = new iTextSharp.text.Font(bfArialI, 12);
			PdfPCell cell = new PdfPCell(new Phrase(celltext, fArial));
			cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
			cell.HorizontalAlignment = horz_alig;
			cell.Padding = InMM(2f);
			return cell;
		}

		/// <summary>
		/// Создает штамп основной надписи
		/// </summary>
		/// <returns>Созданный PdfTemplate</returns>
		private PdfTemplate CreateTemplOT() {
			PdfTemplate tpl = null;
			try {
				tpl = cb.CreateTemplate(InMM(185f), InMM(15f));
				tpl.SetLineWidth(InMM(0.6f));
				tpl.Rectangle(InMM(0.3f), InMM(0.3f), InMM(185f - 0.6f), InMM(15f - 0.6f));
				float[] shtampcol = { 7f, 10f, 10f, 13f, 15f, 10f, 110f };
				float tempX = 0.3f;
				for (int i = 0; i < shtampcol.Length; i++) {
					tempX += shtampcol[i];
					tpl.MoveTo(InMM(tempX), InMM(0));
					tpl.LineTo(InMM(tempX), InMM(14.7f));
				}
				tpl.MoveTo(InMM(0.3f), InMM(4.7f));
				tpl.LineTo(InMM(65f), InMM(4.7f));
				tpl.MoveTo(InMM(184.7f), InMM(8f));
				tpl.LineTo(InMM(185f - 9.7f), InMM(8f));
				tpl.Stroke();
				tpl.SetLineWidth(0.2f);
				tpl.MoveTo(InMM(0.1f), InMM(9.9f));
				tpl.LineTo(InMM(65f), InMM(9.9f));
				tpl.Stroke();

				tpl.SetFontAndSize(bfArialI, 9);
				string[] shtamptext = { "Изм.", "Кол.уч.", "Лист", "№ док.", "Подп.", "Дата" };
				tempX = 0.3f;
				tpl.BeginText();
				for (int i = 0; i < shtamptext.Length; i++) {
					float offsetX = shtampcol[i] / 2;
					tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, shtamptext[i], InMM(tempX + offsetX), InMM(1.3f), 0);
					tempX += shtampcol[i];
				}
				tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Лист", InMM(185f - 5f), InMM(11f), 0);
				tpl.EndText();

				tpl.SetFontAndSize(bfArialBI, 14);
				tpl.BeginText();
				tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tbProjName.Text.Trim(), InMM(tempX + 55f), InMM(7f), 0);
				tpl.EndText();

			}
			catch (Exception ex) {
				PrintMess("Ошибка создания основной надписи. " + ex.Message, 2);
				tpl = null;
			}      
			return tpl;
		}

		/// <summary>
		/// Создает левый штамп
		/// </summary>
		/// <returns>Созданный PdfTemplate</returns>
		private PdfTemplate CreateTemplLeft() {
			PdfTemplate tpl = null;
			try {
				tpl = cb.CreateTemplate(InMM(85f), InMM(5f + 7f));
				tpl.SetLineWidth(InMM(0.6f));
				tpl.MoveTo(InMM(0.3f), InMM(0f));
				tpl.LineTo(InMM(0.3f), InMM(12f - 0.3f));
				tpl.LineTo(InMM(85f - 0.3f), InMM(12f - 0.3f));
				tpl.LineTo(InMM(85f - 0.3f), InMM(0f));
				tpl.MoveTo(InMM(0.3f), InMM(6.7f));
				tpl.LineTo(InMM(85f), InMM(6.7f));
				tpl.MoveTo(InMM(22.7f), InMM(0f));
				tpl.LineTo(InMM(22.7f), InMM(11.7f));
				tpl.MoveTo(InMM(52.7f), InMM(0f));
				tpl.LineTo(InMM(52.7f), InMM(11.7f));
				tpl.Stroke();
				tpl.SetFontAndSize(bfArialI, 9);
				tpl.BeginText();
				tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Инв. № подл.", InMM(11.5f), InMM(8f), 0);
				tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Подп. и дата", InMM(23f + 15f), InMM(8f), 0);
				tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Взам. инв. №", InMM(53f + 16f), InMM(8f), 0);
				tpl.EndText();
			}
			catch (Exception ex) {
				PrintMess("Ошибка создания основной надписи. " + ex.Message, 2);
				tpl = null;
			}
			return tpl;
		}

		/// <summary>
		/// Создает шапку основной таблицы
		/// </summary>
		/// <returns>Созданный PdfTemplate</returns>
		private PdfTemplate CreateTemplTabHeader() {
			PdfTemplate tpl = null;
			try {
				tpl = cb.CreateTemplate(InMM(394f), InMM(37f));
				tpl.SetLineWidth(InMM(0.6f));
				tpl.MoveTo(InMM(0.3f), InMM(6.7f));
				tpl.LineTo(InMM(393.7f), InMM(6.7f));
				tpl.Stroke();
				tpl.SetLineWidth(InMM(0.2f));
				tpl.MoveTo(InMM(0.1f), InMM(0.1f));
				tpl.LineTo(InMM(393.9f), InMM(0.1f));
				tpl.Stroke();
				string[] tabtext = { "ПОЗИ-|ЦИЯ", 
														 "НАИМЕНОВАНИЕ И ТЕХНИЧЕСКАЯ ХАРАКТЕРИСТИКА", 
														 "ТИП, МАРКА,|ОБОЗНАЧЕНИЕ|ДОКУМЕНТА,|ОПРОСНОГО ЛИСТА",
														 "КОД|ОБОРУДОВАНИЯ,|ИЗДЕЛИЯ,|МАТЕРИАЛА",
														 "ЗАВОД-|ИЗГОТОВИТЕЛЬ",
														 "ЕДИ-|НИЦА|ИЗМЕ-|РЕНИЯ",
														 "КОЛИ-|ЧЕСТВО",
														 "МАССА|ЕДИНИЦЫ,|КГ",
														 "ПРИМЕЧАНИЕ"};
				float tempX = 0f;
				tpl.SetFontAndSize(bfArialI, 12);
				tpl.BeginText();
				for (int i = 0; i < tabcol.Length; i++) {
					float offset = tabcol[i] / 2;
					{
						string[] warr = tabtext[i].Split('|');
						if (warr.Length > 1) {
							float yoffset = 30f - 4.5f * warr.Length + 1.5f;
							for (int j = 0; j < warr.Length; j++) {
								tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, warr[warr.Length-1-j], InMM(tempX + offset), InMM(yoffset), 0);
								yoffset += 4.5f;
							}
						}
						else {
							tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tabtext[i], InMM(tempX + offset), InMM(20f), 0);
						}
					}
					tpl.ShowTextAligned(PdfContentByte.ALIGN_CENTER, (i+1).ToString(), InMM(tempX + offset), InMM(1.5f), 0);
					tempX += tabcol[i];
				}
				tpl.EndText();
			}
			catch (Exception ex) {
				PrintMess("Ошибка создания основной надписи. " + ex.Message, 2);
				tpl = null;
			}
			return tpl;
		}

		#endregion

		//========================================================================

		#region Excel

		/// <summary>
		/// Открывает файл Excel для чтения
		/// </summary>
		/// <returns>При успехе возвращает true</returns>
		private bool OpenExcelFile() {
			bool Result = false;
			//Подключение к Excel
			try {
				exapp.Visible = false;
			}
			catch //(System.NullReferenceException ex)
			{
				exapp = new Excel.Application();
			}

			try {
				wbSource = exapp.Workbooks.Open(ExcelFileName,
					Type.Missing, true, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing);
				AllData = new List<GroupData>();
				Result = true;
			}
			catch (Exception ex) {
				PrintMess("Ошибка открытия файла Excel " + ExcelFileName, 2);
				PrintMess(ex.Message, 2);
			}
			return Result;
		}

		/// <summary>
		/// Считывает данные группы со страницы Excel
		/// </summary>
		/// <param name="exPageNum">Номер страницы Excel</param>
		/// <returns>При отсутствии ошибок возвращает true</returns>
		private bool ReadExcelPage(int exPageNum) {
			bool Result = false;
			if (!excelOpened)
				return Result;
			try {
				wsSource = (Excel.Worksheet)wbSource.Worksheets[exPageNum];
				GroupData nodeGD = new GroupData();
				List<ElementData>  nodeArr = new List<ElementData>();
				
				Excel.Range rng;
				rng = wsSource.get_Range("K1", Type.Missing);
				string GN = rng.Value2.ToString().Trim();
				if (GN.Length > 0)
					nodeGD.GroupName = GN;
				else  
					throw new Exception("Не найдено название группы!");

				rng = wsSource.get_Range("A1", Type.Missing);
				bool CanRead = false;
				if (rng.Value2 != null)
					CanRead = true;
				//Циклп по элементам группы на листе
				while (CanRead) {
					if (rng.Value2 == null) {
						//Данных на странице больше нет
						CanRead = false;
						continue;
					}
					Excel.Range workRng = rng;
					ElementData nodeElement = new ElementData();
					nodeElement.Pos = "";
					{
						string tempStr = workRng.Value2.ToString().Trim();	//A
						workRng = workRng.get_Offset(0, 1);	//B
						if (workRng.Value2 != null) {
							nodeElement.ElType = workRng.Value2.ToString().Trim();
							//nodeElement.CodeName = workRng.Value2.ToString().Trim();
						}
						else {
							//nodeElement.CodeName = "";
							nodeElement.ElType = string.Empty;
						}
						Regex rgx = new Regex(@"Опросный");
						if (!rgx.Match(tempStr).Success) {
							//Нет опросного листа
							nodeElement.ElName = tempStr;
						}
						else {
							//Выделим опросный лист
							string[] sArr = rgx.Split(tempStr);
							nodeElement.ElName = sArr[0].Trim();
							nodeElement.ElType = string.Format("{0}{1}", ("Опросный" + sArr[1]),
								(nodeElement.ElType == string.Empty) ? string.Empty : (Environment.NewLine + nodeElement.ElType));
						}
					}
					//nodeElement.CodeName = "";
					nodeElement.ElMaker = "";

					workRng = workRng.get_Offset(0, 1);	//C
					if (workRng.Value2 != null)
						nodeElement.ElUnit = workRng.Value2.ToString().Trim();
					else
						PrintMess("Ошибка: нет ед. изм. - " + nodeElement.ElName, 2);

					workRng = workRng.get_Offset(0, 1);	//D
					if (workRng.Value2 != null)
						nodeElement.ElTotal = workRng.Value2.ToString().Trim().Replace(".00", string.Empty) ;
					else
						PrintMess("Ошибка: нет количества - " + nodeElement.ElName, 2);

					workRng = workRng.get_Offset(0, 1);	//E
					if (workRng.Value2 != null)
						nodeElement.ElMass = workRng.Value2.ToString().Trim();
					else {
						nodeElement.ElMass = string.Empty;
					}

					workRng = workRng.get_Offset(0, 1);	//F
					if (workRng.Value2 != null)
						nodeElement.CodeName = workRng.Value2.ToString().Trim();
					else {
						nodeElement.CodeName = string.Empty;
					}

					workRng = workRng.get_Offset(0, 1);	//G
					if (workRng.Value2 != null)
						nodeElement.Notes = workRng.Value2.ToString().Trim();
					else {
						nodeElement.Notes = string.Empty;
					}

					nodeArr.Add(nodeElement);

					//На следующую строке
					rng = rng.get_Offset(1, 0);
				}
				nodeGD.elArray = nodeArr;
				AllData.Add(nodeGD);
				Result = true;
			}
			catch (Exception ex) {
				PrintMess("Ошибка чтения данных с листа " + exPageNum.ToString(), 2);
				PrintMess(ex.Message, 2);
			}
			return Result;
		}

		#endregion
//========================================================================
https://github.com/microsoft/XamlBehaviorsWpf/wiki/InvokeCommandAction
//========================================================================
https://kb.itextpdf.com/home/it7kb/ebooks/itext-jump-start-tutorial-for-java/chapter-6-reusing-existing-pdf-documents

	}
}

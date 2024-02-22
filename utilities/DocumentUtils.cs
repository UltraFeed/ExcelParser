#pragma warning disable CS8602

using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using ExcelParser.modules;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.Util;
using NPOI.XWPF.Model;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser.utilities;

internal static class DocumentUtils
{
	internal static void AddParagraph (XWPFDocument doc, string text, int fontSize = 13, bool isBold = false, bool tab = false, bool spacing = false, ParagraphAlignment alignment = ParagraphAlignment.LEFT)
	{
		XWPFParagraph paragraph = doc.CreateParagraph();
		paragraph.Alignment = alignment;

		XWPFRun run = paragraph.CreateRun();
		run.SetText(text);
		run.FontFamily = "Arial";
		run.FontSize = fontSize;
		run.IsBold = isBold;
		if (tab)
		{
			run.AddTab();
		}

		if (spacing)
		{
			paragraph.SpacingAfterLines = 0;
			paragraph.SpacingBeforeLines = 0;
			paragraph.SpacingBefore = 0;
			paragraph.SpacingAfter = 0;
		}
	}

	internal static void SetColorfulBlock (XWPFDocument doc, string message1, string message2, string color)
	{
		XWPFParagraph paragraph = doc.CreateParagraph();
		XWPFRun text41 = paragraph.CreateRun();
		text41.SetText(message1);
		text41.FontFamily = "Arial";
		text41.FontSize = 13;
		text41.IsBold = true;
		text41.SetColor(color);

		XWPFRun text42 = paragraph.CreateRun();
		text42.SetText(message2);
		text42.FontFamily = "Arial";
		text42.FontSize = 13;
	}

	internal static void SetPcNumbers (XWPFDocument doc, string message1, string message2)
	{
		XWPFParagraph paragraph = doc.CreateParagraph();
		XWPFRun text41 = paragraph.CreateRun();
		text41.SetText(message1);
		text41.FontFamily = "Arial";
		text41.FontSize = 13;

		XWPFRun text42 = paragraph.CreateRun();
		text42.SetText(message2);
		text42.FontFamily = "Arial";
		text42.FontSize = 13;
		text42.IsBold = true;
	}

	internal static void CreateNullParagraphs (XWPFDocument doc, int counter)
	{
		for (int i = 0; i < counter; i++)
		{
			_ = doc.CreateParagraph();
		}
	}

	// Функция для создания верхнего колонтитула
	private static void CreateHeader (XWPFDocument doc)
	{
		// Создание экземпляра для управления верхним и нижним колонтитулами
		XWPFHeaderFooterPolicy headerFooterPolicy = doc.CreateHeaderFooterPolicy();

		// Создание верхнего колонтитула
		XWPFHeader header = headerFooterPolicy.CreateHeader(XWPFHeaderFooterPolicy.DEFAULT);

		// Создание параграфа в верхнем колонтитуле
		XWPFParagraph headerParagraph = header.CreateParagraph();
		headerParagraph.Alignment = ParagraphAlignment.LEFT;
		headerParagraph.IndentationLeft = -300;

		headerParagraph.IndentationFirstLine = 0;
		headerParagraph.IndentationHanging = 0;
		headerParagraph.IsWordWrapped = false;
		headerParagraph.SpacingBefore = 0;
		headerParagraph.SpacingBeforeLines = 0;
		headerParagraph.SpacingAfter = 0;
		headerParagraph.SpacingAfterLines = 0;
		headerParagraph.FirstLineIndent = 0;
		headerParagraph.IndentationFirstLine = 0;
		headerParagraph.IndentationHanging = 0;

		// Добавление изображения в верхний колонтитул
		string logoSmall = "ExcelParser.resources.logoSmall.png";
		using Stream? imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(logoSmall);
		{
			byte [] imageBytes = NPOI.Util.IOUtils.ToByteArray(imageStream, (int) imageStream.Length);
			_ = headerParagraph.CreateRun().AddPicture(new MemoryStream(imageBytes), (int) PictureType.PNG, logoSmall, Units.ToEMU(30), Units.ToEMU(30));
		}
	}

	// Функция для создания нижнего колонтитула
	private static void CreateFooter (XWPFDocument doc)
	{
		// Создание экземпляра для управления верхним и нижним колонтитулами
		XWPFHeaderFooterPolicy headerFooterPolicy = doc.CreateHeaderFooterPolicy();

		// Создание нижнего колонтитула
		XWPFFooter footer = headerFooterPolicy.CreateFooter(XWPFHeaderFooterPolicy.DEFAULT);
		XWPFParagraph paragraph = footer.GetParagraphArray(0);
		paragraph ??= footer.CreateParagraph();
		paragraph.Alignment = ParagraphAlignment.CENTER;
		XWPFRun run = paragraph.CreateRun();
		run.SetText("г. Омск");

		paragraph.IndentationFirstLine = 0;
		paragraph.IndentationHanging = 0;
		paragraph.IsWordWrapped = false;
		paragraph.SpacingBefore = 0;
		paragraph.SpacingAfter = 0;
		paragraph.SpacingAfterLines = 0;
		paragraph.FirstLineIndent = 0;
		paragraph.IndentationFirstLine = 0;
		paragraph.IndentationHanging = 0;

		paragraph = footer.GetParagraphArray(1);
		paragraph ??= footer.CreateParagraph();
		paragraph.Alignment = ParagraphAlignment.CENTER;
		run = paragraph.CreateRun();
		run.SetText("8 800 775 75 91");

		// Установка отступа от нижнего края через объект CTSectPr
		CT_SectPr sectPr = doc.Document.body.sectPr ?? doc.Document.body.AddNewSectPr();
		CT_PageMar pageMar = sectPr.AddPageMar();
		pageMar.footer = 565;
		pageMar.header = 5;
	}

	private static void AddMainPage (ExcelWorksheet worksheet, XWPFDocument doc)
	{
		Debug.WriteLine("\nSTART DEBUG MESSAGES\n");
		List<string> pcNumbers = ConstantsUtils.GetPcList(worksheet);

		CreateNullParagraphs(doc, 4);

		// Добавляем параграф с изображением
		ParagraphAlignment alignment = ParagraphAlignment.CENTER;
		Debug.WriteLine("Добавляю лого");
		XWPFParagraph pictureParagraph = doc.CreateParagraph();
		pictureParagraph.IndentationLeft = -500;
		pictureParagraph.Alignment = alignment;

		// Добавляем само изображение
		string logoLarge = "ExcelParser.resources.logoLarge.png";
		using Stream? imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(logoLarge);
		{
			byte [] imageBytes = NPOI.Util.IOUtils.ToByteArray(imageStream, (int) imageStream.Length);
			_ = pictureParagraph.CreateRun().AddPicture(new MemoryStream(imageBytes), (int) PictureType.PNG, logoLarge, Units.ToEMU(490), Units.ToEMU(185));
		}

		CreateNullParagraphs(doc, 1);

		string message1 = $"Отчёт";
		string message2 = $"ИТ Аудит";
		string message3 = $"{Constants.companyName}";
		string message4 = $"Филиал на {Constants.companyAddress}";
		string message5 = $"{DateTime.Now.ToString("MMMM yyyy", new CultureInfo("ru-RU"))}";
		string message6 = $"1. Цели аудита";
		string message7 = $"Выдача рекомендаций по повышению отказоустойчивости рабочих станций и локальной сети, обеспечения сохранности и защите информации, организационным вопросам.";
		string message8 = $"1.1. Объекты аудита";
		string message9 = $"Рабочие станции (РС), периферия, доступ в Интернет, локальная сеть.";
		string message10 = $"2. Общее описание объекта аудита";
		string message11 = $"- Рабочих Станций: {pcNumbers.Count}";
		string message12 = $"- Сетевое оборудование";
		string message13 = $"3. Анализ и рекомендации";

		Debug.WriteLine(message1);
		Debug.WriteLine(message2);
		Debug.WriteLine(message3);
		Debug.WriteLine(message4);
		Debug.WriteLine(message5);
		Debug.WriteLine(message6);
		Debug.WriteLine(message7);
		Debug.WriteLine(message8);
		Debug.WriteLine(message9);
		Debug.WriteLine(message10);
		Debug.WriteLine(message11);
		Debug.WriteLine(message12);
		Debug.WriteLine(message13);

		// Добавляем параграфы и тексты
		AddParagraph(doc, message1, fontSize: 48, isBold: true, alignment: alignment);
		AddParagraph(doc, message2, fontSize: 48, isBold: true, alignment: alignment);
		AddParagraph(doc, message3, fontSize: 28, alignment: alignment);
		AddParagraph(doc, message4, fontSize: 28, alignment: alignment);

		CreateNullParagraphs(doc, 4);

		// Добавляем параграф с датой
		AddParagraph(doc, message5, fontSize: 24, isBold: true, alignment: alignment);

		AddParagraph(doc, message6, fontSize: 16, isBold: true); // 1. Цели аудита
		AddParagraph(doc, message7); // Выдача рекомендаций по...
		CreateNullParagraphs(doc, 1);
		AddParagraph(doc, message8, fontSize: 16, isBold: true); // 1.1. Объекты аудита
		AddParagraph(doc, message9); // Рабочие станции (РС), периферия, доступ в Интернет...
		CreateNullParagraphs(doc, 1);
		AddParagraph(doc, message10, fontSize: 16, isBold: true); // 2. Общее описание объекта аудита
		AddParagraph(doc, message11, tab: true); // - {pcCounter} Рабочих Станций
		AddParagraph(doc, message12, tab: true); // Сетевое оборудование
		CreateNullParagraphs(doc, 1);
		AddParagraph(doc, message13, fontSize: 16, isBold: true); // 3. Анализ и рекомендации

		Debug.WriteLine("\nEND DEBUG MESSAGES\n");
	}

	internal static void FillDoc (XWPFDocument doc, ExcelWorksheet worksheet)
	{
		// создаем нижний и верхний колонтикул
		DocumentUtils.CreateFooter(doc);
		DocumentUtils.CreateHeader(doc);

		// добавляем главную страницу документа
		DocumentUtils.AddMainPage(worksheet, doc);

		// поиск и вывод информации по модулям
		List<string> badDefenderPcNumbers = Defender.SearchAndPrintDefender(worksheet, doc);
		List<string> badPowerPcNumbers = Power.SearchAndPrintPower(worksheet, doc);
		List<string> badInternetPcNumbers = Internet.SearchAndPrintInternet(worksheet, doc);
		List<string> badFileServerPcNumbers = FileServer.SearchAndPrintFileServer(worksheet, doc);

		DocumentUtils.CreateNullParagraphs(doc, 1);
		DocumentUtils.AddParagraph(doc, "4. Модернизация", fontSize: 16, isBold: true);

		List<string> badSystemDrivePcNumbers = SystemDrive.SearchAndPrintSystemDrive(worksheet, doc);
		List<string> badDisplayPcNumbers = Display.SearchAndPrintDisplay(worksheet, doc);
		List<string> badMouseKeyboardPcNumbers = MouseKeyboard.SearchAndPrintMouseKeyboard(worksheet, doc);

		// выводим блок с результатами по каждому ПК
		DocumentUtils.Summarize(worksheet, doc,
			badDefenderPcNumbers,
			badPowerPcNumbers,
			badInternetPcNumbers,
			badSystemDrivePcNumbers,
			badFileServerPcNumbers,
			badDisplayPcNumbers,
			badMouseKeyboardPcNumbers);
	}

	// Функция для вывода информации по каждому ПК
	private static void Summarize (ExcelWorksheet worksheet, XWPFDocument doc,
	List<string> badDefenderPcNumbers,
	List<string> badPowerPcNumbers,
	List<string> badInternetPcNumbers,
	List<string> badSystemDrivePcNumbers,
	List<string> badFileServerPcNumbers,
	List<string> badDisplayPcNumbers,
	List<string> badMouseKeyboardPcNumbers)
	{
		Debug.WriteLine($"\nSTART DEBUG MESSAGES\n");

		CreateNullParagraphs(doc, 1);
		AddParagraph(doc, $"5. Сводка по каждому ПК:", fontSize: 16, isBold: true);
		CreateNullParagraphs(doc, 1);

		foreach (string pcNumber in ConstantsUtils.GetPcList(worksheet))
		{
			Debug.WriteLine("");
			string currentPcNumber = $"Номер ПК: {pcNumber}";
			string currentUserName = $"Пользователь: {ConstantsUtils.GetUserForPcNumber(worksheet, pcNumber)}";
			string currentUserPosition = $"Должность: {ConstantsUtils.GetUserPositionForPcNumber(worksheet, pcNumber)}";

			AddParagraph(doc, currentPcNumber, spacing: true);
			AddParagraph(doc, currentUserName, spacing: true);
			AddParagraph(doc, currentUserPosition, spacing: true);

			Debug.WriteLine(currentPcNumber);
			Debug.WriteLine(currentUserName);
			Debug.WriteLine(currentUserPosition);

			AddParagraph(doc, $"Проблемы:", spacing: true);

			// Defender
			if (badDefenderPcNumbers.Equals(pcNumber))
			{
				string message = $"Отсутствует Антивирус";
				Debug.WriteLine(message);
				AddParagraph(doc, message, spacing: true);
			}

			// Power
			if (badPowerPcNumbers.Equals(pcNumber))
			{
				string message = $"Отсутствует ИБП";
				Debug.WriteLine(message);
				AddParagraph(doc, message, spacing: true);
			}

			// Internet
			if (badInternetPcNumbers.Equals(pcNumber))
			{
				string message = $"Плохая работа интернета";
				Debug.WriteLine(message);
				AddParagraph(doc, message, spacing: true);
			}

			// SystemDrive
			if (badSystemDrivePcNumbers.Equals(pcNumber))
			{
				string message = $"В качестве системного диска должен быть установлен SSD";
				Debug.WriteLine(message);
				AddParagraph(doc, message, spacing: true);
			}

			// FileServer
			if (badFileServerPcNumbers.Equals(pcNumber))
			{
				string message = $"Является файловым сервером";
				Debug.WriteLine(message);
				AddParagraph(doc, message, spacing: true, isBold: true);
			}

			// Display
			if (badDisplayPcNumbers.Equals(pcNumber))
			{
				string message = $"С монитором";
				Debug.WriteLine(message);
				AddParagraph(doc, message, spacing: true);
			}

			// MouseKeyboard
			if (badMouseKeyboardPcNumbers.Equals(pcNumber))
			{
				string message = $"С клавиатурой и мышью";
				Debug.WriteLine(message);
				AddParagraph(doc, message, spacing: true);
			}

			// Пустая строка для разделения информации по разным ПК
			CreateNullParagraphs(doc, 1);
		}

		Debug.WriteLine($"\nEND DEBUG MESSAGES\n");
	}
}
using System.Diagnostics;
using ExcelParser.utilities;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser.modules;

internal static class Display
{

	internal static List<string> SearchAndPrintDisplay (ExcelWorksheet worksheet, XWPFDocument doc)
	{
		Debug.WriteLine($"\nSTART DEBUG MESSAGES\n");
		List<string> troubledPcNumbers = [];

		// Перебор строк в столбце
		for (int row = Constants.firstDataRow; row <= worksheet.Dimension.End.Row; row++)
		{
			// Получение значения ячейки в столбце компаний
			string currentCompany = worksheet.Cells [row, Constants.companiesNamesColumn].Text;

			//получаем значение ячейки в столбце адресов филиалов
			string currentAddress = worksheet.Cells [row, Constants.companiesAddressesColumn].Text;

			//получаем значение ячейки в столбце дат
			string currentDate = worksheet.Cells [row, Constants.dateColumn].Text;

			// Проверка соответствия заданному названию компании и адресу филиала и дате
			if (currentCompany.Equals(Constants.companyName, StringComparison.OrdinalIgnoreCase) &&
				currentAddress.Equals(Constants.companyAddress, StringComparison.OrdinalIgnoreCase) &&
				ConstantsUtils.IsDateInRange(currentDate))
			{
				// Получение значения номера ПК
				string pcNumberCell = worksheet.Cells [row, Constants.pcNumbersColumn].Text;

				// Получение значения ячейки в столбце
				string currentCellValue = worksheet.Cells [row, Constants.displayColumn].Text;

				// Проверка наличия подстроки "отсутствует"
				if (!currentCellValue.Contains("отсутствует", StringComparison.OrdinalIgnoreCase))
				{
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} есть проблемы с монитором: {currentCellValue}");
					troubledPcNumbers.Add(pcNumberCell);
				}
				else
				{
					// если нет проблем с монитором
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} нет проблем с монитором: {currentCellValue}");
				}
			}
		}

		Debug.WriteLine("");
		DocumentUtils.CreateNullParagraphs(doc, 1);
		string message1 = $"4.2 Мониторы";
		Debug.WriteLine(message1);
		DocumentUtils.AddParagraph(doc, message1, fontSize: 16, isBold: true);

		if (troubledPcNumbers.Count != 0)
		{
			string message21 = $"Выявлено: ";
			string message22 = $"не на всех ПК мониторы работают корректно";
			string message31 = $"Риски: ";
			string message32 = $"Медленная работа сотрудников";
			string message41 = $"Рекомендации: ";
			string message42 = $"приобрести новые или отремонтировать мониторы на {troubledPcNumbers.Count} компьютере(ах).";
			string message51 = $"Номера ПК, имеющих проблемы с мониторами: ";
			string message52 = string.Join(", ", troubledPcNumbers);

			Debug.WriteLine(message21);
			Debug.WriteLine(message22);
			Debug.WriteLine(message31);
			Debug.WriteLine(message32);
			Debug.WriteLine(message41);
			Debug.WriteLine(message42);
			Debug.WriteLine(message51);
			Debug.WriteLine(message52);

			DocumentUtils.SetColorfulBlock(doc, message21, message22, "BLACK"); //Выявлено:
			DocumentUtils.SetColorfulBlock(doc, message31, message32, "RED"); //Риски:
			DocumentUtils.SetColorfulBlock(doc, message41, message42, "GREEN"); //Рекомендации:
			DocumentUtils.SetPcNumbers(doc, message51, message52); // номера ПК
		}
		else
		{
			string message0 = $"На ваших ПК нет проблем с мониторами";
			Debug.WriteLine(message0);
			DocumentUtils.AddParagraph(doc, message0);
		}
		Debug.WriteLine($"\nEND DEBUG MESSAGES\n");

		return troubledPcNumbers;
	}
}
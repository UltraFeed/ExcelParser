using System.Diagnostics;
using ExcelParser.utilities;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser.modules;

internal static class Power
{

	internal static List<string> SearchAndPrintPower (ExcelWorksheet worksheet, XWPFDocument doc)
	{
		Debug.WriteLine($"\nSTART DEBUG MESSAGES\n");
		List<string> troubledPCNumbers = [];

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
				string currentCellValue = worksheet.Cells [row, Constants.powerSupplyColumn].Text;

				// Проверка наличия подстроки "отсутствует" или "не держит" в значении столбца ИБП
				if (currentCellValue.Contains("отсутствует", StringComparison.OrdinalIgnoreCase) || currentCellValue.Contains("не держит", StringComparison.OrdinalIgnoreCase))
				{
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} ИБП {currentCellValue}");
					troubledPCNumbers.Add(pcNumberCell);
				}
				else
				{
					// Если ИБП есть и работает
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} ИБП {currentCellValue}");
				}
			}
		}

		Debug.WriteLine("");
		DocumentUtils.CreateNullParagraphs(doc, 1);
		string message1 = $"3.2 Отказоустойчивость рабочих станций";
		Debug.WriteLine(message1);
		DocumentUtils.AddParagraph(doc, message1, fontSize: 16, isBold: true);

		if (troubledPCNumbers.Count != 0)
		{
			string message21 = $"Выявлено: ";
			string message22 = $"не каждый ПК подключён к источнику бесперебойного питания.";
			string message31 = $"Риски: ";
			string message32 = $"потеря данных и отсутствие возможности сохранить работу при внезапных перебоях с электричеством.";
			string message41 = $"Рекомендации: ";
			string message42 = $"Использование для защиты рабочих станций источников бесперебойного питания на {troubledPCNumbers.Count} компьютер(ах).";
			string message51 = $"Номера ПК без ИБП: ";
			string message52 = string.Join(", ", troubledPCNumbers);

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
			string message0 = $"Каждый ПК подключён к источнику бесперебойного питания.";
			DocumentUtils.AddParagraph(doc, message0);
		}

		Debug.WriteLine($"\nEND DEBUG MESSAGES\n");

		return troubledPCNumbers;
	}
}
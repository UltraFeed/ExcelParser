﻿using System.Diagnostics;
using ExcelParser.utilities;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser.modules;

internal static class MouseKeyboard
{

	internal static List<string> SearchAndPrintMouseKeyboard (ExcelWorksheet worksheet, XWPFDocument doc)
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
				string currentCellValue = worksheet.Cells [row, Constants.mouseKeyboardColumn].Text;

				// Проверка наличия подстроки "отсутствуют"
				if (!currentCellValue.Contains("отсутствуют", StringComparison.OrdinalIgnoreCase))
				{
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} проблемы с мышью/клавиатурой: {currentCellValue}");
					troubledPCNumbers.Add(pcNumberCell);
				}
				else
				{
					// если нет проблем с монитором
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} нет проблем с мышью/клавиатурой: {currentCellValue}");
				}
			}
		}

		Debug.WriteLine("");
		DocumentUtils.CreateNullParagraphs(doc, 1);
		string message1 = $"4.3 Клавиатуры и мыши";
		Debug.WriteLine(message1);
		DocumentUtils.AddParagraph(doc, message1, fontSize: 16, isBold: true);

		if (troubledPCNumbers.Count != 0)
		{
			string message21 = $"Выявлено: ";
			string message22 = $"не на всех ПК мыши и клавиатуры работают корректно";
			string message31 = $"Риски: ";
			string message32 = $"Медленная работа сотрудников";
			string message41 = $"Рекомендации: ";
			string message42 = $"приобрести новые мыши и клавиатуры для {troubledPCNumbers.Count} компьютера(ов).";
			string message51 = $"Номера ПК, имеющих проблемы с клавиатурой или мышью: ";
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
			string message0 = $"На ваших ПК нет проблем с клавиатурами и мышками";
			Debug.WriteLine(message0);
			DocumentUtils.AddParagraph(doc, message0);
		}
		Debug.WriteLine($"\nEND DEBUG MESSAGES\n");

		return troubledPCNumbers;
	}
}
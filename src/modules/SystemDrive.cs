using System.Diagnostics;
using ExcelParser.utilities;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser.modules;

internal static class SystemDrive
{
	internal static List<string> SearchAndPrintSystemDrive (ExcelWorksheet worksheet, XWPFDocument doc)
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
				string currentCellValue = worksheet.Cells [row, Constants.systemDriveColumn].Text;

				// Проверка наличия подстроки "HDD" в ячейке интернета
				if (currentCellValue.Contains("HDD", StringComparison.OrdinalIgnoreCase))
				{
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} системный диск - {currentCellValue}");
					troubledPCNumbers.Add(pcNumberCell);
				}
				else
				{
					// если диск нормальный
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} системный диск - {currentCellValue}");
				}
			}
		}

		Debug.WriteLine("");
		DocumentUtils.CreateNullParagraphs(doc, 1);
		string message1 = $"4.1 Системные диски";
		Debug.WriteLine(message1);
		DocumentUtils.AddParagraph(doc, message1, fontSize: 16, isBold: true);

		if (troubledPCNumbers.Count != 0)
		{
			string message21 = $"Выявлено: ";
			string message22 = $"не на всех ПК установлены SSD-накопителей в качестве системных.";
			string message31 = $"Риски: ";
			string message32 = $"потеря данных из-за износа накопителей и возникновение существенных затруднений при попытке восстановления данных.";
			string message41 = $"Рекомендации: ";
			string message42 = $"Установка SSD-накопителей на {troubledPCNumbers.Count} компьютер(ов).";
			string message51 = $"Номера ПК c носителями плохого качества: ";
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
			string message0 = $"На ваших ПК нет проблем с системными дисками.";
			DocumentUtils.AddParagraph(doc, message0);
		}

		Debug.WriteLine($"\nEND DEBUG MESSAGES\n");

		return troubledPCNumbers;
	}
}

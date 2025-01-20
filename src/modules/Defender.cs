using System.Diagnostics;
using ExcelParser.utilities;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser.modules;

internal static class Defender
{

	internal static List<string> SearchAndPrintDefender (ExcelWorksheet worksheet, XWPFDocument doc)
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
				string currentCellValue = worksheet.Cells [row, Constants.defenderTypesColumn].Text;

				// Проверка наличия подстроки "отсутствует" или "бесплатный" в типе антивируса
				if (currentCellValue.Contains("отсутствует", StringComparison.OrdinalIgnoreCase) || currentCellValue.Contains("бесплатный", StringComparison.OrdinalIgnoreCase))
				{
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} Антивирус {currentCellValue}");
					troubledPcNumbers.Add(pcNumberCell);
				}
				else
				{
					// если антивирус не бесплатный или не отсутствует
					Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. На ПК №:{pcNumberCell} {currentCellValue} (качественный)");
				}
			}
		}

		Debug.WriteLine("");
		DocumentUtils.CreateNullParagraphs(doc, 1);
		string message1 = $"3.1 Антивирусная защита и защита от кражи информации";
		Debug.WriteLine(message1);
		DocumentUtils.AddParagraph(doc, message1, fontSize: 16, isBold: true);

		if (troubledPcNumbers.Count != 0)
		{
			string message21 = $"Выявлено: ";
			string message22 = $"не на всех ПК установлен антивирус и настроен Firewall.";
			string message31 = $"Риски: ";
			string message32 = $"потеря данных, утечка информации в интернет, а также нарушение правильной работоспособности компьютеров.";
			string message41 = $"Рекомендации: ";
			string message42 = $"приобрести антивирус Dr. Web для {troubledPcNumbers.Count} компьютер(ов) (Dr. Web не поддерживается на Windows 7 и старее).";
			string message51 = $"Номера ПК без антивируса: ";
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
			DocumentUtils.SetPcNumbers(doc, message51, message52); // Номера ПК
		}
		else
		{
			string message0 = $"На всех ваших ПК установлены качественные антивирусы и настроен Firewall.";
			Debug.WriteLine(message0);
			DocumentUtils.AddParagraph(doc, message0);
		}

		Debug.WriteLine($"\nEND DEBUG MESSAGES\n");

		return troubledPcNumbers;
	}
}

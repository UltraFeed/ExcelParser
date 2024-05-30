using System.Diagnostics;
using ExcelParser.utilities;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser.modules;

internal static class FileServer
{
    internal static List<string> SearchAndPrintFileServer (ExcelWorksheet worksheet, XWPFDocument doc)
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
                string currentCellValue = worksheet.Cells [row, Constants.fileServerColumn].Text;

                // Проверка наличия подстроки "сервер" в ячейке типа носителя
                if (currentCellValue.Contains("сервер", StringComparison.OrdinalIgnoreCase))
                {
                    Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. ПК №:{pcNumberCell} - {currentCellValue}");
                    troubledPcNumbers.Add(pcNumberCell);
                }
                else
                {
                    // если не файловый сервер
                    Debug.WriteLine($"Найдено значение для {Constants.companyName}(адрес: {currentAddress}) в строке {row}. ПК №:{pcNumberCell} не файловый сервер. значение: {currentCellValue}");
                }
            }
        }

        Debug.WriteLine("");
        DocumentUtils.CreateNullParagraphs(doc, 1);
        string message1 = $"3.4 Резервирование и хранение данных";
        Debug.WriteLine(message1);
        DocumentUtils.AddParagraph(doc, message1, fontSize: 16, isBold: true);

        if (troubledPcNumbers.Count != 0)
        {
            string message21 = $"Выявлено: ";
            string message22 = $"отсутствует централизованный сервер резервирования данных. В роли файлового сервера выступает ПК №{troubledPcNumbers [0]} Сотрудника {ConstantsUtils.GetUserForPcNumber(worksheet, troubledPcNumbers [0])}. Должность сотрудника - {ConstantsUtils.GetUserPositionForPcNumber(worksheet, troubledPcNumbers [0])}.";
            string message31 = $"Риски: ";
            string message32 = $"В случае выхода из строя ПК №{troubledPcNumbers [0]} доступ к общим папкам будет утрачен. При случайном удалении или потере данных, восстановление будет невозможным.";
            string message41 = $"Рекомендации: ";
            string message42 = $"приобретение отдельного централизованного сервера для выполнения роли 'обменника файлами' и настройка на нем резервирования данных (бэкапов).";

            Debug.WriteLine(message21);
            Debug.WriteLine(message22);
            Debug.WriteLine(message31);
            Debug.WriteLine(message32);
            Debug.WriteLine(message41);
            Debug.WriteLine(message42);

            DocumentUtils.SetColorfulBlock(doc, message21, message22, "BLACK"); //Выявлено:
            DocumentUtils.SetColorfulBlock(doc, message31, message32, "RED"); //Риски:
            DocumentUtils.SetColorfulBlock(doc, message41, message42, "GREEN"); //Рекомендации:
        }
        else
        {
            string message0 = $"C Резервированием и хранением данных проблем не обнаружено";
            DocumentUtils.AddParagraph(doc, message0);
        }

        Debug.WriteLine($"\nEND DEBUG MESSAGES\n");

        return troubledPcNumbers;
    }
}
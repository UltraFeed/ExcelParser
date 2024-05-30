using System.Globalization;
using OfficeOpenXml;

namespace ExcelParser.utilities;

internal static class ConstantsUtils
{
    internal static List<string> GetCompanyNames (ExcelWorksheet worksheet)
    {
        return worksheet
            .Cells [Constants.companiesNamesColumn, Constants.companiesNamesColumn, worksheet.Dimension.End.Row, Constants.companiesNamesColumn]
            .Select(cell => cell.Text)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    internal static List<string> GetCompanyAddresses (ExcelWorksheet worksheet, string companyName)
    {
        return worksheet
            .Cells [Constants.companiesAddressesColumn, Constants.companiesAddressesColumn, worksheet.Dimension.End.Row, Constants.companiesAddressesColumn]
            .Where(cell => worksheet.Cells [cell.Start.Row, Constants.companiesNamesColumn].Text.Equals(companyName, StringComparison.OrdinalIgnoreCase))
            .Select(cell => cell.Text)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    internal static bool CheckOccupation (string filePath)
    {
        try
        {
            // Попытка открыть файл на запись с блокировкой
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Write);
            // Если успешно открыт, то файл не занят другой программой
            return false;
        }
        catch (IOException)
        {
            // Если возникло исключение IOException, то файл занят другой программой
            return true;
        }
    }

    internal static bool IsDateInRange (string inputDate)
    {
        if (DateTime.TryParseExact(inputDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
        {
            return parsedDate >= Constants.firstDate && parsedDate <= Constants.secondDate;
        }
        else
        {
            return false;
        }
    }

    // функция для очистки названия и адреса от недопустимых символов, нужна для имени выходного .docx файла
    internal static string Sanitize (string input)
    {
        return Path.GetInvalidFileNameChars().Aggregate(input, (current, invalidChar) => current.Replace(invalidChar, '_'));
    }

    internal static string GetUserForPcNumber (ExcelWorksheet worksheet, string targetPcNumber)
    {
        for (int row = Constants.firstDataRow; row <= worksheet.Dimension.End.Row; row++)
        {
            // Получение значения ячейки в столбце компаний
            string currentCompany = worksheet.Cells [row, Constants.companiesNamesColumn].Text;

            //получаем значение ячейки в столбце адресов филиалов
            string currentAddress = worksheet.Cells [row, Constants.companiesAddressesColumn].Text;

            //получаем значение ячейки в столбце дат
            string currentDate = worksheet.Cells [row, Constants.dateColumn].Text;

            // получаем номер ПК
            string pcNumberCell = worksheet.Cells [row, Constants.pcNumbersColumn].Text;

            if (currentCompany.Equals(Constants.companyName, StringComparison.OrdinalIgnoreCase) &&
                currentAddress.Equals(Constants.companyAddress, StringComparison.OrdinalIgnoreCase) &&
                IsDateInRange(currentDate) &&
                pcNumberCell.Equals(targetPcNumber, StringComparison.OrdinalIgnoreCase))
            {
                return worksheet.Cells [row, Constants.userNamesColumn].Text;
            }
        }
        // если не найдено соответствие
        return $"ФИО не указано";
    }

    internal static string GetUserPositionForPcNumber (ExcelWorksheet worksheet, string targetPcNumber)
    {
        for (int row = Constants.firstDataRow; row <= worksheet.Dimension.End.Row; row++)
        {
            // Получение значения ячейки в столбце компаний
            string currentCompany = worksheet.Cells [row, Constants.companiesNamesColumn].Text;

            //получаем значение ячейки в столбце адресов филиалов
            string currentAddress = worksheet.Cells [row, Constants.companiesAddressesColumn].Text;

            //получаем значение ячейки в столбце дат
            string currentDate = worksheet.Cells [row, Constants.dateColumn].Text;

            // получаем номер ПК
            string pcNumberCell = worksheet.Cells [row, Constants.pcNumbersColumn].Text;

            if (currentCompany.Equals(Constants.companyName, StringComparison.OrdinalIgnoreCase) &&
                currentAddress.Equals(Constants.companyAddress, StringComparison.OrdinalIgnoreCase) &&
                IsDateInRange(currentDate) &&
                pcNumberCell.Equals(targetPcNumber, StringComparison.OrdinalIgnoreCase))
            {
                return worksheet.Cells [row, Constants.userPositionsColumn].Text;
            }
        }
        // если не найдено соответствие
        return $"Должность не указана";
    }

    internal static List<string> GetPcList (ExcelWorksheet worksheet)
    {
        List<string> pcNumbers = [];

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
                IsDateInRange(currentDate))
            {
                // Получение значения номера ПК
                string pcNumberCell = worksheet.Cells [row, Constants.pcNumbersColumn].Text;
                pcNumbers.Add(pcNumberCell);
            }
        }

        pcNumbers.Sort();
        return pcNumbers;
    }
}
#pragma warning disable CS8604
#pragma warning disable CA1305
#pragma warning disable CA1812
#pragma warning disable IDE1006

using System.Reflection;
using System.Text;
using System.Text.Json;

namespace ExcelParser;

internal static class Constants
{
	internal static int firstDataRow;               // строка, с которой начинается информация (1 - названия столбцов)
	internal static int dateColumn;                 // столбец дат
	internal static int companiesNamesColumn;       // столбец названий компаний
	internal static int companiesAddressesColumn;   // столбец адресов компаний
	internal static int userNamesColumn;            // столбец ФИО пользователей
	internal static int userPositionsColumn;        // столбец должностей пользователей
	internal static int pcNumbersColumn;            // столбец номеров ПК
	internal static int defenderTypesColumn;        // столбец антивирусов
	internal static int powerSupplyColumn;          // столбец ИБП
	internal static int systemDriveColumn;          // столбец типов дисков, на который установлена ОС
	internal static int displayColumn;              // столбец проблем с мониторами
	internal static int mouseKeyboardColumn;        // столбец с проблемами с клавиатурой/мышкой
	internal static int internetStabilityColumn;    // столбец с проблемами с интернетом
	internal static int fileServerColumn;           // столбец, в котором ищем файловый сервер

	// присваиваем значения переменным, которые в любом случае поменяются, чтобы избежать предупреждений
	internal static string companyName = string.Empty;
	internal static string companyAddress = string.Empty;
	internal static DateTime firstDate = DateTime.Now;
	internal static DateTime secondDate = DateTime.Now;

	internal static void LoadConstantsFromJson ()
	{
		string jsonFilePath = "settings.json";

		// Проверяем существование файла настроек. Если отсутствует, то создаём новый на основе значений по умолчанию
		if (!File.Exists(jsonFilePath))
		{
			_ = MessageBox.Show($"Файл настроек {jsonFilePath} не найден");

			string defaultSettingsPath = "ExcelParser.resources.default_settings.json";

			using (StreamReader reader = new(Assembly.GetExecutingAssembly().GetManifestResourceStream(defaultSettingsPath)))
			{
				// Записываем в файл содержимое настроек по умолчанию из встроенного ресурса
				File.WriteAllText(jsonFilePath, reader.ReadToEnd());
			}

			_ = MessageBox.Show($"Файл {jsonFilePath} создан на основе настроек по умолчанию");
		}

		string jsonContent = File.ReadAllText(jsonFilePath);
		Settings? settings = null;
		try
		{
			settings = JsonSerializer.Deserialize<Settings>(jsonContent);
		}
		catch (JsonException ex)
		{
			_ = MessageBox.Show($"Ошибка десериализации {Path.GetFileName(jsonFilePath)}: {ex.Message}");
			Environment.Exit(0);
		}

		// проверяем корректность файла настроек
		if (settings == null || !CheckSettingsValid(settings))
		{
			_ = MessageBox.Show($"Ошибка при загрузке настроек. {Path.GetFileName(jsonFilePath)} задан некорректно");
			Environment.Exit(0);
		}

		// присваиваем значения из объекта settings соответствующим константам
		firstDataRow = settings.firstDataRow;
		dateColumn = settings.dateColumn;
		companiesNamesColumn = settings.companiesNamesColumn;
		companiesAddressesColumn = settings.companiesAddressesColumn;
		userNamesColumn = settings.userNamesColumn;
		userPositionsColumn = settings.userPositionsColumn;
		pcNumbersColumn = settings.pcNumbersColumn;
		defenderTypesColumn = settings.defenderTypesColumn;
		powerSupplyColumn = settings.powerSupplyColumn;
		systemDriveColumn = settings.systemDriveColumn;
		displayColumn = settings.displayColumn;
		mouseKeyboardColumn = settings.mouseKeyboardColumn;
		internetStabilityColumn = settings.internetStabilityColumn;
		fileServerColumn = settings.fileServerColumn;
	}

	// Метод для проверки всех настроек
	private static bool CheckSettingsValid (Settings settings)
	{
		bool isValid = true;
		StringBuilder errorMessage = new();

		isValid &= settings.firstDataRow.IsValid(nameof(settings.firstDataRow), errorMessage);
		isValid &= settings.dateColumn.IsValid(nameof(settings.dateColumn), errorMessage);
		isValid &= settings.companiesNamesColumn.IsValid(nameof(settings.companiesNamesColumn), errorMessage);
		isValid &= settings.companiesAddressesColumn.IsValid(nameof(settings.companiesAddressesColumn), errorMessage);
		isValid &= settings.userNamesColumn.IsValid(nameof(settings.userNamesColumn), errorMessage);
		isValid &= settings.userPositionsColumn.IsValid(nameof(settings.userPositionsColumn), errorMessage);
		isValid &= settings.pcNumbersColumn.IsValid(nameof(settings.pcNumbersColumn), errorMessage);
		isValid &= settings.defenderTypesColumn.IsValid(nameof(settings.defenderTypesColumn), errorMessage);
		isValid &= settings.powerSupplyColumn.IsValid(nameof(settings.powerSupplyColumn), errorMessage);
		isValid &= settings.systemDriveColumn.IsValid(nameof(settings.systemDriveColumn), errorMessage);
		isValid &= settings.displayColumn.IsValid(nameof(settings.displayColumn), errorMessage);
		isValid &= settings.mouseKeyboardColumn.IsValid(nameof(settings.mouseKeyboardColumn), errorMessage);
		isValid &= settings.internetStabilityColumn.IsValid(nameof(settings.internetStabilityColumn), errorMessage);
		isValid &= settings.fileServerColumn.IsValid(nameof(settings.fileServerColumn), errorMessage);

		if (!isValid)
		{
			_ = MessageBox.Show(errorMessage.ToString());
		}

		return isValid;
	}

	// Метод для проверки конкретного значения
	private static bool IsValid (this int value, string name, StringBuilder errorMessage)
	{
		if (value < 1)
		{
			_ = errorMessage.AppendLine($"Неверное значение для {name}: {value}");
			return false;
		}

		return true;
	}

	private sealed class Settings
	{
		public int firstDataRow
		{
			get; set;
		}
		public int dateColumn
		{
			get; set;
		}
		public int companiesNamesColumn
		{
			get; set;
		}
		public int companiesAddressesColumn
		{
			get; set;
		}
		public int userNamesColumn
		{
			get; set;
		}
		public int userPositionsColumn
		{
			get; set;
		}
		public int pcNumbersColumn
		{
			get; set;
		}
		public int defenderTypesColumn
		{
			get; set;
		}
		public int powerSupplyColumn
		{
			get; set;
		}
		public int systemDriveColumn
		{
			get; set;
		}
		public int displayColumn
		{
			get; set;
		}
		public int mouseKeyboardColumn
		{
			get; set;
		}
		public int internetStabilityColumn
		{
			get; set;
		}
		public int fileServerColumn
		{
			get; set;
		}
	}
}
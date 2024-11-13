#pragma warning disable CA1303
#pragma warning disable CA2000
#pragma warning disable CA2213
#pragma warning disable CS8601
#pragma warning disable CS8604
#pragma warning disable CS8618
#pragma warning disable CS8622
#pragma warning disable IDE0305

using System.Diagnostics;
using ExcelParser.utilities;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;

namespace ExcelParser;

internal sealed class Program : Form
{
	private ComboBox companyComboBox;
	private ComboBox addressComboBox;
	private DateTimePicker firstDatePicker;
	private DateTimePicker secondDatePicker;
	private Button generateReportButton;

	private readonly ExcelWorksheet worksheet = GetWorksheet(Directory.GetCurrentDirectory());

	private Program ()
	{
		// подгружаем конфиг столбцов с json файла конфига
		Constants.LoadConstantsFromJson();
		InitializeComponent();
		AutoSize = false;
		AutoSizeMode = AutoSizeMode.GrowAndShrink;
		FormBorderStyle = FormBorderStyle.Sizable;
		//Size = new Size(Size.Width * 2, Size.Height);
	}

	private void InitializeComponent ()
	{
		Text = "ExcelParser";

		TableLayoutPanel panel = new()
		{
			AutoSize = true,
			AutoSizeMode = AutoSizeMode.GrowAndShrink,
			Dock = DockStyle.Fill,
			ColumnCount = 1,
			Padding = new Padding(10)
		};

		companyComboBox = new ComboBox
		{
			Dock = DockStyle.Fill,
			DropDownStyle = ComboBoxStyle.DropDownList,
			IntegralHeight = false,
			MaxDropDownItems = 10
		};

		addressComboBox = new ComboBox
		{
			Dock = DockStyle.Fill,
			DropDownStyle = ComboBoxStyle.DropDownList,
			IntegralHeight = false,
			MaxDropDownItems = 10
		};

		firstDatePicker = new DateTimePicker
		{
			Dock = DockStyle.Fill,
			Value = DateTimePicker.MinimumDateTime,

		};

		secondDatePicker = new DateTimePicker
		{
			Dock = DockStyle.Fill,
			Value = DateTimePicker.MaximumDateTime
		};

		generateReportButton = new Button
		{
			Dock = DockStyle.Fill,
			Text = $"Составить отчет"
		};

		firstDatePicker.ValueChanged += FirstDatePicker_ValueChanged;
		secondDatePicker.ValueChanged += SecondDatePicker_ValueChanged;
		companyComboBox.SelectedIndexChanged += CompanyComboBox_SelectedIndexChanged;
		addressComboBox.SelectedIndexChanged += AddressComboBox_SelectedIndexChanged;
		generateReportButton.Click += GenerateReportButton_Click;
		companyComboBox.Items.AddRange(ConstantsUtils.GetCompanyNames(worksheet).ToArray());

		panel.Controls.Add(firstDatePicker);
		panel.Controls.Add(secondDatePicker);
		panel.Controls.Add(companyComboBox);
		panel.Controls.Add(addressComboBox);
		panel.Controls.Add(generateReportButton);
		Controls.Add(panel);
	}

	[STAThread]
	private static void Main ()
	{
		Application.EnableVisualStyles();
		Application.SetCompatibleTextRenderingDefault(false);
		Program program = new();
		Application.Run(program);
	}

	private void FirstDatePicker_ValueChanged (object sender, EventArgs e)
	{
		Debug.WriteLine(firstDatePicker.Value);
	}

	private void SecondDatePicker_ValueChanged (object sender, EventArgs e)
	{
		Debug.WriteLine(secondDatePicker.Value);
		if (secondDatePicker.Value < firstDatePicker.Value)
		{
			Debug.WriteLine($"{secondDatePicker.Value} = {firstDatePicker.Value}");
			_ = MessageBox.Show($"Дата окончания не может быть меньше даты начала");
			secondDatePicker.Value = firstDatePicker.Value;
		}
	}

	private void CompanyComboBox_SelectedIndexChanged (object sender, EventArgs e)
	{
		addressComboBox.Items.Clear();

		if (companyComboBox.SelectedItem != null)
		{
			Debug.WriteLine(companyComboBox.SelectedItem.ToString());
			List<string> companyAddresses = ConstantsUtils.GetCompanyAddresses(worksheet, companyComboBox.SelectedItem.ToString());

			if (companyAddresses.Count > 0)
			{
				addressComboBox.Items.AddRange(companyAddresses.ToArray());
				addressComboBox.SelectedIndex = 0;
			}
			else
			{
				Debug.WriteLine($"Адреса отсутствуют");
			}
		}
	}

	private void AddressComboBox_SelectedIndexChanged (object sender, EventArgs e)
	{
		if (addressComboBox.SelectedItem != null)
		{
			Debug.WriteLine(addressComboBox.SelectedItem.ToString());
		}
	}

	private void GenerateReportButton_Click (object sender, EventArgs e)
	{

		if (string.IsNullOrWhiteSpace(companyComboBox.SelectedItem?.ToString()) || string.IsNullOrWhiteSpace(addressComboBox.SelectedItem?.ToString()))
		{
			_ = MessageBox.Show($"Выберите компанию и адрес для составления отчета");
			return;
		}

		Constants.companyName = companyComboBox.SelectedItem.ToString();
		Constants.companyAddress = addressComboBox.SelectedItem.ToString();
		Constants.firstDate = firstDatePicker.Value;
		Constants.secondDate = secondDatePicker.Value;

		string reportsDirectoryPath = Path.Combine(Directory.GetCurrentDirectory(), "Reports");
		_ = Directory.CreateDirectory(reportsDirectoryPath);

		// путь к файлу .docx
		string docxFilePath = Path.Combine(reportsDirectoryPath, $"{ConstantsUtils.Sanitize(Constants.companyName)}_{ConstantsUtils.Sanitize(Constants.companyAddress)}.docx");

		if (ConstantsUtils.CheckOccupation(docxFilePath) && File.Exists(docxFilePath))
		{
			_ = MessageBox.Show($"Файл {Path.GetFileName(docxFilePath)} занят другой программой");
			return;
		}

		File.Delete(docxFilePath);
		Debug.WriteLine($"{Path.GetFileName(docxFilePath)} удалён");

		// создаем новый .docx документ
		FileStream fileStreamDocx = new(docxFilePath, FileMode.Create);
		XWPFDocument doc = new();

		DocumentUtils.FillDoc(doc, worksheet);

		// сохранение документа и освобождение ресурсов
		doc.Write(fileStreamDocx);
		doc.Dispose();
		fileStreamDocx.Dispose();
		_ = MessageBox.Show($"Документ создан успешно: {Path.GetFileName(docxFilePath)}");
	}

	// Функция для получения листа Excel
	private static ExcelWorksheet GetWorksheet (string currentDirectory)
	{
		ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

		// поиск файлов .xlsx
		string? xlsxPath = Directory.GetFiles(currentDirectory, "*.xlsx").FirstOrDefault();

		// проверяем наличие файлов .xlsx
		if (string.IsNullOrWhiteSpace(xlsxPath))
		{
			Exit($"В {currentDirectory} нет файлов .xlsx");
		}

		if (ConstantsUtils.CheckOccupation(xlsxPath))
		{
			Exit($"Файл {Path.GetFileName(xlsxPath)} занят другой программой");
		}

		ExcelPackage package = new(new FileInfo(xlsxPath));

		if (package.Workbook.Worksheets.Count == 0)
		{
			Exit($"Файл {Path.GetFileName(xlsxPath)} не содержит листов");
		}

		return package.Workbook.Worksheets [0];
	}

	private static void Exit (string message)
	{

		Debug.WriteLine(message);
		_ = MessageBox.Show($"{message}\nНажмите ОК для выхода");
		Environment.Exit(0);
	}
}

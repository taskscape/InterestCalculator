using Newtonsoft.Json;
using OfficeOpenXml;

namespace Calculator;

public class Config(double annualInterestRate, bool overwriteExistingFile)
{
    public double AnnualInterestRate { get; } = annualInterestRate;
    public string? InputFile { get; set; }
    public string? OutputFile { get; set; }
    public string? InterestRatesFile { get; set; }
    public bool OverwriteExistingFile { get; } = overwriteExistingFile;

    public void Print()
    {
        Console.WriteLine("configuration:");
        Console.WriteLine("- inputFile: " + InputFile);
        Console.WriteLine("- outputFile: " + OutputFile);
        Console.WriteLine("- OverwriteExistingFile: " + OverwriteExistingFile);
        Console.WriteLine("- InterestRatesFile: " + InterestRatesFile);
        Console.WriteLine();
    }
}

/// <summary>
/// Expects input file in Excel format containing a header describing columns:
/// description, date, amount
/// And a list of calculations below the header whereas each
/// date indicates a start date for calculating monthly interests
/// </summary>
internal class InputModel
{
    public string? Description { get; init; }
    public DateTime Date { get; init; }
    public double Amount { get; init; }
}

internal class OutputModel(string date)
{
    public string? Description { get; init; }
    public string Date { get; init; } = date;
    public double Amount { get; init; }
}

internal class InterestModel
{
    public DateTime Date { get; init; }
    public double Amount { get; init; }
}

abstract class Program
{
    private static Config? _config;
    
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        _config = ReadConfig("config.json");

        if (string.IsNullOrEmpty(_config?.InputFile))
        {
            Console.Write("Input file name: ");
            if (_config != null) _config.InputFile = Console.ReadLine();
        }
        if (string.IsNullOrEmpty(_config?.OutputFile))
        {
            Console.Write("Output file name: ");
            if (_config != null) _config.OutputFile = Console.ReadLine();
        }

        _config?.Print();

        if (_config?.InputFile == null) return;
        List<InputModel> data = ReadData(_config.InputFile);

        List<InterestModel>? interestRates = null;
        if (_config.InterestRatesFile != null)
            interestRates = ReadInterestRatesFromFile(_config.InterestRatesFile);

        List<OutputModel> results = CalculateInterest(data, _config.AnnualInterestRate / 100, interestRates);
        if (_config.OutputFile != null) SaveData(_config.OutputFile, results);
    }

    private static List<InterestModel>? ReadInterestRatesFromFile(string inputFile)
    {
        if (!File.Exists(inputFile))
        {
            Console.WriteLine(inputFile + "nie istnieje. wykorzystywana będzie wartość podana w konfiguracji 'config.json': " + _config.AnnualInterestRate);
            return null;
        }

        List<InterestModel> interestRates = new();
        string fileExtension = Path.GetExtension(inputFile).ToLower();

        if (fileExtension == ".csv")
        {
            // Wczytaj dane z pliku CSV
            using (var reader = new StreamReader(inputFile))
            {
                var headerLine = reader.ReadLine(); // Skip header line
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');

                    interestRates.Add(new InterestModel
                    {
                        Date = DateTime.Parse(values[0]),
                        Amount = double.Parse(values[1]) / 100
                    });
                }
            }
        }
        else if (fileExtension == ".xlsx" || fileExtension == ".xls")
        {
            // Wczytaj dane z pliku Excel
            using (var package = new ExcelPackage(new FileInfo(inputFile)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int row = 2; // Start from the second row (assuming the first row is the header)

                while (worksheet.Cells[row, 1].Value != null)
                {
                    interestRates.Add(new InterestModel
                    {
                        Date = DateTime.Parse(worksheet.Cells[row, 1].Text),
                        Amount = double.Parse(worksheet.Cells[row, 2].Text) / 100
                    });
                    row++;
                }
            }
        }
        else
        {
            throw new InvalidOperationException("Unsupported file type. Please use a .csv or .xlsx file.");
        }

        interestRates.Sort((a, b) => a.Date.CompareTo(b.Date));

        return interestRates;
    }

    private static Config? ReadConfig(string configPath)
    {
        using StreamReader reader = new(configPath);
        string json = reader.ReadToEnd();
        return JsonConvert.DeserializeObject<Config>(json);
    }

    private static List<InputModel> ReadData(string filePath)
    {
        List<InputModel> data = [];
        string fileExtension = Path.GetExtension(filePath).ToLower();

        switch (fileExtension)
        {
            case ".csv":
            {
                using StreamReader reader = new(filePath);
                string? headerLine = reader.ReadLine(); // Skip header line
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    string?[]? values = line?.Split(';');

                    InputModel inputData = new()
                    {
                        Description = values?[0],
                        Date = DateTime.Parse(values?[1]),
                        Amount = double.Parse(values?[2])
                    };
                    data.Add(inputData);
                }

                break;
            }
            case ".xlsx":
            case ".xls":
            {
                // Wczytaj dane z pliku Excel
                using ExcelPackage package = new(new FileInfo(filePath));
                ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];
                int row = 2; // Start from the second row (assuming the first row is the header)

                while (worksheet.Cells[row, 1].Value != null)
                {
                    InputModel inputData = new()
                    {
                        Description = worksheet.Cells[row, 1].Text,
                        Date = DateTime.Parse(worksheet.Cells[row, 2].Text),
                        Amount = double.Parse(worksheet.Cells[row, 3].Text)
                    };
                    data.Add(inputData);
                    row++;
                }

                break;
            }
            default:
                throw new InvalidOperationException("Unsupported file type. Please use a .csv or .xlsx file.");
        }

        return data;
    }

    static List<OutputModel> CalculateInterest(List<InputModel> data, double annualInterestRate, List<InterestModel>? InterestRates = null)
    {
        List<OutputModel> results = [];
        double dailyInterestRate = annualInterestRate / 365;

        foreach (InputModel entry in data)
        {
            double amount = entry.Amount;
            DateTime currentDate = entry.Date;
            DateTime presentDate = DateTime.Now;

            while (currentDate.AddMonths(1) < presentDate)
            {
                DateTime nextDate = currentDate.AddMonths(1);
                double numberOfDays = (nextDate - currentDate).TotalDays;

                dailyInterestRate = FindDailyInterestRate(InterestRates, currentDate, dailyInterestRate);
                double accruedInterest = amount * dailyInterestRate * numberOfDays;

                amount += accruedInterest;

                results.Add(new OutputModel(nextDate.ToString("yyyy-MM-dd"))
                {
                    Description = entry.Description,
                    Amount = Math.Round(amount, 2)
                });

                currentDate = nextDate;
            }
        }

        return results;
    }

    private static double FindDailyInterestRate(List<InterestModel>? interests, DateTime currentDate, double defaultValue = 0)
    {
        if(interests == null)
            return defaultValue;

        double ret = defaultValue;
        foreach (var interest in interests)
        {
            if (interest.Date <= currentDate)
            {
                ret = interest.Amount;
            }
            else
            {
                break;
            }
        }
        return ret / 365;
    }

    static void SaveData(string filePath, List<OutputModel> results)
    {
        if (File.Exists(filePath) && _config is { OverwriteExistingFile: false })
            filePath = filePath.Insert(filePath.LastIndexOf('.'), DateTime.Now.ToString("yyyy-MM-dd HH-MM-ss"));
        
        string fileExtension = Path.GetExtension(filePath).ToLower();

        switch (fileExtension)
        {
            case ".csv":
            {
                using StreamWriter writer = new(filePath);
                writer.WriteLine("Opis,Data,Kwota");
                foreach (OutputModel result in results)
                {
                    writer.WriteLine($"{result.Description},{result.Date},{result.Amount}");
                }

                break;
            }
            case ".xlsx":
            case ".xls":
            {
                using ExcelPackage package = new();
                ExcelWorksheet? worksheet = package.Workbook.Worksheets.Add("Result");
                worksheet.Cells[1, 1].Value = "Opis";
                worksheet.Cells[1, 2].Value = "Data";
                worksheet.Cells[1, 3].Value = "Kwota";

                int row = 2;
                foreach (OutputModel result in results)
                {
                    worksheet.Cells[row, 1].Value = result.Description;
                    worksheet.Cells[row, 2].Value = result.Date;
                    worksheet.Cells[row, 3].Value = result.Amount;
                    row++;
                }

                package.SaveAs(new FileInfo(filePath));

                break;
            }
            default:
                throw new InvalidOperationException("Unsupported file type. Please use a .csv or .xlsx file.");
        }

        Console.WriteLine("Results saved in file: " + filePath);
    }
}

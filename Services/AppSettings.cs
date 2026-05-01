using System.Text.Json;

namespace BISTMatriks.Services;

public class AppSettings
{
    public string WorkbookPath { get; set; } = "";

    // CSV yolları exe yanında sabit — kullanıcıya sorulmaz
    public string TarihselCsv =>
        Path.Combine(AppContext.BaseDirectory, "bist_tarihsel.csv");
    public string CanliCsv =>
        Path.Combine(AppContext.BaseDirectory, "bist_canli.csv");

    public bool IsConfigured => !string.IsNullOrWhiteSpace(WorkbookPath);

    private static readonly string FilePath =
        Path.Combine(AppContext.BaseDirectory, "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var s = JsonSerializer.Deserialize<AppSettings>(
                    File.ReadAllText(FilePath));
                if (s != null) return s;
            }
        }
        catch { }
        return new AppSettings();
    }

    public void Save()
    {
        File.WriteAllText(FilePath,
            JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
    }
}

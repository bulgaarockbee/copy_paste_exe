using System;
using System.IO;
using System.Text.Json;

public class AppSettingsModel
{
    public string WindowTitle { get; set; } = "";
    public string WindowClass { get; set; } = "";
}

public static class AppSettings
{
    private static string SettingsFilePath => Path.Combine(AppContext.BaseDirectory, "appsettings.json");
    private static AppSettingsModel cached;

    public static AppSettingsModel Get()
    {
        if (cached != null) return cached;
        if (!File.Exists(SettingsFilePath)) return cached = new AppSettingsModel();
        cached = JsonSerializer.Deserialize<AppSettingsModel>(File.ReadAllText(SettingsFilePath)) ?? new AppSettingsModel();
        return cached;
    }

    public static void Save()
    {
        var dir = Path.GetDirectoryName(SettingsFilePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(SettingsFilePath, JsonSerializer.Serialize(Get(), new JsonSerializerOptions { WriteIndented = true }));
    }
}

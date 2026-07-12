namespace CheckIn.Client.Mobile.Models;

/// <summary>
/// Application global configuration, persisted to config.json.
/// Contains server connection, admin password, and general settings.
/// </summary>
public class AppConfig
{
    public string School { get; set; } = "";
    public string Nj { get; set; } = "";
    public string ClassId { get; set; } = "";
    public string Km { get; set; } = "";
    public int ButtonRows { get; set; } = 6;
    public int ButtonCols { get; set; } = 6;
    public bool OnlineMode { get; set; } = true;
    public string ServerIp { get; set; } = "";
    public int ServerPort { get; set; } = 5250;
    public string ServerPassword { get; set; } = "";
    public string AdminPasswordHash { get; set; } = "";
    public const string Version = "v2.7.1";
}

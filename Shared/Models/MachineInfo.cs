namespace CheckIn.Shared.Models;

public class MachineInfo
{
    public string Uuid { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string PublicKey { get; set; } = string.Empty;
    public string? LastSeen { get; set; }
    public bool Online { get; set; }
}

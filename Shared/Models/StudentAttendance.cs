namespace CheckIn.Shared.Models;

public class StudentAttendance
{
    public string Name { get; set; } = string.Empty;
    public int Count { get; set; }
    public string? FirstTime { get; set; }
    public List<string> History { get; set; } = new();
}

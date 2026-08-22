using System.Text.Json.Serialization;

namespace CheckIn.Client.Models;

// ============ 认证 ============

public class LoginResponse
{
    [JsonPropertyName("token")] public string Token { get; set; } = "";
    [JsonPropertyName("user")] public RemoteUser? User { get; set; }
    [JsonPropertyName("error")] public string? Error { get; set; }
}

public class RemoteUser
{
    [JsonPropertyName("id")] public int Id { get; set; }
    [JsonPropertyName("username")] public string Username { get; set; } = "";
    [JsonPropertyName("role")] public string Role { get; set; } = "";
    [JsonPropertyName("display_name")] public string DisplayName { get; set; } = "";
}

// ============ 仪表盘 ============

public class DashboardResponse
{
    [JsonPropertyName("summary")] public DashboardSummary? Summary { get; set; }
    [JsonPropertyName("devices")] public List<DashboardDevice> Devices { get; set; } = new();
    [JsonPropertyName("active_signin_tasks")] public List<DashboardTask> ActiveSignInTasks { get; set; } = new();
}

public class DashboardSummary
{
    [JsonPropertyName("total_devices")] public int TotalDevices { get; set; }
    [JsonPropertyName("online_devices")] public int OnlineDevices { get; set; }
    [JsonPropertyName("total_users")] public int TotalUsers { get; set; }
    [JsonPropertyName("today_checkins")] public int TodayCheckins { get; set; }
    [JsonPropertyName("active_signin_tasks")] public int ActiveSignInTasks { get; set; }
}

public class DashboardDevice
{
    [JsonPropertyName("uuid")] public string Uuid { get; set; } = "";
    [JsonPropertyName("name")] public string Name { get; set; } = "";
    [JsonPropertyName("task_count")] public int TaskCount { get; set; }
    [JsonPropertyName("last")] public string? Last { get; set; }
    [JsonPropertyName("online")] public bool Online { get; set; }
}

public class DashboardTask
{
    [JsonPropertyName("short_code")] public string ShortCode { get; set; } = "";
    [JsonPropertyName("subject")] public string Subject { get; set; } = "";
    [JsonPropertyName("classroom")] public string Classroom { get; set; } = "";
    [JsonPropertyName("student_count")] public int StudentCount { get; set; }
    [JsonPropertyName("signed_count")] public int SignedCount { get; set; }
    [JsonPropertyName("created_at")] public string CreatedAt { get; set; } = "";
}

// ============ 设备 ============

public class DeviceResponse
{
    [JsonPropertyName("devices")] public List<DeviceItem> Devices { get; set; } = new();
}

public class DeviceItem
{
    [JsonPropertyName("uuid")] public string Uuid { get; set; } = "";
    [JsonPropertyName("name")] public string Name { get; set; } = "";
    [JsonPropertyName("online")] public bool Online { get; set; }
    [JsonPropertyName("last_seen")] public string? LastSeen { get; set; }
    [JsonPropertyName("public_key")] public string PublicKey { get; set; } = "";
}

// ============ 任务 ============

public class TaskListResponse
{
    [JsonPropertyName("tasks")] public List<RemoteTask> Tasks { get; set; } = new();
}

public class RemoteTask
{
    [JsonPropertyName("id")] public int Id { get; set; }
    [JsonPropertyName("short_code")] public string ShortCode { get; set; } = "";
    [JsonPropertyName("subject")] public string Subject { get; set; } = "";
    [JsonPropertyName("classroom")] public string Classroom { get; set; } = "";
    [JsonPropertyName("status")] public string Status { get; set; } = "";
    [JsonPropertyName("student_count")] public int StudentCount { get; set; }
    [JsonPropertyName("signed_count")] public int SignedCount { get; set; }
    [JsonPropertyName("created_at")] public string CreatedAt { get; set; } = "";
    [JsonPropertyName("machine_uuid")] public string MachineUuid { get; set; } = "";
}

// ============ 考勤 ============

public class AttendanceResponse
{
    [JsonPropertyName("tasks")] public List<AttendanceTask> Tasks { get; set; } = new();
}

public class AttendanceTask
{
    [JsonPropertyName("machine_uuid")] public string MachineUuid { get; set; } = "";
    [JsonPropertyName("machine_name")] public string MachineName { get; set; } = "";
    [JsonPropertyName("task_id")] public string TaskId { get; set; } = "";
    [JsonPropertyName("total_students")] public int TotalStudents { get; set; }
    [JsonPropertyName("punched_count")] public int PunchedCount { get; set; }
    [JsonPropertyName("unpunched_count")] public int UnpunchedCount { get; set; }
    [JsonPropertyName("attendance_rate")] public double AttendanceRate { get; set; }
    [JsonPropertyName("last_updated")] public string LastUpdated { get; set; } = "";
    [JsonPropertyName("students")] public List<AttendanceStudent> Students { get; set; } = new();
}

public class AttendanceStudent
{
    [JsonPropertyName("name")] public string Name { get; set; } = "";
    [JsonPropertyName("checked_in")] public bool CheckedIn { get; set; }
    [JsonPropertyName("first_time")] public string? FirstTime { get; set; }
    [JsonPropertyName("count")] public int Count { get; set; }
}

// ============ 签到历史 ============

public class HistoryResponse
{
    [JsonPropertyName("role")] public string Role { get; set; } = "";
    [JsonPropertyName("total_tasks")] public int TotalTasks { get; set; }
    [JsonPropertyName("total_checkins")] public int TotalCheckins { get; set; }
    [JsonPropertyName("history")] public List<HistoryTask> History { get; set; } = new();
}

public class HistoryTask
{
    [JsonPropertyName("subject")] public string Subject { get; set; } = "";
    [JsonPropertyName("classroom")] public string Classroom { get; set; } = "";
    [JsonPropertyName("short_code")] public string ShortCode { get; set; } = "";
    [JsonPropertyName("task_name")] public string TaskName { get; set; } = "";
    [JsonPropertyName("student_count")] public int StudentCount { get; set; }
    [JsonPropertyName("signed_count")] public int SignedCount { get; set; }
    [JsonPropertyName("status")] public string Status { get; set; } = "";
    [JsonPropertyName("created_at")] public string CreatedAt { get; set; } = "";
    [JsonPropertyName("records")] public List<HistoryRecord> Records { get; set; } = new();
}

public class HistoryRecord
{
    [JsonPropertyName("name")] public string Name { get; set; } = "";
    [JsonPropertyName("time")] public string Time { get; set; } = "";
}

// ============ 用户 ============

public class RemoteUserItem
{
    [JsonPropertyName("id")] public int Id { get; set; }
    [JsonPropertyName("username")] public string Username { get; set; } = "";
    [JsonPropertyName("role")] public string Role { get; set; } = "";
    [JsonPropertyName("display_name")] public string DisplayName { get; set; } = "";
    [JsonPropertyName("created_at")] public string CreatedAt { get; set; } = "";
    [JsonPropertyName("is_active")] public bool IsActive { get; set; }
}

// ============ 通用响应 ============

public class StatusResponse
{
    [JsonPropertyName("status")] public string Status { get; set; } = "";
    [JsonPropertyName("message")] public string? Message { get; set; }
    [JsonPropertyName("error")] public string? Error { get; set; }
}

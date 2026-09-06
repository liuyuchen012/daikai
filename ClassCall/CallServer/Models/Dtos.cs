namespace CallServer.Models;

// ---------- 设备相关 ----------

public record RegisterRequest(string Name, string Room);

public record RegisterResponse(string Uuid, string Password);

public record HeartbeatRequest(string Uuid, string Password);

public record DeviceDto(int Id, string Uuid, string Name, string Room, bool IsOnline, DateTime LastHeartbeat);

// ---------- 呼叫相关 ----------

public record SendCallRequest(string? TargetUuid, string Type, string Title, string Message, string Sender);

public record PullRequest(string Uuid, string Password);

public record AckRequest(int Id, string Uuid, string Password);

public record CallDto(int Id, string Type, string Title, string Message, string Sender, DateTime CreatedAt);

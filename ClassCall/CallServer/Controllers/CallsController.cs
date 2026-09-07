using CallServer.Data;
using CallServer.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace CallServer.Controllers;

/// <summary>
/// 呼叫中转：呼出端发送 → 被控端拉取 → 被控端确认
/// </summary>
[ApiController]
[Route("api/calls")]
public class CallsController(AppDbContext db) : ControllerBase
{
    /// <summary>
    /// 呼出端发送呼叫。
    /// targetUuid 为空时广播给全部设备；type: urgent | notice | speech
    /// </summary>
    [HttpPost("send")]
    public async Task<IActionResult> Send([FromBody] SendCallRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Title) && string.IsNullOrWhiteSpace(req.Message))
            return BadRequest(new { error = "呼叫标题和内容不能同时为空" });

        var call = new CallRecord
        {
            Type = req.Type switch
            {
                "urgent" => "urgent",
                "speech" => "speech",
                _ => "notice"
            },
            Title = (req.Title ?? "").Trim(),
            Message = (req.Message ?? "").Trim(),
            Sender = (req.Sender ?? "").Trim(),
            TargetUuid = string.IsNullOrWhiteSpace(req.TargetUuid) ? null : req.TargetUuid,
            CreatedAt = DateTime.Now
        };
        db.CallRecords.Add(call);
        await db.SaveChangesAsync();
        return Ok(new { id = call.Id });
    }

    /// <summary>
    /// 被控端轮询拉取待处理的呼叫：
    /// 广播呼叫（TargetUuid 为空）或发给本设备的呼叫，且尚未确认
    /// </summary>
    [HttpPost("pull")]
    public async Task<IActionResult> Pull([FromBody] PullRequest req)
    {
        var device = await db.Devices.FirstOrDefaultAsync(d => d.Uuid == req.Uuid && d.Password == req.Password);
        if (device == null) return NotFound(new { error = "设备不存在或凭证错误" });

        device.LastHeartbeat = DateTime.Now;

        var calls = await db.CallRecords.AsNoTracking()
            .Where(c => c.AckedAt == null && (c.TargetUuid == null || c.TargetUuid == req.Uuid))
            .OrderBy(c => c.Id)
            .Select(c => new CallDto(c.Id, c.Type, c.Title, c.Message, c.Sender, c.CreatedAt))
            .ToListAsync();

        await db.SaveChangesAsync();
        return Ok(new { calls });
    }

    /// <summary>被控端确认已收到并展示/朗读呼叫，防止重复拉取</summary>
    [HttpPost("ack")]
    public async Task<IActionResult> Ack([FromBody] AckRequest req)
    {
        var device = await db.Devices.FirstOrDefaultAsync(d => d.Uuid == req.Uuid && d.Password == req.Password);
        if (device == null) return NotFound(new { error = "设备不存在或凭证错误" });

        var call = await db.CallRecords.FirstOrDefaultAsync(c => c.Id == req.Id);
        if (call == null) return NotFound(new { error = "呼叫不存在" });

        call.AckedAt = DateTime.Now;
        call.AckedByUuid = req.Uuid;
        await db.SaveChangesAsync();
        return Ok();
    }
}

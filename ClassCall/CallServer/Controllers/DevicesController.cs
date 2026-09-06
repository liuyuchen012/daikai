using CallServer.Data;
using CallServer.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace CallServer.Controllers;

/// <summary>
/// 设备管理：注册（换发 UUID/密码）、心跳（在线状态）、列表（呼出端查询）
/// </summary>
[ApiController]
[Route("api/devices")]
public class DevicesController(AppDbContext db) : ControllerBase
{
    /// <summary>
    /// 插件端注册设备。
    /// 同名设备复用已有凭证（便于教室一体机重装后保持同一身份）；
    /// 新设备生成 UUID + 8 位密码。
    /// </summary>
    [HttpPost("register")]
    public async Task<IActionResult> Register([FromBody] RegisterRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Name))
            return BadRequest(new { error = "设备名称不能为空" });

        var existing = await db.Devices.FirstOrDefaultAsync(d => d.Name == req.Name);
        if (existing != null)
        {
            existing.Room = req.Room ?? "";
            existing.LastHeartbeat = DateTime.Now;
            await db.SaveChangesAsync();
            return Ok(new RegisterResponse(existing.Uuid, existing.Password));
        }

        var device = new Device
        {
            Name = req.Name.Trim(),
            Room = req.Room ?? "",
            Uuid = Guid.NewGuid().ToString("N"),
            Password = Guid.NewGuid().ToString("N")[..8],
            LastHeartbeat = DateTime.Now
        };
        db.Devices.Add(device);
        await db.SaveChangesAsync();
        return Ok(new RegisterResponse(device.Uuid, device.Password));
    }

    /// <summary>插件端心跳，更新在线状态</summary>
    [HttpPost("heartbeat")]
    public async Task<IActionResult> Heartbeat([FromBody] HeartbeatRequest req)
    {
        var device = await db.Devices.FirstOrDefaultAsync(d => d.Uuid == req.Uuid && d.Password == req.Password);
        if (device == null) return NotFound(new { error = "设备不存在或凭证错误" });
        device.LastHeartbeat = DateTime.Now;
        await db.SaveChangesAsync();
        return Ok();
    }

    /// <summary>设备列表（呼出端显示在线状态）</summary>
    [HttpGet]
    public async Task<IActionResult> List()
    {
        // 先取回数据再在内存中计算在线状态（SQLite 不支持 TimeSpan 求值）
        var devices = await db.Devices.AsNoTracking()
            .OrderByDescending(d => d.LastHeartbeat)
            .ToListAsync();

        var dtos = devices
            .Select(d => new DeviceDto(d.Id, d.Uuid, d.Name, d.Room, d.IsOnline, d.LastHeartbeat))
            .ToList();
        return Ok(new { devices = dtos });
    }
}

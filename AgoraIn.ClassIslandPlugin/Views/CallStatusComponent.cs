using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using AgoraIn.ClassIslandPlugin.Models;
using ClassIsland.Core.Abstractions.Controls;
using ClassIsland.Core.Attributes;

namespace AgoraIn.ClassIslandPlugin.Views;

/// <summary>
/// AgoraIn 呼叫状态主界面组件：显示在 ClassIsland 主界面上，
/// 展示插件连接状态与最近一次呼叫（类型 / 标题 / 时间），
/// 可在主界面「编辑」模式下拖拽摆放。
/// </summary>
[ComponentInfo("D7F2B8A4-1E3C-4B6D-9A0E-7C5D2F8E4A1B", "AgoraIn 呼叫状态", "lucide(\ue138)", "显示集控平台呼叫插件的连接状态与最近一次呼叫")]
public class CallStatusComponent : ComponentBase
{
    private readonly TextBlock _statusText;
    private readonly TextBlock _callText;

    public CallStatusComponent()
    {
        var panel = new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing = 4,
            Margin = new Avalonia.Thickness(8),
        };

        _statusText = new TextBlock
        {
            FontSize = 13,
            FontWeight = FontWeight.SemiBold,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };
        _callText = new TextBlock
        {
            FontSize = 12,
            Opacity = 0.7,
            TextWrapping = TextWrapping.Wrap,
            TextTrimming = TextTrimming.CharacterEllipsis,
            MaxHeight = 40,
        };

        panel.Children.Add(_statusText);
        panel.Children.Add(_callText);

        Content = panel;

        // 订阅状态变化，刷新显示
        CallStatusStore.Changed += Refresh;

        // 初始渲染
        Refresh();
    }

    private void Refresh()
    {
        _statusText.Text = CallStatusStore.StatusText;

        var call = CallStatusStore.LastCall;
        if (call == null)
        {
            _callText.Text = "暂无呼叫记录";
        }
        else
        {
            _callText.Text = $"{LabelFor(call.Type)}：{call.Title}";
        }
    }

    private static string LabelFor(string type) => type switch
    {
        "emergency" => "应急",
        "summon"    => "传唤",
        _           => "待下课",
    };
}

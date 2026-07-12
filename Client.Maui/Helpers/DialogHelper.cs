namespace CheckIn.Client.Maui.ViewModels;

/// <summary>
/// 对话框辅助工具类，封装 MAUI DisplayAlert 和 DisplayPromptAsync 方法
/// </summary>
public static class DialogHelper
{
    /// <summary>
    /// 显示确认对话框（含确定/取消按钮），返回用户选择
    /// </summary>
    public static async Task<bool> ConfirmAsync(string message, string title = "确认")
    {
        var page = Application.Current?.Windows.FirstOrDefault()?.Page;
        if (page == null) return false;
        return await page.DisplayAlertAsync(title, message, "确定", "取消");
    }

    /// <summary>
    /// 显示提示对话框（仅含确定按钮）
    /// </summary>
    public static async Task AlertAsync(string message, string title = "提示")
    {
        var page = Application.Current?.Windows.FirstOrDefault()?.Page;
        if (page == null) return;
        await page.DisplayAlertAsync(title, message, "确定");
    }

    /// <summary>
    /// 显示输入对话框，返回用户输入的文本，取消返回 null
    /// </summary>
    public static async Task<string?> PromptAsync(string title, string message, string initialValue = "", bool isPassword = false)
    {
        var page = Application.Current?.Windows.FirstOrDefault()?.Page;
        if (page == null) return null;

        // MAUI DisplayPromptAsync is not available on all platforms, so we use a workaround
        // On Windows/macOS, it's available through the standard API
        try
        {
            return await page.DisplayPromptAsync(title, message, "确定", "取消", initialValue, keyboard: isPassword ? Keyboard.Default : Keyboard.Default);
        }
        catch
        {
            // Fallback: just return initial value
            return initialValue;
        }
    }
}

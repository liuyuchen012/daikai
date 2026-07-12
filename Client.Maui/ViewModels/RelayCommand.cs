using System.Windows.Input;

namespace CheckIn.Client.Maui.ViewModels;

/// <summary>
/// 通用的 ICommand 实现，用于将 ViewModel 方法绑定到 UI 按钮
/// 支持带参数和无参数两种重载形式
/// </summary>
public class RelayCommand : ICommand
{
    private readonly Action<object?> _execute;
    private readonly Func<object?, bool>? _canExecute;

    /// <summary>
    /// 创建带参数的命令
    /// </summary>
    public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
    {
        _execute = execute;
        _canExecute = canExecute;
    }

    /// <summary>
    /// 创建无参数的命令的便捷重载
    /// </summary>
    public RelayCommand(Action execute, Func<bool>? canExecute = null)
        : this(_ => execute(), canExecute != null ? _ => canExecute() : null) { }

    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;

    public void Execute(object? parameter) => _execute(parameter);

    /// <summary>
    /// MAUI 下通过手动触发方式通知命令状态变化
    /// </summary>
    public event EventHandler? CanExecuteChanged;

    /// <summary>
    /// 手动触发 CanExecuteChanged 事件（MAUI 不支持 CommandManager）
    /// </summary>
    public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}

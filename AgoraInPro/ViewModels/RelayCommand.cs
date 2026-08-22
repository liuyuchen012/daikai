using System.Windows.Input;

namespace CheckIn.Client.ViewModels;

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
    /// <param name="execute">执行方法，接收 object? 参数</param>
    /// <param name="canExecute">可执行判断方法，为 null 时始终可执行</param>
    public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
    {
        _execute = execute;
        _canExecute = canExecute;
    }

    /// <summary>
    /// 创建无参数的命令的便捷重载
    /// </summary>
    /// <param name="execute">执行方法，无参数</param>
    /// <param name="canExecute">可执行判断方法，无参数</param>
    public RelayCommand(Action execute, Func<bool>? canExecute = null)
        : this(_ => execute(), canExecute != null ? _ => canExecute() : null) { }

    /// <summary>
    /// 判断命令是否可执行
    /// </summary>
    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;

    /// <summary>
    /// 执行命令
    /// </summary>
    public void Execute(object? parameter) => _execute(parameter);

    /// <summary>
    /// 命令的可执行状态变更事件，委托给 CommandManager.RequerySuggested 自动刷新
    /// </summary>
    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }
}

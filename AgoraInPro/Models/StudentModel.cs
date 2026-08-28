using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CheckIn.Client.Models;

/// <summary>
/// 学生数据模型，实现 INotifyPropertyChanged 以支持 UI 数据绑定
/// 包含学生姓名、打卡次数、首次打卡时间和历史记录
/// </summary>
public class StudentModel : INotifyPropertyChanged
{
    private string _name = "";
    private int _count;
    private DateTime? _firstTime;
    private bool _isCheckedIn;
    private List<string> _history = new();

    /// <summary>学生姓名</summary>
    public string Name { get => _name; set { _name = value; OnPropertyChanged(); } }
    /// <summary>打卡累计次数</summary>
    public int Count { get => _count; set { _count = value; OnPropertyChanged(); } }
    /// <summary>首次打卡时间（强类型 DateTime），为 null 表示尚未打卡</summary>
    public DateTime? FirstTime { get => _firstTime; set { _firstTime = value; IsCheckedIn = value != null; OnPropertyChanged(); } }
    /// <summary>是否已打卡（依赖 FirstTime 自动更新）</summary>
    public bool IsCheckedIn { get => _isCheckedIn; set { _isCheckedIn = value; OnPropertyChanged(); } }
    /// <summary>打卡历史记录列表，每次打卡的时间戳</summary>
    public List<string> History { get => _history; set { _history = value; OnPropertyChanged(); } }

    /// <summary>属性变更事件，通知 UI 绑定刷新</summary>
    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

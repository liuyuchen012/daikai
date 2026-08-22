using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using CheckIn.Client.Models;
using CheckIn.Client.Services;

namespace CheckIn.Client.ViewModels;

/// <summary>
/// 课时划消 ViewModel：学生管理、课时划消记录、排课管理（每日复制、不排课日）
/// </summary>
public class ClassHoursViewModel : INotifyPropertyChanged
{
    private readonly ClassHourStore _store;
    private ClassHourData _data;

    public ObservableCollection<ChStudent> Students { get; } = new();
    public ObservableCollection<ChRecord> Records { get; } = new();

    private ChStudent? _selectedStudent;
    public ChStudent? SelectedStudent
    {
        get => _selectedStudent;
        set
        {
            _selectedStudent = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(SelectedDetail));
            OnPropertyChanged(nameof(CanOperate));
            RefreshRecords();
        }
    }

    /// <summary>选中学生的详情文本（用于状态栏展示）</summary>
    public string SelectedDetail => SelectedStudent == null
        ? "未选择学生"
        : $"{SelectedStudent.Name}：总 {SelectedStudent.TotalHours:0.#} 课时，已划 {SelectedStudent.UsedHours:0.#}，剩余 {SelectedStudent.RemainingHours:0.#}";

    public bool CanOperate => SelectedStudent != null;

    // ---- 排课状态 ----
    private DateTime _currentMonth = new(DateTime.Now.Year, DateTime.Now.Month, 1);
    public DateTime CurrentMonth
    {
        get => _currentMonth;
        set { _currentMonth = value; OnPropertyChanged(); }
    }

    private DateTime _selectedDate = DateTime.Today;
    public DateTime SelectedDate
    {
        get => _selectedDate;
        set { _selectedDate = value; OnPropertyChanged(); OnPropertyChanged(nameof(SelectedDateText)); OnPropertyChanged(nameof(IsOffDayForSelected)); RefreshScheduleDetail(); }
    }

    public string SelectedDateText => SelectedDate.ToString("yyyy年M月d日 dddd");

    private bool _isOffDayForSelected;
    public bool IsOffDayForSelected
    {
        get => _isOffDayForSelected;
        set { _isOffDayForSelected = value; OnPropertyChanged(); }
    }

    /// <summary>月历 42 格（6周×7天）</summary>
    public ObservableCollection<CalendarDayItem> CalendarDays { get; } = new();

    /// <summary>当日排课学生</summary>
    public ObservableCollection<ChStudent> ScheduleStudents { get; } = new();

    // ---- 学生输入 ----
    private string _newStudentName = "";
    public string NewStudentName
    {
        get => _newStudentName;
        set { _newStudentName = value; OnPropertyChanged(); }
    }

    private double _newStudentHours = 0;
    public double NewStudentHours
    {
        get => _newStudentHours;
        set { _newStudentHours = value; OnPropertyChanged(); }
    }

    // ---- 划课输入 ----
    private double _operateHours = 1;
    public double OperateHours
    {
        get => _operateHours;
        set { _operateHours = value; OnPropertyChanged(); }
    }

    private string _operateNote = "";
    public string OperateNote
    {
        get => _operateNote;
        set { _operateNote = value; OnPropertyChanged(); }
    }

    /// <summary>状态消息</summary>
    private string _statusMessage = "就绪";
    public string StatusMessage
    {
        get => _statusMessage;
        set { _statusMessage = value; OnPropertyChanged(); }
    }

    public ClassHoursViewModel()
    {
        _store = new ClassHourStore();
        _data = _store.Load();
        ReloadStudents();
        BuildCalendar();
    }

    // ---- 学生管理 ----

    private void ReloadStudents()
    {
        Students.Clear();
        foreach (var s in _data.Students)
            Students.Add(s);
        OnPropertyChanged(nameof(Students));
    }

    public void AddStudent()
    {
        var name = NewStudentName.Trim();
        if (string.IsNullOrEmpty(name))
        {
            StatusMessage = "请输入学生姓名";
            return;
        }
        if (_data.Students.Any(s => s.Name == name))
        {
            StatusMessage = $"学生「{name}」已存在";
            return;
        }
        // L5：初始课时数限制最多 1 位小数，避免 0.001 这类无意义精度
        if (Math.Round(NewStudentHours, 1) != NewStudentHours)
        {
            StatusMessage = "初始课时数最多支持 1 位小数（如 1.5）";
            return;
        }
        var student = new ChStudent
        {
            Name = name,
            TotalHours = Math.Max(0, NewStudentHours),
            CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };
        _data.Students.Add(student);
        _store.Save(_data);
        ReloadStudents();
        SelectedStudent = student;
        NewStudentName = "";
        NewStudentHours = 0;
        StatusMessage = $"已添加学生「{name}」";
    }

    public void DeleteStudent(ChStudent student)
    {
        if (student == null) return;
        _data.Students.Remove(student);
        _data.Records.RemoveAll(r => r.StudentId == student.Id);
        // 清理排课中的该学生
        foreach (var kv in _data.Schedule)
            kv.Value.RemoveAll(id => id == student.Id);
        _store.Save(_data);
        ReloadStudents();
        SelectedStudent = null;
        RefreshScheduleDetail();
        StatusMessage = $"已删除学生「{student.Name}」";
    }

    // ---- 课时划消 ----

    private void RefreshRecords()
    {
        Records.Clear();
        if (SelectedStudent == null) return;
        foreach (var r in _data.Records.Where(r => r.StudentId == SelectedStudent.Id)
                     .OrderByDescending(r => r.Date).ThenByDescending(r => r.CreatedAt))
            Records.Add(r);
    }

    /// <summary>划消课时（扣减）</summary>
    public void DeductHours()
    {
        if (SelectedStudent == null) return;
        var hours = OperateHours;
        if (hours <= 0) { StatusMessage = "请输入大于 0 的课时数"; return; }
        // L5：限制精度，最多 1 位小数（如 1.5），拒绝 0.001 这类输入
        if (Math.Round(hours, 1) != hours) { StatusMessage = "课时数最多支持 1 位小数（如 1.5）"; return; }
        if (hours > SelectedStudent.RemainingHours)
        {
            StatusMessage = $"剩余课时不足（剩余 {SelectedStudent.RemainingHours:0.#}）";
            return;
        }
        ApplyOperate(-hours, "划消");
    }

    /// <summary>增加课时</summary>
    public void AddHours()
    {
        if (SelectedStudent == null) return;
        var hours = OperateHours;
        if (hours <= 0) { StatusMessage = "请输入大于 0 的课时数"; return; }
        // L5：限制精度，最多 1 位小数（如 1.5），拒绝 0.001 这类输入
        if (Math.Round(hours, 1) != hours) { StatusMessage = "课时数最多支持 1 位小数（如 1.5）"; return; }
        ApplyOperate(hours, "增加");
    }

    private void ApplyOperate(double delta, string typeName)
    {
        var student = SelectedStudent!;
        student.UsedHours += Math.Max(0, -delta);
        student.TotalHours += Math.Max(0, delta);
        _data.Records.Add(new ChRecord
        {
            StudentId = student.Id,
            Date = DateTime.Now.ToString("yyyy-MM-dd"),
            Hours = delta,
            Note = string.IsNullOrWhiteSpace(OperateNote) ? typeName : $"{typeName}：{OperateNote.Trim()}",
            CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        });
        _store.Save(_data);
        OnPropertyChanged(nameof(SelectedDetail));
        RefreshRecords();
        StatusMessage = $"{typeName} {Math.Abs(delta):0.#} 课时成功，剩余 {student.RemainingHours:0.#}";
        OperateHours = 1;
        OperateNote = "";
    }

    // ---- 排课管理 ----

    /// <summary>月历格子项</summary>
    public class CalendarDayItem
    {
        public DateTime Date { get; init; }
        /// <summary>是否当月</summary>
        public bool IsCurrentMonth { get; init; } = true;
        /// <summary>是否今天</summary>
        public bool IsToday { get; set; }
        /// <summary>排课人数</summary>
        public int Count { get; set; }
        /// <summary>是否不排课日</summary>
        public bool IsOff { get; set; }
        public string Hint => IsOff ? "休" : (Count > 0 ? $"{Count}人" : "");
        public string Display => Date.Day.ToString();
    }

    public void BuildCalendar()
    {
        CalendarDays.Clear();
        var firstDay = new DateTime(CurrentMonth.Year, CurrentMonth.Month, 1);
        // 周一为一周起点
        int offset = ((int)firstDay.DayOfWeek + 6) % 7;
        var start = firstDay.AddDays(-offset);
        var today = DateTime.Today;
        for (int i = 0; i < 42; i++)
        {
            var d = start.AddDays(i);
            var key = d.ToString("yyyy-MM-dd");
            var item = new CalendarDayItem
            {
                Date = d,
                IsCurrentMonth = d.Month == CurrentMonth.Month,
                IsToday = d == today,
                Count = _data.Schedule.TryGetValue(key, out var ids) ? ids.Count : 0,
                IsOff = _data.OffDays.Contains(key)
            };
            CalendarDays.Add(item);
        }
    }

    /// <summary>选中某天</summary>
    public void SelectCalendarDay(CalendarDayItem item)
    {
        SelectedDate = item.Date;
    }

    private void RefreshScheduleDetail()
    {
        var key = SelectedDate.ToString("yyyy-MM-dd");
        IsOffDayForSelected = _data.OffDays.Contains(key);
        ScheduleStudents.Clear();
        var scheduled = new HashSet<string>();
        if (_data.Schedule.TryGetValue(key, out var ids))
        {
            scheduled = new HashSet<string>(ids);
            foreach (var id in ids)
            {
                var s = _data.Students.FirstOrDefault(x => x.Id == id);
                if (s != null) ScheduleStudents.Add(s);
            }
        }
        // 刷新学生列表中"已排课"标记
        foreach (var s in Students)
        {
            s.IsInSchedule = scheduled.Contains(s.Id);
            OnPropertyChanged(nameof(s.IsInSchedule));
        }
        OnPropertyChanged(nameof(ScheduleStudents));
        OnPropertyChanged(nameof(ScheduleHint));
    }

    public string ScheduleHint => ScheduleStudents.Count == 0
        ? "该日未排课"
        : $"该日排课 {ScheduleStudents.Count} 人";

    /// <summary>添加/移除排课学生（切换）</summary>
    public void ToggleScheduleStudent(ChStudent student)
    {
        if (student == null) return;
        var key = SelectedDate.ToString("yyyy-MM-dd");
        if (!_data.Schedule.TryGetValue(key, out var ids))
        {
            ids = new List<string>();
            _data.Schedule[key] = ids;
        }
        if (ids.Contains(student.Id))
            ids.Remove(student.Id);
        else
            ids.Add(student.Id);
        _store.Save(_data);
        RefreshScheduleDetail();
        BuildCalendar();
        StatusMessage = IsOffDayForSelected
            ? "该日为不排课日，排课已忽略"
            : $"已更新 {SelectedDateText} 排课";
    }

    /// <summary>设置/取消不排课日</summary>
    public void ToggleOffDay()
    {
        var key = SelectedDate.ToString("yyyy-MM-dd");
        if (_data.OffDays.Contains(key))
            _data.OffDays.Remove(key);
        else
        {
            _data.OffDays.Add(key);
            // 不排课日自动清空排课
            if (_data.Schedule.ContainsKey(key))
                _data.Schedule.Remove(key);
        }
        _store.Save(_data);
        RefreshScheduleDetail();
        BuildCalendar();
        StatusMessage = IsOffDayForSelected ? $"{SelectedDateText} 已设为不排课日" : $"{SelectedDateText} 已恢复排课";
    }

    /// <summary>清空当日排课</summary>
    public void ClearSchedule()
    {
        var key = SelectedDate.ToString("yyyy-MM-dd");
        if (_data.Schedule.ContainsKey(key))
            _data.Schedule.Remove(key);
        _store.Save(_data);
        RefreshScheduleDetail();
        BuildCalendar();
        StatusMessage = $"已清空 {SelectedDateText} 排课";
    }

    /// <summary>复制某日排课到指定日期范围</summary>
    public int CopySchedule(DateTime fromDate, List<DateTime> toDates, bool skipOffDays)
    {
        var fromKey = fromDate.ToString("yyyy-MM-dd");
        if (!_data.Schedule.TryGetValue(fromKey, out var sourceIds) || sourceIds.Count == 0)
        {
            StatusMessage = $"源日期 {fromDate:yyyy-MM-dd} 无排课";
            return 0;
        }
        int copied = 0;
        foreach (var d in toDates)
        {
            var key = d.ToString("yyyy-MM-dd");
            if (skipOffDays && _data.OffDays.Contains(key)) continue;
            _data.Schedule[key] = new List<string>(sourceIds);
            copied++;
        }
        _store.Save(_data);
        RefreshScheduleDetail();
        BuildCalendar();
        StatusMessage = $"已将 {fromDate:yyyy-MM-dd} 的排课复制到 {copied} 天";
        return copied;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

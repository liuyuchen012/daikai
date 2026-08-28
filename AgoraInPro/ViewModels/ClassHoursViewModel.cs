using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Threading;
using CheckIn.Client.Models;
using CheckIn.Client.Services;

namespace CheckIn.Client.ViewModels;

/// <summary>
/// 课时划消 ViewModel：学生管理、课时划消记录、排课管理（每日复制、不排课日、上课时间细分、自动划消）
/// </summary>
public class ClassHoursViewModel : INotifyPropertyChanged
{
    private readonly ClassHourStore _store;
    private ClassHourData _data;
    private readonly DispatcherTimer _autoTimer;

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

    /// <summary>当日排课学生（含上课时间）</summary>
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

    // ---- 设置 ----

    /// <summary>每小时上课消耗课时（支持小数）</summary>
    public double HoursPerHour
    {
        get => _data.HoursPerHour;
        set
        {
            if (value <= 0 || value > 24) return;
            _data.HoursPerHour = value;
            _store.Save(_data);
            OnPropertyChanged();
            OnPropertyChanged(nameof(HoursPerHourText));
            StatusMessage = $"已保存：每小时消耗 {value:0.##} 课时";
        }
    }

    /// <summary>设置页文本框绑定：失焦解析并保存（支持小数）</summary>
    public string HoursPerHourText
    {
        get => _data.HoursPerHour.ToString("0.##");
        set
        {
            if (double.TryParse(value?.Trim(), out var v) && v > 0 && v <= 24)
            {
                _data.HoursPerHour = v;
                _store.Save(_data);
                OnPropertyChanged();
                OnPropertyChanged(nameof(HoursPerHour));
                StatusMessage = $"已保存：每小时消耗 {v:0.##} 课时";
            }
            else
            {
                StatusMessage = "请输入大于 0 的数字（支持小数，如 0.5）";
                OnPropertyChanged(); // 回显旧值
            }
        }
    }

    /// <summary>是否自动划消课时</summary>
    public bool AutoDeduct
    {
        get => _data.AutoDeduct;
        set
        {
            _data.AutoDeduct = value;
            _store.Save(_data);
            OnPropertyChanged();
            StatusMessage = value ? "已开启自动划消：课程结束后自动扣减课时" : "已关闭自动划消课时";
            if (value) AutoDeductCheck();
        }
    }

    public ClassHoursViewModel()
    {
        _store = new ClassHourStore();
        _data = _store.Load();
        ReloadStudents();
        BuildCalendar();
        // 自动划消定时器：每 30 秒检查一次，按排课时间表在课程结束后自动扣减课时（SlotKey 幂等，多实例安全）
        _autoTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
        _autoTimer.Tick += (_, _) => AutoDeductCheck();
        _autoTimer.Start();
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
            kv.Value.RemoveAll(e => e.StudentId == student.Id);
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

    /// <summary>月历格子项（IsSelected 支持属性变更通知，保证蓝框实时跟随选中日期）</summary>
    public class CalendarDayItem : INotifyPropertyChanged
    {
        public DateTime Date { get; init; }
        /// <summary>是否当月</summary>
        public bool IsCurrentMonth { get; init; } = true;
        /// <summary>是否今天</summary>
        public bool IsToday { get; set; }

        private bool _isSelected;
        /// <summary>是否选中日期（蓝框跟随移动）</summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        /// <summary>排课人数</summary>
        public int Count { get; set; }
        /// <summary>是否不排课日</summary>
        public bool IsOff { get; set; }
        public string Hint => IsOff ? "休" : (Count > 0 ? $"{Count}人" : "");
        public string Display => Date.Day.ToString();

        public event PropertyChangedEventHandler? PropertyChanged;
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
                IsSelected = d == SelectedDate,
                Count = _data.Schedule.TryGetValue(key, out var entries) ? entries.Count : 0,
                IsOff = _data.OffDays.Contains(key)
            };
            CalendarDays.Add(item);
        }
    }

    /// <summary>选中某天（蓝框跟随选中日期移动）</summary>
    public void SelectCalendarDay(CalendarDayItem item)
    {
        SelectedDate = item.Date;
        foreach (var day in CalendarDays)
            day.IsSelected = day.Date == item.Date;
    }

    private void RefreshScheduleDetail()
    {
        var key = SelectedDate.ToString("yyyy-MM-dd");
        IsOffDayForSelected = _data.OffDays.Contains(key);
        ScheduleStudents.Clear();
        var scheduled = new HashSet<string>();
        if (_data.Schedule.TryGetValue(key, out var entries))
        {
            scheduled = new HashSet<string>(entries.Select(e => e.StudentId));
            foreach (var entry in entries)
            {
                var s = _data.Students.FirstOrDefault(x => x.Id == entry.StudentId);
                if (s != null)
                {
                    s.ScheduleStartTime = entry.StartTime;
                    s.ScheduleEndTime = entry.EndTime;
                    ScheduleStudents.Add(s);
                }
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

    /// <summary>添加排课学生（必须填写上课/下课时间，无默认值）</summary>
    public bool AddToSchedule(ChStudent student, string startTime, string endTime)
    {
        if (student == null) return false;
        var key = SelectedDate.ToString("yyyy-MM-dd");
        if (IsOffDayForSelected)
        {
            StatusMessage = "该日为不排课日，不能添加排课";
            return false;
        }
        if (!TryNormalizePeriod(startTime, endTime, out var start, out var end))
        {
            StatusMessage = "请正确填写上课与下课时间（HH:mm，下课须晚于上课）";
            return false;
        }
        if (!_data.Schedule.TryGetValue(key, out var entries))
        {
            entries = new List<ScheduleEntry>();
            _data.Schedule[key] = entries;
        }
        if (entries.Any(e => e.StudentId == student.Id))
        {
            StatusMessage = $"「{student.Name}」已在当日排课中";
            return false;
        }
        entries.Add(new ScheduleEntry { StudentId = student.Id, StartTime = start, EndTime = end });
        _store.Save(_data);
        RefreshScheduleDetail();
        BuildCalendar();
        StatusMessage = $"已为「{student.Name}」添加排课 {start} - {end}";
        return true;
    }

    /// <summary>修改当日排课中某学生的上课/下课时间（两个都必填，失焦保存）</summary>
    public void UpdateScheduleTime(ChStudent student, string? startTime, string? endTime)
    {
        if (student == null) return;
        var key = SelectedDate.ToString("yyyy-MM-dd");
        if (!_data.Schedule.TryGetValue(key, out var entries)) return;
        var entry = entries.FirstOrDefault(e => e.StudentId == student.Id);
        if (entry == null) return;
        if (!TryNormalizePeriod(startTime, endTime, out var start, out var end))
        {
            RefreshScheduleDetail(); // 回显原值
            StatusMessage = "请正确填写上课与下课时间（HH:mm，下课须晚于上课）";
            return;
        }
        entry.StartTime = start;
        entry.EndTime = end;
        student.ScheduleStartTime = start;
        student.ScheduleEndTime = end;
        _store.Save(_data);
        OnPropertyChanged(nameof(ScheduleStudents));
        StatusMessage = $"{student.Name} 上课时间已设为 {start} - {end}";
    }

    private static bool TryNormalizeTime(string? time, out string result)
    {
        result = "";
        if (string.IsNullOrWhiteSpace(time)) return false;
        if (DateTime.TryParse(time.Trim(), out var dt))
        {
            result = dt.ToString("HH:mm");
            return true;
        }
        return false;
    }

    /// <summary>校验一组起止时间：均必填、格式 HH:mm 有效、下课不得等于上课（跨天课程自动顺延）</summary>
    private static bool TryNormalizePeriod(string? startTime, string? endTime, out string start, out string end)
    {
        start = "";
        end = "";
        if (!TryNormalizeTime(startTime, out start)) return false;
        if (!TryNormalizeTime(endTime, out end)) return false;
        var s = DateTime.Parse(start);
        var e = DateTime.Parse(end);
        return e != s; // 下课早于上课视为跨天课程（如 23:00-01:00），由 ScheduleEntry.End 自动顺延
    }

    /// <summary>从当日排课移除某学生</summary>
    public void RemoveFromSchedule(ChStudent student)
    {
        if (student == null) return;
        var key = SelectedDate.ToString("yyyy-MM-dd");
        if (_data.Schedule.TryGetValue(key, out var entries))
        {
            entries.RemoveAll(e => e.StudentId == student.Id);
            if (entries.Count == 0) _data.Schedule.Remove(key);
            _store.Save(_data);
            RefreshScheduleDetail();
            BuildCalendar();
            StatusMessage = $"已从 {SelectedDateText} 排课移除「{student.Name}」";
        }
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

    /// <summary>复制某日排课到指定日期范围（含上课时间）</summary>
    public int CopySchedule(DateTime fromDate, List<DateTime> toDates, bool skipOffDays)
    {
        var fromKey = fromDate.ToString("yyyy-MM-dd");
        if (!_data.Schedule.TryGetValue(fromKey, out var sourceEntries) || sourceEntries.Count == 0)
        {
            StatusMessage = $"源日期 {fromDate:yyyy-MM-dd} 无排课";
            return 0;
        }
        int copied = 0;
        foreach (var d in toDates)
        {
            var key = d.ToString("yyyy-MM-dd");
            if (skipOffDays && _data.OffDays.Contains(key)) continue;
            _data.Schedule[key] = sourceEntries
                .Select(e => new ScheduleEntry { StudentId = e.StudentId, StartTime = e.StartTime, EndTime = e.EndTime })
                .ToList();
            copied++;
        }
        _store.Save(_data);
        RefreshScheduleDetail();
        BuildCalendar();
        StatusMessage = $"已将 {fromDate:yyyy-MM-dd} 的排课复制到 {copied} 天";
        return copied;
    }

    // ---- 自动划消 ----

    /// <summary>
    /// 按排课时间表自动划消：课程在下课时间结束后，按实际时长（下课-上课）× 每小时消耗课时扣减。
    /// 起止时间未填写的排课不参与自动划消。通过 SlotKey 去重（日期|学生|上课时间|下课时间），重启后也不会重复扣减。
    /// </summary>
    private void AutoDeductCheck()
    {
        if (!_data.AutoDeduct) return;
        var key = DateTime.Today.ToString("yyyy-MM-dd");
        if (!_data.Schedule.TryGetValue(key, out var entries)) return;
        bool changed = false;
        foreach (var entry in entries)
        {
            if (!entry.IsValidTime) continue; // 未填写起止时间，无法计算课时
            if (DateTime.Now < entry.End) continue;
            var student = _data.Students.FirstOrDefault(s => s.Id == entry.StudentId);
            if (student == null || student.RemainingHours <= 0) continue;
            var slotKey = $"{key}|{entry.StudentId}|{entry.StartTime}|{entry.EndTime}";
            if (_data.Records.Any(r => r.SlotKey == slotKey)) continue;
            var deduct = Math.Min(Math.Round(entry.DurationHours * _data.HoursPerHour, 1), student.RemainingHours);
            if (deduct <= 0) continue;
            student.UsedHours += deduct;
            _data.Records.Add(new ChRecord
            {
                StudentId = student.Id,
                Date = key,
                Hours = -deduct,
                Note = $"自动划消 {entry.StartTime} - {entry.End:HH:mm}",
                SlotKey = slotKey,
                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            });
            changed = true;
        }
        if (!changed) return;
        _store.Save(_data);
        OnPropertyChanged(nameof(SelectedDetail));
        RefreshRecords();
        StatusMessage = "已按排课时间表自动划消课时";
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

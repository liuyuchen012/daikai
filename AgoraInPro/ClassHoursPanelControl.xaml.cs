using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CheckIn.Client.Models;
using CheckIn.Client.ViewModels;

namespace CheckIn.Client;

/// <summary>
/// 课时划消（划课）界面，可作为独立窗口（ClassHoursWindow）或控制中心「划课」页嵌入使用。
/// 自带 ClassHoursViewModel，宿主无需额外设置 DataContext。
/// </summary>
public partial class ClassHoursPanelControl : UserControl
{
    private readonly ClassHoursViewModel _vm;

    public ClassHoursPanelControl()
    {
        InitializeComponent();
        _vm = new ClassHoursViewModel();
        DataContext = _vm;
    }

    /// <summary>内部使用的课时视图模型（宿主可经此访问数据与操作）</summary>
    public ClassHoursViewModel HoursViewModel => _vm;

    // ============ 学生管理 ============

    private void AddStudent_Click(object sender, RoutedEventArgs e)
    {
        _vm.AddStudent();
        if (string.IsNullOrEmpty(_vm.NewStudentName))
            NameBox.Focus();
    }

    private void DeleteStudent_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi && mi.DataContext is ChStudent student)
        {
            if (MessageBox.Show($"确定删除学生「{student.Name}」？其课时记录与排课将一并清除。",
                    "删除学生", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                _vm.DeleteStudent(student);
        }
    }

    /// <summary>菜单：聚焦添加学生输入框</summary>
    public void FocusAddStudent()
    {
        NameBox.Focus();
        NameBox.SelectAll();
    }

    /// <summary>菜单：删除当前选中的学生</summary>
    public void DeleteSelectedStudent()
    {
        if (_vm.SelectedStudent == null)
        {
            MessageBox.Show("请先选择要删除的学生", "课时划消",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (MessageBox.Show($"确定删除学生「{_vm.SelectedStudent.Name}」？其课时记录与排课将一并清除。",
                "删除学生", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            _vm.DeleteStudent(_vm.SelectedStudent);
    }

    /// <summary>菜单：显示使用说明</summary>
    public void ShowHelp()
    {
        MessageBox.Show(
            "【课时划消】\n" +
            "1. 左侧添加学生并设置初始课时。\n" +
            "2. 选中学生后可划消课时或增加课时，并填写备注。\n" +
            "3. 「排课管理」按学生编排每日课程，格子显示人数或「休」。\n" +
            "4. 「当日排课」可为选中日期添加/移除排课学生、设置不排课日、清空排课、复制排课。\n" +
            "5. 所有数据保存在本地 data/classhours.json。",
            "使用说明", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    // ============ 课时操作 ============

    private void Deduct_Click(object sender, RoutedEventArgs e) => _vm.DeductHours();

    private void AddHours_Click(object sender, RoutedEventArgs e) => _vm.AddHours();

    // ============ 排课管理 ============

    private void PrevMonth_Click(object sender, RoutedEventArgs e)
    {
        _vm.CurrentMonth = _vm.CurrentMonth.AddMonths(-1);
        _vm.BuildCalendar();
    }

    private void NextMonth_Click(object sender, RoutedEventArgs e)
    {
        _vm.CurrentMonth = _vm.CurrentMonth.AddMonths(1);
        _vm.BuildCalendar();
    }

    private void Today_Click(object sender, RoutedEventArgs e)
    {
        _vm.CurrentMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
        _vm.BuildCalendar();
        _vm.SelectedDate = DateTime.Today;
    }

    private void CalendarDay_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border b && b.DataContext is ClassHoursViewModel.CalendarDayItem item)
            _vm.SelectCalendarDay(item);
    }

    private void ToggleStudent_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border b && b.DataContext is ChStudent student)
            _vm.ToggleScheduleStudent(student);
    }

    private void ToggleOffDay_Click(object sender, RoutedEventArgs e) => _vm.ToggleOffDay();

    private void ClearSchedule_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show($"确定清空 {_vm.SelectedDate:yyyy-MM-dd} 的排课？",
                "清空排课", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            _vm.ClearSchedule();
    }

    private void CopySchedule_Click(object sender, RoutedEventArgs e) => CopySchedule();

    /// <summary>菜单/按钮：复制当前日期排课到其他日期</summary>
    public void CopySchedule()
    {
        var dialog = new CopyScheduleDialog(_vm.SelectedDate)
        {
            Owner = Window.GetWindow(this)
        };
        if (dialog.ShowDialog() == true)
        {
            _vm.CopySchedule(dialog.FromDate, dialog.ToDates, dialog.SkipOffDays);
        }
    }

    /// <summary>菜单：设为不排课日/恢复</summary>
    public void ToggleOffDay() => _vm.ToggleOffDay();

    /// <summary>菜单：清空当前日期排课</summary>
    public void ClearSchedule()
    {
        if (MessageBox.Show($"确定清空 {_vm.SelectedDate:yyyy-MM-dd} 的排课？",
                "清空排课", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            _vm.ClearSchedule();
    }
}

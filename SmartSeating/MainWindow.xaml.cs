using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SmartSeating
{
    // ===============================================================
    // 数据模型类 (Data Model Classes)
    // ===============================================================
    public class Student
    {
        public string Name { get; set; } = string.Empty;
        public double HeightWeight { get; set; }
        public double ImportanceWeight { get; set; }
        public List<string> AvoidStudents { get; set; } = new();
        public List<string> PreferStudents { get; set; } = new();
        public string PreferArea { get; set; } = string.Empty;
        public bool IsFixed { get; set; }
        public bool IsDisabled { get; set; }
    }

    public class Seat
    {
        public Student? Student { get; set; }
        public bool IsDisabled { get; set; }
        public bool IsFixed { get; set; }
        public string StudentName => Student?.Name ?? string.Empty;
    }

    public class SeatAssignment
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public string StudentName { get; set; } = string.Empty;
        public bool IsFixed { get; set; }
        public bool IsDisabled { get; set; }
    }

    public class SeatingConfiguration
    {
        public int Rows { get; set; }
        public int Cols { get; set; }
        public List<Student> AllStudents { get; set; } = new();
        public List<SeatAssignment> Assignments { get; set; } = new();
    }

    // ===============================================================
    // 主窗口逻辑 (Main Window Logic)
    // ===============================================================
    public partial class MainWindow : Window
    {
        private const double SeatWidth = 120;
        private const double SeatHeight = 48;
        private const double SeatSpacing = 12;
        private const string DefaultConfigFileName = "SeatingConfig.json";

        private readonly ObservableCollection<Student> _students = new();
        private readonly HashSet<(int row, int col)> _selection = new();
        private readonly Random _random = new();
        private ICollectionView? _studentView;

        private Seat[,] _seats = null!;
        private int _rows = 10;
        private int _cols = 10;

        private bool _isDragging;
        private Point _dragStartPoint;
        private (int row, int col)? _activeSeat;

        private Brush _normalBrush = null!;
        private Brush _fixedBrush = null!;
        private Brush _disabledBrush = null!;
        private Brush _strokeBrush = null!;
        private Brush _selectedStrokeBrush = null!;
        private Brush _hoverStrokeBrush = null!;
        private Brush _textBrush = null!;
        private double _seatCornerRadius;
        private double _seatFontSize;

        public MainWindow()
        {
            InitializeComponent();

            studentListView.ItemsSource = _students;
            _studentView = CollectionViewSource.GetDefaultView(_students);
            if (_studentView != null)
            {
                _studentView.Filter = StudentFilter;
            }
            InitializeResourceCache();
            HookCanvasEvents();
            RootGrid.KeyDown += RootGrid_KeyDown;
            RootGrid.Focus();

            GenerateSeats(_rows, _cols, true);

            Loaded += MainWindow_LoadedAsync;
            Closing += MainWindow_ClosingAsync;
        }

        private async void MainWindow_LoadedAsync(object sender, RoutedEventArgs e)
        {
            try
            {
                await LoadConfigurationAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void MainWindow_ClosingAsync(object? sender, CancelEventArgs e)
        {
            try
            {
                await SaveConfigurationAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void InitializeResourceCache()
        {
            _normalBrush = (Brush)RootGrid.Resources["SeatNormalBrush"];
            _fixedBrush = (Brush)RootGrid.Resources["SeatFixedBrush"];
            _disabledBrush = (Brush)RootGrid.Resources["SeatDisabledBrush"];
            _strokeBrush = (Brush)RootGrid.Resources["SeatStrokeBrush"];
            _selectedStrokeBrush = (Brush)RootGrid.Resources["SeatSelectedStrokeBrush"];
            _hoverStrokeBrush = (Brush)RootGrid.Resources["SeatHoverStrokeBrush"];
            _textBrush = (Brush)RootGrid.Resources["SeatTextBrush"];
            _seatCornerRadius = (double)RootGrid.Resources["SeatCornerRadius"];
            _seatFontSize = (double)RootGrid.Resources["SeatFontSize"];
        }

        private void HookCanvasEvents()
        {
            seatCanvas.MouseLeftButtonDown += SeatCanvas_MouseLeftButtonDown;
            seatCanvas.MouseMove += SeatCanvas_MouseMove;
            seatCanvas.MouseLeftButtonUp += SeatCanvas_MouseLeftButtonUp;
        }

        #region 座位生成与绘制
        private void GenerateSeats_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(txtRows.Text, out var newRows)) newRows = _rows;
            if (!int.TryParse(txtCols.Text, out var newCols)) newCols = _cols;
            GenerateSeats(newRows, newCols);
        }

        private void GenerateSeats(int requestedRows, int requestedCols, bool clearAll = false)
        {
            var existingAssignments = new Dictionary<string, (int row, int col)>();
            if (_seats != null)
            {
                for (int r = 0; r < _rows; r++)
                {
                    for (int c = 0; c < _cols; c++)
                    {
                        var student = _seats[r, c].Student;
                        if (student != null)
                        {
                            existingAssignments[student.Name] = (r, c);
                        }
                    }
                }
            }

            _rows = Math.Max(1, requestedRows);
            _cols = Math.Max(1, requestedCols);
            txtRows.Text = _rows.ToString();
            txtCols.Text = _cols.ToString();

            if (clearAll)
            {
                _students.Clear();
            }

            var newSeats = new Seat[_rows, _cols];
            for (int r = 0; r < _rows; r++)
            {
                for (int c = 0; c < _cols; c++)
                {
                    newSeats[r, c] = new Seat();
                }
            }

            if (!clearAll)
            {
                foreach (var student in _students)
                {
                    if (existingAssignments.TryGetValue(student.Name, out var pos) && pos.row < _rows && pos.col < _cols)
                    {
                        if (newSeats[pos.row, pos.col].Student == null)
                        {
                            newSeats[pos.row, pos.col].Student = student;
                        }
                    }
                }
            }

            _seats = newSeats;
            DrawSeatCanvas();

            if (_rows > 0 && _cols > 0 && _selection.Count == 0)
            {
                _selection.Add((0, 0));
                ShowStudentDetailForPosition(0, 0);
                DrawSeatCanvas();
            }
        }

        private void BtnClearAllSeats_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("确定要清空所有学生信息和座位安排吗？此操作不可撤销。", "确认清空", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) return;

            GenerateSeats(_rows, _cols, true);
            _selection.Clear();
            _activeSeat = null;
            lblPosition.Text = "(无)";
            txtStudentName.Text = string.Empty;
            studentDetailsPanel.IsEnabled = false;
            txtHeightWeight.Text = "0";
            txtImportanceWeight.Text = "0";
            txtPreferStudents.Text = string.Empty;
            txtAvoidStudents.Text = string.Empty;
            txtPreferArea.Text = string.Empty;
            chkFixed.IsChecked = false;
            chkDisabled.IsChecked = false;
            txtStatus.Text = "已清空所有学生信息。";
        }

        private void BtnPlaceAllStudents_Click(object sender, RoutedEventArgs e)
        {
            if (_seats == null)
            {
                txtStatus.Text = "请先生成座位。";
                return;
            }

            if (_students.Count == 0)
            {
                txtStatus.Text = "没有学生可放置。";
                return;
            }

            var fixedStudentNames = new HashSet<string>(_seats.Cast<Seat>()
                .Where(seat => seat.IsFixed && seat.Student != null)
                .Select(seat => seat.Student!.Name));

            foreach (var seat in _seats)
            {
                if (!seat.IsFixed)
                {
                    seat.Student = null;
                }
            }

            var queue = new Queue<Student>(_students.Where(s => !fixedStudentNames.Contains(s.Name)));
            for (int r = 0; r < _rows && queue.Count > 0; r++)
            {
                for (int c = 0; c < _cols && queue.Count > 0; c++)
                {
                    var seat = _seats[r, c];
                    if (seat.IsDisabled || seat.IsFixed)
                    {
                        continue;
                    }
                    seat.Student = queue.Dequeue();
                }
            }

            DrawSeatCanvas();
            txtStatus.Text = queue.Count == 0 ? "已完成学生放置。" : "学生数量超过可用座位，部分学生未放置。";
        }

        private void DrawSeatCanvas()
        {
            seatCanvas.Children.Clear();
            seatCanvas.Width = _cols * (SeatWidth + SeatSpacing);
            seatCanvas.Height = _rows * (SeatHeight + SeatSpacing);

            for (int r = 0; r < _rows; r++)
            {
                for (int c = 0; c < _cols; c++)
                {
                    var seat = _seats[r, c];
                    var position = new Tuple<int, int>(r, c);

                    double x = c * (SeatWidth + SeatSpacing);
                    double y = r * (SeatHeight + SeatSpacing);

                    var rect = new Rectangle
                    {
                        Width = SeatWidth,
                        Height = SeatHeight,
                        RadiusX = _seatCornerRadius,
                        RadiusY = _seatCornerRadius,
                        Fill = seat.IsDisabled ? _disabledBrush : (seat.IsFixed ? _fixedBrush : _normalBrush),
                        Stroke = _selection.Contains((r, c)) ? _selectedStrokeBrush : _strokeBrush,
                        StrokeThickness = _selection.Contains((r, c)) ? 2 : 1,
                        Tag = position
                    };

                    Canvas.SetLeft(rect, x);
                    Canvas.SetTop(rect, y);
                    seatCanvas.Children.Add(rect);

                    rect.ToolTip = $"{(string.IsNullOrEmpty(seat.StudentName) ? "空" : seat.StudentName)}\n" +
                                   $"状态: {(seat.IsFixed ? "固定" : "未固定")}, {(seat.IsDisabled ? "禁用" : "可用")}";

                    var label = new TextBlock
                    {
                        Text = string.IsNullOrEmpty(seat.StudentName) ? "空" : seat.StudentName,
                        Width = SeatWidth,
                        Height = SeatHeight,
                        TextAlignment = TextAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        FontSize = _seatFontSize,
                        Foreground = _textBrush,
                        IsHitTestVisible = false
                    };

                    Canvas.SetLeft(label, x);
                    Canvas.SetTop(label, y);
                    seatCanvas.Children.Add(label);

                    rect.MouseLeftButtonDown += Rect_MouseLeftButtonDown;
                    rect.MouseRightButtonUp += Rect_MouseRightButtonUp;
                    rect.MouseEnter += (s, e) => ApplyHoverStroke(rect, true);
                    rect.MouseLeave += (s, e) => ApplyHoverStroke(rect, false);

                    var doubleClick = new MouseGesture(MouseAction.LeftDoubleClick);
                    var toggleFixedBinding = new MouseBinding(new RelayCommand(_ =>
                    {
                        _seats[r, c].IsFixed = !_seats[r, c].IsFixed;
                        DrawSeatCanvas();
                    }), doubleClick);
                    rect.InputBindings.Add(toggleFixedBinding);
                }
            }

            UpdateSeatSummary();
        }

        private void ApplyHoverStroke(Rectangle rect, bool isHover)
        {
            if (rect.Tag is Tuple<int, int> tuple)
            {
                var coords = (tuple.Item1, tuple.Item2);
                if (_selection.Contains(coords))
                {
                    rect.Stroke = isHover ? _hoverStrokeBrush : _selectedStrokeBrush;
                    rect.StrokeThickness = 2;
                }
                else
                {
                    rect.Stroke = isHover ? _hoverStrokeBrush : _strokeBrush;
                    rect.StrokeThickness = isHover ? 2 : 1;
                }
            }
        }
        #endregion

        #region 座位交互
        private void Rect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not FrameworkElement element || element.Tag is not Tuple<int, int> tuple)
            {
                return;
            }

            var coords = (tuple.Item1, tuple.Item2);
            bool shiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

            if (!shiftPressed)
            {
                _selection.Clear();
            }

            if (_selection.Contains(coords))
            {
                if (shiftPressed) // Only remove if shift is pressed and it's already selected
                {
                    _selection.Remove(coords);
                }
            }
            else
            {
                _selection.Add(coords);
            }

            DrawSeatCanvas();

            if (_selection.Count == 1)
            {
                ShowStudentDetailForPosition(coords.Item1, coords.Item2);
            }
            else
            {
                ClearStudentDetailPanel();
            }
            e.Handled = true;
        }


        private void Rect_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.Tag is Tuple<int, int> tuple)
            {
                var coords = (tuple.Item1, tuple.Item2);
                if (!_selection.Contains(coords))
                {
                    _selection.Clear();
                    _selection.Add(coords);
                    DrawSeatCanvas();
                    ShowStudentDetailForPosition(coords.Item1, coords.Item2);
                }
                ShowSeatActionMenu(element);
                e.Handled = true;
            }
        }

        private void ShowSeatActionMenu(FrameworkElement anchor)
        {
            var menu = new ContextMenu();

            var toggleDisabled = new MenuItem { Header = "切换 禁用/启用" };
            toggleDisabled.Click += (_, __) =>
            {
                foreach (var (row, col) in _selection)
                {
                    _seats[row, col].IsDisabled = !_seats[row, col].IsDisabled;
                }
                DrawSeatCanvas();
            };

            var toggleFixed = new MenuItem { Header = "切换 固定/解除" };
            toggleFixed.Click += (_, __) =>
            {
                foreach (var (row, col) in _selection)
                {
                    _seats[row, col].IsFixed = !_seats[row, col].IsFixed;
                }
                DrawSeatCanvas();
            };

            var clearItem = new MenuItem { Header = "清空所选座位学生" };
            clearItem.Click += (_, __) =>
            {
                foreach (var (row, col) in _selection)
                {
                    var student = _seats[row, col].Student;
                    if (student != null)
                    {
                        _students.Remove(student);
                        _seats[row, col].Student = null;
                    }
                }
                DrawSeatCanvas();
                ClearStudentDetailPanel();
            };
            clearItem.IsEnabled = _selection.Any(s => _seats[s.row, s.col].Student != null);

            menu.Items.Add(toggleDisabled);
            menu.Items.Add(toggleFixed);
            menu.Items.Add(new Separator());
            menu.Items.Add(clearItem);
            menu.PlacementTarget = anchor;
            menu.IsOpen = true;
        }

        private void SeatCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.Source == seatCanvas)
            {
                _isDragging = true;
                _dragStartPoint = e.GetPosition(seatCanvas);
                seatCanvas.CaptureMouse();
                _selection.Clear(); // Start with a clear selection
                DrawSeatCanvas(); // Redraw to clear previous selections
                e.Handled = true;
            }
        }

        private void SeatCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging) return;

            var currentPoint = e.GetPosition(seatCanvas);

            int startRow = (int)(_dragStartPoint.Y / (SeatHeight + SeatSpacing));
            int startCol = (int)(_dragStartPoint.X / (SeatWidth + SeatSpacing));
            int endRow = (int)(currentPoint.Y / (SeatHeight + SeatSpacing));
            int endCol = (int)(currentPoint.X / (SeatWidth + SeatSpacing));

            _selection.Clear();
            for (int r = Math.Min(startRow, endRow); r <= Math.Max(startRow, endRow); r++)
            {
                if (r < 0 || r >= _rows) continue;
                for (int c = Math.Min(startCol, endCol); c <= Math.Max(startCol, endCol); c++)
                {
                    if (c < 0 || c >= _cols) continue;
                    _selection.Add((r, c));
                }
            }
            DrawSeatCanvas();
        }

        private void SeatCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDragging)
            {
                _isDragging = false;
                seatCanvas.ReleaseMouseCapture();
                if (_selection.Count == 1)
                {
                    var (r, c) = _selection.First();
                    ShowStudentDetailForPosition(r, c);
                }
                else
                {
                    ClearStudentDetailPanel();
                }
            }
        }

        private void RootGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (_selection.Count == 0) return;
            bool changed = false;

            switch (e.Key)
            {
                case Key.F:
                    foreach (var (row, col) in _selection) _seats[row, col].IsFixed = !_seats[row, col].IsFixed;
                    changed = true;
                    break;
                case Key.D:
                    foreach (var (row, col) in _selection) _seats[row, col].IsDisabled = !_seats[row, col].IsDisabled;
                    changed = true;
                    break;
                case Key.Delete:
                case Key.Back:
                    foreach (var (row, col) in _selection)
                    {
                        var student = _seats[row, col].Student;
                        if (student != null)
                        {
                            _students.Remove(student);
                            _seats[row, col].Student = null;
                        }
                    }
                    changed = true;
                    break;
            }

            if (changed)
            {
                DrawSeatCanvas();
                if (_selection.Count == 1)
                {
                    var (r, c) = _selection.First();
                    ShowStudentDetailForPosition(r, c);
                }
                else
                {
                    ClearStudentDetailPanel();
                }
                e.Handled = true;
            }
        }
        #endregion

        #region 右侧详情面板
        private void ClearStudentDetailPanel()
        {
            _activeSeat = null;
            lblPosition.Text = _selection.Count > 1 ? "(多选)" : "(无)";
            txtStudentName.Text = string.Empty;
            studentDetailsPanel.IsEnabled = false;
            txtHeightWeight.Text = "0";
            txtImportanceWeight.Text = "0";
            txtPreferStudents.Text = string.Empty;
            txtAvoidStudents.Text = string.Empty;
            txtPreferArea.Text = string.Empty;
            chkFixed.IsChecked = false;
            chkDisabled.IsChecked = false;
        }

        private void ShowStudentDetailForPosition(int row, int col)
        {
            _activeSeat = (row, col);
            lblPosition.Text = $"({row + 1}, {col + 1})";
            var seat = _seats[row, col];
            var student = seat.Student;

            chkFixed.IsChecked = seat.IsFixed;
            chkDisabled.IsChecked = seat.IsDisabled;
            studentDetailsPanel.IsEnabled = true;

            if (student == null)
            {
                txtStudentName.Text = string.Empty;
                txtHeightWeight.Text = "0";
                txtImportanceWeight.Text = "0";
                txtPreferStudents.Text = string.Empty;
                txtAvoidStudents.Text = string.Empty;
                txtPreferArea.Text = string.Empty;
            }
            else
            {
                txtStudentName.Text = student.Name;
                txtHeightWeight.Text = student.HeightWeight.ToString();
                txtImportanceWeight.Text = student.ImportanceWeight.ToString();
                txtPreferStudents.Text = string.Join(", ", student.PreferStudents);
                txtAvoidStudents.Text = string.Join(", ", student.AvoidStudents);
                txtPreferArea.Text = student.PreferArea;
            }
        }

        private void BtnSaveStudentToSeat_Click(object sender, RoutedEventArgs e)
        {
            if (_activeSeat == null)
            {
                MessageBox.Show("请先选择一个座位。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var (row, col) = _activeSeat.Value;
            var seat = _seats[row, col];
            var name = txtStudentName.Text?.Trim() ?? string.Empty;

            // Handle student assignment/creation
            if (string.IsNullOrEmpty(name))
            {
                if (seat.Student != null)
                {
                    _students.Remove(seat.Student);
                    seat.Student = null;
                }
            }
            else
            {
                Student? targetStudent;
                var existingStudent = _students.FirstOrDefault(s => s.Name.Equals(name, StringComparison.Ordinal));
                if (seat.Student?.Name == name)
                {
                    targetStudent = seat.Student;
                }
                else if (existingStudent != null)
                {
                    var oldSeat = FindSeatOfStudent(name);
                    if (oldSeat.HasValue) _seats[oldSeat.Value.row, oldSeat.Value.col].Student = null;
                    targetStudent = existingStudent;
                }
                else
                {
                    if (seat.Student != null)
                    {
                        seat.Student.Name = name; // Rename existing student on seat
                        targetStudent = seat.Student;
                    }
                    else
                    {
                        targetStudent = new Student { Name = name };
                        _students.Add(targetStudent);
                    }
                }
                seat.Student = targetStudent;

                // Update student properties
                if (double.TryParse(txtHeightWeight.Text, out var height)) targetStudent.HeightWeight = height;
                if (double.TryParse(txtImportanceWeight.Text, out var importance)) targetStudent.ImportanceWeight = importance;
                targetStudent.PreferArea = txtPreferArea.Text?.Trim() ?? "";
                targetStudent.PreferStudents = txtPreferStudents.Text.Split(new[] { ',', '，', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
                targetStudent.AvoidStudents = txtAvoidStudents.Text.Split(new[] { ',', '，', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
            }

            seat.IsFixed = chkFixed.IsChecked == true;
            seat.IsDisabled = chkDisabled.IsChecked == true;

            DrawSeatCanvas();
            ShowStudentDetailForPosition(row, col); // Refresh details
            txtStatus.Text = "更改已保存。";
        }

        private (int row, int col)? FindSeatOfStudent(string studentName)
        {
            for (int r = 0; r < _rows; r++)
                for (int c = 0; c < _cols; c++)
                    if (_seats[r, c].Student?.Name == studentName)
                        return (r, c);
            return null;
        }

        private void BtnClearStudentFromSeat_Click(object sender, RoutedEventArgs e)
        {
            if (_activeSeat == null) return;
            var (row, col) = _activeSeat.Value;
            var seat = _seats[row, col];

            if (seat.Student != null)
            {
                _students.Remove(seat.Student);
                seat.Student = null;
            }

            DrawSeatCanvas();
            ShowStudentDetailForPosition(row, col);
        }

        private (int row, int col)? FindFirstAvailableSeat()
        {
            for (int r = 0; r < _rows; r++)
            {
                for (int c = 0; c < _cols; c++)
                {
                    var seat = _seats[r, c];
                    if (seat.IsDisabled || seat.IsFixed) continue;
                    if (seat.Student == null) return (r, c);
                }
            }
            return null;
        }

        private (int row, int col)? GetFirstSelectableSeatFromSelection()
        {
            foreach (var coords in _selection.OrderBy(s => s.row).ThenBy(s => s.col))
            {
                var seat = _seats[coords.row, coords.col];
                if (!seat.IsDisabled)
                {
                    return coords;
                }
            }
            return null;
        }

        private void ScrollSeatIntoView((int row, int col) coords)
        {
            if (seatScrollViewer == null) return;
            double x = coords.col * (SeatWidth + SeatSpacing);
            double y = coords.row * (SeatHeight + SeatSpacing);
            double targetX = Math.Max(0, x - seatScrollViewer.ViewportWidth / 2);
            double targetY = Math.Max(0, y - seatScrollViewer.ViewportHeight / 2);
            seatScrollViewer.ScrollToHorizontalOffset(targetX);
            seatScrollViewer.ScrollToVerticalOffset(targetY);
        }

        private void UpdateSeatSummary()
        {
            if (_seats == null)
            {
                txtSeatSummary.Text = string.Empty;
                txtStudentSummary.Text = string.Empty;
                return;
            }

            int totalSeats = _rows * _cols;
            int disabledSeats = _seats.Cast<Seat>().Count(seat => seat.IsDisabled);
            int fixedSeats = _seats.Cast<Seat>().Count(seat => seat.IsFixed);
            var occupiedSeatStudents = _seats.Cast<Seat>().Where(seat => seat.Student != null).Select(seat => seat.Student!.Name).ToList();
            int occupiedSeats = occupiedSeatStudents.Count;
            int availableSeats = totalSeats - disabledSeats - occupiedSeats;
            var assignedStudentNames = new HashSet<string>(occupiedSeatStudents);
            int unassignedStudents = Math.Max(0, _students.Count - assignedStudentNames.Count);

            txtSeatSummary.Text =
                $"总座位: {totalSeats} (固定 {fixedSeats}, 禁用 {disabledSeats})\n" +
                $"已占用: {occupiedSeats}，可用空位: {Math.Max(0, availableSeats)}";
            txtStudentSummary.Text =
                $"学生总数: {_students.Count}，已安排: {assignedStudentNames.Count}，待安排: {unassignedStudents}";
        }

        #region 学生列表增强
        private void TxtStudentSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyStudentFilter();
        }

        private void BtnClearStudentSearch_Click(object sender, RoutedEventArgs e)
        {
            txtStudentSearch.Text = string.Empty;
        }

        private void ApplyStudentFilter()
        {
            _studentView?.Refresh();
        }

        private bool StudentFilter(object obj)
        {
            if (obj is not Student student) return false;
            var keyword = txtStudentSearch?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(keyword)) return true;
            return student.Name.Contains(keyword, StringComparison.OrdinalIgnoreCase);
        }

        private void BtnAssignStudentToSeat_Click(object sender, RoutedEventArgs e)
        {
            if (studentListView.SelectedItem is not Student student)
            {
                MessageBox.Show("请先在左侧选择一名学生。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!TryAssignStudentToSeat(student))
            {
                MessageBox.Show("当前没有可用的座位 (需可用且未固定)。", "无法分配", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void StudentListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (studentListView.SelectedItem is Student student)
            {
                TryAssignStudentToSeat(student);
            }
        }

        private void BtnLocateStudent_Click(object sender, RoutedEventArgs e)
        {
            if (studentListView.SelectedItem is not Student student)
            {
                MessageBox.Show("请先在左侧选择一名学生。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var coords = FindSeatOfStudent(student.Name);
            if (!coords.HasValue)
            {
                txtStatus.Text = $"{student.Name} 尚未安排座位。";
                return;
            }

            _selection.Clear();
            _selection.Add(coords.Value);
            DrawSeatCanvas();
            ShowStudentDetailForPosition(coords.Value.row, coords.Value.col);
            ScrollSeatIntoView(coords.Value);
        }

        private bool TryAssignStudentToSeat(Student student)
        {
            if (_seats == null) return false;

            var target = GetFirstSelectableSeatFromSelection() ?? FindFirstAvailableSeat();
            if (!target.HasValue) return false;

            var (row, col) = target.Value;
            var seat = _seats[row, col];
            if (seat.IsDisabled || seat.IsFixed)
            {
                return false;
            }

            var previousSeat = FindSeatOfStudent(student.Name);
            if (previousSeat.HasValue)
            {
                _seats[previousSeat.Value.row, previousSeat.Value.col].Student = null;
            }

            seat.Student = student;
            _selection.Clear();
            _selection.Add((row, col));
            DrawSeatCanvas();
            ShowStudentDetailForPosition(row, col);
            ScrollSeatIntoView((row, col));
            txtStatus.Text = $"已将 {student.Name} 分配到 ({row + 1}, {col + 1})。";
            return true;
        }
        #endregion
        #endregion

        #region 导入/导出
        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls", Title = "选择学生名单文件" };
            if (dialog.ShowDialog() != true) return;

            try
            {
                var names = new List<string>();
                using (var workbook = new XLWorkbook(dialog.FileName))
                {
                    var sheet = workbook.Worksheets.FirstOrDefault();
                    if (sheet == null)
                    {
                        MessageBox.Show("Excel文件中没有找到工作表。", "导入错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    foreach (var row in sheet.RowsUsed())
                    {
                        var value = row.Cell(1).GetValue<string>()?.Trim();
                        if (!string.IsNullOrEmpty(value) && !names.Contains(value)) names.Add(value);
                    }
                }

                var result = MessageBox.Show($"即将导入 {names.Count} 名学生。这将清空当前所有学生列表和座位安排，是否继续？", "确认导入", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result != MessageBoxResult.Yes) return;

                GenerateSeats(_rows, _cols, true); // Clear everything
                foreach (var name in names) _students.Add(new Student { Name = name });

                BtnPlaceAllStudents_Click(this, new RoutedEventArgs()); // Auto-place them
                txtStatus.Text = $"已成功导入并放置 {names.Count} 名学生。";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入Excel出错: {ex.Message}", "导入失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnExportConfig_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog { Filter = "JSON Files|*.json", Title = "导出配置文件", FileName = "SeatingConfig.json" };
            if (dialog.ShowDialog() != true) return;

            try
            {
                await SaveConfigurationAsync(dialog.FileName);
                txtStatus.Text = $"配置已导出到 {System.IO.Path.GetFileName(dialog.FileName)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnImportConfig_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filter = "JSON Files|*.json", Title = "导入配置文件" };
            if (dialog.ShowDialog() != true) return;

            try
            {
                await LoadConfigurationAsync(dialog.FileName);
                txtStatus.Text = $"已从 {System.IO.Path.GetFileName(dialog.FileName)} 导入配置。";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetDefaultConfigPath() => System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DefaultConfigFileName);

        private async Task SaveConfigurationAsync(string? filePath = null)
        {
            if (_seats == null) return;
            var path = filePath ?? GetDefaultConfigPath();

            var config = new SeatingConfiguration
            {
                Rows = _rows,
                Cols = _cols,
                AllStudents = new List<Student>(_students),
                Assignments = new List<SeatAssignment>()
            };

            for (int r = 0; r < _rows; r++)
            {
                for (int c = 0; c < _cols; c++)
                {
                    var seat = _seats[r, c];
                    config.Assignments.Add(new SeatAssignment
                    {
                        Row = r,
                        Col = c,
                        StudentName = seat.StudentName,
                        IsFixed = seat.IsFixed,
                        IsDisabled = seat.IsDisabled
                    });
                }
            }

            var options = new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
            await using var stream = File.Create(path);
            await JsonSerializer.SerializeAsync(stream, config, options);
        }

        private async Task LoadConfigurationAsync(string? filePath = null)
        {
            var path = filePath ?? GetDefaultConfigPath();
            if (!File.Exists(path))
            {
                GenerateSeats(_rows, _cols, true);
                return;
            }

            var options = new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
            await using var stream = File.OpenRead(path);
            var config = await JsonSerializer.DeserializeAsync<SeatingConfiguration>(stream, options) ?? new SeatingConfiguration();

            _rows = config.Rows > 0 ? config.Rows : 10;
            _cols = config.Cols > 0 ? config.Cols : 10;

            _students.Clear();
            if (config.AllStudents != null)
            {
                foreach (var student in config.AllStudents)
                {
                    student.AvoidStudents ??= new List<string>();
                    student.PreferStudents ??= new List<string>();
                    _students.Add(student);
                }
            }

            GenerateSeats(_rows, _cols, false); // 保留学生列表，仅重建座位
            var studentMap = _students.ToDictionary(s => s.Name, s => s);

            if (config.Assignments != null)
            {
                foreach (var assignment in config.Assignments.Where(a => a.Row < _rows && a.Col < _cols))
                {
                    var seat = _seats[assignment.Row, assignment.Col];
                    seat.IsFixed = assignment.IsFixed;
                    seat.IsDisabled = assignment.IsDisabled;
                    if (!string.IsNullOrWhiteSpace(assignment.StudentName) && studentMap.TryGetValue(assignment.StudentName, out var student))
                    {
                        seat.Student = student;
                    }
                }
            }

            txtRows.Text = _rows.ToString();
            txtCols.Text = _cols.ToString();
            DrawSeatCanvas();
            _selection.Clear();
            ClearStudentDetailPanel();
            txtStatus.Text = "配置已加载。";
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx", Title = "导出座位表", FileName = "seat_layout.xlsx" };
            if (dialog.ShowDialog() != true) return;

            try
            {
                using var workbook = new XLWorkbook();
                var sheet = workbook.Worksheets.Add("座位表");
                for (int r = 0; r < _rows; r++)
                {
                    for (int c = 0; c < _cols; c++)
                    {
                        sheet.Cell(r + 1, c + 1).Value = _seats[r, c].StudentName;
                    }
                }
                sheet.Columns().AdjustToContents();
                workbook.SaveAs(dialog.FileName);
                txtStatus.Text = "已导出座位布局到 Excel。";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExportImage_Click(object sender, RoutedEventArgs e)
        {
            if (seatCanvas.ActualWidth <= 0 || seatCanvas.ActualHeight <= 0)
            {
                txtStatus.Text = "无法导出图片：座位画布尚未渲染或尺寸为零。";
                return;
            }

            var dialog = new SaveFileDialog { Filter = "PNG Image|*.png", Title = "导出图片", FileName = "seat_canvas.png" };
            if (dialog.ShowDialog() != true) return;

            try
            {
                var renderBitmap = new RenderTargetBitmap((int)seatCanvas.ActualWidth, (int)seatCanvas.ActualHeight, 96, 96, PixelFormats.Pbgra32);
                renderBitmap.Render(seatCanvas);

                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(renderBitmap));

                using var fileStream = new FileStream(dialog.FileName, FileMode.Create, FileAccess.Write);
                encoder.Save(fileStream);
                txtStatus.Text = "已成功导出 PNG 图像。";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出图片失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region 智能排座
        private async void BtnRunArrangement_Click(object sender, RoutedEventArgs e)
        {
            var movableStudents = new List<Student>();
            var movablePositions = new List<(int row, int col)>();
            var fixedAssignments = new Dictionary<string, (int row, int col)>();
            var allAssignedStudents = new HashSet<string>();

            for (int r = 0; r < _rows; r++)
            {
                for (int c = 0; c < _cols; c++)
                {
                    var seat = _seats[r, c];
                    if (seat.IsDisabled) continue;

                    if (seat.IsFixed && seat.Student != null)
                    {
                        fixedAssignments[seat.Student.Name] = (r, c);
                        allAssignedStudents.Add(seat.Student.Name);
                    }
                    else
                    {
                        movablePositions.Add((r, c));
                        if (seat.Student != null)
                        {
                            movableStudents.Add(seat.Student);
                            allAssignedStudents.Add(seat.Student.Name);
                        }
                    }
                }
            }

            movableStudents.AddRange(_students.Where(s => !allAssignedStudents.Contains(s.Name)));

            if (movableStudents.Count == 0)
            {
                txtStatus.Text = "没有可安排的学生。";
                return;
            }

            if (movablePositions.Count < movableStudents.Count)
            {
                txtStatus.Text = "没有足够的空余或非固定座位来安排所有学生。";
                return;
            }

            modalOverlay.Visibility = Visibility.Visible;
            progressBar.Visibility = Visibility.Visible;
            progressBar.Value = 0;
            txtStatus.Text = "正在优化座位，请稍候...";

            try
            {
                var bestLayout = await Task.Run(() => PerformSimulatedAnnealing(movableStudents, movablePositions, fixedAssignments));

                for (int r = 0; r < _rows; r++)
                    for (int c = 0; c < _cols; c++)
                        if (!_seats[r, c].IsFixed) _seats[r, c].Student = null;

                foreach (var kvp in bestLayout)
                    if (kvp.Key.row < _rows && kvp.Key.col < _cols)
                        _seats[kvp.Key.row, kvp.Key.col].Student = kvp.Value;

                DrawSeatCanvas();
                txtStatus.Text = "智能排座完成！";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"排座时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                txtStatus.Text = "排座失败。";
            }
            finally
            {
                progressBar.Visibility = Visibility.Collapsed;
                modalOverlay.Visibility = Visibility.Collapsed;
            }
        }

        private Dictionary<(int row, int col), Student> PerformSimulatedAnnealing(List<Student> movableStudents, List<(int, int)> movablePositions, Dictionary<string, (int, int)> fixedAssignments)
        {
            var studentMap = _students.ToDictionary(s => s.Name, s => s);
            var currentLayout = new Dictionary<(int, int), Student>();
            var shuffledStudents = movableStudents.OrderBy(_ => _random.Next()).ToList();
            for (int i = 0; i < shuffledStudents.Count; i++)
            {
                currentLayout[movablePositions[i]] = shuffledStudents[i];
            }

            var bestLayout = new Dictionary<(int, int), Student>(currentLayout);
            var fullLayout = new Dictionary<(int, int), Student>(currentLayout);
            foreach (var kvp in fixedAssignments)
                if (studentMap.TryGetValue(kvp.Key, out var student))
                    fullLayout[kvp.Value] = student;

            double currentCost = CalculateCost(fullLayout);
            double bestCost = currentCost;
            double temperature = 10000.0;
            double coolingRate = 0.9995;
            int maxIterations = Math.Max(20000, movableStudents.Count * 2000);

            var movablePosList = currentLayout.Keys.ToList();

            for (int iteration = 0; iteration < maxIterations; iteration++)
            {
                var candidateLayout = new Dictionary<(int, int), Student>(currentLayout);
                if (movablePosList.Count >= 2)
                {
                    var pos1Index = _random.Next(movablePosList.Count);
                    var pos2Index = _random.Next(movablePosList.Count);

                    if (pos1Index != pos2Index)
                    {
                        var pos1 = movablePosList[pos1Index];
                        var pos2 = movablePosList[pos2Index];
                        (candidateLayout[pos1], candidateLayout[pos2]) = (candidateLayout[pos2], candidateLayout[pos1]);
                    }
                }

                var candidateFullLayout = new Dictionary<(int, int), Student>(candidateLayout);
                foreach (var kvp in fixedAssignments)
                    if (studentMap.TryGetValue(kvp.Key, out var student))
                        candidateFullLayout[kvp.Value] = student;

                double candidateCost = CalculateCost(candidateFullLayout);

                if (candidateCost < currentCost || Math.Exp((currentCost - candidateCost) / temperature) > _random.NextDouble())
                {
                    currentLayout = candidateLayout;
                    currentCost = candidateCost;
                }

                if (currentCost < bestCost)
                {
                    bestLayout = new Dictionary<(int, int), Student>(currentLayout);
                    bestCost = currentCost;
                }

                temperature *= coolingRate;
                if (iteration % (maxIterations / 100) == 0)
                {
                    Dispatcher.Invoke(() => progressBar.Value = (double)iteration / maxIterations * 100);
                }
            }
            Debug.WriteLine($"--- 排座结束 --- 最终最优成本: {bestCost:F2}");
            return bestLayout;
        }

        private (int minRow, int minCol, int maxRow, int maxCol)? ParsePreferredArea(string area)
        {
            if (string.IsNullOrWhiteSpace(area)) return null;
            try
            {
                var parts = area.Split('-', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 1)
                {
                    var coords = parts[0].Split(',', StringSplitOptions.RemoveEmptyEntries);
                    if (coords.Length == 2 && int.TryParse(coords[0].Trim(), out var r) && int.TryParse(coords[1].Trim(), out var c))
                        return (r - 1, c - 1, r - 1, c - 1);
                }
                else if (parts.Length == 2)
                {
                    var start = parts[0].Split(',', StringSplitOptions.RemoveEmptyEntries);
                    var end = parts[1].Split(',', StringSplitOptions.RemoveEmptyEntries);
                    if (start.Length == 2 && end.Length == 2 && int.TryParse(start[0].Trim(), out var r1) && int.TryParse(start[1].Trim(), out var c1) && int.TryParse(end[0].Trim(), out var r2) && int.TryParse(end[1].Trim(), out var c2))
                        return (Math.Min(r1, r2) - 1, Math.Min(c1, c2) - 1, Math.Max(r1, r2) - 1, Math.Max(c1, c2) - 1);
                }
            }
            catch { /* ignored */ }
            return null;
        }

        private double CalculateCost(Dictionary<(int row, int col), Student> layout)
        {
            const double W_AVOID_ADJACENT = 500.0;
            const double W_AVOID_NEAR = 100.0;
            const double W_PREFER_DISTANCE = 10.0;
            const double W_HEIGHT = 50.0;
            const double W_IMPORTANCE = 60.0;
            const double W_AREA = 150.0;

            double total = 0;
            var positionByStudent = layout.ToDictionary(pair => pair.Value.Name, pair => pair.Key);

            foreach (var pair in layout)
            {
                var (row, col) = pair.Key;
                var student = pair.Value;

                double rowFactor = _rows > 1 ? (double)row / (_rows - 1) : 0;
                double reverseRowFactor = _rows > 1 ? (double)(_rows - 1 - row) / (_rows - 1) : 0;

                total += student.HeightWeight > 0 ? student.HeightWeight * reverseRowFactor * W_HEIGHT : Math.Abs(student.HeightWeight) * rowFactor * W_HEIGHT;
                if (student.ImportanceWeight > 0) total += student.ImportanceWeight * rowFactor * W_IMPORTANCE;

                foreach (var avoidName in student.AvoidStudents)
                {
                    if (positionByStudent.TryGetValue(avoidName, out var otherPos))
                    {
                        int dist = Math.Abs(row - otherPos.row) + Math.Abs(col - otherPos.col);
                        if (dist <= 1) total += W_AVOID_ADJACENT;
                        else if (dist == 2) total += W_AVOID_NEAR;
                    }
                }

                foreach (var preferName in student.PreferStudents)
                {
                    if (positionByStudent.TryGetValue(preferName, out var otherPos))
                    {
                        int dist = Math.Abs(row - otherPos.row) + Math.Abs(col - otherPos.col);
                        total += dist * dist * W_PREFER_DISTANCE;
                    }
                }

                if (!string.IsNullOrEmpty(student.PreferArea))
                {
                    var area = ParsePreferredArea(student.PreferArea);
                    if (area.HasValue && (row < area.Value.minRow || row > area.Value.maxRow || col < area.Value.minCol || col > area.Value.maxCol))
                    {
                        total += W_AREA;
                    }
                }
            }
            return total;
        }
        #endregion
    }

    // ===============================================================
    // 辅助命令类 (Helper Command Class)
    // ===============================================================
    internal sealed class RelayCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Predicate<object?>? _canExecute;

        public RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }
        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
        public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object? parameter) => _execute(parameter);
    }
}
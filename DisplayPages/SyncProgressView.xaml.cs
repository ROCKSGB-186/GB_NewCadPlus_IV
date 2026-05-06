using System;
using System.ComponentModel;
using System.Windows;

namespace GB_NewCadPlus_IV
{
    public partial class SyncProgressView : System.Windows.Controls.UserControl, INotifyPropertyChanged
    {
        private readonly DateTime _startedAtUtc = DateTime.UtcNow;
        private string _windowTitle = "正在复制文件到本地";
        private string _statusLine = "请稍候...";
        private string _stageText = "正在准备复制...";
        private string _currentItemText = string.Empty;
        private string _progressValueText = "0%";
        private string _copyCountText = "0 / 0";
        private string _remainingTimeText = "计算中...";
        private string _detailsText = "等待开始...";
        private string _cancelButtonText = "取消同步";
        private bool _isProgressIndeterminate;
        private double _progressValue;
        private bool _isDetailsExpanded;
        private bool _isCancelButtonEnabled = true;

        public SyncProgressView()
        {
            InitializeComponent();
            DataContext = this;
        }

        public event EventHandler? CancelRequested;
        public event PropertyChangedEventHandler? PropertyChanged;

        public string WindowTitle
        {
            get => _windowTitle;
            private set => SetField(ref _windowTitle, value, nameof(WindowTitle));
        }

        public string StatusLine
        {
            get => _statusLine;
            private set => SetField(ref _statusLine, value, nameof(StatusLine));
        }

        public string StageText
        {
            get => _stageText;
            private set => SetField(ref _stageText, value, nameof(StageText));
        }

        public string CurrentItemText
        {
            get => _currentItemText;
            private set => SetField(ref _currentItemText, value, nameof(CurrentItemText));
        }

        public string ProgressValueText
        {
            get => _progressValueText;
            private set => SetField(ref _progressValueText, value, nameof(ProgressValueText));
        }

        public string CopyCountText
        {
            get => _copyCountText;
            private set => SetField(ref _copyCountText, value, nameof(CopyCountText));
        }

        public string RemainingTimeText
        {
            get => _remainingTimeText;
            private set => SetField(ref _remainingTimeText, value, nameof(RemainingTimeText));
        }

        public string DetailsText
        {
            get => _detailsText;
            private set => SetField(ref _detailsText, value, nameof(DetailsText));
        }

        public string CancelButtonText
        {
            get => _cancelButtonText;
            private set => SetField(ref _cancelButtonText, value, nameof(CancelButtonText));
        }

        public bool IsProgressIndeterminate
        {
            get => _isProgressIndeterminate;
            private set => SetField(ref _isProgressIndeterminate, value, nameof(IsProgressIndeterminate));
        }

        public double ProgressValue
        {
            get => _progressValue;
            private set => SetField(ref _progressValue, value, nameof(ProgressValue));
        }

        public bool IsDetailsExpanded
        {
            get => _isDetailsExpanded;
            set => SetField(ref _isDetailsExpanded, value, nameof(IsDetailsExpanded));
        }

        public bool IsCancelButtonEnabled
        {
            get => _isCancelButtonEnabled;
            private set => SetField(ref _isCancelButtonEnabled, value, nameof(IsCancelButtonEnabled));
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelButtonEnabled = false;
            CancelButtonText = "正在取消...";
            CancelRequested?.Invoke(this, EventArgs.Empty);
        }

        public void UpdateProgress(SyncProgressInfo progress)
        {
            if (progress == null)
            {
                return;
            }

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(() => UpdateProgress(progress)));
                return;
            }

            StageText = string.IsNullOrWhiteSpace(progress.Stage) ? "正在同步..." : progress.Stage;
            CurrentItemText = string.IsNullOrWhiteSpace(progress.CurrentItem) ? string.Empty : progress.CurrentItem;
            ProgressValueText = progress.IsIndeterminate ? "..." : $"{progress.Percent}% ({progress.CompletedOperations}/{progress.TotalOperations})";
            CopyCountText = $"{progress.CompletedOperations} / {progress.TotalOperations}";
            RemainingTimeText = GetRemainingTimeText(progress);
            DetailsText = BuildDetailsText(progress, RemainingTimeText);
            StatusLine = progress.IsIndeterminate ? "正在分析可复制文件..." : "正在复制源文件到本地目录...";

            if (progress.IsIndeterminate)
            {
                IsProgressIndeterminate = true;
                return;
            }

            IsProgressIndeterminate = false;
            ProgressValue = progress.Percent;
        }

        private string GetRemainingTimeText(SyncProgressInfo progress)
        {
            if (progress.IsIndeterminate || progress.TotalOperations <= 0 || progress.CompletedOperations <= 0)
            {
                return "计算中...";
            }

            var percent = progress.Percent;
            if (percent <= 0 || percent >= 100)
            {
                return percent >= 100 ? "0 秒" : "计算中...";
            }

            var elapsed = DateTime.UtcNow - _startedAtUtc;
            var estimatedTotal = TimeSpan.FromMilliseconds(elapsed.TotalMilliseconds * 100.0 / percent);
            var remaining = estimatedTotal - elapsed;
            if (remaining < TimeSpan.Zero)
            {
                remaining = TimeSpan.Zero;
            }

            return FormatDuration(remaining);
        }

        private static string BuildDetailsText(SyncProgressInfo progress, string remainingText)
        {
            return string.Join(Environment.NewLine, new[]
            {
                $"阶段：{(string.IsNullOrWhiteSpace(progress.Stage) ? "正在同步" : progress.Stage)}",
                $"阶段说明：{(string.IsNullOrWhiteSpace(progress.StageDetail) ? "-" : progress.StageDetail)}",
                $"当前对象：{(string.IsNullOrWhiteSpace(progress.CurrentItem) ? "-" : progress.CurrentItem)}",
                $"已复制：{progress.CompletedOperations} / {progress.TotalOperations}",
                $"进度：{(progress.IsIndeterminate ? "..." : $"{progress.Percent}%")}",
                $"剩余时间：{remainingText}"
            });
        }

        private static string FormatDuration(TimeSpan duration)
        {
            if (duration < TimeSpan.Zero)
            {
                duration = TimeSpan.Zero;
            }

            if (duration.TotalHours >= 1)
            {
                return $"{(int)duration.TotalHours} 小时 {duration.Minutes} 分 {duration.Seconds} 秒";
            }

            if (duration.TotalMinutes >= 1)
            {
                return $"{duration.Minutes} 分 {duration.Seconds} 秒";
            }

            return $"{Math.Max(0, (int)Math.Ceiling(duration.TotalSeconds))} 秒";
        }

        private bool SetField<T>(ref T field, T value, string propertyName)
        {
            if (Equals(field, value))
            {
                return false;
            }

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            return true;
        }
    }
}

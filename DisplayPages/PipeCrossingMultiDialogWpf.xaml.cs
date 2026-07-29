using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Interop;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace GB_NewCadPlus_IV.DisplayPages
{
    /// <summary>
    /// 管道交叉点多选对话框
    /// </summary>
    public partial class PipeCrossingMultiDialogWpf : Window
    {
        /// <summary>
        /// 交叉点行为：Connect：连接, Cover：覆盖, Under：下层
        /// </summary>
        public enum CrossingAction { Connect, Cover, Under }
        /// <summary>
        /// 交叉点选项视图模型
        /// </summary>
        public class CrossingOptionViewModel : INotifyPropertyChanged
        {
            /// <summary>
            /// 交叉点的行为
            /// </summary>
            private CrossingAction _selectedAction = CrossingAction.Connect;
            /// <summary>
            /// 交叉点索引
            /// </summary>
            public int Index { get; set; }
            /// <summary>
            /// 交叉点
            /// </summary>
            public Point3d Intersection { get; set; }
            /// <summary>
            /// 管道名称
            /// </summary>
            public string PipeName { get; set; }
            /// <summary>
            /// 旧管道的 ID
            /// </summary>
            public ObjectId OldPipeId { get; set; }
            /// <summary>
            /// 新管道的 ID
            /// </summary>
            public ObjectId NewPipeId { get; set; }

            /// <summary>
            /// 交叉点的显示字符串
            /// </summary>
            public string IntersectionDisplay => $"({Intersection.X:F2}, {Intersection.Y:F2})";

            /// <summary>
            /// 获取或设置用户选择的行为
            /// </summary>
            public CrossingAction SelectedAction
            {
                get => _selectedAction;
                set
                {
                    if (_selectedAction == value) return;
                    _selectedAction = value;
                    // 添加临时日志，可输出到 AutoCAD 命令行或弹出消息框
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                        $"\n[调试] 用户选择了：{value}");
                    OnPropertyChanged(nameof(SelectedAction));//选中行为已更改
                    OnPropertyChanged(nameof(IsConnect));//是否连接
                    OnPropertyChanged(nameof(IsCover));//是否覆盖
                    OnPropertyChanged(nameof(IsUnder));//是否下层
                }
            }
            /// <summary>
            /// 是否是连接
            /// </summary>
            public bool IsConnect
            {
                /// <summary>
                /// 获取或设置是否是连接
                /// </summary>
                get => SelectedAction == CrossingAction.Connect;
                /// <summary>
                /// 设置是否是连接
                /// </summary>
                set { if (value) SelectedAction = CrossingAction.Connect; }
            }
            /// <summary>
            /// 是否是覆盖
            /// </summary>
            public bool IsCover
            {
                ///获取或设置是否是覆盖
                get => SelectedAction == CrossingAction.Cover;
                ///设置是否是覆盖
                set { if (value) SelectedAction = CrossingAction.Cover; }
            }
            /// <summary>
            /// 是否是下层
            /// </summary>
            public bool IsUnder
            {
                ///获取或设置是否是下层
                get => SelectedAction == CrossingAction.Under;
                ///设置是否是下层
                set { if (value) SelectedAction = CrossingAction.Under; }
            }
            /// <summary>
            /// 属性更改事件
            /// </summary>
            public event PropertyChangedEventHandler PropertyChanged;
            /// <summary>
            /// 属性已更改
            /// </summary>
            /// <param name="name"></param>
            protected void OnPropertyChanged([CallerMemberName] string name = null) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        /// <summary>
        /// 交叉点选项
        /// </summary>
        public List<CrossingOptionViewModel> Options { get; }
        /// <summary>
        /// 是否取消
        /// </summary>
        public bool IsCancelled { get; private set; } = true;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="options"></param>
        public PipeCrossingMultiDialogWpf(List<CrossingOptionViewModel> options)
        {
            InitializeComponent();
            Options = options ?? new List<CrossingOptionViewModel>();
            for (int i = 0; i < Options.Count; i++)
                Options[i].Index = i + 1;
            DataContext = this;
        }
        /// <summary>
        /// 显示对话框
        /// </summary>
        /// <returns></returns>
        public bool? ShowDialogWithOwner()
        {
            IntPtr ownerHandle = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Handle;
            new WindowInteropHelper(this) { Owner = ownerHandle };
            return ShowDialog();
        }
        /// <summary>
        /// 点击确定按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = false;
            DialogResult = true;
            Close();
        }
        /// <summary>
        /// 点击取消按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            DialogResult = false;
            Close();
        }
               
    }
}

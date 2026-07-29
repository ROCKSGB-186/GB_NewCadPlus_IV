using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Interop;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;

namespace GB_NewCadPlus_IV.DisplayPages
{
    /// <summary>
    /// WPF 多点管道交叉处理对话框，集成遮罩创建与绘图次序调整
    /// </summary>
    public partial class PipeCrossingMultiDialogWpf : Window
    {
        public enum CrossingAction { Connect, Cover, Under }

        /// <summary>
        /// 单个交叉点的 ViewModel
        /// </summary>
        public class CrossingOptionViewModel : INotifyPropertyChanged
        {
            private CrossingAction _selectedAction = CrossingAction.Connect;

            public int Index { get; set; }
            public Point3d Intersection { get; set; }
            public string PipeName { get; set; }
            public ObjectId OldPipeId { get; set; }   // 已有管道的 ObjectId
            public ObjectId NewPipeId { get; set; }   // 新管道的 ObjectId

            public CrossingAction SelectedAction
            {
                get => _selectedAction;
                set
                {
                    _selectedAction = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsConnect));
                    OnPropertyChanged(nameof(IsCover));
                    OnPropertyChanged(nameof(IsUnder));
                }
            }

            public bool IsConnect { get => SelectedAction == CrossingAction.Connect; set { if (value) SelectedAction = CrossingAction.Connect; } }
            public bool IsCover { get => SelectedAction == CrossingAction.Cover; set { if (value) SelectedAction = CrossingAction.Cover; } }
            public bool IsUnder { get => SelectedAction == CrossingAction.Under; set { if (value) SelectedAction = CrossingAction.Under; } }

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged([CallerMemberName] string name = null) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // 依赖资源
        private readonly UnifiedTableGenerator _generator;
        private readonly Transaction _tr;
        private readonly Database _db;
        private readonly string _layer;

        public List<CrossingOptionViewModel> Options { get; }
        public bool IsCancelled { get; private set; } = true;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="generator">UnifiedTableGenerator 实例，用于调用遮罩方法</param>
        /// <param name="tr">当前活动事务</param>
        /// <param name="db">数据库</param>
        /// <param name="layer">遮罩图层（通常来自样例管道）</param>
        /// <param name="options">交叉点选项列表（需预填充 Intersection, PipeName, OldPipeId, NewPipeId）</param>
        public PipeCrossingMultiDialogWpf(
            UnifiedTableGenerator generator,
            Transaction tr,
            Database db,
            string layer,
            List<CrossingOptionViewModel> options)
        {
            InitializeComponent();
            _generator = generator;
            _tr = tr;
            _db = db;
            _layer = layer;
            Options = options ?? new List<CrossingOptionViewModel>();

            // 自动设置序号
            for (int i = 0; i < Options.Count; i++)
                Options[i].Index = i + 1;

            DataContext = this;
        }

        /// <summary>
        /// 以 AutoCAD 主窗口为所有者显示对话框
        /// </summary>
        public bool? ShowDialogWithOwner()
        {
            IntPtr ownerHandle = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Handle;
            new WindowInteropHelper(this) { Owner = ownerHandle };
            return ShowDialog();
        }

        /// <summary>
        /// 确定按钮：执行每个交叉点的遮罩操作
        /// </summary>
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var opt in Options)
            {
                if (opt.SelectedAction == CrossingAction.Connect)
                    continue;   // 相连，无需遮罩

                ObjectId entityBelow, entityAbove;
                if (opt.SelectedAction == CrossingAction.Cover)
                {
                    entityBelow = opt.OldPipeId;   // 旧管道在下
                    entityAbove = opt.NewPipeId;   // 新管道在上
                }
                else // Under
                {
                    entityBelow = opt.NewPipeId;   // 新管道在下
                    entityAbove = opt.OldPipeId;   // 旧管道在上
                }

                // 创建遮罩（Hatch 纯色填充，无边框）
                var maskId = _generator.CreateBackgroundMask(
                    opt.Intersection,
                    2.0 * AutoCadHelper.GetScale(),   // 根据当前比例缩放
                    _tr,
                    _db,
                    _layer);

                // 调整绘图次序：entityBelow → mask → entityAbove
                _generator.SetDrawOrderBetween(_tr, _db, entityBelow, maskId, entityAbove);
            }

            IsCancelled = false;
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// 取消按钮
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            DialogResult = false;
            Close();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Font = System.Drawing.Font;
using FlowDirection = System.Windows.Forms.FlowDirection;


namespace GB_NewCadPlus_IV.DisplayPages
{
    /// <summary>
    /// 多点管道交叉处理对话框
    /// </summary>
    public class PipeCrossingMultiDialog : Form
    {
        public class CrossingOption
        {
            public Point3d Intersection { get; set; }
            public string PipeName { get; set; }
            public ObjectId PipeId { get; set; }
            public CrossingAction SelectedAction { get; set; } = CrossingAction.Connect;
        }

        public enum CrossingAction
        {
            Connect,    // 交叉相连（不做遮罩）
            Cover,      // 覆盖原管线（新管道在上）
            Under       // 在原管线之下（新管道在下）
        }

        private List<CrossingOption> _options;
        private TableLayoutPanel _tableLayout;
        private Button _btnOK;
        private Button _btnCancel;

        public List<CrossingOption> GetOptions() => _options;

        public PipeCrossingMultiDialog(List<CrossingOption> options)
        {
            _options = options;
            InitializeComponents();
            PopulateTable();
        }

        private void InitializeComponents()
        {
            this.Text = "管道交叉处理";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ClientSize = new Size(650, 450);
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblTitle = new Label
            {
                Text = "检测到以下交叉点，请为每个交叉点选择处理方式：",
                Location = new Point(10, 10),
                AutoSize = true,
                Font = new Font("微软雅黑", 9F, FontStyle.Bold)
            };
            this.Controls.Add(lblTitle);

            _tableLayout = new TableLayoutPanel
            {
                Location = new Point(10, 35),
                Size = new Size(620, 330),
                AutoScroll = true,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                ColumnCount = 4,
                RowCount = 1
            };
            _tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30)); // 坐标
            _tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // 管道名称
            _tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15)); // 操作提示
            _tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30)); // 单选按钮

            this.Controls.Add(_tableLayout);

            _btnOK = new Button { Text = "确定", Location = new Point(460, 380), Width = 80, DialogResult = DialogResult.OK };
            _btnCancel = new Button { Text = "取消", Location = new Point(550, 380), Width = 80, DialogResult = DialogResult.Cancel };
            this.Controls.Add(_btnOK);
            this.Controls.Add(_btnCancel);
            this.AcceptButton = _btnOK;
            this.CancelButton = _btnCancel;
        }

        private void PopulateTable()
        {
            // 表头
            _tableLayout.Controls.Add(new Label { Text = "交叉点坐标", Font = new Font("微软雅黑", 9, FontStyle.Bold), Anchor = AnchorStyles.None }, 0, 0);
            _tableLayout.Controls.Add(new Label { Text = "管道名称", Font = new Font("微软雅黑", 9, FontStyle.Bold), Anchor = AnchorStyles.None }, 1, 0);
            _tableLayout.Controls.Add(new Label { Text = "操作选择", Font = new Font("微软雅黑", 9, FontStyle.Bold), Anchor = AnchorStyles.None }, 2, 0);
            _tableLayout.Controls.Add(new Label { Text = "", Font = new Font("微软雅黑", 9, FontStyle.Bold), Anchor = AnchorStyles.None }, 3, 0);

            for (int i = 0; i < _options.Count; i++)
            {
                var opt = _options[i];
                int row = i + 1;
                _tableLayout.RowCount = row + 1;
                _tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));

                // 坐标
                Label coordLabel = new Label
                {
                    Text = $"({opt.Intersection.X:F2}, {opt.Intersection.Y:F2})",
                    Anchor = AnchorStyles.Left,
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Width = 160
                };
                _tableLayout.Controls.Add(coordLabel, 0, row);

                // 管道名称
                Label nameLabel = new Label
                {
                    Text = opt.PipeName,
                    Anchor = AnchorStyles.Left,
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Width = 130
                };
                _tableLayout.Controls.Add(nameLabel, 1, row);

                // 操作提示
                Label actionHint = new Label
                {
                    Text = "操作：",
                    Anchor = AnchorStyles.Right,
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleRight,
                    Width = 60
                };
                _tableLayout.Controls.Add(actionHint, 2, row);

                // 单选按钮组
                FlowLayoutPanel panel = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.LeftToRight,
                    Anchor = AnchorStyles.Left,
                    AutoSize = true
                };
                RadioButton rbConnect = new RadioButton { Text = "相连", Tag = CrossingAction.Connect, AutoSize = true };
                RadioButton rbCover = new RadioButton { Text = "覆盖", Tag = CrossingAction.Cover, AutoSize = true };
                RadioButton rbUnder = new RadioButton { Text = "下方", Tag = CrossingAction.Under, AutoSize = true };
                rbConnect.Checked = true; // 默认相连

                rbConnect.CheckedChanged += (s, e) => { if (rbConnect.Checked) opt.SelectedAction = CrossingAction.Connect; };
                rbCover.CheckedChanged += (s, e) => { if (rbCover.Checked) opt.SelectedAction = CrossingAction.Cover; };
                rbUnder.CheckedChanged += (s, e) => { if (rbUnder.Checked) opt.SelectedAction = CrossingAction.Under; };

                panel.Controls.Add(rbConnect);
                panel.Controls.Add(rbCover);
                panel.Controls.Add(rbUnder);
                _tableLayout.Controls.Add(panel, 3, row);
            }
        }
    }
}

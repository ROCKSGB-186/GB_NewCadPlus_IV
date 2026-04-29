using Autodesk.AutoCAD.Windows;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;
using System.Globalization;
using System.Windows;
using System.Windows.Forms.VisualStyles;

namespace GB_NewCadPlus_IV
{
    public partial class FormMain : Form
    {

        //private bool isTabPageVisible = true; // 联动TabPage的可见状态  DDimLinear
        /// <summary>
        /// 主程序入口 
        /// </summary>
        public FormMain()
        {
            InitializeComponent();
            // 窗口初始化时固定高度
            this.groupBox_工艺.Height = 535;
            //设置面板为透明色；不加这行，容易报异常；
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            // 绑定只允许整数输入的事件
            this.textBox_Scale_比例.KeyPress += textBox_Scale_比例_KeyPress;
            this.textBox_Scale_比例.TextChanged += textBox_Scale_比例_TextChanged;
            this.textBox_Scale_比例.Leave += textBox_Scale_比例_Leave;
            //记录运行时间
            textBox_cmdShow.Text = DateTime.Now.ToString();
            // 注册到统一管理器
            UnifiedUIManager.SetWinFormInstance(this);
            //初始化图层
            NewTjLayer();
            // 添加占位符
            InitializePlaceholders();
            // 修正：将特定逻辑的初始化提取出来，单独调用一次
            InitializeTwoStoreyLogic();

            // 添加可切换的 RadioButton（递归绑定点击事件）
            AttachToggleableRadioButtons(this);
        }
        /// <summary>
        /// 添加占位符
        /// </summary>
        private void InitializePlaceholders()
        {
            AttachPlaceholder(textBox_荷载数据, "输入荷载数据");
            // 示例：对其他 TextBox 启用占位符
            AttachPlaceholder(textBox_排水沟_深, "请输入深");
            AttachPlaceholder(textBox_排水沟_宽, "请输入宽");
            AttachPlaceholder(textBox_排风百分比, "排风百分比");
            AttachPlaceholder(textBox_inputKW, "请输入功率");
            AttachPlaceholder(textBoxA_上下开洞, "请输入功率");
            AttachPlaceholder(textBoxA_左右开洞, "请输入功率");
            AttachPlaceholder(textBoxE_上下开洞, "请输入功率");
            AttachPlaceholder(textBoxE_左右开洞, "请输入功率");
            AttachPlaceholder(textBoxN_上下开洞, "请输入功率");
            AttachPlaceholder(textBoxN_左右开洞, "请输入功率");
            AttachPlaceholder(textBoxP_上下开洞, "请输入功率");
            AttachPlaceholder(textBoxP_左右开洞, "请输入功率");
            AttachPlaceholder(textBoxS_上下开洞, "请输入功率");
            AttachPlaceholder(textBoxS_左右开洞, "请输入功率");
            AttachPlaceholder(textBoxZ_上下开洞, "请输入功率");
            AttachPlaceholder(textBoxZ_左右开洞, "请输入功率");
            AttachPlaceholder(textBox_传递窗_长, "请输入功率");
            AttachPlaceholder(textBox_传递窗_宽, "请输入功率");
            AttachPlaceholder(textBox_传递窗_高, "请输入功率");
            AttachPlaceholder(textBox_设备部位号_1, "请输入功率");
            AttachPlaceholder(textBox_设备部位号_2, "请输入功率");
            AttachPlaceholder(textBoxA_左右开洞, "请输入功率");
            AttachPlaceholder(textBoxA_上下开洞, "请输入功率");
            AttachPlaceholder(textBoxA_左右开洞, "请输入功率");


        }

        // 2. 新增初始化方法 (将原 AttachToggleableRadioButtons 中的头部逻辑移到这里)
        private void InitializeTwoStoreyLogic()
        {
            // 1. 先解绑可能存在的事件，防止重复
            radioButton_是否二层.CheckedChanged -= RadioButton_是否二层_CheckedChanged;

            // 2. 强制初始化状态为【未选中】
            radioButton_是否二层.Checked = false;

            // 3. 重新绑定事件
            radioButton_是否二层.CheckedChanged += RadioButton_是否二层_CheckedChanged;

            // 4. 手动调用一次处理逻辑，确保groupBox_二层开门方向内的控件状态与当前（未选中）一致
            //    此时会把groupBox内的按钮都设为 Disable (灰色)
            RadioButton_是否二层_CheckedChanged(radioButton_是否二层, EventArgs.Empty);
        }

        /// <summary>
        /// 添加可切换的 RadioButton 还原为纯粹的递归绑定方法
        /// </summary>
        /// <param name="parent"></param>
        private void AttachToggleableRadioButtons(Control parent)
        {
            // 5. 遍历绑定点击事件（保持原有逻辑，处理点击切换效果）
            foreach (Control c in parent.Controls)
            {
                if (c is RadioButton rb)
                {
                    // 关闭控件自动切换，交由我们处理
                    rb.AutoCheck = false;
                    rb.Click -= ToggleableRadioButton_Click;
                    rb.Click += ToggleableRadioButton_Click;
                }
                if (c.HasChildren)
                    AttachToggleableRadioButtons(c); // 递归查找
            }
        }
        /// <summary>
        /// 添加可切换的 RadioButton
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToggleableRadioButton_Click(object sender, EventArgs e)
        {
            if (sender is not RadioButton rb) return;

            // 如果当前已选中，再次点击则取消选中
            if (rb.Checked)
            {
                rb.Checked = false;
                return;
            }

            // 否则在同一容器中取消其它 RadioButton（保留 GroupBox/Panel 作为分组）
            var parent = rb.Parent;
            // 取消其它 注意：如果有嵌套的 GroupBox/Panel，这里只会取消同一层级的 RadioButton，符合常规分组逻辑
            if (parent != null)
            {
                // 取消其它 取消同一容器内的其它 RadioButton（保留 GroupBox/Panel 作为分组）
                foreach (Control child in parent.Controls)
                {
                    // 如果是 RadioButton 且不是当前点击的那个，则取消选中
                    if (child is RadioButton other && other != rb)
                        // 取消选中
                        other.Checked = false;
                }
            }
            // 选中当前
            rb.Checked = true;
        }

        #region 系统变量

        // 只允许数字键和控制键（退格等）
        private void textBox_Scale_比例_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        // 修改：将 string -> double 的赋值改为安全解析并赋值（保留原有行为）
        private void textBox_Scale_比例_TextChanged(object sender, EventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;

            int selStart = tb.SelectionStart;
            string filtered = System.Text.RegularExpressions.Regex.Replace(tb.Text ?? string.Empty, @"\D+", "");
            if (tb.Text != filtered)
            {
                tb.Text = filtered;
                tb.SelectionStart = Math.Min(selStart, tb.Text.Length);
            }

            // 同步到变量（尝试解析为 double，解析失败使用默认值 1.0）
            if (double.TryParse(tb.Text, out double val))
            {
                VariableDictionary.textBoxScale = val;
            }
            else
            {
                VariableDictionary.textBoxScale = 1.0;
            }
        }

        private void textBox_Scale_比例_Leave(object sender, EventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;

            if (string.IsNullOrWhiteSpace(tb.Text))
            {
                tb.Text = "1";
            }

            if (!int.TryParse(tb.Text, out int ival) || ival < 0)
            {
                //MessageBox.Show("请输入正整数。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tb.Text = "1";
                ival = 1;
            }

            // 以 double 存储
            VariableDictionary.textBoxScale = Convert.ToDouble(ival);
        }
        public static void NewTjLayer()
        {
            while (true)
            {
                foreach (var item in VariableDictionary.GGtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.GYtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.AtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.StjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.PtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.NtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.EtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.ZKtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.tjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.SBtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                break;
            }
        }

        /// <summary>
        /// 单选按键选中键名
        /// </summary>
        public string? checkRadioButtonsText = null;
        /// <summary>
        /// 单选按键选中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RadioButton_是否二层_CheckedChanged(object sender, EventArgs e)
        {
            // 获取RadioButton_是否二层是否选中 获取当前选中状态
            bool enabled = radioButton_是否二层.Checked;
            if (enabled)
            {
                VariableDictionary.radioButton = true;
            }
            else
            {
                VariableDictionary.radioButton = false;
            }
            // 启用或禁用groupBox_二层开门方向 根据选中状态启用或禁用 groupBox_二层开门方向 内的所有控件
            groupBox_二层开门方向.Enabled = enabled;

            // 额外需求：当变为不可选时，取消内部所有 RadioButton 的选中状态
            if (!enabled)
            {
                RecursivelyUncheckRadioButtons(groupBox_二层开门方向);
            }
            //VariableDictionary.radioButton = false;
        }

        // 4. 新增递归取消选中辅助方法
        private void RecursivelyUncheckRadioButtons(Control parent)
        {
            foreach (Control c in parent.Controls)
            {
                if (c is RadioButton rb)
                {
                    rb.Checked = false;
                }
                // 递归处理可能的嵌套容器
                if (c.HasChildren)
                {
                    RecursivelyUncheckRadioButtons(c);
                }
            }
        }
        /// <summary>
        /// 实现面板
        /// </summary>
        public class GB_CadToolsForm
        {
            #region 窗体实现
            /// <summary>
            /// 创建cad里的一个空的窗体
            /// </summary>
            public static PaletteSet? Cad_PaletteSet = null;      //创建cad里的一个空的窗体
            /// <summary>
            /// 初始化一个容器内的工具实体；
            /// </summary>
            static private FormMain GB_ToolsPanel = new FormMain();   //初始化一个容器内的工具实体；
            /// <summary>
            /// 初始化一个容器GB_ToolPanel
            /// </summary>
            public static GB_CadToolsForm GB_ToolsForm = new GB_CadToolsForm(); //初始化一个容器GB_ToolPanel

            /// <summary>
            /// 一个GB_CadToolPanel容器
            /// </summary>
            /// <returns>返回一个单一窗体</returns>
            static public FormMain GB_CadToolPanel()
            {
                return GB_ToolsPanel;//返回一个工具容器
            }


            /// <summary>
            /// 显示工具panel
            /// </summary>
            static public void ShowToolsPanel()
            {
                if (Cad_PaletteSet == null || Cad_PaletteSet.IsDisposed)
                {
                    Cad_PaletteSet = new PaletteSet("图库管理");//初始化这个图库管理窗体；
                    //Cad_PaletteSet.Size = new Size(350, 700);//初始化窗体的大小
                    Cad_PaletteSet.MinimumSize = new System.Drawing.Size(300, 650);//初始化窗体时最小的尺寸
                    //设置为子窗体；
                    GB_CadToolPanel().Anchor =
                       System.Windows.Forms.AnchorStyles.Left |
                       System.Windows.Forms.AnchorStyles.Right |
                       System.Windows.Forms.AnchorStyles.Top;
                    GB_CadToolPanel().Dock = DockStyle.Fill;       //子面板整体覆盖
                    GB_CadToolPanel().TopLevel = false;//子窗体是不是为顶级窗体；
                    GB_CadToolPanel().Location = new System.Drawing.Point(0, 0);//相对位置，左上角，是有设计图纸的区域里；
                    Cad_PaletteSet.Add("屏幕菜单", GB_CadToolPanel());
                }
                Cad_PaletteSet.Visible = true;//显示面板；
                Cad_PaletteSet.Dock = DockSides.Left;//绑定在左侧；
            }
            /// <summary>
            /// 打开的文件列表
            /// </summary>
            private GB_CadToolsForm()
            {
                GetPath.ListDwgFile = new List<string>();
                Load();
            }

            #endregion

            /// <summary>
            /// 拿到本程序的自己路径
            /// </summary>
            /// <returns></returns>
            static private string Path()
            {
                string path = GetPath.GetSelftUserPath();//实例化本程序的自己的路径
                System.IO.Directory.CreateDirectory(path);//在本程序下创建文件夹
                return System.IO.Path.Combine(GetPath.GetSelftUserPath(), "TuKu.txt");//返回本程序的自己路径； 
            }
            /// <summary>
            /// 读取本地设置路径下的配置文件
            /// </summary>
            public void Load()
            {
                string[]? lines = null;
                try
                {
                    lines = System.IO.File.ReadAllLines(Path());//按每一行为一个DWG文件读进来； 
                    GetPath.ListDwgFile.AddRange(lines);//把本程序下添加的文件都显示在列表里；
                }
                catch
                {
                }
            }
            /// <summary>
            /// 保存添加的图库文件与写入配置文件中
            /// </summary>
            public void Save()
            {
                try
                {
                    using (var sr = new StreamWriter(Path())) //useing调用后主动释放文件
                    {
                        foreach (var item in GetPath.ListDwgFile)
                        {
                            sr.WriteLine(item);
                        }
                    }
                }
                catch (System.Exception)
                {
                }
            }
        }

        /// <summary>
        /// 通用占位符：把下面方法加入到 FormMain 类中（替换掉原来的两个专用方法）
        /// </summary>
        /// <param name="tb"></param>
        /// <param name="placeholder"></param>
        private void AttachPlaceholder(TextBox tb, string placeholder)
        {
            if (tb == null) return;
            tb.Tag = placeholder;

            // 初始状态：如果为空则显示占位文本（灰色）
            if (string.IsNullOrEmpty(tb.Text))
            {
                tb.Text = placeholder;
                tb.ForeColor = System.Drawing.SystemColors.GrayText;
            }

            // 解绑以防重复订阅
            tb.MouseDown -= GenericTextBox_MouseDown;
            tb.MouseLeave -= GenericTextBox_MouseLeave;
            tb.GotFocus -= GenericTextBox_GotFocus;
            tb.LostFocus -= GenericTextBox_LostFocus;

            // 绑定通用事件（支持鼠标与键盘焦点两类交互）
            tb.MouseDown += GenericTextBox_MouseDown;
            tb.MouseLeave += GenericTextBox_MouseLeave;
            tb.GotFocus += GenericTextBox_GotFocus;
            tb.LostFocus += GenericTextBox_LostFocus;
        }
        /// <summary>
        /// 鼠标点击或键盘获取焦点时，如果文本框内容为占位符则清空并显示为正常文本
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenericTextBox_MouseDown(object sender, MouseEventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;
            var placeholder = tb.Tag as string;
            if (!string.IsNullOrEmpty(placeholder) && tb.Text == placeholder)
            {
                tb.Clear();
                tb.ForeColor = System.Drawing.SystemColors.ControlText;
            }
        }
        /// <summary>
        /// 鼠标移出文本框时，如果文本框内容为空则显示占位符
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenericTextBox_MouseLeave(object sender, EventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;
            var placeholder = tb.Tag as string;
            if (!string.IsNullOrEmpty(placeholder) && string.IsNullOrWhiteSpace(tb.Text))
            {
                tb.Text = placeholder;
                tb.ForeColor = System.Drawing.SystemColors.GrayText;
            }
        }

        /// <summary>
        /// 获取焦点时，如果文本框内容为占位符则清空并显示为正常文本
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenericTextBox_GotFocus(object sender, EventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;
            var placeholder = tb.Tag as string;
            if (!string.IsNullOrEmpty(placeholder) && tb.Text == placeholder)
            {
                tb.Clear();
                tb.ForeColor = System.Drawing.SystemColors.ControlText;
            }
        }
        /// <summary>
        /// 失去焦点时，如果文本框内容为空则显示占位符
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenericTextBox_LostFocus(object sender, EventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;
            var placeholder = tb.Tag as string;
            if (!string.IsNullOrEmpty(placeholder) && string.IsNullOrWhiteSpace(tb.Text))
            {
                tb.Text = placeholder;
                tb.ForeColor = System.Drawing.SystemColors.GrayText;
            }
        }
        #endregion

        #region 方向按键
        // 1.75 = 顺时针45    
        // 1.5  = 顺时针90    
        // 1.25 = 顺时针135   
        // 1    =       180
        // 0.75 = 逆时针135
        // 0.5  = 逆时针90
        // 0.25 = 逆时针45
        public void button_向下_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("下");
            command?.Invoke();
        }

        public void button_向左下_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("左下");
            command?.Invoke();
        }

        public void button_向左_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("左");
            command?.Invoke();
        }

        public void button_向左上_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("左上");
            command?.Invoke();
        }

        public void button_向上_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("上");
            command?.Invoke();

        }

        public void button_向右上_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("右上");
            command?.Invoke();
        }

        public void button_向右_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("右");
            command?.Invoke();
        }

        public void button_向右下_Click(object sender, EventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("右下");
            command?.Invoke();
        }
        #endregion

        #region 检查专业图元
        /// <summary>
        /// 查检区域填充图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_onOff_QY_Layer_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button_onOff_QY_Layer.ForeColor.Name == "Black" || button_onOff_QY_Layer.ForeColor.Name == "ControlText")
            {
                button_onOff_QY_Layer.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_onOff_QY_Layer.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Add("QY");
            VariableDictionary.selectTjtLayer.Add("qy");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 检查设备图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_SB_onOff_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button_SB_onOff.ForeColor.Name == "Black" || button_SB_onOff.ForeColor.Name == "ControlText")
            {
                button_SB_onOff.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_SB_onOff.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            foreach (var item in VariableDictionary.SBtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 检查工艺条件按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_GY_检查工艺_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button_检查工艺.ForeColor.Name == "Black" || button_检查工艺.ForeColor.Name == "ControlText")
            {
                button_检查工艺.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_检查工艺.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            //Env.Document.SendStringToExecute("CloseAllLayer ", false, false, false);
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 关闭工艺图
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_GY_关闭工艺_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            //VariableDictionary.tjtBtn = VariableDictionary.tjtBtn;
            if (button_关闭工艺.ForeColor.Name == "Black" || button_关闭工艺.ForeColor.Name == "ControlText")
            {
                button_关闭工艺.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭工艺.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                //VariableDictionary.allTjtLayer.Add(item);
                VariableDictionary.selectTjtLayer.Add(item);
            }
            //Env.Document.SendStringToExecute("OpenLayer ", false, false, false);
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 关闭工艺外的图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_GY_保留工艺_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button_保留工艺.ForeColor.Name == "Black" || button_保留工艺.ForeColor.Name == "ControlText")
            {
                button_保留工艺.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_保留工艺.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.allTjtLayer)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 保留建筑，关闭其它专业图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_JZ_保留建筑_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.AtjtBtn;
            if (button_保留建筑.ForeColor.Name == "Black" || button_保留建筑.ForeColor.Name == "ControlText")
            {
                button_保留建筑.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_保留建筑.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }

            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.AtjtBtn)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.allTjtLayer)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }
        /// <summary>
        /// 关闭建筑打开其它专业图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_JZ_关闭建筑_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.AtjtBtn;
            if (button_关闭建筑.ForeColor.Name == "Black" || button_关闭建筑.ForeColor.Name == "ControlText")
            {
                button_关闭建筑.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭建筑.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            foreach (var item in VariableDictionary.AtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            //Env.Document.SendStringToExecute("OpenLayer ", false, false, false);
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 关闭其它
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_JZ_检查建筑_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.AtjtBtn;
            if (button_检查建筑条件.ForeColor.Name == "Black" || button_检查建筑条件.ForeColor.Name == "ControlText")
            {
                button_检查建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_检查建筑条件.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.AtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 保留结构关闭其它专业图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_检查结构_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.StjtBtn;
            if (button_检查结构.ForeColor.Name == "Black" || button_检查结构.ForeColor.Name == "ControlText")
            {
                button_检查结构.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_检查结构.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.StjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 关闭结构图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_关闭结构_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.StjtBtn;
            if (button_关闭结构.ForeColor.Name == "Black" || button_关闭结构.ForeColor.Name == "ControlText")
            {
                button_关闭结构.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭结构.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            foreach (var item in VariableDictionary.StjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 保留结构图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_保留结构_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.StjtBtn;
            if (button_保留结构.ForeColor.Name == "Black" || button_保留结构.ForeColor.Name == "ControlText")
            {
                button_保留结构.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_保留结构.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.StjtBtn)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.allTjtLayer)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }
        /// <summary>
        /// 检查给排水
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_P_检查给排水_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.PtjtBtn;

            if (button_检查给排水.ForeColor.Name == "Black" || button_检查给排水.ForeColor.Name == "ControlText")
            {
                button_检查给排水.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_检查给排水.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.PtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        public void button_P_关闭给排水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.PtjtBtn;
            if (button_关闭给排水.ForeColor.Name == "Black" || button_关闭给排水.ForeColor.Name == "ControlText")
            {
                button_关闭给排水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭给排水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            foreach (var item in VariableDictionary.PtjtBtn)
            {
                //VariableDictionary.allTjtLayer.Add(item);
                VariableDictionary.selectTjtLayer.Add(item);
            }
            //Env.Document.SendStringToExecute("OpenLayer ", false, false, false);
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        public void button_P_保留给排水_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.PtjtBtn;

            if (button_保留给排水.ForeColor.Name == "Black" || button_保留给排水.ForeColor.Name == "ControlText")
            {
                button_保留给排水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_保留给排水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.PtjtBtn)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.allTjtLayer)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
            //Env.Document.SendStringToExecute("CloseLayer ", false, false, false);
        }
        public void button_NT_检查暖通_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.NtjtBtn;

            if (button_检查暖通.ForeColor.Name == "Black" || button_检查暖通.ForeColor.Name == "ControlText")
            {
                button_检查暖通.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_检查暖通.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.NtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }
        public void button_NT_关闭暖通_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.NtjtBtn;
            if (button_关闭暖通.ForeColor.Name == "Black" || button_关闭暖通.ForeColor.Name == "ControlText")
            {
                button_关闭暖通.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭暖通.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            foreach (var item in VariableDictionary.NtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        public void button_NT_保留暖通_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.NtjtBtn;
            if (button_保留暖通.ForeColor.Name == "Black" || button_保留暖通.ForeColor.Name == "ControlText")
            {
                button_保留暖通.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_保留暖通.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.NtjtBtn)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.allTjtLayer)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }
        /// <summary>
        /// 只保留电气，关闭其它所有图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_E_检查电气_Click(object sender, EventArgs e)
        {
            if (button_检查电气.ForeColor.Name == "Black" || button_检查电气.ForeColor.Name == "ControlText")
            {
                button_检查电气.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_检查电气.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.EtjtBtn)
            {
                //if (!VariableDictionary.EtjtBtn.Contains(item))
                VariableDictionary.selectTjtLayer.Add(item);
            }
            //Env.Document.SendStringToExecute("CloseAllLayer ", false, false, false);
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 只关闭电气，保留其它所有图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_E_关闭电气_Click(object sender, EventArgs e)
        {
            //VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.EtjtBtn;
            if (button_关闭电气.ForeColor.Name == "Black" || button_关闭电气.ForeColor.Name == "ControlText")
            {
                button_关闭电气.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭电气.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            foreach (var item in VariableDictionary.EtjtBtn)
            {
                //VariableDictionary.allTjtLayer.Add(item);
                VariableDictionary.selectTjtLayer.Add(item);
            }
            //Env.Document.SendStringToExecute("OpenLayer ", false, false, false);
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        /// <summary>
        /// 只保留电气，关闭其它条件图层，但不是条件的图层不改变状态
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_E_保留电气_Click(object sender, EventArgs e)
        {
            if (!VariableDictionary.btnState && VariableDictionary.tjtBtn != VariableDictionary.tjtBtnNull)
                VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.EtjtBtn;
            if (button_保留电气.ForeColor.Name == "Black" || button_保留电气.ForeColor.Name == "ControlText")
            {
                button_保留电气.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_保留电气.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.EtjtBtn)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.allTjtLayer)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        public void button_ZK_检查自控_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.ZKtjtBtn;

            if (button_检查自控.ForeColor.Name == "Black" || button_检查自控.ForeColor.Name == "ControlText")
            {
                button_检查自控.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_检查自控.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            foreach (var item in VariableDictionary.ZKtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }
        public void button_ZK_关闭自控_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.ZKtjtBtn;
            if (button_关闭自控.ForeColor.Name == "Black" || button_关闭自控.ForeColor.Name == "ControlText")
            {
                button_关闭自控.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭自控.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            foreach (var item in VariableDictionary.ZKtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
                //VariableDictionary.allTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        public void button_ZK_保留自控_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            VariableDictionary.tjtBtn = VariableDictionary.ZKtjtBtn;

            if (button_保留自控.ForeColor.Name == "Black" || button_保留自控.ForeColor.Name == "ControlText")
            {
                button_保留自控.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                VariableDictionary.btnState = true;

            }
            else
            {
                button_保留自控.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ZKtjtBtn)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.allTjtLayer)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        //public void button关闭总图_Click(object sender, EventArgs e)
        //{
        //    for (int i = 0; i < VariableDictionary.tjtBtn .Length; i++)
        //    {
        //        VariableDictionary.tjtBtn [i] = "";
        //    }
        //    VariableDictionary.tjtBtn [0] = "STJ";
        //    Env.Document.SendStringToExecute("OpenLayer ", false, false, false);
        //}

        #endregion

        #region  共用条件
        public void button_共用条件说明_Click(object sender, EventArgs e)
        {
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "共用条件说明";
            VariableDictionary.btnBlockLayer = "TJ(共用条件)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 161;//设置为被插入的图层颜色
            VariableDictionary.layerName = "TJ(共用条件)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
        }

        #endregion
        /// <summary>
        /// 通用处理按键命令
        /// </summary>
        /// <param name="commandName"></param>
        /// <param name="btnFileName"></param>
        /// <param name="btnBlockLayer"></param>
        /// <param name="layerColorIndex"></param>
        /// <param name="rotateAngle"></param>
        private void ExecuteProcessCommand(string commandName, string btnFileName, string btnBlockLayer, int layerColorIndex, double rotateAngle)
        {
            try
            {
                VariableDictionary.entityRotateAngle = rotateAngle;
                VariableDictionary.btnFileName = btnFileName;
                VariableDictionary.btnBlockLayer = btnBlockLayer;//设置为被插入的图层名
                VariableDictionary.layerColorIndex = layerColorIndex;//设置为被插入的图层颜色
                VariableDictionary.layerName = btnBlockLayer;
                VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);

                Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行命令时出错: {ex.Message}");
            }
        }


        #region 工艺
        public void button_GY_房间号_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("房间号", "RWS", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_低温循环上水_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("低温循环上水", "RWS,DN??,??m³/h", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_低压蒸汽_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("低压蒸汽", "LS,DN??,??MPa,??kg/h,", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_二氧化碳_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("二氧化碳", "CO2,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_氮气_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("氮气", "N2,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_氧气_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("氧气", "O2,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_常温循环上水_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("常温循环上水", "CWS,DN??,??m³/h", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_注射用水_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("注射用水", "WFI,DN??,??℃,??L/h,使用量??L/h", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_纯蒸汽_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("纯蒸汽", "LS,DN??,??MPa,??kg/h,", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_纯化水_Click(object sender, EventArgs e)
        {
            //var command = UnifiedCommandManager.GetCommand("纯化水");
            //command?.Invoke();
            ExecuteProcessCommand("纯化水", "PW,DN??,??L/h", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_仪表压缩空气_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("仪表压缩空气", "IA,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_无菌压缩空气_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("无菌压缩空气", "CA,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_热水上水_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("热水上水", "HWS,DN??,??m³/h", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_凝结回水_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("凝结回水", "SC,DN??", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_液体物料_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("液体物料", "PL", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_乙二醇冷却上液_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("乙二醇冷却上液", "EGS,DN??,??m³/h", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_软化水_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("软化水", "SW,DN??,??m³/h", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_真空_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("真空", "VE", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_放空管_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("放空管", "VT", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_氨水_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("氨水", "AW", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_乙醇_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("乙醇", "AH", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_酸液_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("酸液", "AL", "TJ(工艺专业GY)", 40, 0);
        }
        public void button_G_碱液_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("碱液", "SL", "TJ(工艺专业GY)", 40, 0);
        }

        private void button_设备部位号_Click(object sender, EventArgs e)
        {
            // 设置设备位号标注的基础上下文（图层、颜色、比例等）
            VariableDictionary.btnFileName = "TJ(设备位号)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.layerName = "TJ(设备位号)";
            VariableDictionary.layerColorIndex = 241; // 设置为被插入的图层颜色
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);

            // 读取当前主编码（例如 X0001），为空时兜底为 X0001
            string currentMainCode = (textBox_设备部位号_1.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(currentMainCode))
            {
                currentMainCode = "X0001";
                textBox_设备部位号_1.Text = currentMainCode;
            }

            // 读取副编码（例如 xx），用于拼接第二段
            string secondCode = (textBox_设备部位号_2.Text ?? string.Empty).Trim();

            using var tr = new DBTrans();

            // 按现有逻辑创建标注：当副编码是“xx”或为空时，只标注主编码
            if (string.Equals(secondCode, "xx", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(secondCode))
            {
                Command.DDimLinear(tr, currentMainCode, null);
            }
            else
            {
                Command.DDimLinear(tr, currentMainCode, secondCode);
            }

            // 提交事务，确保标注写入成功
            tr.Commit();

            // 标注成功后，主编码自动+1（X0001 -> X0002）
            textBox_设备部位号_1.Text = GetNextDeviceCode(currentMainCode);
        }
        /// <summary>
        /// 计算下一个设备编码：
        /// 例如 X0001 -> X0002，A0099 -> A0100
        /// </summary>
        /// <param name="currentCode">当前编码</param>
        /// <returns>递增后的编码</returns>
        private static string GetNextDeviceCode(string? currentCode)
        {
            // 空值兜底
            if (string.IsNullOrWhiteSpace(currentCode)) return "X0001";

            string code = currentCode.Trim();

            // 规则：前缀字母 + 数字（如 X0001）
            var match = System.Text.RegularExpressions.Regex.Match(code, @"^(?<prefix>[A-Za-z]+)(?<num>\d+)$");
            if (!match.Success)
            {
                // 不符合规则时，回退为默认起始值
                return "X0001";
            }

            // 取前缀与数字部分
            string prefix = match.Groups["prefix"].Value;
            string numText = match.Groups["num"].Value;

            // 数字解析失败则回到前缀+0001
            if (!long.TryParse(numText, out long value))
            {
                return $"{prefix}0001";
            }

            // 数字递增
            long next = value + 1;

            // 保持原有位数，不足补0；若进位导致位数增加，自动扩展
            int width = Math.Max(numText.Length, next.ToString().Length);
            return $"{prefix}{next.ToString(new string('0', width))}";
        }
        public void textBox_设备部位号_1_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox_设备部位号_1.Text == "输入参数")
                textBox_设备部位号_1.Clear();
        }

        public void textBox_设备部位号_1_MouseLeave(object sender, EventArgs e)
        {
            // 失焦时若为空，恢复为默认编码 X0001（避免出现 XXXX）
            if (string.IsNullOrWhiteSpace(textBox_设备部位号_1.Text))
                textBox_设备部位号_1.Text = "X0001";
        }

        public void button工艺开闭工艺条件_Click(object sender, EventArgs e)
        {
            List<string> GGtjtBtn = new List<string>
        {
            "TJ(工艺专业GY)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button工艺开闭工艺条件.ForeColor.Name == "Black" || button工艺开闭工艺条件.ForeColor.Name == "ControlText")
            {
                button工艺开闭工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button工艺开闭工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button工艺开闭建筑条件_Click(object sender, EventArgs e)
        {
            List<string> GAtjtBtn = new List<string>
        {
            "TJ(建筑专业J)Y",
            "TJ(建筑专业J)N",
            "TJ(房间编号)",
            "QY",
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button工艺开闭建筑条件.ForeColor.Name == "Black" || button工艺开闭建筑条件.ForeColor.Name == "ControlText")
            {
                button工艺开闭建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button工艺开闭建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in GAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }

        public void button工艺收工艺条件图层_Click(object sender, EventArgs e)
        {
            List<string> GGtjtBtn = new List<string>
        {
            "TJ(工艺专业GY)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button工艺收工艺条件图层.ForeColor.Name == "Black" || button工艺收工艺条件图层.ForeColor.Name == "ControlText")
            {
                button工艺收工艺条件图层.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button工艺收工艺条件图层.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button工艺收建筑过工艺条件_Click(object sender, EventArgs e)
        {
            List<string> GAtjtBtn = new List<string>
        {
            "TJ(建筑专业J)Y",
            "TJ(建筑专业J)N",
            "TJ(房间编号)",
            "QY",
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button工艺收建筑过工艺条件.ForeColor.Name == "Black" || button工艺收建筑过工艺条件.ForeColor.Name == "ControlText")
            {
                button工艺收建筑过工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button工艺收建筑过工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in GAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button工艺收建筑吊顶高度_Click(object sender, EventArgs e)
        {
            List<string> GAtjtBtn = new List<string>
        {
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button工艺收建筑吊顶高度.ForeColor.Name == "Black" || button工艺收建筑吊顶高度.ForeColor.Name == "ControlText")
            {
                button工艺收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button工艺收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in GAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button工艺收建筑房间编号_Click(object sender, EventArgs e)
        {
            List<string> GAtjtBtn = new List<string>
        {
            "TJ(房间编号)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button工艺收建筑房间编号.ForeColor.Name == "Black" || button工艺收建筑房间编号.ForeColor.Name == "ControlText")
            {
                button工艺收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button工艺收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in GAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button_生成_插入传递窗_Click(object sender, EventArgs e)
        {
            VariableDictionary.layerName = "WALL-PARAPET";
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            // 读取尺寸并验证
            if (!double.TryParse(textBox_传递窗_长.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var coreLen) || coreLen <= 0)
            {
                System.Windows.Forms.MessageBox.Show("请输入有效的传递窗长度。");
                return;
            }
            if (!double.TryParse(textBox_传递窗_宽.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var coreWid) || coreWid <= 0)
            {
                System.Windows.Forms.MessageBox.Show("请输入有效的传递窗宽度。");
                return;
            }
            double.TryParse(textBox_传递窗_高.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var coreHeightVal);

            // 收集被选中的 RadioButton 名称（支持多选）
            List<string> selectedRadioNames = new List<string>();
            // 递归获取所有 RadioButton 递归方法获取所有 RadioButton 控件
            IEnumerable<RadioButton> GetAllRadioButtons(Control parent)
            {
                // 迭代所有子控件 遍历当前控件的子控件
                foreach (Control c in parent.Controls)
                {
                    // 判断当前控件是否 RadioButton 如果是 RadioButton 就返回
                    if (c is RadioButton rb) yield return rb;
                    // 递归处理子控件 如果当前控件有子控件 就递归调用获取子控件中的 RadioButton
                    if (c.HasChildren)
                    {
                        // 递归调用获取子控件中的 RadioButton 递归调用获取子控件中的 RadioButton
                        foreach (var child in GetAllRadioButtons(c))
                            // 返回 返回子控件中的 RadioButton
                            yield return child;
                    }
                }
            }
            // 获取所有 RadioButton 控件 遍历所有 RadioButton 获取被选中的 RadioButton 的名称
            foreach (var rb in GetAllRadioButtons(this))
            {
                if (rb.Checked)// 被选中如果当前 RadioButton 被选中 就把它的名称添加到列表中
                    selectedRadioNames.Add(rb.Name);// 添加被选中的 RadioButton 的名称 添加名称
            }

            // 文字内容
            string typeVal = comboBox_传递窗类型.Text ?? string.Empty;
            string paramVal = comboBox_传递窗参数.Text ?? string.Empty;

            // 调用 Command 层处理（统一管理）
            try
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                // 调用 Command 层处理
                GB_NewCadPlus_IV.FunctionalMethod.Command.GenerateAndInsertTransferWindow(doc, coreLen, coreWid, coreHeightVal, typeVal, paramVal, selectedRadioNames);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"调用生成传递窗命令失败：{ex.Message}");
            }
        }


        #endregion

        #region 工艺暖通

        public void button_NT_排潮_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("排潮", "(排潮)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_排尘_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("排尘", "(排尘)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_排热_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("排热", "(排热)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_直排_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("直排", "(直排)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_除味_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("除味", "(除味)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_A级高度_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("A级高度", "(A级高度？米)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_设备取风量_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("设备取风量", "(设备取风量 ？m³/h)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_设备排风量_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("设备排风量", "(设备排风量 ？m³/h)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_排风百分比_Click(object sender, EventArgs e)
        {
            if (textBox_排风百分比.Text == "排风百分比")
            {
                VariableDictionary.btnFileName = "(排风 ？ %)";
            }
            else
            {
                VariableDictionary.btnFileName = "(排风 " + textBox_排风百分比.Text + " %)";
            }

            Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
            ExecuteProcessCommand("排风百分比", $"{VariableDictionary.btnFileName}", "TJ(暖通专业N)", 6, 0);

        }
        public void button_NT_温度_Click(object sender, EventArgs e)
        {

            ExecuteProcessCommand("温度", "(温度 ？℃±？℃)", "TJ(暖通专业N)", 6, 0);
        }
        public void button_NT_湿度_Click(object sender, EventArgs e)
        {
            ExecuteProcessCommand("湿度", "(湿度 ？%±？%)", "TJ(暖通专业N)", 6, 0);
        }

        private void button_风口_F1_Click(object sender, EventArgs e)
        {
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DUCT_风口_F1";
            VariableDictionary.btnFileName_blockName = "$风口$00000224";
            VariableDictionary.btnBlockLayer = "TJ(暖通专业FFU)";
            VariableDictionary.blockScale = 1;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.DUCT_风口_F1_600x1200_TCH;
            VariableDictionary.textBoxScale = Convert.ToInt32(textBox_Scale_比例.Text) / 100;
            VariableDictionary.winForm_Status = true;
            // 生成临时文件名（放在系统临时目录，带 .dwg 扩展名）
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        private void button_风口_F2_Click(object sender, EventArgs e)
        {
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DUCT_风口_F2";
            VariableDictionary.btnFileName_blockName = "$风口$00000224";
            VariableDictionary.btnBlockLayer = "TJ(暖通专业FFU)";
            VariableDictionary.blockScale = 1;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.DUCT_风口_F2_900x600_TCH;
            VariableDictionary.textBoxScale = Convert.ToInt32(textBox_Scale_比例.Text) / 100;
            VariableDictionary.winForm_Status = true;
            // 生成临时文件名（放在系统临时目录，带 .dwg 扩展名）
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        private void button_风口_F3_Click(object sender, EventArgs e)
        {
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DUCT_风口_F3";
            VariableDictionary.btnFileName_blockName = "$风口$00000224";
            VariableDictionary.btnBlockLayer = "TJ(暖通专业FFU)";
            VariableDictionary.blockScale = 1;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.DUCT_风口_F3_600x600_TCH;
            VariableDictionary.textBoxScale = Convert.ToInt32(textBox_Scale_比例.Text) / 100;
            VariableDictionary.winForm_Status = true;
            // 生成临时文件名（放在系统临时目录，带 .dwg 扩展名）
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        private void button_风口_F4_Click(object sender, EventArgs e)
        {
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DUCT_风口_F4";
            VariableDictionary.btnFileName_blockName = "$风口$00000224";
            VariableDictionary.btnBlockLayer = "TJ(暖通专业FFU)";
            VariableDictionary.blockScale = 1;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.DUCT_风口_F4_900x1200_TCH;
            VariableDictionary.textBoxScale = Convert.ToInt32(textBox_Scale_比例.Text) / 100;
            VariableDictionary.winForm_Status = true;
            // 生成临时文件名（放在系统临时目录，带 .dwg 扩展名）
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        private void button_清空图层_Click(object sender, EventArgs e)
        {
            try
            {
                VariableDictionary.winForm_Status = true;

                Command.DeleteSelectLayerContent();
                RunPurgeAfterExplode();
                VariableDictionary.winForm_Status = false;
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"启动分解图层块时发生错误：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogManager.Instance.LogError($"启动分解图层块时出错: {ex.Message}");
            }
        }

        #endregion

        #region 建筑
        public void button_JZ_吊顶_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            if (VariableDictionary.winForm_Status)
                VariableDictionary.winFormDiaoDingHeight = textBox_TJ_JZ_height.Text;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            var command = UnifiedCommandManager.GetCommand("吊顶");
            command?.Invoke();
            //VariableDictionary.winForm_Status = false;
        }

        public void button_JZ_不吊顶_Click(object sender, EventArgs e)
        {

            VariableDictionary.winForm_Status = true;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            var command = UnifiedCommandManager.GetCommand("不吊顶");
            command?.Invoke();
        }

        public void button_JZ_防撞护板_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "防撞护板";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(建筑专业J)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 30;//设置为被插入的图层颜色
            VariableDictionary.layerName = "TJ(建筑专业J)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            VariableDictionary.textbox_Gap = Convert.ToDouble(textBox_距离墙值.Text);
            Env.Document.SendStringToExecute("ParallelLines ", false, false, false);

        }

        public void button_冷藏库降板_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "冷藏库降板（270）";
            //VariableDictionary.btnBlockLayer = "GYTJ-碱液";
            VariableDictionary.btnBlockLayer = "TJ(建筑专业J)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 30;//设置为被插入的图层颜色
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            VariableDictionary.layerName = "TJ(建筑专业J)";
            Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
            var command = UnifiedCommandManager.GetCommand("冷藏库降板");
            command?.Invoke();
        }

        public void button_冷冻库降板_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "冷冻库降板（390）";
            //VariableDictionary.btnBlockLayer = "GYTJ-碱液";
            VariableDictionary.btnBlockLayer = "TJ(建筑专业J)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 30;//设置为被插入的图层颜色
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            VariableDictionary.layerName = "TJ(建筑专业J)";
            Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
            var command = UnifiedCommandManager.GetCommand("冷藏库降板");
            command?.Invoke();
        }

        public void button_排水沟_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            //VariableDictionary.btnFileName = "JZTJ_排水沟";
            //VariableDictionary.buttonText = "JZTJ_排水沟";
            //VariableDictionary.btnFileName_blockName = "$TWTSYS$00000508";
            //VariableDictionary.btnBlockLayer = "TJ(建筑专业J)";
            //VariableDictionary.layerColorIndex = 30;//设置为被插入的图层颜色
            VariableDictionary.dimString_JZ_宽 = Convert.ToDouble(textBox_排水沟_宽.Text);
            VariableDictionary.dimString_JZ_深 = Convert.ToDouble(textBox_排水沟_深.Text);
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            //VariableDictionary.resourcesFile = Resources.PTJ_消火栓;
            //Env.Document.SendStringToExecute("Rec2PolyLine_3 ", false, false, false);
            var command = UnifiedCommandManager.GetCommand("排水沟");
            command?.Invoke();
            VariableDictionary.winForm_Status = false;
        }

        public void button_JZ_房间号_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = textBox_楼层号.Text + "-" + textBox_洁净区1_2.Text + textBox_系统分区.Text + textBox_房间号副号.Text;
            VariableDictionary.btnBlockLayer = "TJ(房间编号)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(房间编号)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            if (textBox_楼层号.Text == "1" && textBox_洁净区1_2.Text == "1" && textBox_系统分区.Text == "1")
            {
                VariableDictionary.layerColorIndex = 64;
                VariableDictionary.jjqInt = 1;
                VariableDictionary.xtqInt = 1;
            }
            else if (Convert.ToInt32(textBox_洁净区1_2.Text) != VariableDictionary.jjqInt)
            {
                var layerColorTest = VariableDictionary.jjqLayerColorIndex[Convert.ToInt32(textBox_洁净区1_2.Text)];
                VariableDictionary.layerColorIndex = Convert.ToInt16(layerColorTest);//设置为被插入的图层颜色
                VariableDictionary.jjqInt = Convert.ToInt32(textBox_洁净区1_2.Text);
            }
            else if (Convert.ToInt32(textBox_系统分区.Text) != VariableDictionary.xtqInt)
            {
                var layerColorTest = VariableDictionary.xtqLayerColorIndex[Convert.ToInt32(textBox_系统分区.Text)];
                VariableDictionary.layerColorIndex = Convert.ToInt16(layerColorTest);//设置为被插入的图层颜色
                VariableDictionary.xtqInt = Convert.ToInt32(textBox_系统分区.Text);
            }
            //Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
            if (Convert.ToInt32(textBox_房间号副号.Text) < 9)
            {
                textBox_房间号副号.Text = "0" + (Convert.ToInt32(textBox_房间号副号.Text) + 1).ToString();
            }
            else
            {
                textBox_房间号副号.Text = (Convert.ToInt32(textBox_房间号副号.Text) + 1).ToString();
            }

            var command = UnifiedCommandManager.GetCommand("房间编号");
            command?.Invoke();
        }
        #endregion

        #region 自控
        public void button_ZK_无线AP_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_无线AP";
            VariableDictionary.btnFileName_blockName = "$equip$00001857";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = GB_NewCadPlus_IV.Resources.ZKTJ_EQUIP_无线AP;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_电话插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_电话插座";
            VariableDictionary.btnFileName_blockName = "$equip$00001867";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_电话插座;
            //VariableDictionary.blockScale = 500;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_网络插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_网络插座";
            VariableDictionary.btnFileName_blockName = "$equip$00001847";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_网络插座;
            //VariableDictionary.blockScale = 500;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_电话网络插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_电话网络插座";
            VariableDictionary.btnFileName_blockName = "ZKTJ-电话网络插座";
            // VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_电话网络插座;
            //VariableDictionary.blockScale = 500;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) * 5;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_安防监控_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_安防监控";
            VariableDictionary.btnFileName_blockName = "HC002695005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_安防监控;
            //VariableDictionary.blockScale = 500;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_眼纹识别器_Click(object sender, EventArgs e)
        {

            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_眼纹识别器";
            VariableDictionary.btnFileName_blockName = "$equip$00002616";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_眼纹识别器;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_无线网络接入点_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_无线AP";
            VariableDictionary.btnFileName_blockName = "$equip$00003217";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_无线AP;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        private void button_无线AP_吊顶_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_吸顶式无线AP";
            //VariableDictionary.btnFileName_blockName = "$equip$00003217";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_吸顶式无线AP;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_室外彩色云台摄像机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_室外彩色云台摄像机";
            VariableDictionary.btnFileName_blockName = "$equip$00002970";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_室外彩色云台摄像机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_外线电话插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_外线电话插座";
            VariableDictionary.btnFileName_blockName = "$Equip$00003196";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_外线电话插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_网络交换机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_网络交换机";
            VariableDictionary.btnFileName_blockName = "$equip$00002332";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_网络交换机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_室外彩色摄像机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_室外彩色摄像机";
            VariableDictionary.btnFileName_blockName = "$equip$00002969";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_室外彩色摄像机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_人像识别器_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_人像识别器";
            VariableDictionary.btnFileName_blockName = "$equip$00002496";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_人像识别器;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        private void button_ZK_人脸识别一体机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_人脸识别一体机";
            //VariableDictionary.btnFileName_blockName = "$equip$00002496";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_人脸识别一体机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_内线电话插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_内线电话插座";
            VariableDictionary.btnFileName_blockName = "$Equip$00003195";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_内线电话插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_门磁开关_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_门磁开关";
            VariableDictionary.btnFileName_blockName = "$equip$00002621";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_门磁开关;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_局域网插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_局域网插座";
            VariableDictionary.btnFileName_blockName = "$Equip$00003198";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_局域网插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_门禁控制器_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_门禁控制器";
            VariableDictionary.btnFileName_blockName = "$equip_U$00000028";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_门禁控制器;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_读卡器_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_读卡器";
            VariableDictionary.btnFileName_blockName = "$equip$00002617";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_读卡器;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_带扬声器电话机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_带扬声器电话机";
            VariableDictionary.btnFileName_blockName = "$equip$00003042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_带扬声器电话机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_互联网插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_互联网插座";
            VariableDictionary.btnFileName_blockName = "$Equip$00003197";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_互联网插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_广角彩色摄像机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_广角彩色摄像机";
            VariableDictionary.btnFileName_blockName = "$equip$00002731";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_广角彩色摄像机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_防爆型网络摄像机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_防爆型网络摄像机";
            VariableDictionary.btnFileName_blockName = "$equip$00002975";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_防爆型网络摄像机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_防爆型电话机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_防爆型电话机";
            VariableDictionary.btnFileName_blockName = "$equip$00003047";
            // VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-通讯";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_防爆型电话机;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_半球彩色摄像机_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_半球彩色摄像机";
            VariableDictionary.btnFileName_blockName = "$equip$00002353";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_半球彩色摄像机;
            //VariableDictionary.blockScale = 500;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 125;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_ZK_电锁按键_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_电锁按键";
            VariableDictionary.btnFileName_blockName = "$equip$00002375";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_电锁按键;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_ZK_电控锁_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_电控锁";
            VariableDictionary.btnFileName_blockName = "$equip$00002474";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_电控锁;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_ZK_监控文字_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "ZKTJ_EQUIP_监控文字";
            VariableDictionary.btnFileName_blockName = "ZKTJ-EQUIP-监控";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "EQUIP-安防";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 3;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.ZKTJ_EQUIP_监控文字;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        #endregion

        #region 结构
        /// <summary>
        /// 结构插入受力点
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_结构受力点_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "TJ(结构专业JG)";
            VariableDictionary.layerName = "TJ(结构专业JG)";
            VariableDictionary.btnFileName_blockName = "A$C9bff4efc";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;//设置为被插入的图层名
            VariableDictionary.buttonText = "STJ_受力点";
            VariableDictionary.layerColorIndex = 231;//设置为被插入的图层颜色\
            VariableDictionary.dimString = textBox_荷载数据.Text;
            VariableDictionary.resourcesFile = Resources.STJ_受力点;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("GB_InsertBlock_5 ", false, false, false);
        }
        /// <summary>
        /// 结构水平荷载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_水平荷载_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "TJ(结构专业JG)";
            VariableDictionary.layerName = "TJ(结构专业JG)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "STJ_水平荷载";
            VariableDictionary.dimString = textBox_荷载数据.Text;
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //VariableDictionary.resourcesFile = Resources.STJ_水平荷载;
            Env.Document.SendStringToExecute("NLinePolyline_N ", false, false, false);
        }

        /// <summary>
        /// 面着地   DrawPolyline
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_面着地_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnFileName = "TJ(结构专业JG)";
            VariableDictionary.layerName = "TJ(结构专业JG)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "STJ_面着地";
            VariableDictionary.dimString = textBox_荷载数据.Text;
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("NLinePolyline ", false, false, false);
        }

        /// <summary>
        /// 结构框着地
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_框着地_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnFileName = "TJ(结构专业JG)";
            VariableDictionary.layerName = "TJ(结构专业JG)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "STJ_框着地";
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.dimString = textBox_荷载数据.Text;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("NLinePolyline_Not ", false, false, false);
        }


        /// <summary>
        /// 结构直径圆形开洞
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_圆形开洞_Click(object sender, EventArgs e)
        {

            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnFileName = "TJ(结构洞口)";
            VariableDictionary.layerName = "TJ(结构洞口)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "STJ_圆形开洞";
            //VariableDictionary.btnBlockLayer = "TJ(结构专业JG)";
            VariableDictionary.textBox_S_CirDiameter = Convert.ToDouble(textBox_S_直径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_cirDiameter_Plus.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            if (VariableDictionary.textBox_S_CirDiameter == 0)
            {
                Env.Document.SendStringToExecute("CirDiameter ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirDiameter_2 ", false, false, false);
            }
            ;
        }
        /// <summary>
        /// 结构半径开圆洞口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_半径开圆洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnFileName = "TJ(结构洞口)";
            VariableDictionary.layerName = "TJ(结构洞口)";
            VariableDictionary.buttonText = "STJ_圆形开洞";
            //VariableDictionary.btnBlockLayer = "TJ(结构专业JG)";
            VariableDictionary.textbox_S_Cirradius = Convert.ToDouble(textBox_S_半径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_cirRadius_Plus.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            if (VariableDictionary.textbox_S_Cirradius == 0)
            {
                Env.Document.SendStringToExecute("CirRadius ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirRadius_2 ", false, false, false);
            }
            ;
        }
        /// <summary>
        /// 矩形开洞
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_S_矩形开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.textbox_Height = Convert.ToDouble(textBox_S_高.Text);
            VariableDictionary.textbox_Width = Convert.ToDouble(textBox_S_宽.Text);
            VariableDictionary.btnBlockLayer = "TJ(结构洞口)";
            VariableDictionary.layerName = "TJ(结构洞口)";
            VariableDictionary.buttonText = "矩形开洞";
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.textColorIndex = 3;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            if (Convert.ToDouble(VariableDictionary.textbox_Height) > 0 && Convert.ToDouble(VariableDictionary.textbox_Width) > 0)
            {
                recAndMRec = 0;
                VariableDictionary.btnFileName = "TJ(结构洞口)";
                VariableDictionary.textbox_RecPlus_Text = textBox2_RectangleExpansion.Text;
                Env.Document.SendStringToExecute("DrawRec ", false, false, false);
            }
            else
            {
                VariableDictionary.btnFileName = "TJ(结构洞口)";
                VariableDictionary.textbox_RecPlus_Text = textBox2_RectangleExpansion.Text;

                Env.Document.SendStringToExecute("Rec2PolyLine ", false, false, false);
                //Env.Editor.Regen();
            }



        }


        /// <summary>
        /// 工艺内结构画矩形为0，结构内画矩形为1
        /// </summary>
        public static int recAndMRec = 0;
        /// <summary>
        /// 结构指定长宽画矩形
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_Rectangle_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            recAndMRec = 1;
            VariableDictionary.textbox_Height = Convert.ToDouble(textBox2_height.Text);
            VariableDictionary.textbox_Width = Convert.ToDouble(textBox2_width.Text);
            VariableDictionary.buttonText = "TJ(结构洞口)";
            VariableDictionary.layerName = "TJ(结构洞口)";
            VariableDictionary.btnBlockLayer = "TJ(结构洞口)";
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.btnFileName = "TJ(结构洞口)";
            VariableDictionary.textbox_RecPlus_Text = "0";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("DrawRec ", false, false, false);
        }
        public void button_MRectangle_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            recAndMRec = 0;
            VariableDictionary.textbox_Height = Convert.ToDouble(textBox_S_高.Text);
            VariableDictionary.textbox_Width = Convert.ToDouble(textBox_S_宽.Text);
            VariableDictionary.layerColorIndex = 231;
            VariableDictionary.layerName = "TJ(结构洞口)";
            VariableDictionary.buttonText = "TJ(结构专业JG)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("DrawRec ", false, false, false);
        }

        #endregion

        #region 给排水

        public void button_P_洗眼器_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_洗眼器";
            VariableDictionary.btnFileName_blockName = "$TWTSYS$00000604";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.PTJ_洗眼器;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_不给饮用水_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "不给饮用水";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.layerColorIndex = 7;//设置为被插入的图层颜色
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            Env.Document.SendStringToExecute("DDimLinearP ", false, false, false);
        }

        public void button_P_小便器给水_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_小便器给水";
            VariableDictionary.btnFileName_blockName = "$TWTSYS$00000603";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_小便器给水;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_大洗涤池_Click_1(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_大洗涤池";
            VariableDictionary.btnFileName_blockName = "$equip$00003217";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_大洗涤池_1x0_5m;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_大便器给水_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_大便器给水";
            VariableDictionary.btnFileName_blockName = "$TWTSYS$00000602";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_大便器给水;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_洗涤盆_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_洗涤盆";
            VariableDictionary.btnFileName_blockName = "普通区洗涤盆";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 0;
            // VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_洗涤盆;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock_Ptj ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_水池给水_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_水池给水";
            VariableDictionary.btnFileName_blockName = "$TWTSYS$00000605";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.blockScale = 1.5;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_水池给水;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_热直排管_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_热直排管";
            VariableDictionary.btnFileName_blockName = "$TwtSys$00000328";
            VariableDictionary.btnBlockLayer = "EQUIP_地漏";
            VariableDictionary.blockScale = 1.5;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_热直排管;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_冷直排管_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_冷直排管";
            VariableDictionary.btnFileName_blockName = "$TwtSys$00000327";
            VariableDictionary.btnBlockLayer = "EQUIP_地漏";
            VariableDictionary.blockScale = 1.5;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_冷直排管;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_地漏_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_地漏";
            VariableDictionary.btnFileName_blockName = "$TwtSys$00000141";
            VariableDictionary.btnBlockLayer = "PTJ_地漏";
            VariableDictionary.blockScale = 1.5;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_地漏;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_给水点_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_给水点";
            VariableDictionary.btnFileName_blockName = "普通区给水";
            VariableDictionary.btnBlockLayer = "EQUIP_给水";
            VariableDictionary.blockScale = 1;
            VariableDictionary.TCH_Ptj_No = 0;
            VariableDictionary.resourcesFile = Resources.PTJ_给水点;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_洗脸盆_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_洗脸盆";
            //VariableDictionary.btnFileName_blockName = "$TWTSYS$00000600";
            VariableDictionary.btnFileName_blockName = "普通区洗脸盆";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 10;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_洗脸盆;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_冷不带压直排_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_冷不带压直排";
            VariableDictionary.btnFileName_blockName = "$TWTSYS$00000622";
            VariableDictionary.btnBlockLayer = "EQUIP_地漏";
            VariableDictionary.blockScale = 1;
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_冷不带压直排;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_热不带压直排_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_热不带压直排";
            VariableDictionary.btnFileName_blockName = "$TwtSys$00000138";
            VariableDictionary.btnBlockLayer = "EQUIP_地漏";
            VariableDictionary.blockScale = 1;
            //TCH_Ptj = true;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_热不带压直排;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_拖布池_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_拖布池";
            VariableDictionary.btnFileName_blockName = "A$C32361FA1";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_拖布池;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_洗涤池1x05m_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_大洗涤池_1x0_5m";
            VariableDictionary.btnFileName_blockName = "A$C5C905366";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 10;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_大洗涤池_1x0_5m;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_洗涤池12x05m_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_大洗涤池_1_2x0_5m";
            VariableDictionary.btnFileName_blockName = "A$C18325CD1";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 10;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_大洗涤池_1_2x0_5m;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_洗涤池15x05m_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_大洗涤池_1_5x0_5m";
            VariableDictionary.btnFileName_blockName = "A$C5A6D4801";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 10;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_大洗涤池_1_5x0_5m;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_洗涤池18x05m_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_大洗涤池_1_8x0_5m";
            VariableDictionary.btnFileName_blockName = "A$C3C1E07D8";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 0;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_大洗涤池_1_8x0_5m;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        public void button_P_洗涤池2x05m_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "PTJ_大洗涤池_2_0x0_5m";
            VariableDictionary.btnFileName_blockName = "A$C0AB663B7";
            VariableDictionary.btnBlockLayer = "TJ(给排水专业S)";
            VariableDictionary.TCH_Ptj_No = 10;
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.resourcesFile = Resources.PTJ_大洗涤池_2_0x0_5m;
            VariableDictionary.layerName = "TJ(给排水专业S)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        #endregion

        #region 电气
        public void button_DQ_220V插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            //VariableDictionary.btnFileName = "DQTJ_EQUIP_单相插座";
            //VariableDictionary.btnFileName_blockName = "HC002694005706";
            ////VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            //VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);

        }
        public void button_DQ_三相380V插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            //VariableDictionary.btnFileName = "DQTJ_EQUIP_三相380V插座";
            //VariableDictionary.btnFileName_blockName = "HC002696005706";
            ////VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            //VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_三相380V插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_潮湿插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            //VariableDictionary.btnFileName = "DQTJ_EQUIP_潮湿插座";
            //VariableDictionary.btnFileName_blockName = "HC002695005706";
            ////VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            //VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_潮湿插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_三相潮湿插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            //VariableDictionary.btnFileName = "DQTJ_EQUIP_三相潮湿插座";
            //VariableDictionary.btnFileName_blockName = "HC002697005706";
            ////VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            //VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_三相潮湿插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_空调插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_空调插座";
            VariableDictionary.btnFileName_blockName = "HC003131100042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_空调插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_设备用电点位_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_设备用电点位";
            VariableDictionary.btnFileName_blockName = "HC002694005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_设备用电点位;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相夹层_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相夹层插座";
            VariableDictionary.btnFileName_blockName = "HC002698005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相夹层插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_插座箱_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_插座箱";
            VariableDictionary.btnFileName_blockName = "DQTJ-插座箱";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.blockScale = 1;
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_插座箱;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_应急插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_应急插座";
            VariableDictionary.btnFileName_blockName = "HC002997005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 4;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_应急插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_应急16A电源_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_应急16A插座";
            VariableDictionary.btnFileName_blockName = "DQTJ-UPS16A电源";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 4;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_应急16A插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_UPS插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_UPS插座";
            VariableDictionary.btnFileName_blockName = "DQTJ-应急UPS电源";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 4;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_UPS插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_UPS16A插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_UPS16A插座";
            VariableDictionary.btnFileName_blockName = "DQTJ-应急UPS16A电源";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 4;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_UPS16A插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_传递窗电源插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_传递窗电源插座";
            VariableDictionary.btnFileName_blockName = "HC003001006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_传递窗电源插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_门禁插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_门禁插座";
            VariableDictionary.btnFileName_blockName = "A$C16EA1F35";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_门禁插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_红外感应门插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_红外感应门插座";
            VariableDictionary.btnFileName_blockName = "DQTJ-红外感应门插座";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_红外感应门插座;
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_紫外灯_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "紫外灯";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
        }
        public void button_DQ_三联插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_三联插座";
            VariableDictionary.btnFileName_blockName = "$equip$00001992";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.blockScale = 1.5;
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_三联插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_四联插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_四联插座";
            VariableDictionary.btnFileName_blockName = "$equip$00002163";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_四联插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 500;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_互锁插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_互锁插座";
            VariableDictionary.btnFileName_blockName = "HC002698005707";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_互锁插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_两点互锁_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_两点互锁";
            VariableDictionary.btnFileName_blockName = "DQTJ-两点互锁";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            //VariableDictionary.blockScale = 0.8;
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_两点互锁;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_三点互锁_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_三点互锁";
            VariableDictionary.btnFileName_blockName = "A$C0664bbbd";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.blockScale = 0.8;
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_三点互锁;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_立式空调插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_立式空调插座";
            VariableDictionary.btnFileName_blockName = "HC003131000042";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_立式空调插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_壁挂空调插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_壁挂空调插座";
            VariableDictionary.btnFileName_blockName = "HC003130000042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_壁挂空调插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_手消毒插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_手消毒插座";
            VariableDictionary.btnFileName_blockName = "HC003007006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_手消毒插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_视孔灯_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_视孔灯";
            VariableDictionary.btnFileName_blockName = "$Equip$00003237";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_视孔灯;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 200;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_烘手器插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_烘手器插座";
            VariableDictionary.btnFileName_blockName = "A$C21791F4C";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_烘手器插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_实验台功能柱插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_实验台功能柱插座";
            VariableDictionary.btnFileName_blockName = "HC002694005706N";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名HC002694005706
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_实验台功能柱插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_实验台UPS功能柱插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_实验台UPS功能柱电源";
            VariableDictionary.btnFileName_blockName = "HC003210000042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D1)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 4;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_实验台UPS功能柱电源;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_电热水器插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_电热水器安全型插座";
            VariableDictionary.btnFileName_blockName = "HC003021006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_电热水器安全型插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_厨宝插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_厨宝安全型插座";
            VariableDictionary.btnFileName_blockName = "A$C1E63194F";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_厨宝安全型插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_烘手器_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_烘手器";
            VariableDictionary.btnFileName_blockName = "$Equip$00003233";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_烘手器;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 200;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_驱鼠器插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_驱鼠器插座";
            VariableDictionary.btnFileName_blockName = "HC003076006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_驱鼠器插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_灭蝇灯插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_灭蝇灯插座";
            VariableDictionary.btnFileName_blockName = "HC003076006336";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_灭蝇灯插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_灭蝇灯插座_底边_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_灭蝇灯插座_底边";
            VariableDictionary.btnFileName_blockName = "HC002694005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_灭蝇灯插座_底边;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_实验台UPS功能柱电源_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_实验台UPS功能柱电源";
            VariableDictionary.btnFileName_blockName = "HC003210000042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_实验台UPS功能柱电源;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_实验台上方220V插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_实验台上方220V插座";
            VariableDictionary.btnFileName_blockName = "HC003212000042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_实验台上方220V插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_380V用电设备_点或配电柜_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_380V用电设备点或配电柜";
            VariableDictionary.btnFileName_blockName = "380V用电设备点或配电柜";
            VariableDictionary.buttonText = "380V用电设备点或配电柜";
            VariableDictionary.dimString = "设备名称\n" + textBox_inputKW.Text + "kw";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_380V用电设备点或配电柜;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            VariableDictionary.winForm_Status = true;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);

        }
        public void button_DQ_380V用电设备大于10KW_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_380V用电设备大于10KW";
            VariableDictionary.btnFileName_blockName = "380V用电设备大于10KW";
            VariableDictionary.buttonText = "380V用电设备大于10KW";
            VariableDictionary.dimString = "设备名称\n" + textBox_input10KW.Text + "kw";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D2)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D2)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 110;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_380V用电设备大于10KW;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            VariableDictionary.winForm_Status = true;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
            //VariableDictionary.dimString = null;
        }
        public void button_DQ_220V用电设备_点或配电柜_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.buttonText = "220V用电设备点或配电柜";
            VariableDictionary.btnFileName = "DQTJ_EQUIP_220V用电设备点或配电柜";
            VariableDictionary.btnFileName_blockName = "220V用电设备点或配电柜";
            VariableDictionary.dimString = "设备名称\n" + "220V," + textBox_inputKW.Text + "kw";
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名

            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_220V用电设备点或配电柜;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            VariableDictionary.blockScale = VariableDictionary.textBoxScale;
            VariableDictionary.winForm_Status = true;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);

        }
        public void button_DQ_单相插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相插座";
            VariableDictionary.btnFileName_blockName = "HC002694005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相地面插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相地面插座";
            VariableDictionary.btnFileName_blockName = "HC003202000042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相地面插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相三孔插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相三孔插座";
            VariableDictionary.btnFileName_blockName = "HC002696005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相三孔插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相空调插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相空调插座";
            VariableDictionary.btnFileName_blockName = "HC003130000042";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相空调插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相16A三孔插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相16A三孔插座";
            VariableDictionary.btnFileName_blockName = "HC002805006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相16A三孔插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相20A三孔插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相20A三孔插座";
            VariableDictionary.btnFileName_blockName = "HC002944006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相20A三孔插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相25A三孔插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相25A三孔插座";
            VariableDictionary.btnFileName_blockName = "HC002806006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D2)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相25A三孔插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相32A三孔_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相32A三孔插座";
            VariableDictionary.btnFileName_blockName = "HC002957006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相32A三孔插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相五孔岛型插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相五孔岛型插座";
            VariableDictionary.btnFileName_blockName = "$equip_U$00000168";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相五孔岛型插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 500;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相三孔岛型插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相三孔岛型插座";
            VariableDictionary.btnFileName_blockName = "$equip_U$00000169";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相三孔岛型插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 500;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_三相岛型插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_三相岛型插座";
            VariableDictionary.btnFileName_blockName = "$equip_U$00000167";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_三相岛型插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 500;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_带保护极的单相防爆插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_带保护极的单相防爆插座";
            VariableDictionary.btnFileName_blockName = "HC002820006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_带保护极的单相防爆插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_带保护极的三相防爆插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_带保护极的三相防爆插座";
            VariableDictionary.btnFileName_blockName = "HC002821006335";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_带保护极的三相防爆插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相防爆岛型插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_单相防爆岛型插座";
            VariableDictionary.btnFileName_blockName = "$equip_U$00000170";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_单相防爆岛型插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 500;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相暗敷插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_带保护极的单相暗敷插座";
            VariableDictionary.btnFileName_blockName = "DQTJ-带保护极的单相暗敷插座";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName; 
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_带保护极的单相暗敷插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_单相密闭插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_带保护极的单相密闭插座";
            VariableDictionary.btnFileName_blockName = "HC002695005706";
            //VariableDictionary.btnFileName_blockName = "HC002697005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_带保护极的单相密闭插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_三相密闭插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_带保护极的三相密闭插座";
            VariableDictionary.btnFileName_blockName = "HC002697005706";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_带保护极的三相密闭插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text) / 100;
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }
        public void button_DQ_三相暗敷插座_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnFileName = "DQTJ_EQUIP_带保护极的三相暗敷插座";
            VariableDictionary.btnFileName_blockName = "DQTJ-带保护极的三相暗敷插座";
            //VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.btnBlockLayer = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerName = "TJ(电气专业D)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.resourcesFile = Resources.DQTJ_EQUIP_带保护极的三相暗敷插座;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            //Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
            System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
            GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
        }

        #endregion

        #region 外参图元

        public void button_获取外参图元_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("CopyAndSync1 ", false, false, false);
        }

        public void button获取外参图元2_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("ReferenceCopy ", false, false, false);
        }

        public void button获取外参图元三_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("CopyAndSync3 ", false, false, false);
        }

        public void button获取外参图元四_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("CopyAndSync6 ", false, false, false);
        }

        #endregion

        #region  外参相关按键、 选择外参并从中选图元复制到当前空间内
        /// <summary>
        /// 外参实体objectid列表
        /// </summary>
        private List<ObjectId> xrefEntities;
        /// <summary>
        /// 选中实体objectid列表
        /// </summary>
        public static List<ObjectId> selectedEntities = new List<ObjectId>();
        /// <summary>
        /// 选择的外参
        /// </summary>
        public static string? selectItem;
        /// <summary>
        /// 选择全部外参图元
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_获取外参全部图元_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            Env.Document.SendStringToExecute("CopyXrefAllEntity ", false, false, false);
        }

        /// <summary>
        /// 选择外参并从中选图元列入列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_SelectReference_Click(object sender, EventArgs e)
        {
            #region
            ReferenceEntity.Items.Clear();
            SelectEntity.Items.Clear();
            Reference.Items.Clear();

            VariableDictionary.winForm_Status = true;
            using var tr = new DBTrans();
            // 选择外部参照
            PromptEntityOptions opt = new PromptEntityOptions("选择一个外部参照：");
            opt.SetRejectMessage("您必须选择一个外部参照。");
            opt.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult res = Env.Editor.GetEntity(opt);
            if (res.Status != PromptStatus.OK) return;
            // 获取外部参照中的图元
            xrefEntities = Command.GetXrefEntities(res.ObjectId);
            // 第五步：获取外部参照名称
            string xrefName = Command.getXrefName(tr, res.ObjectId);
            if (xrefName != null && xrefName != "")
                Reference.Items.Add(xrefName);
            // 添加图元到左侧列表
            foreach (ObjectId entityId in xrefEntities)
            {
                if (entityId != null)
                {
                    Entity entity = Command.GetEntity(entityId);
                    if (entity is not null)
                    {
                        selectedEntities.Add(entityId);
                        ReferenceEntity.Items.Add(Command.getXrefName(tr, entityId));
                    }
                }
            }
            tr.Commit();
            #endregion
            //Env.Document.SendStringToExecute("ReferenceCopy ", false, false, false);
            /// 发送命令到AutoCAD执行
            //Env.Document.SendStringToExecute("CopyXrefEntity2 ", false, false, false);

        }


        /// <summary>
        /// 选择实体，拿到这个实体的详细参数；
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void selectEntity(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            Env.Document.SendStringToExecute("tzData ", false, false, false);
            Env.Editor.Redraw();
        }
        /// <summary>
        /// 引用实体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void referenceEntity_SelectedIndexChanged(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            ListBox listBox = (ListBox)sender;

            if (listBox.SelectedItem != null)
            {
                // 从左侧列表移除选中的图元并添加到右侧列表
                selectItem = listBox.SelectedItem.ToString();
                SelectEntity.Items.Add(selectItem);
                ReferenceEntity.Items.Remove(selectItem);
            }
        }
        /// <summary>
        /// 选择图元
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void selectEntity_SelectedIndexChanged(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            ListBox listBox = (ListBox)sender;

            if (listBox.SelectedItem != null)
            {
                // 从右侧列表移除选中的图元并添加到左侧列表
                selectItem = listBox.SelectedItem.ToString();
                ReferenceEntity.Items.Add(selectItem);
                SelectEntity.Items.Remove(selectItem);
            }
        }

        /// <summary>
        /// 复制选中外参图元
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void Btn_selectEntity_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            /// 发送命令到AutoCAD执行  CopyXrefAllEntity
            //Env.Document.SendStringToExecute("CopyXrefAllEntity ", false, false, false);
            Env.Document.SendStringToExecute("CopyXrefEntity ", false, false, false);
        }
        /// <summary>
        /// 显示天正
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_ShowDatas_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            Env.Document.SendStringToExecute("(vlax-dump-object (vlax-ename->vla-object (car (entsel )))T) ", false, false, false);
        }
        public void buttonTEST_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            Env.Document.SendStringToExecute("CopyAtSamePosition ", false, false, false);
        }
        #endregion

        #region 开洞

        public void button_JZ_左右开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(建筑过结构洞口)";
            VariableDictionary.layerColorIndex = 64;//设置图层颜色
            VariableDictionary.textbox_Width = Convert.ToDouble(textBoxA_左右开洞.Text);
            VariableDictionary.textbox_Height = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerName = "TJ(建筑过结构洞口)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_JZ_上下开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(建筑过结构洞口)";
            VariableDictionary.layerColorIndex = 64;//设置图层颜色
            VariableDictionary.textbox_Width = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.textbox_Height = Convert.ToDouble(textBoxA_左右开洞.Text);
            VariableDictionary.layerName = "TJ(建筑过结构洞口)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_S_左右开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            //VariableDictionary.btnBlockLayer = "HOLE";
            VariableDictionary.btnBlockLayer = "TJ(结构洞口)";
            VariableDictionary.layerColorIndex = 231;//设置图层颜色
            VariableDictionary.textbox_Width = Convert.ToDouble(textBoxS_左右开洞.Text);
            VariableDictionary.textbox_Height = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerName = "TJ(结构洞口)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_S_上下开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(结构开洞)";
            VariableDictionary.layerColorIndex = 231;//设置图层颜色
            VariableDictionary.textbox_Width = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.textbox_Height = Convert.ToDouble(textBoxS_上下开洞.Text);
            VariableDictionary.layerName = "TJ(结构开洞)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_P_左右开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(给排水过建筑)";
            VariableDictionary.layerColorIndex = 7;//设置图层颜色
            VariableDictionary.textbox_Width = Convert.ToDouble(textBoxP_左右开洞.Text);
            VariableDictionary.textbox_Height = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerName = "TJ(给排水过建筑)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_P_上下开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(给排水过建筑)";
            VariableDictionary.textbox_Width = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerColorIndex = 7;//设置图层颜色
            VariableDictionary.textbox_Height = Convert.ToDouble(textBoxP_上下开洞.Text);
            VariableDictionary.layerName = "TJ(给排水过建筑)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_NT_左右开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(暖通过建筑)";//设置为被插入的图层名
            VariableDictionary.buttonText = "TJ(暖通过建筑)";
            VariableDictionary.layerColorIndex = 6;//设置图层颜色
            VariableDictionary.textbox_Width = Convert.ToDouble(textBoxN_左右开洞.Text);
            VariableDictionary.textbox_Height = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerName = "TJ(暖通过建筑)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
        }

        public void button_NT_上下开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(暖通过建筑)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 6;//设置为被插入的图层颜色
            VariableDictionary.buttonText = "TJ(暖通过建筑)";
            VariableDictionary.textbox_Height = Convert.ToDouble(textBoxN_上下开洞.Text);
            VariableDictionary.textbox_Width = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerName = "TJ(暖通过建筑)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
        }

        public void button_DQ_左右开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(电气过建筑孔洞D)";//设置为被插入的图层名
            VariableDictionary.buttonText = "TJ(电气过建筑孔洞D)";
            VariableDictionary.layerName = "TJ(电气过建筑孔洞D)";
            VariableDictionary.layerColorIndex = 142;//设置为被插入的图层颜色
            VariableDictionary.textbox_Width = Convert.ToDouble(textBoxP_左右开洞.Text);
            VariableDictionary.textbox_Height = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_DQ_上下开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(电气过建筑孔洞D)";//设置为被插入的图层名
            VariableDictionary.buttonText = "TJ(电气过建筑孔洞D)";
            VariableDictionary.textbox_Width = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerColorIndex = 142;//设置图层颜色
            VariableDictionary.textbox_Height = Convert.ToDouble(textBoxE_左右开洞.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_ZK_左右开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(自控过建筑)";
            VariableDictionary.textbox_Width = Convert.ToDouble(textBoxZ_左右开洞.Text);
            VariableDictionary.textbox_Height = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerColorIndex = 3;//设置图层颜色
            VariableDictionary.layerName = "TJ(自控过建筑)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        public void button_ZK_上下开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.btnBlockLayer = "TJ(自控过建筑)";
            VariableDictionary.textbox_Width = 0.15 * VariableDictionary.textBoxScale;
            VariableDictionary.layerColorIndex = 3;//设置图层颜色
            VariableDictionary.textbox_Height = Convert.ToDouble(textBoxZ_左右开洞.Text);
            VariableDictionary.layerName = "TJ(自控过建筑)";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            //Env.Editor.Redraw();
        }

        #endregion


        private bool isMax = false;
        public void button_GY_工艺更多_Click(object sender, EventArgs e)
        {
            //237, 278
            //237, 170
            if (!isMax)
            {
                this.groupBox_工艺.Height = 630;
                //this.panel工艺.Height = 278;
                isMax = true;
            }
            else
            {
                this.groupBox_工艺.Height = 535;
                //this.panel工艺.Height = 170;
                isMax = false;
            }
        }

        public void button_清理_Click(object sender, EventArgs e)
        {
            //Env.Document.SendStringToExecute("pu ", false, false, false);
            // -PURGE -> All -> * -> No(不逐项确认)，全自动无人工输入
            Env.Document.SendStringToExecute("_.-PURGE\n_A\n*\n_N\n", false, false, false);

            // 1) -PURGE -> All -> * -> No(不逐项确认)
            // 2) AUDIT  -> Yes(自动修复)
            //Env.Document.SendStringToExecute(
            //"_.-PURGE\n_A\n*\n_N\n_.AUDIT\n_Y\n",
            //false, false, false);
        }
        public void button_Audit_Click(object sender, EventArgs e)
        {

            // 自动执行 AUDIT，并自动回答 Yes（修复错误）
            //Env.Document.SendStringToExecute("_.AUDIT\n_Y\n", false, false, false);

            try
            {
                // 1) -PURGE -> All -> * -> No(不逐项确认)
                // 2) AUDIT  -> Yes(自动修复)
                Env.Document.SendStringToExecute(
                    "_.-PURGE\n_A\n*\n_N\n_.AUDIT\n_Y\n",
                    false, false, false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"执行清理+审计时出错: {ex.Message}");
            }
        }
        public void button_DICTS_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("DICTS ", false, false, false);
        }
        public void button_CLEANUPDWG_Click(object sender, EventArgs e)
        {

            //Env.Document.SendStringToExecute("CLEANUPDWG ", false, false, false);
            try
            {
                VariableDictionary.winForm_Status = true;
                // 执行交互式分解图层块命令
                //doc.SendStringToExecute("ExplodeBlocksInLayerInteractive ", false, false, false);
                Command.ExplodeBlocksInLayerInteractive(Convert.ToDouble(textBox_分解块小于值.Text));
                VariableDictionary.winForm_Status = false;
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"启动分解图层块时发生错误：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogManager.Instance.LogError($"启动分解图层块时出错: {ex.Message}");
            }

        }
        public void button_分解图层内所有块_Click(object sender, EventArgs e)
        {

            //Env.Document.SendStringToExecute("CLEANUPDWG ", false, false, false);
            try
            {
                VariableDictionary.winForm_Status = true;
                // 执行交互式分解图层块命令
                //doc.SendStringToExecute("ExplodeBlocksInLayerInteractive ", false, false, false);
                Command.ExplodeBlocksInLayerInteractive(Convert.ToDouble(textBox_分解块小于值.Text));
                VariableDictionary.winForm_Status = false;
                // 等价于命令行：-PURGE -> All -> * -> No(不逐项确认)
                RunPurgeAfterExplode();
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"启动分解图层块时发生错误：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogManager.Instance.LogError($"启动分解图层块时出错: {ex.Message}");
            }

        }
        public void button_分解块_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            Command.ExplodeNestedBlock(Convert.ToDouble(textBox_分解块小于值.Text));
            VariableDictionary.winForm_Status = false;
            // 等价于命令行：-PURGE -> All -> * -> No(不逐项确认)
            RunPurgeAfterExplode();
        }
        /// <summary>
        /// 分解后执行清理（类似 CAD 的 PU）
        /// </summary>
        private void RunPurgeAfterExplode()
        {
            try
            {
                // 等价于命令行：-PURGE -> All -> * -> No(不逐项确认)
                //Env.Document.SendStringToExecute("_.-PURGE _A * _N ", false, false, false);
                // -PURGE -> All -> * -> No(不逐项确认)，全自动无人工输入
                //Env.Document.SendStringToExecute("_.-PURGE\n_A\n*\n_N\n", false, false, false);
                try
                {
                    // 1) -PURGE -> All -> * -> No(不逐项确认)
                    // 2) AUDIT  -> Yes(自动修复)
                    Env.Document.SendStringToExecute(
            "_.AUDIT\n_Y\n_.-PURGE\n_A\n*\n_N\n",
            false, false, false);
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogError($"执行清理+审计时出错: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"执行PU清理时出错: {ex.Message}");
            }
        }
        /// <summary>
        /// 输入整数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void textBox_Int_TextChanged(object sender, KeyPressEventArgs e)
        {
            // 检查输入的字符是否是数字或控制字符（如退格键）
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                // 如果不是数字，阻止输入并显示提示消息
                e.Handled = true;
                //MessageBox.Show("只能输入数字，请重新输入！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        /// <summary>
        /// 输入小数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void textBox_Double_TextChanged(object sender, KeyPressEventArgs e)
        {
            TextBox? textBox = sender as TextBox;

            // 允许数字、控制字符（如退格键）、小数点
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && e.KeyChar != '.')
            {
                // 如果不是数字、控制字符或小数点，阻止输入并显示提示消息
                e.Handled = true;
                //MessageBox.Show("只能输入数字和小数点，请重新输入！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // 确保小数点只能输入一次
            if (e.KeyChar == '.' && textBox.Text.Contains("."))
            {
                e.Handled = true; // 阻止输入
            }
        }
        public void textBox_inputKW_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox_inputKW.Text == "请输入功率")
                textBox_inputKW.Clear();
        }
        public void textBox_inputKW_MouseLeave(object sender, EventArgs e)
        {
            if (textBox_inputKW.Text.Length == 0)
                textBox_inputKW.Text = "请输入功率";
        }
        public void textBox_input10KW_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox_input10KW.Text == "请输入功率")
                textBox_input10KW.Clear();
            if (textBox_E_input10KW.Text == "请输入功率")
                textBox_E_input10KW.Clear();
        }
        public void textBox_input10KW_MouseLeave(object sender, EventArgs e)
        {
            if (textBox_input10KW.Text.Length == 0)
                textBox_input10KW.Text = "请输入功率";
            if (textBox_E_input10KW.Text.Length == 0)
                textBox_E_input10KW.Text = "请输入功率";
        }
        public void textBox_排水沟_深_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox_排水沟_深.Text == "请输入深")
                textBox_排水沟_深.Clear();
        }
        public void textBox_排水沟_深_MouseLeave(object sender, EventArgs e)
        {
            if (textBox_排水沟_深.Text.Length == 0)
                textBox_排水沟_深.Text = "请输入深";
        }
        public void textBox_排水沟_宽_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox_排水沟_宽.Text == "请输入宽")
                textBox_排水沟_宽.Clear();
        }
        public void textBox_排水沟_宽_MouseLeave(object sender, EventArgs e)
        {
            if (textBox_排水沟_宽.Text.Length == 0)
                textBox_排水沟_宽.Text = "请输入宽";
        }
        public void textBox_排风百分比_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox_排风百分比.Text == "排风百分比")
                textBox_排风百分比.Clear();
        }
        public void textBox_排风百分比_MouseLeave(object sender, EventArgs e)
        {
            if (textBox_排风百分比.Text.Length == 0)
                textBox_排风百分比.Text = "排风百分比";
        }
        public void textBox_荷载数据_MouseDown(object sender, MouseEventArgs e)
        {
            if (textBox_荷载数据.Text == "输入荷载数据")
                textBox_荷载数据.Clear();
        }
        public void textBox_荷载数据_MouseLeave(object sender, EventArgs e)
        {
            if (textBox_荷载数据.Text.Length == 0)
                textBox_荷载数据.Text = "输入荷载数据";
        }

        public void button_RECOVERY_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("DRAWINGRECOVERY ", false, false, false);
        }

        public void button_特殊TEXT_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            VariableDictionary.entityRotateAngle = 0;
            VariableDictionary.btnBlockLayer = "TJ(建筑专业J)";//设置为被插入的图层名
            VariableDictionary.layerColorIndex = 30;//设置为被插入的图层颜色
            VariableDictionary.btnFileName = "特殊地面做法要求";
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);
            VariableDictionary.layerName = "TJ(建筑专业J)";
            Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
        }

        public void button_checkNo_Click(object sender, EventArgs e)
        {
            VariableDictionary.winForm_Status = true;
            if (VariableDictionary.btnState == false) { VariableDictionary.btnState = true; } else { VariableDictionary.btnState = false; }
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();//初始化allTjLayer
            foreach (var item in VariableDictionary.allTjtLayer)//清空全部图层列表
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            VariableDictionary.selectTjtLayer.Remove("TJ(房间编号)");
            VariableDictionary.selectTjtLayer.Remove("房间编号");

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
            //Env.Document.SendStringToExecute("CloseLayer ", false, false, false);
        }
        /// <summary>
        /// 关闭或打开所有条件图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_closeAllTJ_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            // 设置按钮颜色以反映状态
            if (button_closeAllTJ.ForeColor.Name == "Black" || button_closeAllTJ.ForeColor.Name == "ControlText")
            {
                button_closeAllTJ.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_关闭工艺.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_关闭建筑.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_关闭结构.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_关闭给排水.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_关闭暖通.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_关闭电气.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_关闭自控.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_closeAllTJ.ForeColor = System.Drawing.SystemColors.ControlText;
                button_关闭工艺.ForeColor = System.Drawing.SystemColors.ControlText;
                button_关闭建筑.ForeColor = System.Drawing.SystemColors.ControlText;
                button_关闭结构.ForeColor = System.Drawing.SystemColors.ControlText;
                button_关闭给排水.ForeColor = System.Drawing.SystemColors.ControlText;
                button_关闭暖通.ForeColor = System.Drawing.SystemColors.ControlText;
                button_关闭电气.ForeColor = System.Drawing.SystemColors.ControlText;
                button_关闭自控.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            VariableDictionary.allTjtLayer.Clear();//清空全部图层列表
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.allTjtLayer)//清空全部图层列表
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in VariableDictionary.SBtjtBtn)//清空全部图层列表
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }
            foreach (var item in VariableDictionary.ApmtTjBtn)//清空全部图层列表
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
            //Env.Document.SendStringToExecute("CloseLayer ", false, false, false);


        }
        /// <summary>
        /// 生成外轮廓线
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OutlineGenerator_Click(object sender, EventArgs e)
        {

            Env.Document.SendStringToExecute("SMARTOUTLINE ", false, false, false);
        }

        /// <summary>
        /// 只看共用图层
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_onlyAllTJ_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button_closeAllTJ.ForeColor.Name == "Black" || button_closeAllTJ.ForeColor.Name == "ControlText")
            {
                button_closeAllTJ.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                VariableDictionary.btnState = true;
            }
            else
            {
                button_closeAllTJ.ForeColor = System.Drawing.SystemColors.ControlText;
                VariableDictionary.btnState = false;
            }
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("CloseLayer ", false, false, false);

        }

        public void button_test_Off_Btn_Click(object sender, EventArgs e)
        {
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("1");
            VariableDictionary.selectTjtLayer.Add("SB");
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        public void button_test_On_Btn_Click(object sender, EventArgs e)
        {
            VariableDictionary.allTjtLayer.Clear();
            NewTjLayer();
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("1");
            VariableDictionary.selectTjtLayer.Add("SB");
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }


        #region 冻结图层
        /// <summary>
        /// 打开工艺视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenGYXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button_OpenGYXref.ForeColor.Name == "Black" || button_OpenGYXref.ForeColor.Name == "ControlText")
            {
                //button_OpenGYXref.ForeColor = System.Drawing.SystemColors.ActiveCaption;
                button_OpenGYXref.Enabled = false;
                button_OffGYXref.Enabled = true;
                VariableDictionary.btnState = true;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭工艺视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffGYXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button_OffGYXref.ForeColor.Name == "Black" || button_OffGYXref.ForeColor.Name == "ControlText")
            {
                button_OffGYXref.Enabled = false;
                button_OpenGYXref.Enabled = true;
                VariableDictionary.btnState = true;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        /// <summary>
        /// 打开建筑视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenJZXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OpenJZXref.Enabled = false;
            button_OffJZXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.AtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭建筑视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffJZXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OffJZXref.Enabled = false;
            button_OpenJZXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.AtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        /// <summary>
        /// 打开结构视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenJGXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OpenJGXref.Enabled = false;
            button_OffJGXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.StjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭结构视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffJGXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OffJGXref.Enabled = false;
            button_OpenJGXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.StjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        /// <summary>
        /// 打开暖通视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenNTXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OpenNTXref.Enabled = false;
            button_OffNTXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.NtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭暖通视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffNTXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OffNTXref.Enabled = false;
            button_OpenNTXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.NtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        /// <summary>
        /// 打开给排水视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenJPSXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OpenJPSXref.Enabled = false;
            button_OffJPSXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.PtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭给排水视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffJPSXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OffJPSXref.Enabled = false;
            button_OpenJPSXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.PtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        /// <summary>
        /// 打开电气视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenDQXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OpenDQXref.Enabled = false;
            button_OffDQXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.EtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭电气视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffDQXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OffDQXref.Enabled = false;
            button_OpenDQXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.EtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        /// <summary>
        /// 打开自控视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenZKXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OpenZKXref.Enabled = false;
            button_OffZKref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ZKtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭自控视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffZKref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_OffZKref.Enabled = false;
            button_OpenZKXref.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ZKtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }

        /// <summary>
        /// 打开共用图层视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OpenGGXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_开关设备图层.Enabled = false;
            button_开关区域图层.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewportOpen ", false, false, false);
        }

        /// <summary>
        /// 关闭共用图层视口外参
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_OffGGXref_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            button_开关区域图层.Enabled = false;
            button_开关设备图层.Enabled = true;
            VariableDictionary.btnState = true;

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("FindXrefLayersInViewport ", false, false, false);
        }


        #endregion
        /// <summary>
        /// 插入表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void btn_InputExcel_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("InsertExcelTableToCAD ", false, false, false);
        }
        /// <summary>
        /// 导出表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button_outExcel_Click(object sender, EventArgs e)
        {
            Env.Document.SendStringToExecute("ExportCADTable ", false, false, false);
        }



        #region 电气
        public void button电气开关工艺条件_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气开关工艺条件.ForeColor.Name == "Black" || button电气开关工艺条件.ForeColor.Name == "ControlText")
            {
                button电气开关工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气开关工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收工艺条件_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
            {
                "TJ(电气专业D)",
                "TJ(电气专业D)",
                "TJ(电气专业D2)",
                "SB(工艺设备)",
                "SB(设备名称)",
                "S_设备名称",
                "SB(设备外框)",
                "QY",
            };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button电气收工艺条件.ForeColor.Name == "Black" || button电气收工艺条件.ForeColor.Name == "ControlText")
            {
                button电气收工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收工艺屋面放散管位置_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(电气专业D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收工艺屋面放散管位置.ForeColor.Name == "Black" || button电气收工艺屋面放散管位置.ForeColor.Name == "ControlText")
            {
                button电气收工艺屋面放散管位置.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收工艺屋面放散管位置.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收工艺电气大于10kw_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(电气专业D2)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收工艺电气大于10kw.ForeColor.Name == "Black" || button电气收工艺电气大于10kw.ForeColor.Name == "ControlText")
            {
                button电气收工艺电气大于10kw.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收工艺电气大于10kw.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收工艺双电源条件_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(电气专业D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收工艺双电源条件.ForeColor.Name == "Black" || button电气收工艺双电源条件.ForeColor.Name == "ControlText")
            {
                button电气收工艺双电源条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收工艺双电源条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收工艺设备名称_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "SB(设备名称)",
            "S_设备名称",
            "SB(工艺设备)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收工艺设备名称.ForeColor.Name == "Black" || button电气收工艺设备名称.ForeColor.Name == "ControlText")
            {
                button电气收工艺设备名称.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收工艺设备名称.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button洁净区域_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "QY"
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button洁净区域.ForeColor.Name == "Black" || button洁净区域.ForeColor.Name == "ControlText")
            {
                button洁净区域.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button洁净区域.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收工艺设备外框_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "SB(设备外框)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收工艺设备外框.ForeColor.Name == "Black" || button电气收工艺设备外框.ForeColor.Name == "ControlText")
            {
                button电气收工艺设备外框.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收工艺设备外框.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气开关建筑条件_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(房间编号)",
            "TJ(房间编号)",
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气开关建筑条件.ForeColor.Name == "Black" || button电气开关建筑条件.ForeColor.Name == "ControlText")
            {
                button电气开关建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气开关建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气开关给排水条件_Click(object sender, EventArgs e)
        {
            List<string> EPtjtBtn = new List<string>
        {
            "EQUIP_消火栓",
            "TJ(给排水过电气动力条件)",
            "TJ(给排水过电气喷淋条件)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气开关给排水条件.ForeColor.Name == "Black" || button电气开关给排水条件.ForeColor.Name == "ControlText")
            {
                button电气开关给排水条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气开关给排水条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EPtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气开关暖通条件_Click(object sender, EventArgs e)
        {

            List<string> ENtjtBtn = new List<string>
        {
           "TJ(暖通过电气)",
           "暖通过电气其他条件"
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气开关暖通条件.ForeColor.Name == "Black" || button电气开关暖通条件.ForeColor.Name == "ControlText")
            {
                button电气开关暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气开关暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ENtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收建筑底图_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收建筑底图.ForeColor.Name == "Black" || button电气收建筑底图.ForeColor.Name == "ControlText")
            {
                button电气收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收建筑电动卷帘门_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "电动卷帘门",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收建筑电动卷帘门.ForeColor.Name == "Black" || button电气收建筑电动卷帘门.ForeColor.Name == "ControlText")
            {
                button电气收建筑电动卷帘门.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收建筑电动卷帘门.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收建筑防火卷帘门_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "防火卷帘门",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收建筑防火卷帘门.ForeColor.Name == "Black" || button电气收建筑防火卷帘门.ForeColor.Name == "ControlText")
            {
                button电气收建筑防火卷帘门.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收建筑防火卷帘门.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收建筑电动排烟窗_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "电动排烟窗",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收建筑电动排烟窗.ForeColor.Name == "Black" || button电气收建筑电动排烟窗.ForeColor.Name == "ControlText")
            {
                button电气收建筑电动排烟窗.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收建筑电动排烟窗.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收建筑房间编号_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(房间编号)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收建筑房间编号.ForeColor.Name == "Black" || button电气收建筑房间编号.ForeColor.Name == "ControlText")
            {
                button电气收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收建筑吊顶高度_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收建筑吊顶高度.ForeColor.Name == "Black" || button电气收建筑吊顶高度.ForeColor.Name == "ControlText")
            {
                button电气收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收给排水消火栓_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "EQUIP_消火栓",
            "EQUIP-消火栓",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收给排水消火栓.ForeColor.Name == "Black" || button电气收给排水消火栓.ForeColor.Name == "ControlText")
            {
                button电气收给排水消火栓.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收给排水消火栓.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收给排水动力条件_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(给排水过电气动力条件)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收给排水动力条件.ForeColor.Name == "Black" || button电气收给排水动力条件.ForeColor.Name == "ControlText")
            {
                button电气收给排水动力条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收给排水动力条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收给排水喷淋有关条件_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(给排水过电气喷淋条件)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收给排水喷淋有关条件.ForeColor.Name == "Black" || button电气收给排水喷淋有关条件.ForeColor.Name == "ControlText")
            {
                button电气收给排水喷淋有关条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收给排水喷淋有关条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收暖通文字条件_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "TJ(暖通过电气)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收暖通文字条件.ForeColor.Name == "Black" || button电气收暖通文字条件.ForeColor.Name == "ControlText")
            {
                button电气收暖通文字条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收暖通文字条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button电气收暖通其他条件_Click(object sender, EventArgs e)
        {
            List<string> EAtjtBtn = new List<string>
        {
            "暖通其他条件",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button电气收暖通其他条件.ForeColor.Name == "Black" || button电气收暖通其他条件.ForeColor.Name == "ControlText")
            {
                button电气收暖通其他条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button电气收暖通其他条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in EAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        #endregion

        #region 给排水
        public void button给排水开闭工艺条件_Click(object sender, EventArgs e)
        {
            //    List<string> PGtjtBtn = new List<string>
            //{
            //    "TJ(给排水专业S)",
            //    "SB(工艺设备)",

            //    "S_工艺设备",
            //    "SB(设备名称)",
            //    "SB(设备外框)",
            //    "QY",
            //};
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水开闭工艺条件.ForeColor.Name == "Black" || button给排水开闭工艺条件.ForeColor.Name == "ControlText")
            {
                button给排水开闭工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水开闭工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            //foreach (var item in PGtjtBtn)
            //{
            //    VariableDictionary.selectTjtLayer.Remove(item);
            //}
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收工艺条件_Click(object sender, EventArgs e)
        {
            List<string> PGtjtBtn = new List<string>
        {
            "TJ(给排水专业S)",
            "EQUIP_地漏",
            "EQUIP_给水",
            "SB(工艺设备)",
            "SB(设备名称)",
            "S_设备名称",
            "SB(设备外框)",
            "QY",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收工艺条件.ForeColor.Name == "Black" || button给排水收工艺条件.ForeColor.Name == "ControlText")
            {
                button给排水收工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in PGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收工艺设备_Click(object sender, EventArgs e)
        {
            List<string> PGtjtBtn = new List<string>
        {
            "SB(工艺设备)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收工艺设备.ForeColor.Name == "Black" || button给排水收工艺设备.ForeColor.Name == "ControlText")
            {
                button给排水收工艺设备.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收工艺设备.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收工艺洁净区域划分_Click(object sender, EventArgs e)
        {
            List<string> PGtjtBtn = new List<string>
        {
            "QY",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收工艺洁净区域划分.ForeColor.Name == "Black" || button给排水收工艺洁净区域划分.ForeColor.Name == "ControlText")
            {
                button给排水收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收工艺设备名称_Click(object sender, EventArgs e)
        {
            List<string> PGtjtBtn = new List<string>
        {
            "SB(设备名称)",
            "S_设备名称",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收工艺设备名称.ForeColor.Name == "Black" || button给排水收工艺设备名称.ForeColor.Name == "ControlText")
            {
                button给排水收工艺设备名称.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收工艺设备名称.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }

        public void button给排水收工艺设备外框_Click(object sender, EventArgs e)
        {
            List<string> PGtjtBtn = new List<string>
        {
            "SB(设备外框)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收工艺设备外框.ForeColor.Name == "Black" || button给排水收工艺设备外框.ForeColor.Name == "ControlText")
            {
                button给排水收工艺设备外框.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收工艺设备外框.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }

        public void button给排水收建筑底图_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收建筑底图.ForeColor.Name == "Black" || button给排水收建筑底图.ForeColor.Name == "ControlText")
            {
                button给排水收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收建筑房间编号_Click(object sender, EventArgs e)
        {
            List<string> PAtjtBtn = new List<string>
        {
             "TJ(房间编号)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收建筑房间编号.ForeColor.Name == "Black" || button给排水收建筑房间编号.ForeColor.Name == "ControlText")
            {
                button给排水收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收建筑吊顶高度_Click(object sender, EventArgs e)
        {
            List<string> PAtjtBtn = new List<string>
        {
             "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收建筑吊顶高度.ForeColor.Name == "Black" || button给排水收建筑吊顶高度.ForeColor.Name == "ControlText")
            {
                button给排水收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水开闭建筑条件_Click(object sender, EventArgs e)
        {
            List<string> PAtjtBtn = new List<string>
        {
             "TJ(房间编号)",
             "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水开闭建筑条件.ForeColor.Name == "Black" || button给排水开闭建筑条件.ForeColor.Name == "ControlText")
            {
                button给排水开闭建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button给排水收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button给排水收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button给排水收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水开闭建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button给排水收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button给排水收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button给排水收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收暖通过条件_Click(object sender, EventArgs e)
        {
            List<string> PNtjtBtn = new List<string>
        {
             "TJ(暖通过给排水)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收暖通过条件.ForeColor.Name == "Black" || button给排水收暖通过条件.ForeColor.Name == "ControlText")
            {
                button给排水收暖通过条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水收暖通过条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水收暖通文字条件_Click(object sender, EventArgs e)
        {
            List<string> PNtjtBtn = new List<string>
        {
             "TJ(暖通过给排水)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水收暖通文字条件.ForeColor.Name == "Black" || button给排水收暖通文字条件.ForeColor.Name == "ControlText")
            {
                button给排水收暖通文字条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;

            }
            else
            {
                button给排水收暖通文字条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in PNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button给排水开闭暖通条件_Click(object sender, EventArgs e)
        {
            List<string> ENtjtBtn = new List<string>
        {
           "TJ(暖通过给排水)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button给排水开闭暖通条件.ForeColor.Name == "Black" || button给排水开闭暖通条件.ForeColor.Name == "ControlText")
            {
                button给排水开闭暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button给排水收暖通过条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button给排水收暖通文字条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button给排水开闭暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button给排水收暖通过条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button给排水收暖通文字条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ENtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        #endregion

        #region 自控

        public void button自控开关工艺条件_Click(object sender, EventArgs e)
        {
            //    List<string> ZGtjtBtn = new List<string>
            //{
            //    "EQUIP-通讯",
            //    "EQUIP-安防",
            //    "SB(工艺设备)",
            //    "SB(设备名称)",
            //    "SB(设备外框)",
            //    "QY",
            //};
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控开关工艺条件.ForeColor.Name == "Black" || button自控开关工艺条件.ForeColor.Name == "ControlText")
            {
                button自控开关工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控开关工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ZKtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控开关给排水条件_Click(object sender, EventArgs e)
        {
            List<string> ZPtjtBtn = new List<string>
                {
                    "EQUIP-通讯",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控开关给排水条件.ForeColor.Name == "Black" || button自控开关给排水条件.ForeColor.Name == "ControlText")
            {
                button自控开关给排水条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控开关给排水条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZPtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控开关暖通条件_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "TJ(暖通过自控)",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控开关暖通条件.ForeColor.Name == "Black" || button自控开关暖通条件.ForeColor.Name == "ControlText")
            {
                button自控开关暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控开关暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收工艺通讯条件_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "EQUIP-安防",
            "EQUIP-通讯",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收工艺通讯条件.ForeColor.Name == "Black" || button自控收工艺通讯条件.ForeColor.Name == "ControlText")
            {
                button自控收工艺通讯条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button自控收工艺安防条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收工艺通讯条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button自控收工艺安防条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收工艺安防条件_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "EQUIP-安防",
            "EQUIP-通讯",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收工艺安防条件.ForeColor.Name == "Black" || button自控收工艺安防条件.ForeColor.Name == "ControlText")
            {
                button自控收工艺安防条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button自控收工艺通讯条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收工艺安防条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button自控收工艺通讯条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收工艺设备名称_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "SB(设备名称)",
            "S_设备名称",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收工艺设备名称.ForeColor.Name == "Black" || button自控收工艺设备名称.ForeColor.Name == "ControlText")
            {
                button自控收工艺设备名称.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收工艺设备名称.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收工艺设备外框_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "SB(设备外框)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收工艺设备外框.ForeColor.Name == "Black" || button自控收工艺设备外框.ForeColor.Name == "ControlText")
            {
                button自控收工艺设备外框.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收工艺设备外框.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收工艺洁净区域划分_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "QY",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收工艺洁净区域划分.ForeColor.Name == "Black" || button自控收工艺洁净区域划分.ForeColor.Name == "ControlText")
            {
                button自控收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控开关建筑条件_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "TJ(房间编号)",
             "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控开关建筑条件.ForeColor.Name == "Black" || button自控开关建筑条件.ForeColor.Name == "ControlText")
            {
                button自控开关建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控开关建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收建筑底图_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收建筑底图.ForeColor.Name == "Black" || button自控收建筑底图.ForeColor.Name == "ControlText")
            {
                button自控收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收建筑房间编号_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "TJ(房间编号)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收建筑房间编号.ForeColor.Name == "Black" || button自控收建筑房间编号.ForeColor.Name == "ControlText")
            {
                button自控收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收建筑吊顶高度_Click(object sender, EventArgs e)
        {
            List<string> ZGtjtBtn = new List<string>
        {
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收建筑吊顶高度.ForeColor.Name == "Black" || button自控收建筑吊顶高度.ForeColor.Name == "ControlText")
            {
                button自控收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收给排水通讯条件_Click(object sender, EventArgs e)
        {
            List<string> ZPtjtBtn = new List<string>
                {
                    "EQUIP-通讯",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收给排水通讯条件.ForeColor.Name == "Black" || button自控收给排水通讯条件.ForeColor.Name == "ControlText")
            {
                button自控收给排水通讯条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收给排水通讯条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZPtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通文字条件_Click(object sender, EventArgs e)
        {
            List<string> ZPtjtBtn = new List<string>
                {
                    "TJ(暖通过自控)",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通文字条件.ForeColor.Name == "Black" || button自控收暖通文字条件.ForeColor.Name == "ControlText")
            {
                button自控收暖通文字条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通文字条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZPtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通高中效过滤排风_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "暖通专业原有图层",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通高中效过滤排风.ForeColor.Name == "Black" || button自控收暖通高中效过滤排风.ForeColor.Name == "ControlText")
            {
                button自控收暖通高中效过滤排风.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通高中效过滤排风.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通风阀_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "暖通专业原有图层",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通风阀.ForeColor.Name == "Black" || button自控收暖通风阀.ForeColor.Name == "ControlText")
            {
                button自控收暖通风阀.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通风阀.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通FFU_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "暖通专业原有图层",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通FFU.ForeColor.Name == "Black" || button自控收暖通FFU.ForeColor.Name == "ControlText")
            {
                button自控收暖通FFU.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通FFU.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通VAV阀_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "暖通专业原有图层",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通VAV阀.ForeColor.Name == "Black" || button自控收暖通VAV阀.ForeColor.Name == "ControlText")
            {
                button自控收暖通VAV阀.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通VAV阀.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通空调机组_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "暖通专业原有图层",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通空调机组.ForeColor.Name == "Black" || button自控收暖通空调机组.ForeColor.Name == "ControlText")
            {
                button自控收暖通空调机组.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通空调机组.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通流程图_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "暖通专业原有图层",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通流程图.ForeColor.Name == "Black" || button自控收暖通流程图.ForeColor.Name == "ControlText")
            {
                button自控收暖通流程图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通流程图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收暖通压差梯度_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "暖通专业原有图层",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收暖通压差梯度.ForeColor.Name == "Black" || button自控收暖通压差梯度.ForeColor.Name == "ControlText")
            {
                button自控收暖通压差梯度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收暖通压差梯度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        public void button自控开关电气条件_Click(object sender, EventArgs e)
        {
            List<string> ZNtjtBtn = new List<string>
                {
                    "TEL_CABINET",
                    "EQUIP-照明",
                    "WIRE-厂区消防",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控开关电气条件.ForeColor.Name == "Black" || button自控开关电气条件.ForeColor.Name == "ControlText")
            {
                button自控开关电气条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button自控收电气机房配电柜.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button自控收电气路灯.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button自控收电气厂区消防线.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控开关电气条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button自控收电气机房配电柜.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button自控收电气路灯.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button自控收电气厂区消防线.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收电气机房配电柜_Click(object sender, EventArgs e)
        {
            List<string> ZEtjtBtn = new List<string>
                {
                    "TEL_CABINET",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收电气机房配电柜.ForeColor.Name == "Black" || button自控收电气机房配电柜.ForeColor.Name == "ControlText")
            {
                button自控收电气机房配电柜.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收电气机房配电柜.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收电气路灯_Click(object sender, EventArgs e)
        {
            List<string> ZEtjtBtn = new List<string>
                {
                    "EQUIP-照明",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收电气路灯.ForeColor.Name == "Black" || button自控收电气路灯.ForeColor.Name == "ControlText")
            {
                button自控收电气路灯.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收电气路灯.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button自控收电气厂区消防线_Click(object sender, EventArgs e)
        {
            List<string> ZEtjtBtn = new List<string>
                {
                    "WIRE-厂区消防",
                };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button自控收电气厂区消防线.ForeColor.Name == "Black" || button自控收电气厂区消防线.ForeColor.Name == "ControlText")
            {
                button自控收电气厂区消防线.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button自控收电气厂区消防线.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ZEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        #endregion

        #region 结构
        public void button结构开关给排水条件_Click(object sender, EventArgs e)
        {
            List<string> SPtjtBtn = new List<string>
        {
            "TJ(给排水过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构开关给排水条件.ForeColor.Name == "Black" || button结构开关给排水条件.ForeColor.Name == "ControlText")
            {
                button结构开关给排水条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收给排水设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收给排水套管.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关给排水条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收给排水设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收给排水套管.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SPtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收给排水设备基础_Click(object sender, EventArgs e)
        {

            List<string> SPtjtBtn = new List<string>
        {
            "TJ(给排水过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收给排水设备基础.ForeColor.Name == "Black" || button结构收给排水设备基础.ForeColor.Name == "ControlText")
            {
                button结构开关给排水条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收给排水设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收给排水套管.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关给排水条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收给排水设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收给排水套管.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SPtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收给排水套管_Click(object sender, EventArgs e)
        {

            List<string> SPtjtBtn = new List<string>
        {
            "TJ(给排水过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收给排水套管.ForeColor.Name == "Black" || button结构收给排水套管.ForeColor.Name == "ControlText")
            {
                button结构开关给排水条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收给排水设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收给排水套管.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关给排水条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收给排水设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收给排水套管.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SPtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构开关暖通条件_Click(object sender, EventArgs e)
        {
            List<string> SNtjtBtn = new List<string>
        {
            "TJ(暖通过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构开关暖通条件.ForeColor.Name == "Black" || button结构开关暖通条件.ForeColor.Name == "ControlText")
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收暖通楼板洞口_Click(object sender, EventArgs e)
        {
            List<string> SNtjtBtn = new List<string>
        {
            "TJ(暖通过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收暖通楼板洞口.ForeColor.Name == "Black" || button结构收暖通楼板洞口.ForeColor.Name == "ControlText")
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收暖通地沟_Click(object sender, EventArgs e)
        {
            List<string> SNtjtBtn = new List<string>
        {
            "TJ(暖通过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收暖通地沟.ForeColor.Name == "Black" || button结构收暖通地沟.ForeColor.Name == "ControlText")
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收暖通设备基础_Click(object sender, EventArgs e)
        {
            List<string> SNtjtBtn = new List<string>
        {
            "TJ(暖通过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收暖通设备基础.ForeColor.Name == "Black" || button结构收暖通设备基础.ForeColor.Name == "ControlText")
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收暖通吊挂风机及荷载条件_Click(object sender, EventArgs e)
        {
            List<string> SNtjtBtn = new List<string>
        {
            "TJ(暖通过结构)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收暖通吊挂风机及荷载条件.ForeColor.Name == "Black" || button结构收暖通吊挂风机及荷载条件.ForeColor.Name == "ControlText")
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通地沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通设备基础.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收暖通吊挂风机及荷载条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SNtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构开关电气条件_Click(object sender, EventArgs e)
        {
            List<string> SEtjtBtn = new List<string>
        {
            "TJ(电气过结构)",
            "TJ(电气过结构楼板洞D)",
            "TJ(电气过结构电缆沟D)",
            "TJ(电气过结构活荷载D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构开关电气条件.ForeColor.Name == "Black" || button结构开关电气条件.ForeColor.Name == "ControlText")
            {
                button结构开关电气条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收电气楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收电气电缆沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收电气荷载条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关电气条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收电气楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收电气电缆沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收电气荷载条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收电气楼板洞口_Click(object sender, EventArgs e)
        {
            List<string> SEtjtBtn = new List<string>
        {
            "TJ(电气过结构楼板洞D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收电气楼板洞口.ForeColor.Name == "Black" || button结构收电气楼板洞口.ForeColor.Name == "ControlText")
            {
                button结构收电气楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收电气楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收电气电缆沟_Click(object sender, EventArgs e)
        {
            List<string> SEtjtBtn = new List<string>
        {
            "TJ(电气过结构电缆沟D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收电气电缆沟.ForeColor.Name == "Black" || button结构收电气电缆沟.ForeColor.Name == "ControlText")
            {
                button结构收电气电缆沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收电气电缆沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收电气荷载条件_Click(object sender, EventArgs e)
        {
            List<string> SEtjtBtn = new List<string>
        {

            "TJ(电气过结构活荷载D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收电气荷载条件.ForeColor.Name == "Black" || button结构收电气荷载条件.ForeColor.Name == "ControlText")
            {
                button结构收电气荷载条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收电气荷载条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构开关自控条件_Click(object sender, EventArgs e)
        {
            List<string> SZtjtBtn = new List<string>
        {
            "TJ(自控过结构)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构开关自控条件.ForeColor.Name == "Black" || button结构开关自控条件.ForeColor.Name == "ControlText")
            {
                button结构开关自控条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收自控楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关自控条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收自控楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SZtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收自控楼板洞口_Click(object sender, EventArgs e)
        {
            List<string> SZtjtBtn = new List<string>
        {
            "TJ(自控过结构)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收自控楼板洞口.ForeColor.Name == "Black" || button结构收自控楼板洞口.ForeColor.Name == "ControlText")
            {
                button结构开关自控条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收自控楼板洞口.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关自控条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收自控楼板洞口.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SZtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收工艺设备名称_Click(object sender, EventArgs e)
        {

            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收工艺设备名称.ForeColor.Name == "Black" || button结构收工艺设备名称.ForeColor.Name == "ControlText")
            {
                button结构收工艺设备名称.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收工艺设备名称.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.SBtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收工艺设备外框_Click(object sender, EventArgs e)
        {
            List<string> SGtjtBtn = new List<string>
        {
            "SB(设备外框)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收工艺设备外框.ForeColor.Name == "Black" || button结构收工艺设备外框.ForeColor.Name == "ControlText")
            {
                button结构收工艺设备外框.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收工艺设备外框.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收工艺设备_Click(object sender, EventArgs e)
        {
            List<string> SGtjtBtn = new List<string>
        {
            "SB(工艺设备)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收工艺设备.ForeColor.Name == "Black" || button结构收工艺设备.ForeColor.Name == "ControlText")
            {
                button结构收工艺设备.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收工艺设备.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收工艺过结构条件_Click(object sender, EventArgs e)
        {
            List<string> SGtjtBtn = new List<string>
        {
            "TJ(结构专业JG)",
            "SB(工艺设备)",
            "SB(设备名称)",
            "S_设备名称",
            "SB(设备外框)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收工艺过结构条件.ForeColor.Name == "Black" || button结构收工艺过结构条件.ForeColor.Name == "ControlText")
            {
                button结构收工艺过结构条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收工艺过结构条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in SGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构开关工艺条件_Click(object sender, EventArgs e)
        {
            //    List<string> SGtjtBtn = new List<string>
            //{
            //    "TJ(结构专业JG)",
            //    "SB(工艺设备)",
            //    "SB(设备名称)",
            //    "SB(设备外框)",
            //};
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构开关工艺条件.ForeColor.Name == "Black" || button结构开关工艺条件.ForeColor.Name == "ControlText")
            {
                button结构开关工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收工艺过结构条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收工艺设备.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收工艺设备外框.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收工艺设备名称.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构开关工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收工艺过结构条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收工艺设备.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收工艺设备外框.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收工艺设备名称.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收建筑底图_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收建筑底图.ForeColor.Name == "Black" || button结构收建筑底图.ForeColor.Name == "ControlText")
            {
                button结构收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button结构收建筑房间编号_Click(object sender, EventArgs e)
        {
            List<string> SAtjtBtn = new List<string>
        {
            "TJ(房间编号)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button结构收建筑房间编号.ForeColor.Name == "Black" || button结构收建筑房间编号.ForeColor.Name == "ControlText")
            {
                button结构收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button结构收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in SAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button开闭建筑条件_Click(object sender, EventArgs e)
        {
            //    List<string> SAtjtBtn = new List<string>
            //{
            //    "TJ(房间编号)",
            //};
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button开闭建筑条件.ForeColor.Name == "Black" || button开闭建筑条件.ForeColor.Name == "ControlText")
            {
                button开闭建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button结构收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button开闭建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button结构收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            //foreach (var item in SAtjtBtn)
            //{
            //    VariableDictionary.selectTjtLayer.Add(item);
            //}
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        #endregion

        #region 暖通

        public void button_测量房间面积_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnBlockLayer = "暖通房间面积";//设置为被插入的图层名
            VariableDictionary.buttonText = "暖通房间面积";
            VariableDictionary.layerColorIndex = 2;//设置图层颜色
            VariableDictionary.textBoxScale = Convert.ToDouble(textBox_Scale_比例.Text);//设置文本的比例

            Env.Document.SendStringToExecute("AreaByPoints ", false, false, false);
        }

        public void button_暖通开关工艺条件_Click(object sender, EventArgs e)
        {
            //    List<string> NGtjtBtn = new List<string>
            //{
            //    "TJ(暖通专业N)",
            //    "SB(工艺设备)",
            //    "SB(设备名称)",
            //    "SB(设备外框)",
            //    "QY",
            //};
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button_暖通开关工艺条件.ForeColor.Name == "Black" || button_暖通开关工艺条件.ForeColor.Name == "ControlText")
            {
                button_暖通开关工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收工艺设备.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收工艺设备名称.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收工艺设备外框.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_暖通开关工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收工艺设备.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收工艺设备名称.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收工艺设备外框.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收工艺条件_Click(object sender, EventArgs e)
        {
            List<string> NGtjtBtn = new List<string>
        {
            "TJ(暖通专业N)",
            "SB(工艺设备)",
            "SB(设备名称)",
            "S_设备名称",
            "SB(设备外框)",
            "QY",
            "SB",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收工艺条件.ForeColor.Name == "Black" || button暖通收工艺条件.ForeColor.Name == "ControlText")
            {
                button暖通收工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in NGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收工艺设备_Click(object sender, EventArgs e)
        {
            List<string> NGtjtBtn = new List<string>
        {
            "SB(工艺设备)",
            "SB",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收工艺设备.ForeColor.Name == "Black" || button暖通收工艺设备.ForeColor.Name == "ControlText")
            {
                button暖通收工艺设备.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收工艺设备.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in NGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收工艺设备名称_Click(object sender, EventArgs e)
        {
            List<string> NGtjtBtn = new List<string>
        {
            "SB(设备名称)",
            "S_设备名称",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收工艺设备名称.ForeColor.Name == "Black" || button暖通收工艺设备名称.ForeColor.Name == "ControlText")
            {
                button暖通收工艺设备名称.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收工艺设备名称.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in NGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收工艺设备外框_Click(object sender, EventArgs e)
        {
            List<string> NGtjtBtn = new List<string>
        {
            "SB(设备外框)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收工艺设备外框.ForeColor.Name == "Black" || button暖通收工艺设备外框.ForeColor.Name == "ControlText")
            {
                button暖通收工艺设备外框.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收工艺设备外框.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in NGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收工艺洁净区域划分_Click(object sender, EventArgs e)
        {
            List<string> NGtjtBtn = new List<string>
        {
            "QY",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收工艺洁净区域划分.ForeColor.Name == "Black" || button暖通收工艺洁净区域划分.ForeColor.Name == "ControlText")
            {
                button暖通收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in NGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通开关建筑条件_Click(object sender, EventArgs e)
        {
            //    List<string> NAtjtBtn = new List<string>
            //{
            //    "TJ(房间编号)",
            //    "TJ(建筑专业J)Y",
            //    "TJ(建筑吊顶)",
            //};
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通开关建筑条件.ForeColor.Name == "Black" || button暖通开关建筑条件.ForeColor.Name == "ControlText")
            {
                button暖通开关建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button暖通收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通开关建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button暖通收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            //foreach (var item in NAtjtBtn)
            //{
            //    VariableDictionary.selectTjtLayer.Add(item);
            //}
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收建筑底图_Click(object sender, EventArgs e)
        {
            List<string> NAtjtBtn = new List<string>
            {
                "TJ(房间编号)",
                "TJ(建筑专业J)Y",
                "TJ(建筑吊顶)",
            };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收建筑底图.ForeColor.Name == "Black" || button暖通收建筑底图.ForeColor.Name == "ControlText")
            {
                button暖通收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();

            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收建筑房间编号_Click(object sender, EventArgs e)
        {
            List<string> NAtjtBtn = new List<string>
        {
            "TJ(房间编号)",

        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收建筑房间编号.ForeColor.Name == "Black" || button暖通收建筑房间编号.ForeColor.Name == "ControlText")
            {
                button暖通收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in NAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button暖通收建筑吊顶高度_Click(object sender, EventArgs e)
        {
            List<string> NAtjtBtn = new List<string>
        {
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button暖通收建筑吊顶高度.ForeColor.Name == "Black" || button暖通收建筑吊顶高度.ForeColor.Name == "ControlText")
            {
                button暖通收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button暖通收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in NAtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        #endregion

        #region 建筑

        public void button建筑开闭工艺条件_Click(object sender, EventArgs e)
        {
            //List<string> AGtjtBtn = new List<string>();
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑开闭工艺条件.ForeColor.Name == "Black" || button建筑开闭工艺条件.ForeColor.Name == "ControlText")
            {
                button建筑开闭工艺条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑开闭工艺条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);

        }

        public void button建筑收工艺房间编号_Click(object sender, EventArgs e)
        {
            List<string> AGtjtBtn = new List<string>
        {
            "TJ(房间编号)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收工艺房间编号.ForeColor.Name == "Black" || button建筑收工艺房间编号.ForeColor.Name == "ControlText")
            {
                button建筑收工艺房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收工艺房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收工艺吊顶高度_Click(object sender, EventArgs e)
        {
            List<string> AGtjtBtn = new List<string>
        {
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收工艺吊顶高度.ForeColor.Name == "Black" || button建筑收工艺吊顶高度.ForeColor.Name == "ControlText")
            {
                button建筑收工艺吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收工艺吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收工艺洁净区域划分_Click(object sender, EventArgs e)
        {
            List<string> AGtjtBtn = new List<string>
        {
            "QY",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收工艺洁净区域划分.ForeColor.Name == "Black" || button建筑收工艺洁净区域划分.ForeColor.Name == "ControlText")
            {
                button建筑收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收工艺洁净区域划分.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收工艺过建筑条件_Click(object sender, EventArgs e)
        {
            List<string> AGtjtBtn = new List<string>
        {
            "TJ(建筑专业J)",
            "TJ(建筑专业J)Y",
            "TJ(建筑专业J)N",
            "TJ(建筑吊顶)",
            "QY",
            "TJ(房间编号)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收工艺过建筑条件.ForeColor.Name == "Black" || button建筑收工艺过建筑条件.ForeColor.Name == "ControlText")
            {
                button建筑收工艺过建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收工艺过建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in VariableDictionary.GYtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in AGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Remove(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }


        public void button建筑收建筑底图_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收建筑底图.ForeColor.Name == "Black" || button建筑收建筑底图.ForeColor.Name == "ControlText")
            {
                button建筑收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();

            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收建筑房间编号_Click(object sender, EventArgs e)
        {
            List<string> AGtjtBtn = new List<string>
        {
            "TJ(房间编号)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收建筑房间编号.ForeColor.Name == "Black" || button建筑收建筑房间编号.ForeColor.Name == "ControlText")
            {
                button建筑收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收建筑吊顶高度_Click(object sender, EventArgs e)
        {
            List<string> AGtjtBtn = new List<string>
        {
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收建筑吊顶高度.ForeColor.Name == "Black" || button建筑收建筑吊顶高度.ForeColor.Name == "ControlText")
            {
                button建筑收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑开闭建筑条件_Click(object sender, EventArgs e)
        {
            List<string> AGtjtBtn = new List<string>
        {
            "TJ(建筑专业J)",
            "TJ(建筑专业J)Y",
            "TJ(建筑专业J)N",
            "TJ(房间编号)",
            "TJ(建筑吊顶)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑开闭建筑条件.ForeColor.Name == "Black" || button建筑开闭建筑条件.ForeColor.Name == "ControlText")
            {
                button建筑开闭建筑条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收建筑房间编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收建筑底图.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑开闭建筑条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收建筑吊顶高度.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收建筑房间编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收建筑底图.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AGtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            foreach (var item in VariableDictionary.ApmtTjBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }


        public void button建筑收结构柱_Click(object sender, EventArgs e)
        {
            List<string> AStjtBtn = new List<string>
        {
            "COLU",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收结构柱.ForeColor.Name == "Black" || button建筑收结构柱.ForeColor.Name == "ControlText")
            {
                button建筑收结构柱.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收结构柱.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AStjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收结构楼板洞_Click(object sender, EventArgs e)
        {
            List<string> AStjtBtn = new List<string>
        {
            "HOLE",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收结构楼板洞.ForeColor.Name == "Black" || button建筑收结构楼板洞.ForeColor.Name == "ControlText")
            {
                button建筑收结构楼板洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收结构楼板洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AStjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收结构混凝土墙_Click(object sender, EventArgs e)
        {
            List<string> AStjtBtn = new List<string>
        {
            "WALL-C",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收结构混凝土墙.ForeColor.Name == "Black" || button建筑收结构混凝土墙.ForeColor.Name == "ControlText")
            {
                button建筑收结构混凝土墙.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收结构混凝土墙.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AStjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑开闭结构条件_Click(object sender, EventArgs e)
        {
            List<string> AStjtBtn = new List<string>
        {
            "WALL-C",
            "HOLE",
            "COLU",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑开闭结构条件.ForeColor.Name == "Black" || button建筑开闭结构条件.ForeColor.Name == "ControlText")
            {
                button建筑收结构柱.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收结构楼板洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收结构混凝土墙.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑开闭结构条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收结构柱.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收结构楼板洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收结构混凝土墙.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑开闭结构条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AStjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收给排水孔洞_Click(object sender, EventArgs e)
        {
            List<string> APtjtBtn = new List<string>
        {
            "TJ(给排水过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收给排水孔洞.ForeColor.Name == "Black" || button建筑收给排水孔洞.ForeColor.Name == "ControlText")
            {
                button建筑收给排水孔洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收给排水孔洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in APtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收给排水消火栓_Click(object sender, EventArgs e)
        {
            List<string> APtjtBtn = new List<string>
        {
            "EQUIP_消火栓",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收给排水消火栓.ForeColor.Name == "Black" || button建筑收给排水消火栓.ForeColor.Name == "ControlText")
            {
                button建筑收给排水消火栓.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收给排水消火栓.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in APtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑开闭给排水条件_Click(object sender, EventArgs e)
        {
            List<string> APtjtBtn = new List<string>
        {
            "TJ(给排水过建筑)",
            "EQUIP_消火栓",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑开闭给排水条件.ForeColor.Name == "Black" || button建筑开闭给排水条件.ForeColor.Name == "ControlText")
            {
                button建筑开闭给排水条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收给排水消火栓.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收给排水孔洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑开闭给排水条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收给排水消火栓.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收给排水孔洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in APtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收暖通孔洞_Click(object sender, EventArgs e)
        {
            List<string> ANtjtBtn = new List<string>
        {
            "TJ(暖通过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收暖通孔洞.ForeColor.Name == "Black" || button建筑收暖通孔洞.ForeColor.Name == "ControlText")
            {
                button建筑收暖通孔洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收暖通孔洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ANtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收暖通自然排烟窗_Click(object sender, EventArgs e)
        {
            List<string> ANtjtBtn = new List<string>
        {
            "TJ(暖通过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收暖通自然排烟窗.ForeColor.Name == "Black" || button建筑收暖通自然排烟窗.ForeColor.Name == "ControlText")
            {
                button建筑收暖通自然排烟窗.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收暖通自然排烟窗.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ANtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收暖通夹墙_Click(object sender, EventArgs e)
        {
            List<string> ANtjtBtn = new List<string>
        {
            "WALL-PARAPET",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收暖通夹墙.ForeColor.Name == "Black" || button建筑收暖通夹墙.ForeColor.Name == "ControlText")
            {
                button建筑收暖通夹墙.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收暖通夹墙.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ANtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收暖通排水沟_Click(object sender, EventArgs e)
        {
            List<string> ANtjtBtn = new List<string>
        {
            "TJ(暖通过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收暖通排水沟.ForeColor.Name == "Black" || button建筑收暖通排水沟.ForeColor.Name == "ControlText")
            {
                button建筑收暖通排水沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收暖通排水沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ANtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑开闭暖通条件_Click(object sender, EventArgs e)
        {
            List<string> ANtjtBtn = new List<string>
        {
            "TJ(暖通过建筑)",
            "WALL-PARAPET",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑开闭暖通条件.ForeColor.Name == "Black" || button建筑开闭暖通条件.ForeColor.Name == "ControlText")
            {
                button建筑开闭暖通条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收暖通排水沟.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收暖通夹墙.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收暖通自然排烟窗.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收暖通孔洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑开闭暖通条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收暖通排水沟.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收暖通夹墙.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收暖通自然排烟窗.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收暖通孔洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in ANtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收电气条件_Click(object sender, EventArgs e)
        {
            List<string> AEtjtBtn = new List<string>
        {
            "TJ(电气过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收电气条件.ForeColor.Name == "Black" || button建筑收电气条件.ForeColor.Name == "ControlText")
            {
                button建筑收电气条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收电气条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收电气夹墙_Click(object sender, EventArgs e)
        {
            List<string> AEtjtBtn = new List<string>
        {

            "TJ(电气过建筑夹墙D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收电气夹墙.ForeColor.Name == "Black" || button建筑收电气夹墙.ForeColor.Name == "ControlText")
            {
                button建筑收电气夹墙.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收电气夹墙.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收电气孔洞_Click(object sender, EventArgs e)
        {
            List<string> AEtjtBtn = new List<string>
        {
            "TJ(电气过建筑孔洞D)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收电气孔洞.ForeColor.Name == "Black" || button建筑收电气孔洞.ForeColor.Name == "ControlText")
            {
                button建筑收电气孔洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收电气孔洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑开闭电气条件_Click(object sender, EventArgs e)
        {
            List<string> AEtjtBtn = new List<string>
        {
            "TJ(电气过建筑孔洞D)",
            "TJ(电气过建筑夹墙D)",
            "TJ(电气过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑开闭电气条件.ForeColor.Name == "Black" || button建筑开闭电气条件.ForeColor.Name == "ControlText")
            {
                button建筑收电气条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑开闭电气条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收电气夹墙.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收电气孔洞.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收电气条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑开闭电气条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收电气夹墙.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收电气孔洞.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AEtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑收自控条件_Click(object sender, EventArgs e)
        {
            List<string> AZtjtBtn = new List<string>
        {
            "TJ(自控过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑收自控条件.ForeColor.Name == "Black" || button建筑收自控条件.ForeColor.Name == "ControlText")
            {
                button建筑收自控条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑收自控条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AZtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        public void button建筑开闭自控条件_Click(object sender, EventArgs e)
        {
            List<string> AZtjtBtn = new List<string>
        {
            "TJ(自控过建筑)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button建筑开闭自控条件.ForeColor.Name == "Black" || button建筑开闭自控条件.ForeColor.Name == "ControlText")
            {
                button建筑开闭自控条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
                button建筑收自控条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button建筑开闭自控条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
                button建筑收自控条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in AZtjtBtn)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }


        #endregion




        public void button_给排水过结构矩形开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.textbox_Height = Convert.ToDouble(textBox_给排水过结构矩形开洞Y.Text);
            VariableDictionary.textbox_Width = Convert.ToDouble(textBox_给排水过结构矩形开洞X.Text);
            VariableDictionary.btnBlockLayer = "TJ(给排水过结构)";
            VariableDictionary.buttonText = "PTJ_矩形开洞";
            VariableDictionary.layerColorIndex = 7;
            VariableDictionary.btnFileName = "TJ(给排水过结构)";
            VariableDictionary.layerName = "TJ(给排水过结构)";
            VariableDictionary.textbox_RecPlus_Text = textBox_给排水过结构矩形外扩.Text;
            if (Convert.ToDouble(VariableDictionary.textbox_Height) > 0 && Convert.ToDouble(VariableDictionary.textbox_Width) > 0)
            {
                recAndMRec = 0;
                Env.Document.SendStringToExecute("DrawRec ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("Rec2PolyLine ", false, false, false);
            }
        }

        public void button_给排水过结构直径开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(给排水过结构)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "PTJ_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(给排水过结构)";
            VariableDictionary.layerName = "TJ(给排水过结构)";
            VariableDictionary.textBox_S_CirDiameter = Convert.ToDouble(textBox_给排水过结构直径开洞直径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_给排水过结构直径外扩.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 7;
            if (VariableDictionary.textBox_S_CirDiameter == 0)
            {
                Env.Document.SendStringToExecute("CirDiameter ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirDiameter_2 ", false, false, false);
            }
        }

        public void button_给排水过结构半径开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(给排水过结构)";
            VariableDictionary.buttonText = "PTJ_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(给排水过结构)";
            VariableDictionary.layerName = "TJ(给排水过结构)";
            VariableDictionary.textbox_S_Cirradius = Convert.ToDouble(textBox_给排水过结构半径开洞半径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_给排水过结构半径外扩.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 7;
            if (VariableDictionary.textbox_S_Cirradius == 0)
            {
                Env.Document.SendStringToExecute("CirRadius ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirRadius_2 ", false, false, false);
            }
        }

        public void button_自控过结构矩形开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.textbox_Height = Convert.ToDouble(textBox_自控过结构矩形开洞Y.Text);
            VariableDictionary.textbox_Width = Convert.ToDouble(textBox_自控过结构矩形开洞X.Text);
            VariableDictionary.btnBlockLayer = "TJ(自控过结构)";
            VariableDictionary.buttonText = "ZK_矩形开洞";
            VariableDictionary.layerColorIndex = 3;
            VariableDictionary.btnFileName = "TJ(自控过结构)";
            VariableDictionary.layerName = "TJ(自控过结构)";
            VariableDictionary.textbox_RecPlus_Text = textBox_自控过结构矩形开洞外扩.Text;
            if (Convert.ToDouble(VariableDictionary.textbox_Height) > 0 && Convert.ToDouble(VariableDictionary.textbox_Width) > 0)
            {
                recAndMRec = 0;
                Env.Document.SendStringToExecute("DrawRec ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("Rec2PolyLine ", false, false, false);
            }
        }

        public void button_自控过结构直径开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(自控过结构)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "ZK_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(自控过结构)";
            VariableDictionary.layerName = "TJ(自控过结构)";
            VariableDictionary.textBox_S_CirDiameter = Convert.ToDouble(textBox_自控过结构直径开洞直径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_自控过结构直径开洞外扩.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 3;
            if (VariableDictionary.textBox_S_CirDiameter == 0)
            {
                Env.Document.SendStringToExecute("CirDiameter ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirDiameter_2 ", false, false, false);
            }
        }

        public void button_自控过结构半径开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(自控过结构)";
            VariableDictionary.buttonText = "ZK_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(自控过结构)";
            VariableDictionary.textbox_S_Cirradius = Convert.ToDouble(textBox_自控过结构半径开洞半径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_自控过结构半径开洞外扩.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 3;
            if (VariableDictionary.textbox_S_Cirradius == 0)
            {
                Env.Document.SendStringToExecute("CirRadius ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirRadius_2 ", false, false, false);
            }
        }

        public void button_暖通_矩形开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.textbox_Height = Convert.ToDouble(textBox_暖通_矩形Y.Text);
            VariableDictionary.textbox_Width = Convert.ToDouble(textBox_暖通_矩形X.Text);
            VariableDictionary.btnBlockLayer = "TJ(暖通过结构)";
            VariableDictionary.buttonText = "NTJ_矩形开洞";
            VariableDictionary.layerColorIndex = 6;
            VariableDictionary.btnFileName = "TJ(暖通过结构)";
            VariableDictionary.layerName = "TJ(暖通过结构)";
            VariableDictionary.textbox_RecPlus_Text = textBox_矩形外扩值.Text;
            if (Convert.ToDouble(VariableDictionary.textbox_Height) > 0 && Convert.ToDouble(VariableDictionary.textbox_Width) > 0)
            {
                recAndMRec = 0;
                Env.Document.SendStringToExecute("DrawRec ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("Rec2PolyLine ", false, false, false);
            }

        }

        public void button_暖通_直径开圆洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(暖通过结构)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "NTJ_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(暖通过结构)";
            VariableDictionary.textBox_S_CirDiameter = Convert.ToDouble(textBox_暖通_直径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_暖通_直径外扩值.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 6;
            if (VariableDictionary.textBox_S_CirDiameter == 0)
            {
                Env.Document.SendStringToExecute("CirDiameter ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirDiameter_2 ", false, false, false);
            }
            ;
        }

        public void button_暖通_半径开圆洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(暖通过结构)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "NTJ_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(暖通过结构)";
            VariableDictionary.textbox_S_Cirradius = Convert.ToDouble(textBox_暖通_半径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_暖通_半径外扩值.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 6;
            if (VariableDictionary.textbox_S_Cirradius == 0)
            {
                Env.Document.SendStringToExecute("CirRadius ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirRadius_2 ", false, false, false);
            }
        }

        public void button_电气过结构矩形开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.textbox_Height = Convert.ToDouble(textBox_电气过结构矩形Y.Text);
            VariableDictionary.textbox_Width = Convert.ToDouble(textBox_电气过结构矩形X.Text);
            VariableDictionary.btnBlockLayer = "TJ(电气过结构楼板洞D)";
            VariableDictionary.buttonText = "ETJ_矩形开洞";
            VariableDictionary.layerColorIndex = 142;
            VariableDictionary.btnFileName = "TJ(电气过结构楼板洞D)";
            VariableDictionary.layerName = "TJ(电气过结构楼板洞D)";
            VariableDictionary.textbox_RecPlus_Text = textBox_电气过结构矩形外扩.Text;
            if (Convert.ToDouble(VariableDictionary.textbox_Height) > 0 && Convert.ToDouble(VariableDictionary.textbox_Width) > 0)
            {
                recAndMRec = 0;
                Env.Document.SendStringToExecute("DrawRec ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("Rec2PolyLine ", false, false, false);
            }
        }

        public void button_电气过结构半径开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(电气过结构楼板洞D)";
            VariableDictionary.buttonText = "ETJ_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(电气过结构楼板洞D)";
            VariableDictionary.layerName = "TJ(电气过结构楼板洞D)";
            VariableDictionary.textbox_S_Cirradius = Convert.ToDouble(textBox_电气过结构半径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_电气过结构半径外扩.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 142;
            if (VariableDictionary.textbox_S_Cirradius == 0)
            {
                Env.Document.SendStringToExecute("CirRadius ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirRadius_2 ", false, false, false);
            }
        }

        public void button_电气过结构直径开洞_Click(object sender, EventArgs e)
        {
            VariableDictionary.btnFileName = "TJ(电气过结构楼板洞D)";
            VariableDictionary.btnBlockLayer = VariableDictionary.btnFileName;
            VariableDictionary.buttonText = "ETJ_圆形开洞";
            VariableDictionary.btnBlockLayer = "TJ(电气过结构楼板洞D)";
            VariableDictionary.layerName = "TJ(电气过结构楼板洞D)";
            VariableDictionary.textBox_S_CirDiameter = Convert.ToDouble(textBox_电气过结构直径.Text);//拿到指定圆的直径
            VariableDictionary.textbox_CirPlus_Text = textBox_电气过结构直径外扩.Text;//拿到指定圆的外扩量
            VariableDictionary.layerColorIndex = 142;
            if (VariableDictionary.textBox_S_CirDiameter == 0)
            {
                Env.Document.SendStringToExecute("CirDiameter ", false, false, false);
            }
            else
            {
                Env.Document.SendStringToExecute("CirDiameter_2 ", false, false, false);
            }
        }


        public void button_绘图_Click(object sender, EventArgs e)
        {
            // 初始状态：隐藏TabPage（可选）
            //if (isTabPageVisible)
            //{
            //    tabCtl_Main.TabPages.Remove(linkedTabPage);
            //    isTabPageVisible = false; 
            //}
            //else
            //{
            //    tabCtl_Main.TabPages.Add(linkedTabPage);
            //    isTabPageVisible = true;
            //}
        }

        #region 图纸检查Tabitem

        private void button_保留共用条件_Click(object sender, EventArgs e)
        {
            //"TJ(共用条件)"
            List<string> GGTJ = new List<string>
        {
            "TJ(共用条件)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button_保留共用条件.ForeColor.Name == "Black" || button_保留共用条件.ForeColor.Name == "ControlText")
            {
                button_保留共用条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_保留共用条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in GGTJ)
            {
                VariableDictionary.allTjtLayer.Remove(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button_关闭共用条件_Click(object sender, EventArgs e)
        {
            //"TJ(共用条件)"
            List<string> GGTJ = new List<string>
        {
            "TJ(共用条件)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button_关闭共用条件.ForeColor.Name == "Black" || button_关闭共用条件.ForeColor.Name == "ControlText")
            {
                button_关闭共用条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_关闭共用条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in GGTJ)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button_检查共用条件_Click(object sender, EventArgs e)
        {
            //"TJ(共用条件)"
            List<string> GGTJ = new List<string>
        {
            "TJ(共用条件)",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button_检查共用条件.ForeColor.Name == "Black" || button_检查共用条件.ForeColor.Name == "ControlText")
            {
                button_检查共用条件.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_检查共用条件.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in GGTJ)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button_开关门窗编号_Click(object sender, EventArgs e)
        {
            List<string> MCBH = new List<string>
        {
            "WINDOW_TEXT",
        };
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;

            if (button_开关门窗编号.ForeColor.Name == "Black" || button_开关门窗编号.ForeColor.Name == "ControlText")
            {
                button_开关门窗编号.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button_开关门窗编号.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }

            NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            foreach (var item in MCBH)
            {
                VariableDictionary.selectTjtLayer.Add(item);
            }

            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        #endregion

        #region

        private void button纯化水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button纯化水.ForeColor.Name == "Black" || button纯化水.ForeColor.Name == "ControlText")
            {
                button纯化水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button纯化水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-PW");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button注射用水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button注射用水.ForeColor.Name == "Black" || button注射用水.ForeColor.Name == "ControlText")
            {
                button注射用水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button注射用水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-WFI");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button循环水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button循环水.ForeColor.Name == "Black" || button循环水.ForeColor.Name == "ControlText")
            {
                button循环水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button循环水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-CWS");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button软化水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button软化水.ForeColor.Name == "Black" || button软化水.ForeColor.Name == "ControlText")
            {
                button软化水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button软化水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-SW");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button乙二醇_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button乙二醇.ForeColor.Name == "Black" || button乙二醇.ForeColor.Name == "ControlText")
            {
                button乙二醇.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button乙二醇.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-EGS");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button凝结水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button凝结水.ForeColor.Name == "Black" || button凝结水.ForeColor.Name == "ControlText")
            {
                button凝结水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button凝结水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-SC");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button冷冻水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button冷冻水.ForeColor.Name == "Black" || button冷冻水.ForeColor.Name == "ControlText")
            {
                button冷冻水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button冷冻水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-RWS");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button热水_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button热水.ForeColor.Name == "Black" || button热水.ForeColor.Name == "ControlText")
            {
                button热水.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button热水.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-HWS");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void btn洁净压缩空气_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (btn洁净压缩空气.ForeColor.Name == "Black" || btn洁净压缩空气.ForeColor.Name == "ControlText")
            {
                btn洁净压缩空气.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                btn洁净压缩空气.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-CA");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void btn仪表压缩空气_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (btn仪表压缩空气.ForeColor.Name == "Black" || btn仪表压缩空气.ForeColor.Name == "ControlText")
            {
                btn仪表压缩空气.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                btn仪表压缩空气.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-IA");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button氧气_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button氧气.ForeColor.Name == "Black" || button氧气.ForeColor.Name == "ControlText")
            {
                button氧气.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button氧气.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-O2");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button二氧化碳_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button二氧化碳.ForeColor.Name == "Black" || button二氧化碳.ForeColor.Name == "ControlText")
            {
                button二氧化碳.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button二氧化碳.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-CO2");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button氮气_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button氮气.ForeColor.Name == "Black" || button氮气.ForeColor.Name == "ControlText")
            {
                button氮气.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button氮气.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-N2");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button工业蒸汽_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button工业蒸汽.ForeColor.Name == "Black" || button工业蒸汽.ForeColor.Name == "ControlText")
            {
                button工业蒸汽.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button工业蒸汽.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-LS");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button纯蒸汽_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button纯蒸汽.ForeColor.Name == "Black" || button纯蒸汽.ForeColor.Name == "ControlText")
            {
                button纯蒸汽.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button纯蒸汽.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-PS");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button真空_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button真空.ForeColor.Name == "Black" || button真空.ForeColor.Name == "ControlText")
            {
                button真空.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button真空.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-VE");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button放空_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button放空.ForeColor.Name == "Black" || button放空.ForeColor.Name == "ControlText")
            {
                button放空.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button放空.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-VT");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button物料_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button物料.ForeColor.Name == "Black" || button物料.ForeColor.Name == "ControlText")
            {
                button物料.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button物料.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-PL");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button其他_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button其他.ForeColor.Name == "Black" || button其他.ForeColor.Name == "ControlText")
            {
                button其他.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button其他.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-其他");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button标管1_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button标管1.ForeColor.Name == "Black" || button标管1.ForeColor.Name == "ControlText")
            {
                button标管1.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button标管1.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-标管1");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }

        private void button标管2_Click(object sender, EventArgs e)
        {
            VariableDictionary.tjtBtn = VariableDictionary.tjtBtnNull;
            if (button标管2.ForeColor.Name == "Black" || button标管2.ForeColor.Name == "ControlText")
            {
                button标管2.ForeColor = System.Drawing.SystemColors.ActiveCaption; VariableDictionary.btnState = true;
            }
            else
            {
                button标管2.ForeColor = System.Drawing.SystemColors.ControlText; VariableDictionary.btnState = false;
            }
            //NewTjLayer();//初始化allTjLayer
            VariableDictionary.selectTjtLayer.Clear();
            VariableDictionary.selectTjtLayer.Add("ZG-标管2");
            Env.Document.SendStringToExecute("IsFrozenLayer ", false, false, false);
        }
        #endregion
    }
}


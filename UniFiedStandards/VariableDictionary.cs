using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 变量
    /// </summary>
    internal class VariableDictionary
    {
        #region 服务器端字段和属性
        public static string? _serverIP;//服务器IP地址
        public static int _serverPort;//服务器端口号
        public static string _dataBaseName = "cad_sw_library";//数据库名称
        public static string? _userName;//用户名
        public static string? _passWord;//密码
        public static string? _storagePath;//数据库存储路径
        public static bool _useDPath;//是否使用DPath
        public static bool _autoSync;//是否自动同步
        public static int _syncInterval;//同步间隔
        public static string? _newConnectionString;//新的数据库连接字符串
        public static bool winForm_Status = false;//winForm状态
        public static bool radioButton = false;//二层传递窗单选按钮状态
        #endregion
        #region 变量
        /// <summary>
        /// 洁净区层颜色索引
        /// </summary>
        public static List<int> jjqLayerColorIndex = new List<int>
     {
          10,
        31,
        50,
        80,
        130,
        140,
        160,
        200,
        220,
        211,
        161,
        121,
        190,
        205,
        235,
        250,
         11,
        21,
        132,
        142,
        163,
        203,
        223,
        223,
        232,
        241,
        243

     };
        /// <summary>
        /// 系统分区颜色索引
        /// </summary>
        public static List<int> xtqLayerColorIndex = new List<int>
     {
          21,
         200,
         220,
          83,
         211,
         132,
         161,
         142,
          10,
         121,
         243,
          31,
         241,
          50,
         232,
          80,
         223,
         130,
         203,
         140,
          11,
         160
     };
        public static int jjqInt = 1;//洁净区层索引
        public static int xtqInt = 1;//系统分区层索引
        public static string? deviceNo_1;//设备编号
        public static string? deviceNo_2;//设备编号
        public static string? layerName;//图层名称
        /// <summary>
        /// 圆形按键文字 
        /// </summary>
        public static string? buttonCirText;
        /// <summary>
        /// 圆形外扩量
        /// </summary>
        public static string? textbox_CirPlus_Text;
        /// <summary>
        /// 按键文字
        /// </summary>
        public static string? buttonText;
        /// <summary>
        /// 方框文字扩展
        /// </summary>
        public static string? textbox_RecPlus_Text;
        /// <summary>
        /// 建筑防撞板与墙的距离
        /// </summary>
        public static double? textbox_Gap;
        /// <summary>
        /// 结构标注负荷文字
        /// </summary>
        public static string? dimString;
        /// <summary>
        /// 建筑专业排水沟的宽
        /// </summary>
        public static double? dimString_JZ_宽;
        /// <summary>
        /// 建筑专业排水沟的深
        /// </summary>
        public static double? dimString_JZ_深;
        /// <summary>
        /// 结构的矩形宽度
        /// </summary>
        public static double? textbox_S_Width;
        /// <summary>
        /// 结构的矩形高度
        /// </summary>
        public static double? textbox_S_Height;
        /// <summary>
        /// 结构圆直径
        /// </summary>
        public static double? textBox_S_CirDiameter;
        /// <summary>
        /// 结构圆半径
        /// </summary>
        public static double? textbox_S_Cirradius;
        /// <summary>
        /// 文字高
        /// </summary>
        public static double? textbox_Height;
        /// <summary>
        /// 文字宽
        /// </summary>
        public static double? textbox_Width;
        /// <summary>
        /// 图元角度
        /// </summary>
        public static double entityRotateAngle = 0;
        /// <summary>
        /// 图元名称
        /// </summary>
        public static string? btnFileName = null;
        /// <summary>
        /// 图元文件内的块名
        /// </summary>
        public static string btnFileName_blockName = "";
        /// <summary>
        /// 水专业天正图元
        /// </summary>
        public static int TCH_Ptj_No = 0;
        /// <summary>
        /// 图元图层
        /// </summary>
        public static string? btnBlockLayer = null;
        /// <summary>
        /// 图元文件
        /// </summary>
        public static byte[]? resourcesFile = null;
        /// <summary>
        /// 图元、图块比例
        /// </summary>
        public static double blockScale = 1;
        /// <summary>
        /// 文本框输入比例
        /// </summary>
        public static double textBoxScale = 1;
        /// <summary>
        /// 文本框输入比例
        /// </summary>
        public static double wpfTextBoxScale = 1;
        /// <summary>
        /// 条件按键
        /// </summary>
        public static List<string> tjtBtn = new List<string>();
        /// <summary>
        /// 条件按键空
        /// </summary>
        public static List<string> tjtBtnNull = new List<string>();
        /// <summary>
        /// 吊顶高
        /// </summary>
        public static string? wpfDiaoDingHeight = null;
        /// <summary>
        /// 吊顶高
        /// </summary>
        public static string? winFormDiaoDingHeight = null;
        /// <summary>
        /// 图层颜色
        /// </summary>
        public static int layerColorIndex = 0;
        /// <summary>
        /// 标注文字颜色
        /// </summary>
        public static Int16 textColorIndex = 0;
        /// <summary>
        /// 按键状态
        /// </summary>
        public static bool btnState = false;
        /// <summary>
        /// 清理参数
        /// </summary>
        public static double cleanupParameter = 0;
        /// <summary>
        /// 所有条件图图层
        /// </summary>
        public static List<string> allTjtLayer = new List<string>();
        /// <summary>
        /// 选择的条件图图层
        /// </summary>
        public static List<string> selectTjtLayer = new List<string>();
        /// <summary>
        /// 共用条件检查
        /// </summary>
        public static List<string> GGtjtBtn = new List<string>
        {
            "TJ(共用条件)"
        };
        /// <summary>
        /// 设备条件检查
        /// </summary>
        public static List<string> SBtjtBtn = new List<string>
        {
           "SB",
           "设备",
           "S_设备",
           "设备层",
           "TJ(设备位号)",
           "设备位号",
           "SB(工艺设备)",
           "S_工艺设备",
           "SB(设备名称)",
           "设备名称",
           "SB(设备外框)",
           "S_办公家具"
        };
        /// <summary>
        /// 工艺条件检查
        /// </summary>
        public static List<string> GYtjtBtn = new List<string>
        {
            "TJ(工艺专业GY)",
            //"1",
            //"SB",
            //"设备",
            //"设备层",
            //"TJ(设备位号)",
            //"QY",
            //"SB(设备外框)",
            //"SB(工艺设备)",
            //"SB(设备名称)",
            //"TJ(建筑专业J)Y",
            //"TJ(建筑专业J)N",
            //"TJ(结构专业JG)",
            //"TJ(给排水专业S)",
            //"EQUIP_地漏",
            //"EQUIP_给水",
            //"TJ(暖通专业N)",
            //"TJ(电气专业D)",
            //"TJ(电气专业D)",
            //"TJ(电气专业D2)",
            //"EQUIP-通讯",
            //"EQUIP-安防",
        };
        /// <summary>
        /// 建筑平面图层
        /// </summary>
        public static List<string> ApmtTjBtn = new List<string>
        {
             "TJ(房间编号)",
            "AXIS",
            "AXIS_TEXT",
            "BASE",
            "COLUMN",
            "DEFPOINTS",
            "DIM_ELEV",
            "DIM_IDEN",
            "DIM_MODI",
            "DIM_SYMB",
            "DOTE",
            "GROUND",
            "PUB_WALL",
            "PUB_TEXT",
            "SUTE-BOTL",
            "STAIR",
            "SURFACE",
            "WALL",
            "WALL-MOVE",
            "WALL-PARAPET",
            "CURTWALL",
            "LATRINE",
            "WINDOW",
            "WINDOW_TEXT",
            "DOOR_FIRE",
            "DOOR_FIRE_TEXT",
            "DIM_SYMB",
            "EVTR",
            "柱",
            "厨卫",
            "洁具",
            "家具"
        };
        /// <summary>
        /// 建筑条件检查
        /// </summary>
        public static List<string> AtjtBtn = new List<string>
        {
             "TJ(建筑专业)",
            "TJ(建筑专业J)",
            "TJ(建筑专业J)Y",
            "TJ(建筑吊顶)",
            "TJ(建筑过结构洞口)",            
        };

        /// <summary>
        /// 结构条件检查
        /// </summary>
        public static List<string> StjtBtn = new List<string>
        {
            "TJ(结构专业JG)",
            //"COLUMN",
            //"COLU",
            //"HOLE",
            //"WALL-C",
            //"PUB_HATCH",
            //"柱",
            "TJ(结构洞口)",
            "TJ(结构开洞)",
           
            //"STJ_受力点",
            //"STJ_矩形开洞",
            //"STJ_圆形开洞",
            //"STJ_设备着地点",
            //"STJ_结构开洞"
        };
        /// <summary>
        /// 水条件检查
        /// </summary>
        public static List<string> PtjtBtn = new List<string>
        {
            "TJ(给排水专业S)",
            "EQUIP_地漏",
            "EQUIP_给水",
            "TJ(给排水过建筑)",
            "TJ(给排水过结构)",
            "TJ(给排水过电气动力条件)",
            "TJ(给排水过电气喷淋条件)",
            "EQUIP_消火栓",
            //"PTJ_大便器给水",
            //"PTJ_大洗涤池",
            //"PTJ_给水点",
            //"PTJ_冷直排",
            //"PTJ_热直排",
            //"PTJ_水池给水",
            //"PTJ_水池排水",
            //"PTJ_洗涤盆",
            //"PTJ_洗脸盆",
            //"PTJ_洗眼器",
            //"EQUIP_消火栓",
            //"PTJ_小便器给水",
            //"PTJ-WALL-MOVE",
            //"PTJ-给排水开洞"
        };
        /// <summary>
        /// 暖通条件检查
        /// </summary>
        public static List<string> NtjtBtn = new List<string>
        {
            "TJ(暖通专业N)",
            "TJ(暖通过工艺)",
            "TJ(暖通过建筑)",
            //"WALL-PARAPET",
            "TJ(暖通过结构)",
            "TJ(暖通过给排水)",
            "TJ(暖通过电气)",
            "TJ(暖通过自控)",
            "TJ(暖通专业FFU)",
            //"NTJ-排潮",
            //"NTJ-排尘",
            //"NTJ-排热",
            //"NTJ-直排",
            //"NTJ-除味",
            //"NTJ-WALL-MOVE",
            //"NTJ-暖通开洞"
           
        };
        /// <summary>
        /// 电气条件检查
        /// </summary>
        public static List<string> EtjtBtn = new List<string>
        {
            "TJ(电气专业D)",
            "TJ(电气专业D1)",
            "TJ(电气专业D2)",
            "TJ(电气过建筑)",
            "TJ(电气过建筑孔洞D)",
            "TJ(电气过建筑夹墙D)",
            "TJ(电气过结构)",
            "TJ(电气过结构洞口)",
            "TJ(电气过结构楼板洞D)",
            "TJ(电气过结构电缆沟D)",
            "TJ(电气过结构活荷载D)",
            "TEL_CABINET",
            "EQUIP-照明",
            "WIRE-厂区消防",
            /*
            "DQTJ-EQUIP-220V插座",
            "DQTJ-EQUIP-动力",
            "DQTJ-EQUIP-UPS16A插座",
            "DQTJ-EQUIP-UPS插座",
            "DQTJ-EQUIP-UPS电源",
            "DQTJ-EQUIP-插座箱",
            "DQTJ-EQUIP-插座",
            "DQTJ-EQUIP-潮湿插座",
            "DQTJ-EQUIP-厨宝电热水器安全型插座",
            "DQTJ-EQUIP-传递窗电源插座",
            "DQTJ-EQUIP-带保护极的单相暗敷插座",
            "DQTJ-EQUIP-带保护极的单相防爆插座",
            "DQTJ-EQUIP-带保护极的单相密闭插座",
            "DQTJ-EQUIP-带保护极的三相暗敷插座",
            "DQTJ-EQUIP-带保护极的三相防爆插座",
            "DQTJ-EQUIP-带保护极的三相密闭插座",
            "DQTJ-EQUIP-单相16A三孔插座",
            "DQTJ-EQUIP-单相20A三孔插座",
            "DQTJ-EQUIP-单相25A三孔插座",
            "DQTJ-EQUIP-单相32A三孔插座",
            "DQTJ-EQUIP-单相插座",
            "DQTJ-EQUIP-单相地面插座",
            "DQTJ-EQUIP-单相防爆岛型插座",
            "DQTJ-EQUIP-单相空调插座",
            "DQTJ-EQUIP-单相三孔插座",
            "DQTJ-EQUIP-单相三孔岛型插座",
            "DQTJ-EQUIP-单相五孔岛型插座",
            "DQTJ-EQUIP-地面插座",
            "DQTJ-EQUIP-烘手器",
            "DQTJ-EQUIP-红外感应门插座",
            "DQTJ-EQUIP-互锁插座",
            "DQTJ-EQUIP-空调插座",
            "DQTJ-EQUIP-两点互锁",
            "DQTJ-EQUIP-三点互锁",
            "DQTJ-EQUIP-三联插座",
            "DQTJ-EQUIP-三相380V插座",
            "DQTJ-EQUIP-三相潮湿插座",
            "DQTJ-EQUIP-三相岛型插座",
            "DQTJ-EQUIP-设备用电点位",
            "DQTJ-EQUIP-实验台UPS功能柱插座",
            "DQTJ-EQUIP-实验台功能柱插座",
            "DQTJ-EQUIP-视孔灯",
            "DQTJ-EQUIP-手消毒插座",
            "DQTJ-EQUIP-四联插座",
            "DQTJ-EQUIP-应急16A插座",
            "DQTJ-EQUIP-应急插座",
            "DQTJ-WALL-MOVE",
            "ETJ-电气开洞",
            "DQTJ-EQUIP-门禁插座",
            "DQTJ-EQUIP-立式空调插座",
            "DQTJ-EQUIP-壁挂空调插座",
            "DQTJ-EQUIP-灭蝇灯插座",
            "DQTJ-EQUIP-灭蝇灯插座(底边)",
            "DQTJ-EQUIP-220V用电设备(点或配电柜)",
            "DQTJ-EQUIP-380V用电设备(点或配电柜)",
            "DQTJ-EQUIP-380V用电设备大于10KW(点或配电柜)",
            "DQTJ-EQUIP-实验台UPS功能柱电源",
            "DQTJ-EQUIP-实验台上方220V插座"
            */
            
        };
        /// <summary>
        /// 自控条件检查
        /// </summary>
        public static List<string> ZKtjtBtn = new List<string>
        {
            "TJ(自控专业ZK)",
            "EQUIP-通讯",
            "EQUIP-安防",
            "TJ(自控过建筑)",
            "TJ(自控过结构)",
            //"EQUIP-楼控",
            //"EQUIP-电话",
            //"EQUIP-箱柜",
            //"EQUIP-消防",
            //"EQUIP-通讯",
            //"ZKTJ-EQUIP-安防-安防监控",
            //"ZKTJ-EQUIP-安防-半球彩色摄像机",
            //"ZKTJ-EQUIP-安防-眼纹识别器",
            //"ZKTJ-EQUIP-安防-人像识别器",
            //"ZKTJ-EQUIP-安防-室外彩色摄像机",
            //"ZKTJ-EQUIP-安防-室外彩色云台摄像机",
            //"ZKTJ-EQUIP-安防-电控锁",
            //"ZKTJ-EQUIP-安防-电锁按键",
            //"ZKTJ-EQUIP-安防-读卡器",
            //"ZKTJ-EQUIP-安防-门磁开关",
            //"ZKTJ-EQUIP-安防-门禁控制器",
            //"ZKTJ-EQUIP-安防-防爆型网络摄像机",
            //"ZKTJ-EQUIP-安防-广角彩色摄像机",
            //"ZKTJ-EQUIP-监控文字",
            //"ZKTJ-EQUIP-电话-洁净电话机",
            //"ZKTJ-EQUIP-电话-防爆型电话机",
            //"ZKTJ-EQUIP-通讯-电话、网络插座",
            //"ZKTJ-EQUIP-通讯-局域网插座",
            //"ZKTJ-EQUIP-通讯-互联网插座",
            //"ZKTJ-EQUIP-通讯-内线电话插座",
            //"ZKTJ-EQUIP-通讯-外线电话插座",
            //"ZKTJ-EQUIP-通讯-网络插座",
            //"ZKTJ-EQUIP-通讯-网络交换机",
            //"ZKTJ-EQUIP-通讯-无线网络接入点",
            //"ZKTJ-EQUIP-通讯-电话插座",
            //"ZKTJ-EQUIP-通讯-电话网络插座",
            //"ZKTJ-WALL-MOVE",
            //"ZKTJ-自控开洞"
        
        };

        #endregion
    }
}

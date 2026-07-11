using System;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 同步进度信息类
    /// </summary>
    public sealed class SyncProgressInfo
    {
        /// <summary>
        /// 获取或设置当前同步阶段的名称
        /// </summary>
        public string Stage { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置当前同步阶段的详细信息
        /// </summary>
        public string StageDetail { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置当前正在处理的项目的名称
        /// </summary>
        public string CurrentItem { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置已完成的操作数
        /// </summary>
        public int CompletedOperations { get; set; }
        /// <summary>
        /// 获取或设置总操作数
        /// </summary>
        public int TotalOperations { get; set; }
        /// <summary>
        /// 获取或设置一个值，指示进度是否为不确定状态
        /// </summary>
        public bool IsIndeterminate { get; set; }
        /// <summary>
        /// 获取当前同步进度的百分比（0 到 100 之间的整数）
        /// </summary>
        public int Percent
        {
            get
            {
                if (TotalOperations <= 0)
                {
                    return 0;
                }

                var value = (int)Math.Round(CompletedOperations * 100.0 / TotalOperations);
                return value < 0 ? 0 : (value > 100 ? 100 : value);
            }
        }
    }
}

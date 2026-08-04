using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Models
{ 
    /// <summary>
    /// 参数行数据模型，实现属性变更通知接口，用于UI数据绑定
    /// </summary>
    public class ParameterRow : INotifyPropertyChanged
    {
        /// <summary>参数分组名称</summary>
        public string? Group { get; set; }

        /// <summary>参数名称</summary>
        public string? Name { get; set; }

        // 私有字段，用于存储参数值，并触发属性变更通知
        private string? _value;

        /// <summary>
        /// 参数值，当值发生变化时自动触发 PropertyChanged 事件
        /// </summary>
        public string? Value
        {
            get => _value;
            set
            {
                _value = value;
                OnPropertyChanged();   // 通知 UI 值已更改
            }
        }

        /// <summary>参数单位（如 mm, kg 等）</summary>
        public string? Unit { get; set; }

        /// <summary>备注说明</summary>
        public string? Remark { get; set; }

        /// <summary>
        /// 是否允许用户输入/编辑
        /// true 表示可编辑，false 表示只读
        /// </summary>
        public bool IsInput { get; set; }

        /// <summary>
        /// 实现 INotifyPropertyChanged 接口所需的事件
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// 触发属性变更通知的受保护方法
        /// </summary>
        /// <param name="name">属性名称，默认由 CallerMemberName 自动提供</param>
        protected void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}

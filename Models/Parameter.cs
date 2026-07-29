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
    /// 行数据模型
    /// </summary>
    public class ParameterRow : INotifyPropertyChanged
    {
        public string? Group { get; set; }
        public string? Name { get; set; }
        private string? _value;
        public string? Value { get => _value; set { _value = value; OnPropertyChanged(); } }
        public string? Unit { get; set; }
        public string? Remark { get; set; }
        public bool IsInput { get; set; }   // 是否可编辑
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}

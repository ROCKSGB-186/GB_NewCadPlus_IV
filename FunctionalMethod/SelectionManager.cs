using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace GB_NewCadPlus_LM.FunctionalMethod
{
    /// <summary>
    /// 附加属性：表示控件是否被选中（用于样式触发高亮）
    /// </summary>
    public static class SelectionManager
    {
        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.RegisterAttached(
                "IsSelected",
                typeof(bool),
                typeof(SelectionManager),
                new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender)
            );

        public static void SetIsSelected(DependencyObject element, bool value)
        {
            if (element == null) return;
            element.SetValue(IsSelectedProperty, value);
        }

        public static bool GetIsSelected(DependencyObject element)
        {
            if (element == null) return false;
            return (bool)element.GetValue(IsSelectedProperty);
        }
    }
}

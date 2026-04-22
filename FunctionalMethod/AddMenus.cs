using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    public static class AddMenus
    {
        /// <summary>
        /// 添加菜单
        /// </summary>
        public static void AddMenu()
        {
            //加入菜单
            var menus = Application.MenuGroups.invokeMethod("item", "Acad").GetProperty("menus");
            try
            {
                menus.invokeMethod("Add", "GB_CadTools");//加入菜单，名为GB_CadTools
            }
            catch { }
            var menu = menus.invokeMethod("item", "GB_CadTools");//加入菜单，名为GB_CadTools
            while (Convert.ToInt32(menu.GetProperty("Count")) > 0) //判断是不是有菜单项
            { menu.invokeMethod("item", 0).invokeMethod("delete"); }
            ;//删除所有菜单项
            var menuItems = menu.invokeMethod("AddSubMenu", 0, "工具组");//添加子菜单
            menuItems.invokeMethod("AddMenuItem", 0, "Line", "\u0003_Line ");
            menu.invokeMethod("AddSeparator", 1);//添加分隔符一条分割线
            menu.invokeMethod("AddMenuItem", 2, "加载窗体", "\u0003_ffff ");
            try
            {
                menu.invokeMethod("RemoveFromMenuBar");//删除菜单
            }
            catch { }
            menu.invokeMethod("InsertInMenuBar", "");//插入新建的菜单到cad的产品菜单栏中
        }
        /// <summary>
        /// 获取属性
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static object GetProperty(this object obj, string key)
        {
            // 获取对象类型
            var comtype = Type.GetTypeFromHandle(Type.GetTypeHandle(obj));
            // 调用对象属性
            return comtype.InvokeMember(key, BindingFlags.GetProperty, null, obj, null);
        }
        /// <summary>
        /// 设置属性
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public static void SetProperty(this object obj, string key, object value)
        {
            // 获取对象类型
            var comtype = Type.GetTypeFromHandle(Type.GetTypeHandle(obj));
            // 调用对象属性
            comtype.InvokeMember(key, BindingFlags.SetProperty, null, obj, new object[] { value });
        }
        /// <summary>
        /// 调用方法
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="key"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static object invokeMethod(this object obj, string key, params object[] args)
        {
            // 获取对象类型
            var comtype = Type.GetTypeFromHandle(Type.GetTypeHandle(obj));
            // 调用对象方法
            return comtype.InvokeMember(key, BindingFlags.InvokeMethod, null, obj, args);
        }
    }
}

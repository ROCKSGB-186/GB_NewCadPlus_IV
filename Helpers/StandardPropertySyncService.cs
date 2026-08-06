using Autodesk.AutoCAD.DatabaseServices;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 规范参数双写服务：同步到图元 JSON，并同步到已有 AutoCAD 属性 Tag。
    /// </summary>
    public sealed class StandardPropertySyncService
    {
        /// <summary>
        /// 将服务器返回的规范属性合并到图元 JSON 字典。
        /// </summary>
        public void ApplyToAttributesJson(
            ImportEntityDto entity,
            FlangeStandardMatchResponse response)
        {
            // 校验图元对象和服务器结果，避免查询失败时覆盖原有属性。
            if (entity == null || response == null || !response.Success) return;

            // 兼容历史数据中的 null 字典。
            entity.AttributesJson = entity.AttributesJson
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 只写入服务器明确返回的属性，保留其他业务字段。
            foreach (KeyValuePair<string, string> attribute in response.Attributes)
            {
                // 忽略空标签，避免产生无法使用的 JSON 键。
                if (string.IsNullOrWhiteSpace(attribute.Key)) continue;

                // 规范属性以服务器返回值为准，确保 JSON 与当前选定系列一致。
                entity.AttributesJson[attribute.Key] = attribute.Value ?? string.Empty;
            }
        }

        /// <summary>
        /// 将服务器返回的规范属性同步到已有块参照的 AttributeReference。
        /// </summary>
        public int ApplyToBlockReference(
            Transaction transaction,
            BlockReference blockReference,
            FlangeStandardMatchResponse response)
        {
            // 校验事务、块参照和服务器结果，避免无效写入。
            if (transaction == null || blockReference == null || response == null || !response.Success)
            {
                return 0;
            }

            // 记录实际写入的属性数量，便于调用方日志和测试验证。
            int updatedCount = 0;

            // 遍历块参照已有的属性，不创建新的 Tag，避免破坏图块定义。
            foreach (ObjectId attributeId in blockReference.AttributeCollection)
            {
                // 以写模式打开已有属性引用。
                AttributeReference attribute = transaction.GetObject(attributeId, OpenMode.ForWrite) as AttributeReference;
                if (attribute == null) continue;

                // 读取并归一化 Tag，兼容大小写和常见分隔符差异。
                string tag = NormalizeTag(attribute.Tag);
                if (string.IsNullOrWhiteSpace(tag)) continue;

                // 按归一化 Tag 查找服务器返回的规范值。
                string value = FindAttributeValue(response.Attributes, tag);
                if (value == null) continue;

                // 只在值发生变化时写入 AutoCAD 属性。
                string oldValue = attribute.TextString ?? string.Empty;
                if (string.Equals(oldValue, value, StringComparison.Ordinal)) continue;

                // 写入当前选定系列对应的显示值。
                attribute.TextString = value;
                updatedCount++;

                // 记录规范参数实际写入的 Tag、旧值、新值和目标图元标识。
                LogManager.Instance.LogInfo(
                    $"[规范参数赋值][AttributeReference] Tag={attribute.Tag}, OldValue={oldValue}, NewValue={value}, TargetObjectId={blockReference.ObjectId}");
            }

            // 返回写入数量，调用方可据此判断双写是否命中图块 Tag。
            return updatedCount;
        }

        /// <summary>
        /// 归一化 AutoCAD 属性 Tag。
        /// </summary>
        private static string NormalizeTag(string tag)
        {
            // 去除空白并统一大小写，保持与服务器属性键比较稳定。
            return (tag ?? string.Empty).Trim().ToUpperInvariant();
        }

        /// <summary>
        /// 从服务器属性字典中按归一化 Tag 查找值。
        /// </summary>
        private static string FindAttributeValue(
            IDictionary<string, string> attributes,
            string normalizedTag)
        {
            // 空字典或空 Tag 不参与匹配。
            if (attributes == null || string.IsNullOrWhiteSpace(normalizedTag)) return null;

            // 优先使用字典自身的不区分大小写匹配能力。
            foreach (KeyValuePair<string, string> pair in attributes)
            {
                // 归一化双方键后再比较，兼容服务器和图块 Tag 的大小写差异。
                if (string.Equals(NormalizeTag(pair.Key), normalizedTag, StringComparison.Ordinal))
                {
                    return pair.Value ?? string.Empty;
                }
            }

            // 没有同名 Tag 时不写入，避免误改其他属性。
            return null;
        }
    }
}

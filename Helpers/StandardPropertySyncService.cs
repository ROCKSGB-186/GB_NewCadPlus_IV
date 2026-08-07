using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
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
            FlangeStandardMatchResponse response,
            bool createMissingAttributes = true)
        {
            // 校验事务、块参照和服务器结果，避免无效写入。
            if (transaction == null || blockReference == null || response == null || !response.Success)
            {
                return 0;
            }

            // 记录实际写入的属性数量，便于调用方日志和测试验证。
            int updatedCount = 0;
            var existingTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 遍历块参照已有的属性，优先更新现有属性并保留其原始 Prompt。
            foreach (ObjectId attributeId in blockReference.AttributeCollection)
            {
                // 以写模式打开已有属性引用。
                AttributeReference attribute = transaction.GetObject(attributeId, OpenMode.ForWrite) as AttributeReference;
                if (attribute == null) continue;

                // 读取并归一化 Tag，兼容大小写和常见分隔符差异。
                string tag = NormalizeTag(attribute.Tag);
                if (string.IsNullOrWhiteSpace(tag)) continue;
                existingTags.Add(tag);

                // 按归一化 Tag 查找服务器返回的规范值。
                string value = FindAttributeValue(response.Attributes, tag);
                if (value == null && IsFlangeStandardTag(tag))
                {
                    value = FindAttributeValue(response.Attributes, "FLG_STD");
                }
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

            // 炸开后的目标块如果缺少规范字段，则把隐藏属性添加到块定义和块参照中。
            // 炸开前调用本方法时传入 false，避免新增属性被炸开为当前图纸空间中的独立实体。
            if (createMissingAttributes)
            {
                foreach (KeyValuePair<string, string> serverAttribute in response.Attributes)
                {
                    string tag = NormalizeTag(serverAttribute.Key);
                    string value = serverAttribute.Value ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(tag) || existingTags.Contains(tag)) continue;

                    if (AddHiddenAttribute(transaction, blockReference, serverAttribute.Key, value))
                    {
                        updatedCount++;
                        existingTags.Add(tag);
                    }
                }
            }

            // 返回已有属性和新增隐藏属性的实际写入数量。
            return updatedCount;
        }

        /// <summary>
        /// 判断属性 Tag 是否属于法兰标准字段。
        /// </summary>
        private static bool IsFlangeStandardTag(string normalizedTag)
        {
            return string.Equals(normalizedTag, NormalizeTag("FLG_STD"), StringComparison.Ordinal) ||
                   string.Equals(normalizedTag, NormalizeTag("法兰标准"), StringComparison.Ordinal) ||
                   string.Equals(normalizedTag, NormalizeTag("DRAWINGNO.STANDARDNO"), StringComparison.Ordinal);
        }

        /// <summary>
        /// 将缺失规范字段添加到目标块定义及当前块参照，不直接写入当前图纸空间。
        /// </summary>
        private static bool AddHiddenAttribute(
            Transaction transaction,
            BlockReference blockReference,
            string tag,
            string value)
        {
            // 空 Tag 不创建属性，避免产生不可用的块字段。
            if (string.IsNullOrWhiteSpace(tag)) return false;

            // 以写模式打开块定义，确保新增属性属于目标块而不是当前图纸空间。
            BlockTableRecord blockDefinition = transaction.GetObject(
                blockReference.BlockTableRecord,
                OpenMode.ForWrite) as BlockTableRecord;
            if (blockDefinition == null) return false;

            // 创建隐藏属性定义；服务器当前没有独立 Prompt 字段，因此提示留空，不能使用 Tag 冒充。
            var attributeDefinition = new AttributeDefinition
            {
                Position = Point3d.Origin,
                Tag = tag.Trim(),
                Prompt = string.Empty,
                TextString = value,
                Height = 1.0,
                Invisible = true
            };
            attributeDefinition.SetDatabaseDefaults();
            blockDefinition.AppendEntity(attributeDefinition);
            transaction.AddNewlyCreatedDBObject(attributeDefinition, true);

            // 按块变换创建当前块参照对应的隐藏属性引用。
            var attributeReference = new AttributeReference();
            attributeReference.SetAttributeFromBlock(attributeDefinition, blockReference.BlockTransform);
            attributeReference.TextString = value;
            blockReference.AttributeCollection.AppendAttribute(attributeReference);
            transaction.AddNewlyCreatedDBObject(attributeReference, true);

            // 记录新增字段及其值；Prompt 为空是因为接口没有返回数据库提示元数据。
            LogManager.Instance.LogInfo(
                $"[规范参数新增][AttributeReference] Tag={tag}, Prompt=, OldValue=, NewValue={value}, TargetObjectId={blockReference.ObjectId}");
            return true;
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

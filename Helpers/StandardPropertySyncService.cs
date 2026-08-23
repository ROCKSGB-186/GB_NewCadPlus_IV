using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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

            ApplyToAttributesJson(entity, response.Attributes);
        }

        /// <summary>
        /// 将通用规范属性合并到图元 JSON 字典，供管子、管件和阀门复用。
        /// </summary>
        public void ApplyToAttributesJson(
            ImportEntityDto entity,
            IDictionary<string, string> attributes)
        {
            if (entity == null || attributes == null) return;

            // 兼容历史数据中的 null 字典。
            entity.AttributesJson = entity.AttributesJson
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 只写入服务器明确返回的属性，保留其他业务字段。
            foreach (KeyValuePair<string, string> attribute in attributes)
            {
                // 忽略空标签，避免产生无法使用的 JSON 键。
                if (string.IsNullOrWhiteSpace(attribute.Key)) continue;

                // 规范属性以服务器返回值为准，确保 JSON 与当前选定系列一致。
                entity.AttributesJson[attribute.Key] = attribute.Value ?? string.Empty;
            }
        }

        /// <summary>
        /// 将通用规范记录转换为 CAD 属性后写入图元 JSON。
        /// </summary>
        public void ApplyToAttributesJson(ImportEntityDto entity, StandardItemClient item)
        {
            if (entity == null || item == null) return;
            ApplyToAttributesJson(entity, StandardCadAttributeMapper.ToAttributes(item));
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
            string standardDn = FindAttributeValue(response.Attributes, NormalizeTag("DN"));
            string standardPn = FindAttributeValue(response.Attributes, NormalizeTag("PN"));
            LogManager.Instance.LogInfo(
                $"[规范回写][开始] TargetObjectId={blockReference.ObjectId}, createMissingAttributes={createMissingAttributes}, " +
                $"目标块AttributeCount={blockReference.AttributeCollection.Count}, 服务器属性数量={response.Attributes?.Count ?? 0}, " +
                $"服务器DN={standardDn ?? "<缺失>"}, 服务器PN={standardPn ?? "<缺失>"}");

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
                if (value == null && IsModelSpecificationTag(tag))
                {
                    value = SynchronizeModelSpecification(attribute.TextString, standardDn, standardPn);
                }
                if (value == null)
                {
                    if (IsFlangeStandardTag(tag) || IsModelSpecificationTag(tag))
                    {
                        LogManager.Instance.LogWarning(
                            $"[规范回写][跳过] 目标属性未找到对应服务器值：Tag={attribute.Tag}, NormalizedTag={tag}, TargetObjectId={blockReference.ObjectId}");
                    }
                    continue;
                }

                // 只在值发生变化时写入 AutoCAD 属性。
                string oldValue = attribute.TextString ?? string.Empty;
                if (string.Equals(oldValue, value, StringComparison.Ordinal))
                {
                    if (IsFlangeStandardTag(tag) || IsModelSpecificationTag(tag))
                    {
                        LogManager.Instance.LogInfo(
                            $"[规范回写][无需修改] Tag={attribute.Tag}, Value={value}, TargetObjectId={blockReference.ObjectId}");
                    }
                    continue;
                }

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
                    if (string.IsNullOrWhiteSpace(tag)) continue;
                    if (existingTags.Contains(tag)) continue;

                    if (IsFlangeStandardTag(tag) || IsModelSpecificationTag(tag))
                    {
                        LogManager.Instance.LogWarning(
                            $"[规范回写][目标缺少Tag] 服务器返回字段未在目标块中找到：ServerTag={serverAttribute.Key}, " +
                            $"Value={value}, TargetObjectId={blockReference.ObjectId}, createMissingAttributes={createMissingAttributes}");
                    }

                    if (AddHiddenAttribute(transaction, blockReference, serverAttribute.Key, value))
                    {
                        updatedCount++;
                        existingTags.Add(tag);
                    }
                }
            }

            // 返回已有属性和新增隐藏属性的实际写入数量。
            LogManager.Instance.LogInfo(
                $"[规范回写][完成] TargetObjectId={blockReference.ObjectId}, 实际写入数量={updatedCount}, " +
                $"目标块原有Tag数量={existingTags.Count}, createMissingAttributes={createMissingAttributes}");
            return updatedCount;
        }

        /// <summary>
        /// 将用户确认的属性补充为当前块参照的隐藏 AttributeReference。
        /// </summary>
        public int EnsureEditedAttributes(
            Transaction transaction,
            BlockReference blockReference,
            IDictionary<string, string> editedProperties)
        {
            // 参数无效时不创建属性，避免误修改当前图纸。
            if (transaction == null || blockReference == null || editedProperties == null || editedProperties.Count == 0)
            {
                return 0;
            }

            // 先收集当前块参照已经存在的属性 Tag，避免重复创建块属性定义。
            var existingTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId attributeId in blockReference.AttributeCollection)
            {
                if (transaction.GetObject(attributeId, OpenMode.ForRead) is AttributeReference attribute &&
                    !string.IsNullOrWhiteSpace(attribute.Tag))
                {
                    existingTags.Add(NormalizeTag(attribute.Tag));
                }
            }

            // 仅为缺失字段创建隐藏属性；已有字段仍由原有回写逻辑更新值。
            int createdCount = 0;
            foreach (KeyValuePair<string, string> property in editedProperties)
            {
                string tag = property.Key?.Trim() ?? string.Empty;
                string normalizedTag = NormalizeTag(tag);
                if (string.IsNullOrWhiteSpace(tag) ||
                    string.IsNullOrWhiteSpace(normalizedTag) ||
                    existingTags.Contains(normalizedTag))
                {
                    continue;
                }

                if (AddHiddenAttribute(transaction, blockReference, tag, property.Value ?? string.Empty))
                {
                    existingTags.Add(normalizedTag);
                    createdCount++;
                    LogManager.Instance.LogInfo(
                        $"[插入前属性编辑新增AttributeReference] Tag={tag}, OldValue=<不存在>, NewValue={property.Value ?? string.Empty}, TargetObjectId={blockReference.ObjectId}");
                }
            }

            return createdCount;
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
        /// 判断属性 Tag 是否为需要根据法兰规范 DN、PN 更新参数的型号字段。
        /// </summary>
        private static bool IsModelSpecificationTag(string normalizedTag)
        {
            return string.Equals(normalizedTag, NormalizeTag("MODEL"), StringComparison.Ordinal) ||
                   string.Equals(normalizedTag, NormalizeTag("SW_MODEL"), StringComparison.Ordinal);
        }

        /// <summary>
        /// 保留现有型号文本，仅替换其中的 DN、PN 参数。
        /// 兼容 DN150、PN10 及 法兰_PL100_PN10_RF.SLDPRT 这类法兰三维文件名。
        /// </summary>
        private static string SynchronizeModelSpecification(string currentValue, string standardDn, string standardPn)
        {
            if (string.IsNullOrWhiteSpace(currentValue) ||
                string.IsNullOrWhiteSpace(standardDn) ||
                string.IsNullOrWhiteSpace(standardPn))
            {
                return null;
            }

            string dnNumber = Regex.Replace(standardDn.Trim(), @"^DN\s*", string.Empty, RegexOptions.IgnoreCase);
            string pnNumber = Regex.Replace(standardPn.Trim(), @"^PN\s*", string.Empty, RegexOptions.IgnoreCase);
            if (string.IsNullOrWhiteSpace(dnNumber) || string.IsNullOrWhiteSpace(pnNumber)) return null;

            string synchronized = Regex.Replace(
                currentValue,
                @"DN\s*\d+(?:\.\d+)?",
                $"DN{dnNumber}",
                RegexOptions.IgnoreCase);
            synchronized = Regex.Replace(
                synchronized,
                @"PN\s*\d+(?:\.\d+)?",
                $"PN{pnNumber}",
                RegexOptions.IgnoreCase);
            synchronized = Regex.Replace(
                synchronized,
                @"(?<prefix>法兰_[^_\d]*)(?<dn>\d+(?:\.\d+)?)(?=_PN)",
                match => $"{match.Groups["prefix"].Value}{dnNumber}",
                RegexOptions.IgnoreCase);

            return string.Equals(currentValue, synchronized, StringComparison.Ordinal) ? null : synchronized;
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
            // 先还原管道属性使用的编码 Tag，再统一大小写和常见分隔符。
            string decoded = PipelineCadPropertyKeyHelper.Decode(tag ?? string.Empty);
            if (string.IsNullOrWhiteSpace(decoded)) return string.Empty;

            var builder = new System.Text.StringBuilder(decoded.Length);
            foreach (char character in decoded.Trim().ToUpperInvariant())
            {
                if (char.IsLetterOrDigit(character))
                {
                    builder.Append(character);
                }
            }

            return builder.ToString();
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

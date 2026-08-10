using System;
using System.Globalization;
using System.Text;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 管道 CAD 扩展字典键编码助手。
    /// AutoCAD 的 DBDictionary 键有字符限制，业务 Tag 不能直接作为键保存。
    /// </summary>
    public static class PipelineCadPropertyKeyHelper
    {
        /// <summary>
        /// 管道属性统一存储使用的固定扩展字典键。
        /// </summary>
        public const string StorageKey = "GBPIPE_DATA";

        private const string EncodedPrefix = "GBPIPE_";

        /// <summary>
        /// 将业务 Tag 编码为只包含 ASCII 字母、数字和下划线的 CAD 字典键。
        /// </summary>
        public static string Encode(string businessTag)
        {
            string value = businessTag ?? string.Empty;
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            StringBuilder builder = new StringBuilder(EncodedPrefix.Length + bytes.Length * 2);
            builder.Append(EncodedPrefix);

            foreach (byte item in bytes)
            {
                builder.Append(item.ToString("X2", CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        /// <summary>
        /// 将 CAD 字典键还原为业务 Tag；历史未编码键原样返回。
        /// </summary>
        public static string Decode(string cadKey)
        {
            if (string.IsNullOrWhiteSpace(cadKey) ||
                !cadKey.StartsWith(EncodedPrefix, StringComparison.Ordinal) ||
                (cadKey.Length - EncodedPrefix.Length) % 2 != 0)
            {
                return cadKey ?? string.Empty;
            }

            string hex = cadKey.Substring(EncodedPrefix.Length);
            byte[] bytes = new byte[hex.Length / 2];
            for (int index = 0; index < bytes.Length; index++)
            {
                if (!byte.TryParse(
                    hex.Substring(index * 2, 2),
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out byte value))
                {
                    return cadKey;
                }

                bytes[index] = value;
            }

            try
            {
                return Encoding.UTF8.GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                return cadKey;
            }
        }
    }
}

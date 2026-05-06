using System;

namespace GB_NewCadPlus_IV
{
    public sealed class SyncProgressInfo
    {
        public string Stage { get; set; } = string.Empty;
        public string StageDetail { get; set; } = string.Empty;
        public string CurrentItem { get; set; } = string.Empty;
        public int CompletedOperations { get; set; }
        public int TotalOperations { get; set; }
        public bool IsIndeterminate { get; set; }

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

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 分类ID生成器（线程安全）
    /// </summary>
    public static class CategoryIdGenerator
    {
        /// <summary>
        /// 生成ID时的进程内互斥锁，避免并发下重复ID
        /// </summary>
        private static readonly SemaphoreSlim _idLock = new SemaphoreSlim(1, 1);

        /// <summary>
        /// 生成主分类ID（注意：当前主分类表是自增主键，此方法主要用于兼容旧调用）
        /// </summary>
        //public static async Task<int> GenerateMainCategoryIdAsync(DatabaseManager databaseManager)
        //{
        //    // 参数判空，避免空引用
        //    if (databaseManager == null) throw new ArgumentNullException(nameof(databaseManager));

        //    // 进入互斥区，防止并发重复
        //    await _idLock.WaitAsync().ConfigureAwait(false);
        //    try
        //    {
        //        // 查询现有主分类列表
        //        var categories = await databaseManager.GetAllCadCategoriesAsync().ConfigureAwait(false);

        //        // 计算下一个主分类ID（最小从1开始）
        //        var nextId = (categories == null || categories.Count == 0) ? 1 : categories.Max(c => c.Id) + 1;

        //        // 返回新ID
        //        return nextId;
        //    }
        //    finally
        //    {
        //        // 释放互斥锁
        //        _idLock.Release();
        //    }
        //}

        /// <summary>
        /// 生成子分类ID（统一保证 >=10000，以兼容你现有 parentId>=10000 的层级判断逻辑）
        /// </summary>
        public static async Task<int> GenerateSubcategoryIdAsync(DatabaseManager databaseManager, int parentId)
        {
            // 参数判空
            if (databaseManager == null) throw new ArgumentNullException(nameof(databaseManager));

            // 父级ID校验
            if (parentId <= 0) throw new ArgumentOutOfRangeException(nameof(parentId), "父分类ID必须大于0");

            // 进入互斥区，避免并发重复
            await _idLock.WaitAsync().ConfigureAwait(false);
            try
            {
                // 查询所有子分类
                var subcategories = await databaseManager.GetAllCadSubcategoriesAsync().ConfigureAwait(false);

                // 没有子分类时，从10000开始
                if (subcategories == null || subcategories.Count == 0)
                {
                    return 10000;
                }

                // 取当前最大子分类ID并+1
                var maxId = subcategories.Max(s => s.Id);
                var nextId = maxId + 1;

                // 兜底：确保子分类ID始终>=10000
                if (nextId < 10000)
                {
                    nextId = 10000;
                }

                // 返回新ID
                return nextId;
            }
            finally
            {
                // 释放互斥锁
                _idLock.Release();
            }
        }
    }
}

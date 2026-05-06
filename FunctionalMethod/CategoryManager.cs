using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using static GB_NewCadPlus_IV.WpfMainWindow;
using MessageBox = System.Windows.MessageBox;
using CadCategory = GB_NewCadPlus_IV.FunctionalMethod.CadCategory;
using CadSubcategory = GB_NewCadPlus_IV.FunctionalMethod.CadSubcategory;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 分类管理器
    /// </summary>
    public class CategoryManager
    {
        /// <summary>
        /// 数据库管理器
        /// </summary>
        private readonly DatabaseManager _databaseManager;
        /// <summary>
        /// 数据库管理器
        /// </summary>
        /// <param name="databaseManager"></param>
        public CategoryManager(DatabaseManager databaseManager)
        {
            _databaseManager = databaseManager;
        }

        /// <summary>
        /// 应用分类属性（返回bool值）
        /// </summary>
        /// <returns></returns>
        public async Task<bool> ApplyCategoryPropertiesAsync(ItemsControl categoryPropertiesDataGrid)
        {
            try
            {
                // 从界面数据源读取属性列表
                var properties = categoryPropertiesDataGrid.ItemsSource as List<CategoryPropertyEditModel>;

                // 判空检查，避免空数据写入
                if (properties == null || properties.Count == 0)
                {
                    throw new Exception("没有可应用的属性");
                }

                // 解析属性数据（名称、显示名、排序序号）
                var (name, displayName, sortOrder) = ParseCategoryProperties(properties);

                // 分类名称必填校验
                if (string.IsNullOrEmpty(name))
                {
                    throw new Exception("分类名称不能为空");
                }

                // 如果排序序号未提供或<=0，则自动取最大值+1
                if (sortOrder <= 0)
                {
                    sortOrder = await _databaseManager.GetMaxCadCategorySortOrderAsync() + 1;
                }

                // 创建主分类对象（主分类ID由数据库自增生成，因此此处不手动赋Id）
                var category = new CadCategory
                {
                    Name = name,
                    DisplayName = displayName ?? name,
                    SortOrder = sortOrder,
                    SubcategoryIds = "",
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now
                };

                // 插入数据库
                int result = await _databaseManager.AddCadCategoryAsync(category);

                // 返回是否插入成功
                return result > 0;
            }
            catch (Exception ex)
            {
                // 记录日志并继续抛出，便于上层统一处理
                LogManager.Instance.LogInfo($"应用分类属性时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 应用子分类属性（返回bool值）
        /// </summary>
        /// <returns></returns>
        public async Task<bool> ApplySubcategoryPropertiesAsync( ItemsControl CategoryPropertiesDataGrid)
        {
            try
            {
                var properties = CategoryPropertiesDataGrid.ItemsSource as List<CategoryPropertyEditModel>;
                if (properties == null || properties.Count == 0)
                {
                    throw new Exception("没有可应用的属性");
                }

                // 解析属性数据
                var (parentId, name, displayName, sortOrder) = ParseSubcategoryProperties(properties);

                if (parentId <= 0)
                {
                    throw new Exception("父分类ID必须大于0");
                }

                if (string.IsNullOrEmpty(name))
                {
                    throw new Exception("子分类名称不能为空");
                }

                // 自动生成排序序号（如果未提供或为0）
                if (sortOrder <= 0)
                {
                    sortOrder = await _databaseManager.GetMaxCadSubcategorySortOrderAsync(parentId) + 1;
                }

                // 生成子分类ID
                int subcategoryId = await CategoryIdGenerator.GenerateSubcategoryIdAsync(_databaseManager, parentId);

                // 确定层级
                int level = await DetermineCategoryLevelAsync(parentId);

                // 创建子分类对象
                var subcategory = new CadSubcategory
                {
                    Id = subcategoryId,
                    ParentId = parentId,
                    Name = name,
                    DisplayName = displayName ?? name,
                    SortOrder = sortOrder,
                    Level = level,
                    SubcategoryIds = "", // 新建时为空
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now
                };

                // 添加到数据库
                int result = await _databaseManager.AddCadSubcategoryAsync(subcategory);

                if (result > 0)
                {
                    // 更新父级的子分类列表
                    await _databaseManager.UpdateParentSubcategoryListAsync(parentId, subcategoryId);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"应用子分类属性时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 更新分类属性（编辑功能）
        /// </summary>
        /// <returns></returns>
        public async Task<bool> UpdateCategoryPropertiesAsync( ItemsControl CategoryPropertiesDataGrid, CategoryTreeNode _selectedCategoryNode)
        {
            try
            {
                if (_selectedCategoryNode == null)
                    throw new Exception("没有选中的分类");

                var properties = CategoryPropertiesDataGrid.ItemsSource as List<CategoryPropertyEditModel>;
                if (properties == null || properties.Count == 0)
                    throw new Exception("没有可更新的属性");

                // 解析更新的属性
                var (name, displayName, sortOrder) = WpfMainWindow.ParseUpdatedCategoryProperties(properties);

                if (string.IsNullOrEmpty(name))
                    throw new Exception("分类名称不能为空");

                // 根据节点类型更新相应的记录
                if (_selectedCategoryNode.Level == 0 && _selectedCategoryNode.Data is CadCategory category)
                {
                    // 更新主分类
                    category.Name = name;
                    category.DisplayName = displayName ?? name;
                    category.SortOrder = sortOrder;
                    category.UpdatedAt = DateTime.Now;

                    int result = await _databaseManager.UpdateCadCategoryAsync(category);
                    return result > 0;
                }
                else if (_selectedCategoryNode.Data is CadSubcategory subcategory)
                {
                    // 更新子分类
                    subcategory.Name = name;
                    subcategory.DisplayName = displayName ?? name;
                    subcategory.SortOrder = sortOrder;
                    subcategory.UpdatedAt = DateTime.Now;

                    int result = await _databaseManager.UpdateCadSubcategoryAsync(subcategory);
                    return result > 0;
                }
                else
                {
                    throw new Exception("不支持的分类类型");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"更新分类属性时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 解析分类属性数据
        /// </summary>
        /// <param name="properties"></param>
        /// <returns></returns>
        private (string Name, string DisplayName, int SortOrder) ParseCategoryProperties(List<CategoryPropertyEditModel> properties)
        {
            string name = "";
            string displayName = "";
            int sortOrder = 0; // 默认排序序号

            foreach (var property in properties)
            {
                // 处理第一列
                ProcessCategoryProperty(property.PropertyName1, property.PropertyValue1, ref name, ref displayName, ref sortOrder);

                // 处理第二列
                ProcessCategoryProperty(property.PropertyName2, property.PropertyValue2, ref name, ref displayName, ref sortOrder);
            }

            return (name, displayName, sortOrder);
        }

        /// <summary>
        /// 处理单个分类属性
        /// </summary>
        /// <param name="propertyName">分类名</param>
        /// <param name="propertyValue">分类名对应的值</param>
        /// <param name="name">返回的名称</param>
        /// <param name="displayName">返回的显示名称</param>
        /// <param name="sortOrder">返回的排列顺序</param>
        public static void ProcessCategoryProperty(string propertyName, string propertyValue, ref string name, ref string displayName, ref int sortOrder)
        {
            if (string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(propertyValue))
                return;

            switch (propertyName.ToLower().Trim())
            {
                case "名称":
                case "name":
                    name = propertyValue.Trim();
                    break;
                case "显示名称":
                case "displayname":
                case "显示名":
                    displayName = propertyValue.Trim();
                    break;
                case "排序序号":
                case "sortorder":
                    // 排序序号现在是可选的，如果提供了就使用，否则自动生成
                    if (int.TryParse(propertyValue.Trim(), out int sort))
                        sortOrder = sort;
                    break;
            }
        }

        /// <summary>
        /// 解析子分类属性数据
        /// </summary>
        /// <param name="properties"></param>
        /// <returns></returns>
        private (int ParentId, string Name, string DisplayName, int SortOrder) ParseSubcategoryProperties(List<CategoryPropertyEditModel> properties)
        {
            int parentId = 0;
            string name = "";
            string displayName = "";
            int sortOrder = 0; // 默认排序序号

            foreach (var property in properties)
            {
                // 处理第一列
                ProcessSubcategoryProperty(property.PropertyName1, property.PropertyValue1, ref parentId, ref name, ref displayName, ref sortOrder);

                // 处理第二列
                ProcessSubcategoryProperty(property.PropertyName2, property.PropertyValue2, ref parentId, ref name, ref displayName, ref sortOrder);
            }

            return (parentId, name, displayName, sortOrder);
        }

        /// <summary>
        /// 处理单个子分类属性
        /// </summary>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        /// <param name="parentId"></param>
        /// <param name="name"></param>
        /// <param name="displayName"></param>
        /// <param name="sortOrder"></param>
        private void ProcessSubcategoryProperty(string propertyName, string propertyValue, ref int parentId, ref string name, ref string displayName, ref int sortOrder)
        {
            if (string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(propertyValue))
                return;

            switch (propertyName.ToLower().Trim())
            {
                case "父分类id":
                case "parentid":
                case "父id":
                    if (int.TryParse(propertyValue.Trim(), out int pid))
                        parentId = pid;
                    break;
                case "名称":
                case "name":
                    name = propertyValue.Trim();
                    break;
                case "显示名称":
                case "displayname":
                case "显示名":
                    displayName = propertyValue.Trim();
                    break;
                case "排序序号":
                case "sortorder":
                    // 排序序号现在是可选的
                    if (int.TryParse(propertyValue.Trim(), out int sort))
                        sortOrder = sort;
                    break;
            }
        }

        /// <summary>
        /// 删除主分类
        /// </summary>
        /// <param name="categoryNode"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public async Task<bool> DeleteMainCategoryAsync(CategoryTreeNode categoryNode)
        {
            try
            {
                // 检查是否有子分类
                if (categoryNode.Children.Count > 0)
                {
                    if (MessageBox.Show("该主分类下还有子分类，确定要全部删除吗？",
                                      "警告", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                    {
                        return false;
                    }
                }

                // 从数据库删除主分类
                int result = await _databaseManager.DeleteCadCategoryAsync(categoryNode.Id);
                return result > 0;
            }
            catch (Exception ex)
            {
                throw new Exception($"删除主分类失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 删除子分类
        /// </summary>
        /// <param name="subcategoryNode"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public async Task<bool> DeleteSubcategoryAsync(CategoryTreeNode subcategoryNode)
        {
            try
            {
                // 检查是否有子分类
                if (subcategoryNode.Children.Count > 0)
                {
                    if (MessageBox.Show("该子分类下还有子分类，确定要全部删除吗？",
                                      "警告", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                    {
                        return false;
                    }
                }

                // 从数据库删除子分类
                int result = await _databaseManager.DeleteCadSubcategoryAsync(subcategoryNode.Id);
                if (result > 0)
                {
                    // 更新父级的子分类列表
                    await UpdateParentSubcategoryListAfterDeleteAsync(subcategoryNode.ParentId, subcategoryNode.Id);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception($"删除子分类失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 删除后更新父级子分类列表
        /// </summary>
        /// <param name="parentId"></param>
        /// <param name="deletedSubcategoryId"></param>
        /// <returns></returns>
        private async Task UpdateParentSubcategoryListAfterDeleteAsync(int parentId, int deletedSubcategoryId)
        {
            try
            {
                // 获取父级记录
                string currentSubcategoryIds = "";
                if (parentId >= 10000)
                {
                    // 父级是子分类
                    var parentSubcategory = await _databaseManager.GetCadSubcategoryByIdAsync(parentId);
                    currentSubcategoryIds = parentSubcategory?.SubcategoryIds ?? "";
                }
                else
                {
                    // 父级是主分类
                    var categories = await _databaseManager.GetAllCadCategoriesAsync();
                    var parentCategory = categories.FirstOrDefault(c => c.Id == parentId);
                    currentSubcategoryIds = parentCategory?.SubcategoryIds ?? "";
                }

                // 更新子分类列表（移除已删除的ID）
                if (!string.IsNullOrEmpty(currentSubcategoryIds))
                {
                    var ids = currentSubcategoryIds.Split(',').Select(id => id.Trim()).ToList();
                    ids.Remove(deletedSubcategoryId.ToString());
                    string newSubcategoryIds = string.Join(",", ids);

                    // 更新数据库
                    await _databaseManager.UpdateParentSubcategoryListAsync(parentId, newSubcategoryIds);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"更新父级子分类列表失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新架构树显示（修正版） RefreshCategoryTreeAsync
        /// </summary>
        /// <returns></returns>
        public async Task RefreshCategoryTreeAsync(CategoryTreeNode _selectedCategoryNode,ItemsControl _categoryTreeView,
            List<CategoryTreeNode> _categoryTreeNodes,DatabaseManager databaseManager)
        {
            try
            {
                // 重新加载分类和子分类数据 LoadCategoryTreeAsync
                await LoadCategoryTreeAsync(_categoryTreeNodes, databaseManager);

                // 更新UI显示
                DisplayCategoryTree(_categoryTreeView, _categoryTreeNodes);

                // 展开当前选中的节点
                if (_selectedCategoryNode != null && _categoryTreeView != null)
                {
                    ExpandTreeNodeToSelectedNode(_categoryTreeView, _selectedCategoryNode);
                }

                LogManager.Instance.LogInfo("架构树刷新完成");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"刷新架构树时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 加载架构树数据
        /// </summary>
        /// <returns></returns>
        public async Task LoadCategoryTreeAsync(List<CategoryTreeNode> _categoryTreeNodes, DatabaseManager databaseManager)
        {
            try
            {
                _categoryTreeNodes.Clear();

                // 获取所有分类和子分类
                var categories = await databaseManager.GetAllCadCategoriesAsync();
                var subcategories = await databaseManager.GetAllCadSubcategoriesAsync();

                // 构建树结构
                BuildCategoryTree(categories, subcategories, _categoryTreeNodes);

            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载架构树数据失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 构建分类树结构
        /// </summary>
        /// <param name="categories"></param>
        /// <param name="subcategories"></param>
        private static void BuildCategoryTree(List<CadCategory> categories, List<CadSubcategory> subcategories,List<CategoryTreeNode> _categoryTreeNodes)
        {
            // 清空现有节点
            _categoryTreeNodes.Clear();

            // 1. 创建主分类节点
            var mainCategoryNodes = categories
                .OrderBy(c => c.SortOrder)
                .Select(c => new CategoryTreeNode(c.Id, c.Name, c.DisplayName, 0, 0, c))
                .ToList();

            // 2. 创建子分类节点字典，便于快速查找
            var subcategoryDict = subcategories
                .ToDictionary(s => s.Id, s => new CategoryTreeNode(
                    s.Id, s.Name, s.DisplayName, s.Level, s.ParentId, s));

            // 3. 构建父子关系
            BuildTreeRelationships(mainCategoryNodes, subcategoryDict);

            // 4. 将根节点添加到树节点列表
            _categoryTreeNodes.AddRange(mainCategoryNodes);

        }

        /// <summary>
        /// 构建树的父子关系
        /// </summary>
        /// <param name="mainNodes"></param>
        /// <param name="subcategoryDict"></param>
        private static void BuildTreeRelationships(List<CategoryTreeNode> mainNodes, Dictionary<int, CategoryTreeNode> subcategoryDict)
        {
            // 创建所有节点的查找字典
            var allNodesDict = new Dictionary<int, CategoryTreeNode>();

            // 添加主分类节点
            foreach (var node in mainNodes)
            {
                allNodesDict[node.Id] = node;
            }

            // 添加子分类节点
            foreach (var kvp in subcategoryDict)
            {
                allNodesDict[kvp.Key] = kvp.Value;
            }

            // 建立父子关系
            foreach (var node in subcategoryDict.Values)
            {
                if (allNodesDict.ContainsKey(node.ParentId))
                {
                    var parentNode = allNodesDict[node.ParentId];
                    parentNode.Children.Add(node);
                }
                else if (node.Level == 1)
                {
                    // 二级子分类，父级是主分类
                    var mainCategoryNode = mainNodes.FirstOrDefault(n => n.Id == node.ParentId);
                    if (mainCategoryNode != null)
                    {
                        mainCategoryNode.Children.Add(node);
                    }
                }
            }

            // 对所有节点的子节点按排序序号排序
            foreach (var node in allNodesDict.Values)
            {
                node.Children = node.Children.OrderBy(child => GetSortOrder(child)).ToList();
            }
        }

        /// <summary>
        /// 获取节点的排序序号
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        private static int GetSortOrder(CategoryTreeNode node)
        {
            if (node.Data is CadCategory category)
                return category.SortOrder;
            else if (node.Data is CadSubcategory subcategory)
                return subcategory.SortOrder;
            return 0;
        }

        /// <summary>
        /// 显示架构树
        /// </summary>
        public  void DisplayCategoryTree(ItemsControl CategoryTreeView, List<CategoryTreeNode> _categoryTreeNodes)
        {
            try
            {
                if (CategoryTreeView != null)
                {
                    CategoryTreeView.ItemsSource = null;
                    CategoryTreeView.ItemsSource = _categoryTreeNodes;
                    // 添加TreeView的选择事件
                    //CategoryTreeView.SelectedItemChanged += CategoryTreeView_SelectedItemChanged;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"显示架构树失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 展开树节点到选中的节点
        /// </summary>
        /// <param name="selectedNode">选中的节点</param>
        private static void ExpandTreeNodeToSelectedNode(ItemsControl _categoryTreeView,CategoryTreeNode selectedNode)
        {
            try
            {
                // 这里可以实现展开树节点到指定节点的逻辑
                // 例如：展开父节点，选中指定节点等
                if (_categoryTreeView != null && _categoryTreeView.Items.Count > 0)
                {
                    // 可以通过遍历TreeViewItem来展开到指定节点
                    // 这里简化处理，实际可以根据需要完善
                    _categoryTreeView.UpdateLayout();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"展开树节点时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 递归展开所有子节点
        /// </summary>
        /// <param name="item"></param>
        private void ExpandAllChildren(TreeViewItem item)
        {
            if (item == null) return;

            item.IsExpanded = true;
            foreach (var child in item.Items)
            {
                var childItem = GetTreeViewItem(item, child);
                if (childItem != null)
                {
                    ExpandAllChildren(childItem);
                }
            }
        }

        /// <summary>
        /// 获取TreeViewItem的辅助方法（增强版）
        /// </summary>
        /// <param name="container"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        private TreeViewItem GetTreeViewItem(ItemsControl container, object item)
        {
            if (container == null) return null;

            // 首先尝试直接获取
            var directlyFound = container.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
            if (directlyFound != null)
                return directlyFound;

            // 如果直接获取失败，遍历所有子项
            if (container.Items != null)
            {
                foreach (var containerItem in container.Items)
                {
                    var treeViewItem = container.ItemContainerGenerator.ContainerFromItem(containerItem) as TreeViewItem;
                    if (treeViewItem != null)
                    {
                        if (treeViewItem.DataContext == item)
                        {
                            return treeViewItem;
                        }

                        // 递归查找子项
                        var child = GetTreeViewItem(treeViewItem, item);
                        if (child != null)
                        {
                            return child;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 初始化分类属性编辑网格
        /// </summary>
        private void InitializeCategoryPropertyGrid(ItemsControl categoryPropertiesDataGrid)
        {
            var initialRows = new List<CategoryPropertyEditModel>
                {
                    new CategoryPropertyEditModel(),
                    new CategoryPropertyEditModel(),
                    new CategoryPropertyEditModel()
                };

            categoryPropertiesDataGrid.ItemsSource = initialRows;
            LogManager.Instance.LogInfo("初始化分类属性编辑网格成功:InitializeCategoryPropertyGrid()");
        }

        /// <summary>
        /// 初始化子分类属性编辑界面
        /// </summary>
        /// <param name="parentNode"></param>
        private void InitializeSubcategoryPropertiesForEditing(CategoryTreeNode parentNode,ItemsControl categoryPropertiesDataGrid)
        {
            var subcategoryProperties = new List<CategoryPropertyEditModel>
            {
                new CategoryPropertyEditModel { PropertyName1 = "父分类ID", PropertyValue1 = parentNode.Id.ToString(), PropertyName2 = "名称", PropertyValue2 = "" },
                new CategoryPropertyEditModel { PropertyName1 = "显示名称", PropertyValue1 = "", PropertyName2 = "排序序号", PropertyValue2 = "自动生成" } // 留空，表示自动生成
            };

            // 添加参考信息
            subcategoryProperties.Add(new CategoryPropertyEditModel
            {
                PropertyName1 = "父级名称",
                PropertyValue1 = parentNode.DisplayText,
                PropertyName2 = "",
                PropertyValue2 = ""
            });

            // 添加空行用于用户输入
            subcategoryProperties.Add(new CategoryPropertyEditModel());
            subcategoryProperties.Add(new CategoryPropertyEditModel());

            categoryPropertiesDataGrid.ItemsSource = subcategoryProperties;
        }

        /// <summary>
        /// 初始化主分类属性编辑界面
        /// </summary>
        private void InitializeCategoryPropertiesForCategory(ItemsControl categoryPropertiesDataGrid)
        {
            var categoryProperties = new List<CategoryPropertyEditModel>
            {
                new CategoryPropertyEditModel { PropertyName1 = "名称", PropertyValue1 = "", PropertyName2 = "显示名称", PropertyValue2 = "" },
                new CategoryPropertyEditModel { PropertyName1 = "排序序号", PropertyValue1 = "自动生成", PropertyName2 = "", PropertyValue2 = "" } // 留空，表示自动生成
            };

            // 添加空行用于用户输入
            categoryProperties.Add(new CategoryPropertyEditModel());
            categoryProperties.Add(new CategoryPropertyEditModel());

            categoryPropertiesDataGrid.ItemsSource = categoryProperties;
        }

        /// <summary>
        /// 显示节点属性用于编辑
        /// </summary>
        /// <param name="node"></param>
        private void DisplayNodePropertiesForEditing(CategoryTreeNode node,ItemsControl categoryPropertiesDataGrid)
        {
            try
            {
                var propertyRows = new List<CategoryPropertyEditModel>();

                if (node.Level == 0 && node.Data is CadCategory category)
                {
                    // 主分类
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "ID",
                        PropertyValue1 = category.Id.ToString(),
                        PropertyName2 = "名称",
                        PropertyValue2 = category.Name
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "显示名称",
                        PropertyValue1 = category.DisplayName,
                        PropertyName2 = "排序序号",
                        PropertyValue2 = category.SortOrder.ToString()
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "子分类数",
                        PropertyValue1 = GetSubcategoryCount(category).ToString(),
                        PropertyName2 = "",
                        PropertyValue2 = ""
                    });
                }
                else if (node.Data is CadSubcategory subcategory)
                {
                    // 子分类
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "ID",
                        PropertyValue1 = subcategory.Id.ToString(),
                        PropertyName2 = "父ID",
                        PropertyValue2 = subcategory.ParentId.ToString()
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "名称",
                        PropertyValue1 = subcategory.Name,
                        PropertyName2 = "显示名称",
                        PropertyValue2 = subcategory.DisplayName
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "排序序号",
                        PropertyValue1 = subcategory.SortOrder.ToString(),
                        PropertyName2 = "层级",
                        PropertyValue2 = subcategory.Level.ToString()
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "子分类数",
                        PropertyValue1 = subcategory.SubcategoryIds.Split(',').Length.ToString(),
                        PropertyName2 = "",
                        PropertyValue2 = ""
                    });
                }

                // 添加空行用于编辑
                propertyRows.Add(new CategoryPropertyEditModel());
                propertyRows.Add(new CategoryPropertyEditModel());

                categoryPropertiesDataGrid.ItemsSource = propertyRows;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"显示节点属性失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取分类数量
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        private int GetSubcategoryCount(CadCategory category)
        {
            if (string.IsNullOrEmpty(category.SubcategoryIds))
                return 0;

            return category.SubcategoryIds.Split(',').Length;
        }

        /// <summary>
        /// 确定分类层级
        /// </summary>
        /// <param name="parentId"></param>
        /// <returns></returns>
        private async Task<int> DetermineCategoryLevelAsync(int parentId)
        {
            if (parentId < 10000)
            {
                // 父级是主分类（4位），这是二级子分类
                return 1;
            }
            else
            {
                // 父级是子分类，需要确定是几级子分类
                var parentSubcategory = await _databaseManager.GetCadSubcategoryByIdAsync(parentId);
                if (parentSubcategory != null)
                {
                    return parentSubcategory.Level + 1;
                }
                else
                {
                    // 默认为二级分类
                    return 1;
                }
            }
        }


    }
}

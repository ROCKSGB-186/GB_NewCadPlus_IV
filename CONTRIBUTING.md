# 贡献指南 (CONTRIBUTING.md)

## 目的
本文件规定了项目的协作流程、代码风格与提交规范，帮助贡献者快速上手并保持代码库一致性。项目为 WPF 桌面应用，目标运行时为 .NET Framework 4.8，开发环境建议使用 Visual Studio 2026 或兼容版本。

## 开发环境
- 操作系统：Windows 10/11
- IDE：__Visual Studio__（推荐使用 __Visual Studio 2026__）
- 目标框架：.NET Framework 4.8
- 代码仓库：Git（使用集中式分支策略或 GitFlow，详见下文）

## 分支与提交策略
- 主分支：`main` — 始终保持可发布状态。任何合并到 `main` 的变更必须通过 CI 和代码审查。
- 功能分支：`feature/<描述>`，每个功能或修复应在独立分支上开发。
- 修复分支：`hotfix/<描述>` 用于紧急修复。
- 提交信息：使用简洁的英文或中英双语说明，首行不超过 72 字符，示例：`fix: 修复文件选择对话框的异常`。

## Pull Request (PR) 要求
- 每个 PR 应包含清晰的描述、关联的 issue（如有）和变更截图（UI 变更时）。
- 指定至少一位代码评审者。
- 通过静态检查/单元测试后合并。
- PR 模板请包含：目的、变更点、回归测试步骤、影响范围。

## 代码风格（C#）
- 使用 4 个空格缩进（不使用 Tab）。
- 文件编码 UTF-8 无 BOM。
- 命名规范：
  - 类型 (class/enum/interface/struct) 使用 PascalCase。
  - 方法和属性使用 PascalCase。
  - 私有字段使用 `_camelCase`（下划线前缀 + 驼峰）。
  - 局部变量使用 camelCase。
  - 事件使用 PascalCase，并以语义化名称结尾。
- 使用表达式体、模式匹配等现代 C# 特性保持简洁（在兼容范围内）。
- 避免代码中硬编码密码或敏感信息。
- 异步方法以 `Async` 后缀命名并使用 `async/await`。

## XAML 与 WPF 规范
- 将样式、模板与资源放入独立的 ResourceDictionary（例如 `Resources/`），并在 `App.xaml` 中合并。
- 避免在单一 `UserControl` 或 Window 中放入过多控件；通过拆分子控件提高可维护性与编译速度。
- 使用数据绑定和 `ICommand` 代替代码后置 Click 事件（遵循 MVVM 模式）。
- 控件命名使用英文 ASCII（尽量避免中文标识符），采用 PascalCase，例如 `BtnSave`、`TxtFilePath`。
- 可重用的控件样式使用静态资源键（例如 `PrimaryButtonStyle`），颜色使用资源键命名（例如 `AccentBrush`）。
- 对大型列表或多元素容器使用虚拟化（VirtualizingStackPanel/ItemsControl + virtualization）以提升性能。
- Image 资源请使用打包资源或异步加载，避免在 UI 线程同步读取大文件。

## 资源/配置管理
- 不要在源码中保存生产数据库用户名/密码。使用配置文件（`app.config`）或受保护的凭据管理手段。
- 对于可变配置（如默认路径、同步间隔），放入配置文件并提供 UI 配置页面，避免硬编码。

## 日志与异常处理
- 使用集中日志组件（例如 log4net、NLog 或自建轻量日志），捕获未处理异常并友好提示用户。
- 对外部调用（数据库、文件 IO）进行超时、权限和异常处理，并在 UI 层显示可操作的错误信息。

## 性能与可维护性建议
- 将大型 XAML 拆分为多个 UserControl。减少单个文件的复杂度以降低编译时间。
- 使用虚拟化和延迟加载：例如延迟加载 Tab 内容、图片和大列表项。
- 避免在 UI 线程执行耗时操作，使用后台线程或 Task.Run 并反馈进度。

## 测试
- 编写单元测试（业务逻辑尽量与 View 分离），使用 NUnit、xUnit 或 MSTest。
- 关键流程（文件读写、数据库访问、图元导入导出）应有集成测试或手动测试说明。

## CI / 自动化
- 建议添加 CI（GitHub Actions 或 Azure Pipelines）用于：构建、运行单元测试、静态分析、生成发布包。
- 在 CI 中运行静态代码分析（例如 Roslyn 分析器）和格式检查。

## 提交前检查列表（PR Checklist）
- [ ] 代码通过本地构建。
- [ ] 主要变更有单元测试或手动测试说明。
- [ ] 无硬编码敏感信息。
- [ ] XAML 样式/资源已复用而非复制。
- [ ] 命名符合规范，控件名使用英文 ASCII。
- [ ] UI 变更包含截图。

## 联系人与流程
- 新贡献者请在 issue 中说明意图或在 PR 中关联 issue。
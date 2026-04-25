# Copilot Instructions

## General Guidelines
- First general instruction
- Second general instruction
- 在生成或修改代码时添加中文注释，以便于理解。代码方案尽量每一步都添加中文注释，尤其是插入图元与属性同步逻辑，以便于小白理解。用户希望我提供可直接替换的完整方法代码，并且每行都加中文注释，便于理解和自行维护。
- 偏好继续使用注释缩放（annotative），在标注样式中接受原始存储值（Dimtxt 显示为大值），并让实际显示高度由 Dimtxt * Dimlfac 决定；项目基准显示高度设为 3.5。用户要求将文本样式（如 tJText）的 TextSize 设为 0 以避免样式固定高度覆盖实体高度。
- 修改包含中文标识符和中文字符串的代码文件时，必须避免使用会破坏文件编码的写入方式，优先保持原文件编码并避免造成中文乱码。
- 用户偏好按步骤、细致地指导数据库与代码迁移实施方案。

## Code Style
- Use specific formatting rules
- Follow naming conventions

## Project-Specific Rules
- 用户的项目使用附加属性 SelectionManager 放在命名空间 GB_NewCadPlus_IV.FunctionalMethod，并在 XAML 中通过前缀 fm 引用。今后在 XAML/样式修订中应确认 xmlns 包含 assembly 指定并在 ControlTemplate 中为被触发的元素添加 x:Name.
- 保留工艺图元按钮的拖拽命令，仅在“双击”或“单击按住后拖动”时触发，不能在普通悬停时触发。
- 在 AreaByPoints 交互中，按 Z 必须可撤销上一步，并保留选点过程的动态预览连线。
- 首选将功能按文件拆分：线型相关放在 LineTypeStyleHelper.cs，通用系统变量读取放在 AutoCadHelper.cs，命令入口保留在 FunctionalMethod\Command.cs；今后按此约定拆分方法。

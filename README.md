# SmartSeating

SmartSeating 是一个使用 .NET 10 / WPF 构建的智能座位编排桌面应用。它面向教师或行政管理者，支持批量导入学生名单、灵活编辑座位状态以及运行模拟退火算法来生成更合理的座位方案。

## 功能亮点

- **可视化座位画布**：支持任意行列的座位生成，实时拖拽多选、右键菜单快捷操作以及悬停提示。
- **学生信息管理**：可编辑身高、重要度、希望靠近/远离的学生以及偏好区域，支持 Excel 名单导入与 JSON 配置备份/恢复。
- **智能排座**：基于模拟退火的优化引擎，可综合考虑固定/禁用座位、学生偏好与权重，自动给出较优方案，并显示进度条与遮罩避免误操作。
- **增强的学生列表**：新增搜索框快速过滤学生、双击自动定位到其座位，并在左侧面板底部显示实时的座位统计（总座位、已安排、空位、未安排学生数等）。
- **多种导出方式**：一键导出当前布局为 Excel、PNG 图片或 JSON 配置，方便分享或打印。

## 快速开始

1. **准备环境**
   - 安装 [.NET SDK](https://dotnet.microsoft.com/download)（需要支持 `net10.0-windows` 目标框架）。
   - 确保 Windows 上具备 WPF 运行环境。

2. **获取代码**
   ```bash
   git clone <repo-url>
   cd SmartSeating
   ```

3. **构建与运行**
   ```bash
   dotnet build
   dotnet run --project SmartSeating/SmartSeating.csproj
   ```

4. **导入与编排**
   - 通过顶部工具条导入 Excel 名单或 JSON 配置。
   - 使用左侧搜索框快速定位学生，右侧面板可编辑详细权重与偏好。
   - 点击“智能排座”让系统自动寻找较优方案；也可以手动拖拽、右键菜单或快捷键（`F` 固定、`D` 禁用、`Delete` 清空）调整。

## 目录结构

- `SmartSeating/`：WPF 主项目，包含 `MainWindow`、样式与排座逻辑。
- `README.md`：项目说明（即本文件）。
- `SmartSeating.slnx`：解决方案文件，方便使用 Visual Studio / Rider 打开。

## 贡献 & 反馈

欢迎通过 Issue 或 Pull Request 提交优化建议、Bug 反馈以及新的功能点。如果在使用过程中有任何问题，也可以附带配置文件或复现步骤，方便我们定位。

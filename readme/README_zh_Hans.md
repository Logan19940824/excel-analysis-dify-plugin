## Excel Analysis

**作者：** modsdom  
**版本：** 0.1.0  
**类型：** Tool

### 功能

本插件用于分析 Dify 工作流中提供的 Excel 或 CSV 文件，提供两个无状态工具：

- `get_workbook_info`：读取工作表、列名、数据类型、行数和示例数据。
- `query_workbook`：使用 DuckDB 对文件执行只读 `SELECT` 或 `CTE` 查询。

插件会在每次调用时下载 Dify 提供的 HTTP(S) 文件 URL，在临时目录中分析文件，返回结果后删除临时文件，不会持久化工作簿。

### 使用方法

1. 在 Dify 中安装并启用 `excel_analysis` 插件。
2. 先调用 `get_workbook_info`，获取工作表名称、列名和文件 `sha256`。
3. 调用 `query_workbook` 时，继续使用同一个文件 URL，并传入上一步返回的 `sha256`。
4. SQL 必须使用实际返回的工作表名和列名，只能执行一条只读 `SELECT` 或 `CTE` 语句。

### 输入说明

- `file_url`：Dify 可访问的 Excel 或 CSV 文件 HTTP(S) 下载地址。
- `sha256`：`get_workbook_info` 返回的文件摘要，用于确认查询期间文件未被替换。
- `sql`：单条只读 DuckDB `SELECT` 或 `CTE` 查询。
- `limit`：返回行数上限，范围为 `1` 到 `1000`，默认值为 `1000`。

### 限制

- 文件 URL 必须在两次调用之间保持有效且指向同一个文件。
- 查询不能写入数据、执行多条语句或访问外部资源。
- 大型工作簿的处理时间和内存占用取决于 Dify 插件运行环境。

### 隐私

详见根目录的 `PRIVACY.md`。

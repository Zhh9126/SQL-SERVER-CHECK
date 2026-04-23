# SQL Server 数据库巡检报告生成工具

一款功能强大的 SQL Server 数据库健康巡检工具，专为 Windows 服务器环境设计。自动采集数据库配置、性能指标、主机资源等信息，生成专业的 Word 巡检报告，支持单库和批量巡检模式。

## ✨ 功能特性

- **全面巡检**：收集版本信息、配置参数、数据库大小、活动会话、等待统计、阻塞进程、备份信息、锁与死锁等 20+ 项指标
- **主机资源监控**：通过 `psutil` 获取 Windows 主机的 CPU、内存、磁盘使用率
- **专业报告**：自动生成格式规范、内容详尽的 Word 报告（`.docx`）
- **批量巡检**：支持通过 Excel 文件配置多个数据库连接，一键批量巡检并生成汇总报告
- **灵活部署**：支持单库手动输入连接信息或使用配置文件
- **许可证管理**：内置试用期验证机制，保护商业使用

## 🖥️ 环境要求

- **操作系统**：Windows Server / Windows 10/11
- **Python**：3.6 及以上
- **ODBC 驱动**：SQL Server ODBC Driver（如 `ODBC Driver 17 for SQL Server` 或 `SQL Server`）
- **数据库权限**：建议使用具有 `VIEW SERVER STATE`、`VIEW ANY DEFINITION`、`msdb` 读权限的账户

## 📦 安装步骤

1. **克隆仓库**
   ```bash
   git clone https://github.com/Zhh9126/SQL-SERVER-CHECK.git
   cd SQL-SERVER-CHECK
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```
   `requirements.txt` 内容：
   ```
   pyodbc
   python-docx
   psutil
   openpyxl
   configparser
   ```

3. **配置 ODBC 驱动**  
   确保系统已安装 SQL Server ODBC 驱动。可通过 `ODBC 数据源管理器` 查看或使用以下 PowerShell 命令检查：
   ```powershell
   Get-OdbcDriver | Where-Object {$_.Name -like "*SQL Server*"}
   ```

## ⚙️ 配置文件

### SQL 模板文件（自动生成）
首次运行时，程序会自动在 `templates/sqltemplates_sqlserver.ini` 生成默认的 SQL 查询模板。您可以根据需要修改其中的查询语句。

### 数据库连接配置（INI 格式）
创建 `conf/sqlserver.ini` 文件，内容示例：
```ini
[PROD_DB]
name = AdventureWorks
server = 192.168.1.100
port = 1433
user = inspector
password = YourPassword
driver = {ODBC Driver 17 for SQL Server}

[TEST_DB]
name = master
server = localhost
port = 1433
user = sa
password = SaPassword
driver = {SQL Server}
```

### 批量巡检 Excel 模板
运行以下命令生成模板：
```bash
python sqlserver_inspector.py -G connections_template.xlsx
```
模板包含以下列：
| Label | Server | Port | Database | User | Password | Driver |
|-------|--------|------|----------|------|----------|--------|

## 🚀 使用方法

### 1. 单库巡检（交互式）
```bash
python sqlserver_inspector.py
# 选择菜单 1，按提示输入连接信息
```

### 2. 使用配置文件巡检
```bash
# 巡检配置文件中指定的标签（例如 PROD_DB）
python sqlserver_inspector.py -L PROD_DB

# 巡检配置文件中的所有数据库
python sqlserver_inspector.py
# 选择菜单 1 并选择“使用配置文件”选项（需修改代码支持，或直接使用 -L 参数遍历）
```

### 3. 批量巡检（Excel 方式）
```bash
# 生成 Excel 模板后填写连接信息
python sqlserver_inspector.py -E connections.xlsx -S
```
参数说明：
- `-E`：指定 Excel 连接文件
- `-S`：生成汇总 Excel 报告（可选）

### 4. 生成 Excel 模板
```bash
python sqlserver_inspector.py -G my_template.xlsx
```

### 5. 查看帮助
```bash
python sqlserver_inspector.py -h
```

## 📄 输出报告

- **Word 报告**：保存在 `reports/` 目录，命名格式为 `SQLServer_{标签}_report_{时间戳}.docx`
- **汇总 Excel**（批量模式）：`summary_report_{时间戳}.xlsx`，包含各数据库关键指标摘要

报告内容包括：
- 数据库版本、配置参数
- 数据库大小、状态、恢复模式
- 活动会话、最大连接数
- 等待统计 TOP10、阻塞进程
- 锁数量、死锁统计
- 备份信息（最近一次）
- 主机 CPU、内存、磁盘使用率
- 综合健康评分与告警

## 📁 项目结构

```
sqlserver-inspector/
├── main.py      # 主程序
├── license_manager.py          # 许可证验证模块（需自行实现或移除）
├── requirements.txt            # 依赖列表
├── conf/                       # 配置文件目录
│   └── sqlserver.ini           # 数据库连接配置（可选）
├── templates/                  # SQL 模板目录
│   └── sqltemplates_sqlserver.ini  # 自动生成的 SQL 查询模板
├── reports/                    # 生成报告的输出目录
└── README.md
```

## ⚠️ 注意事项

1. **许可证验证**：程序默认包含 `LicenseValidator` 类，您需要实现或移除该逻辑。若移除，请同时删除相关调用代码。
2. **ODBC 驱动名称**：不同环境驱动名称可能略有差异，程序会自动匹配可用的 SQL Server 驱动。
3. **SQL_VARIANT 类型处理**：所有查询结果已做类型安全转换，避免因 `SQL_VARIANT` 导致报告生成失败。
4. **Windows 兼容性**：主机资源获取依赖 `psutil`，在 Windows 下测试通过。若在 Linux 运行，部分命令可能失效。
5. **权限要求**：部分查询（如备份信息、死锁历史）需要 `msdb` 数据库读权限，若权限不足对应章节将显示为空。

## 📜 开源许可证

本项目采用 **MIT 许可证**。详见 [LICENSE](LICENSE) 文件。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request。请确保代码风格一致，并添加必要的注释。

## 📧 联系方式

如有问题或建议，请通过 [GitHub Issues](https://github.com/yourusername/sqlserver-inspector/issues) 联系。

---

**Happy Inspecting! 🚀**

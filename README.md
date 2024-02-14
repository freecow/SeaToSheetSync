# SeaToSheetSync
SeaToSheetSync is a Python tool designed to seamlessly sync data from SeaTable to Excel based on unique identifiers, streamlining the process of exporting structured data for analysis or reporting.

`SeaToSheetSync`是一个专门设计用于从Seatable中提取数据，并根据编号关联字段自动填写到Excel文件中的工具。特别适合日常工作中，既需要将Excel文件导入到Seatable中进行分析，又需要将Seatable结果数据通过关联字段，导出到Excel中进行进一步分析或报告的场景。

## 特性

- **自动数据同步**：从Seatable云端自动提取数据并填充到Excel。
- **配置驱动**：通过简单的JSON配置，定制化需同步的数据字段。
- **安全性**：敏感信息如数据库凭证、API密钥等存放在`.env`文件，避免硬编码。

## 安装

本工具依赖于Python环境。确保你的系统已安装Python 3.6+。使用以下步骤安装所需的依赖包：

1. 克隆本仓库到本地：

```bash
git clone https://github.com/freecow/SeaToSheetSync.git
cd SeaToSheetSync
```

2. 安装所需的Python依赖包：

```bash
pip install -r requirements.txt
```

## 配置

### `.env` 文件

将你的敏感信息如Seatable的`server_url`和`api_token`，以及MySQL的连接参数放在`.env`文件中。示例内容如下：

```plaintext
SEATABLE_SERVER_URL=https://cloud.seatable.cn
SEATABLE_API_TOKEN=your_api_token_here
```

确保`.env`文件不被提交到版本控制系统中（比如通过`.gitignore`）。
SEATABLE_API_TOKEN请参考Seatable官方开发文档。

### `config.json` 文件

在`config.json`中定义你的业务信息，如数据映射和同步规则。示例结构如下：

```json
{
  "table_name": "Seatable表名称",
  "excel_file_path": "Excel XLSX文件名",
  "sheet_name": "Excel表名",
  "relation_field": "关联字段名称",
  "field_mappings": {
    "列名1": 关联字段对应Excel Sheet中的列数,
    "列名2": 需填充字段对应Excel Sheet中的列数,
    "列名3": 需填充字段对应Excel Sheet中的列数
    ...
  }
}
```

### 示例

通过合同编号关联字段，填充Excel表中的下游领域细化列，`合同编号`位于该Sheet的第19列，`下游领域`细化位于该Sheet的第37列。

```json
{
  "table_name": "收入成本大表",
  "excel_file_path": "test1.xlsx",
  "sheet_name": "收入成本大表",
  "relation_field": "合同编号",
  "field_mappings": {
    "合同编号": 19,
    "下游领域细化": 37
  }
}
```

## 使用

配置好`.env`和`config.json`文件后，运行以下命令启动数据同步：

```bash
python main.py
```

确保`main.py`是你的主程序入口文件。

## 安全性提示

- 不要在任何公共或不安全的地方暴露你的`.env`文件或`config.json`中的敏感信息。
- 定期更新你的Seatable API Token以保持安全。

# 📦 库房出入库数据分析与规划工具

基于 **Streamlit** 的 Web 应用，用于分析库房出入库 Excel 数据，输出物流库房规划建议。

## ✨ 功能模块

| 模块 | 输出内容 |
|------|---------|
| 📊 概览 | 出入库总量、笔数、单日峰值/谷值、每日趋势图 |
| ⏱ 频次分析 | 按小时分布柱状图、高峰时段识别 |
| 🧬 批次分布分析 | (生产订单×料号) 涉及批次数量分布、每日多批次发生情况 |
| 🏭 在库时间分析 | FIFO 匹配、平均/中位/最长在库天数、分布直方图 |
| 📋 规划建议 | 托盘级库容、收货区/发货区人员配置 |
| 🤖 AI 智能解读 | 基于大模型自动生成分析结论；无密钥或调用失败时回退规则版 |

## 🚀 本地运行

### 方式一：一键启动（Windows）

1. 创建虚拟环境并安装依赖：
   ```bat
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. 双击 `launch.bat`，浏览器自动打开 `http://localhost:8501`

### 方式二：命令行启动

```bash
pip install -r requirements.txt
streamlit run warehouse_analyzer.py
```

## 🌐 在线部署

### 方式 A：Streamlit Community Cloud（免费，推荐）

1. 将本项目推送到 GitHub 仓库
2. 访问 [share.streamlit.io](https://share.streamlit.io)
3. 登录 GitHub 账号，选择该仓库
4. 设置：
   - **Main file path**: `warehouse_analyzer.py`
   - **Python version**: 3.11
5. 点击 Deploy，等待几分钟即可获得在线访问链接

### 方式 B：Docker 部署

```bash
# 构建镜像
docker build -t warehouse-analyzer .

# 运行容器
docker run -p 8501:8501 warehouse-analyzer
```

访问 `http://localhost:8501`

### 方式 C：云服务器部署

```bash
# 在服务器上
git clone <仓库地址>
cd warehouse_tool
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 后台运行
nohup streamlit run warehouse_analyzer.py --server.port=8501 --server.address=0.0.0.0 &
```

## ⚙️ 配置说明

编辑 `config.json`，将标准字段映射到你表格的实际列名：

```json
{
  "入库": {
    "入库时间": "你的到货日期列名",
    "入库箱数": "你的入库箱数列名"
  },
  "出库": {
    "出库时间": "你的拣货时间列名",
    "出库数量": "你的出库箱数列名"
  }
}
```

**字段映射优先级：** config.json 精确匹配 > 去空格归一化匹配 > 关键词自动识别 > 默认第一/二列

### 大模型配置（可选）

在 `config.json` 的 `"大模型"` 中配置模型、接口地址、超时时间等。推荐通过环境变量 `OPENAI_API_KEY` 提供密钥，也可以直接在 `"密钥"` 字段填写。未配置密钥时，页面会显示规则版解读。

## 📋 使用流程

1. **填写 config.json**：把表格实际列名填到对应字段
2. **上传数据**：在界面选择入库 Excel、出库 Excel（可选物料主数据）
3. **选择 Sheet**：多 Sheet 文件选择对应入库/出库 Sheet
4. **确认列**：侧边栏可手动调整时间列、数量列
5. **筛选分析**：按时间范围、物件、单据种类筛选
6. **查看结果**：概览、频次、批次分布、在库时间、规划建议、AI 智能解读

## ⚠️ 常见问题

1. **必须用虚拟环境**：避免依赖冲突
2. **数量列坑**：若表里有多个含"数量"的列，务必在 config.json 指定确切列名
3. **中文乱码**：bat 文件避免 `chcp 65001` + 中文
4. **大文件**：超过 50 万行加载较慢，FIFO 循环耗时较长

## 📁 项目结构

```
warehouse_tool/
├── warehouse_analyzer.py   # 主程序
├── config.json             # 字段映射配置
├── requirements.txt        # Python 依赖
├── launch.bat              # Windows 一键启动
├── Dockerfile              # Docker 部署
├── .streamlit/
│   └── config.toml         # Streamlit 配置
├── .gitignore
├── .dockerignore
└── README.md
```

## 🛠 技术栈

- **Streamlit** - Web 框架
- **Pandas** - 数据处理
- **Plotly** - 交互式图表
- **openpyxl** - Excel 读取

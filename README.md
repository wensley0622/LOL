# Hacker News Open Data Pipeline

这是一个用于结业作业的开源仓库：选择一个非公众号公开数据源，使用自动化脚本采集不少于 10 个独立文件，对数据进行清洗和主体名称脱敏，并为每个清洗文件增加 YAML Front Matter 元信息。

## 选择该数据源的理由及简介

本项目选择 Hacker News Algolia Search API 作为目标数据源。

选择理由：

- 它不是微信公众号数据源，符合课程作业对数据来源的限制。
- 接口公开、稳定、无需登录，适合用课程中的自动化采集方法复现。
- 数据结构清晰，返回 JSON，便于展示采集、字段筛选、清洗和脱敏流程。
- 每条 story 都能独立保存为一个文件，方便满足“不少于 10 个独立文件”的要求。
- 数据中包含作者、标题、URL、标签、评论数等字段，适合演示主体名称和身份标识的脱敏。

数据源简介：

- 数据源名称：Hacker News Algolia Search API
- 数据类型：公开新闻/讨论元数据
- 接口地址：`https://hn.algolia.com/api/v1/search_by_date?tags=story`
- 本仓库采集对象：按时间倒序返回的最新 story 条目

## 仓库结构

```text
.
├── data/
│   ├── raw/      # 原始采集文件，每条 story 一个 JSON
│   └── clean/    # 清洗脱敏后且带 YAML Front Matter 的 Markdown 文件
├── scripts/
│   ├── collect_hn.py
│   └── clean_anonymize.py
├── tests/
│   └── test_clean_anonymize.py
├── LICENSE
└── README.md
```

## 数据采集方法的详细说明

采集脚本位于 `scripts/collect_hn.py`，只使用 Python 标准库完成 HTTP 请求和文件写入。

采集流程：

1. 使用 `urllib.request` 请求 Hacker News Algolia API。
2. 请求参数设置为 `tags=story` 和 `hitsPerPage=10`，获取最新 10 条 story。
3. 将接口返回的 `hits` 数组逐条拆分。
4. 每条 story 单独保存为一个 JSON 文件，文件名格式为 `序号_storyID.json`。
5. 每个原始文件中额外记录 `source`、`source_url` 和 `collected_at`，便于追踪数据来源和采集时间。

复现采集：

```bash
python3 scripts/collect_hn.py --limit 10 --output-dir data/raw
```

## 数据清洗方法的详细说明

清洗脚本位于 `scripts/clean_anonymize.py`，输入目录为 `data/raw/`，输出目录为 `data/clean/`。

清洗流程：

1. 读取每个原始 JSON 文件。
2. 删除 `_highlightResult` 等接口展示字段，只保留分析所需字段。
3. 保留 story 的基本统计字段，如 `points`、`num_comments`、`created_at`、`updated_at`。
4. 对 HTML 实体进行解码，例如将 `&#x27;` 还原为普通字符。
5. 移除正文中的 HTML 标签，并将多余空白合并为单个空格。
6. 将清洗后的数据写入 Markdown 文件，正文中用 JSON 代码块保留结构化数据。
7. 每个 Markdown 文件顶部加入 YAML Front Matter，记录来源、采集时间、记录 ID、脱敏状态等元信息。

复现清洗：

```bash
python3 scripts/clean_anonymize.py --input-dir data/raw --output-dir data/clean
```

## 脱敏处理方法的详细说明

脱敏目标是移除或替换可能指向具体主体的名称和身份标识。

本项目采用以下方法：

- 作者名称：对 `author` 字段使用加盐 SHA-256 哈希，输出为 `user_xxxxxxxxxxxx` 格式。
- 作者标签：删除 `_tags` 中形如 `author_xxx` 的标签，避免从标签反推出作者名。
- URL 主体：不保留原始 URL，也不保留明文域名；对 URL 域名进行加盐哈希，输出为 `url_domain_hash`。
- 邮箱地址：在标题和正文中替换为 `[EMAIL]`。
- 电话号码：在标题和正文中替换为 `[PHONE]`。
- 完整链接：在标题和正文中替换为 `[URL]`。
- 主体名称：使用规则匹配英文大写开头的主体名称片段，并替换为 `[ENTITY]`，降低标题和正文中组织、产品、项目或人物名称泄露的风险。
- Front Matter：每个清洗文件包含 `contains_raw_subject_names: false`，标记该文件已经过主体名称脱敏处理。

默认脱敏盐值写在脚本中，仅用于课程演示。实际生产环境应通过环境变量或密钥管理系统传入盐值。

## 快速开始

本项目只使用 Python 标准库，建议 Python 3.10+。

```bash
python3 scripts/collect_hn.py --limit 10 --output-dir data/raw
python3 scripts/clean_anonymize.py --input-dir data/raw --output-dir data/clean
```

运行测试：

```bash
python3 -m unittest discover -s tests
```

## 数据说明

`data/raw/` 中包含 10 个自动采集的原始 JSON 文件；`data/clean/` 中包含对应 10 个清洗脱敏后的 Markdown 文件。每个清洗文件都带有 YAML Front Matter，适合做课程展示、字段统计和简单文本分析。

## 复现步骤

1. 克隆或下载本仓库。
2. 运行 `python3 scripts/collect_hn.py --limit 10 --output-dir data/raw` 重新采集 10 条公开 story。
3. 运行 `python3 scripts/clean_anonymize.py --input-dir data/raw --output-dir data/clean` 生成脱敏数据。
4. 检查 `data/clean/` 中是否生成不少于 10 个 `.md` 文件，且每个文件以 YAML Front Matter 开头。

## 合规说明

本项目仅采集公开 API 返回的元数据，用于课程作业演示。仓库中的脱敏样例不保留明文作者名、明文 URL 域名和识别出的主体名称；如需扩展采集范围，应遵守 Hacker News 与 Algolia 的服务条款，并控制请求频率。

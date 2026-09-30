# deliver

面向嵌入式项目的可配置交付自动化工具。根据 Redmine 交付清单，从 GitLab 和钉钉收集代码、设计资料、报告及变更申请，生成交付目录、Excel 清单、ZIP 和 HTML 报告。

## 交付流程

- Redmine：按日期、项目、状态和模块读取交付项，支持常规交付和变更交付。
- GitLab：读取配置分组中的仓库，支持精确匹配、别名回退及可选模糊匹配；检查提交号，拉取指定分支或全部分支。
- 文件整理：架构设计、HAL、Example、Case、Firmware、FTL、引脚功能表、XML、界面表与公共资料。
- 校验：交付物格式、缺失文件、MD5、MCU 系列与选型；比较参考提交，提取引脚变更。
- 报告处理：下载评审表、反馈单和测试报告，根据问题单回填整改信息，匹配文字或图片，保留工作簿中的图片。
- 钉钉：按配置的空间和目录查找变更申请、功能分析及需求资料，选择版本并下载。
- 交付产物：区分 AE / SW，生成 Excel 清单、ZIP、缺失项文本和 HTML 报告；支持 Jenkins 归档和可选 AES256 ZIP 加密。

这是完整交付流程的配置化版本，仍沿用嵌入式资料目录及中文 Excel 模板约定。服务连接、字段名、项目映射和凭据可以替换；其他目录或表格格式需要调整对应处理逻辑。

## 文件

| 文件 | 用途 |
| --- | --- |
| `deliver.py` | 交付调度、仓库拉取、文件整理、校验和 Excel 清单 |
| `delivery_report_packager.py` | Redmine 报告下载、Excel 回填和 HTML 报告 |
| `download_file_from_dingtalk.py` | 钉钉文件查找、版本选择与下载 |
| `package_delivery.py` | 最终文件清理、换行转换及 ZIP 归档 |
| `config.example.json` | 服务、字段、仓库和命名规则配置示例 |
| `mcu_selection_mapping.example.ini` | MCU 选型、工程别名和仓库回退示例 |
| `Jenkinsfile` | Linux Jenkins 流水线 |

## 准备

需要 Python 3.10+ 和 Git。加密 ZIP 另外需要 `7z` 或 `7zz` 位于 PATH；未设置密码时使用 Python 标准 ZIP。

```powershell
python -m pip install -r requirements.txt
Copy-Item config.example.json delivery.local.json
Copy-Item mcu_selection_mapping.example.ini mcu_selection_mapping.local.ini
```

修改 `delivery.local.json` 的地址、项目、字段和仓库映射，将 `selection_mapping` 改为 `mcu_selection_mapping.local.ini`。示例中的 `example.com`、`MCU_DEMO` 均为占位数据。

凭据通过环境变量传入，不写入配置文件：

```powershell
$env:GITLAB_TOKEN = '<GitLab API token>'
$env:REDMINE_API_KEY = '<Redmine API key>'
# 启用钉钉时填写：
$env:DINGTALK_CLIENT_ID = '<appKey>'
$env:DINGTALK_CLIENT_SECRET = '<appSecret>'
$env:DINGTALK_UNION_ID = '<unionId>'
# 可选：加密 ZIP
$env:ZIP_PASSWORD = '<zip password>'
```

`GITLAB_URL`、`REDMINE_URL` 可覆盖 JSON 地址。API token 用于接口访问；Git 克隆需要单独配置 Git 凭据管理器，或将 `gitlab.clone_protocol` 设为 `ssh` 并配置 SSH 密钥。

## 配置约定

- `gitlab.groups`：左边是脚本分组别名，右边是实际 GitLab 分组完整路径或 ID。`arch/{series}` 对应设计仓库分组，`arch/{series}/common` 对应公共文件，`hal` 对应 HAL、Case 和 Example 仓库。
- `repositories.series_aliases`：把 Redmine 项目名映射到设计分组中的系列名；缺省使用项目名的小写形式。
- `repositories.hal/case/example`：按项目名配置仓库名。`example` 中配置的仓库从 HAL 分组读取，未配置时在设计分组中查找名称包含 `example` 的仓库。
- `repositories.libraries/library_roots`：声明中间件交付项对应的文件夹和 HAL 仓库内根路径。
- `naming.series_prefix`：从项目名去掉此前缀后生成 Example/Case 交付项。`naming.hal_pattern` 必须有模块名、`C/H` 两个捕获组。
- `redmine.custom_fields`：字段名精确匹配，括号和空格也要与实际配置一致。
- `module_list`：直接配置模块黑白名单；也可以通过 `issue_id` 读取一个 Redmine 问题单描述中的 `[blacklist]` / `[whitelist]` 列表。
- `dingtalk.enabled`：默认关闭。启用时填写 `space_id` 和 `folder_ids`，后者指向包含系列目录的父目录。特殊目录使用 `series_paths`，如 `{"MCU_DEMO": ["/阶段一/设计资料"]}`；系列名不同可用 `series_aliases` 映射。
- `selection_mapping`：相对路径以 JSON 所在目录为基准。

## Redmine 交付清单格式

在 `redmine.project_id` 对应项目及可查询子项目中创建标题为 `delivery_subject` 的问题单，描述前两条非空行分别填写当天日期和交付模块：

```text
2026-09-30
GPIO,UART,EXAMPLE,FIRMWARE,架构设计
```

`EXAMPLE`、`FIRMWARE`、`架构设计` 表示系列级交付；其他名称与问题单分类（模块）匹配。模块问题单的状态应为配置中的 `status`，自定义字段 `deliverables` 指向的“交付物清单”字段可填写：

```text
架构设计,架构设计@release,HAL_HUART_C,HAL_HUART_H,DEMO_EUART,LIB_USB,系统测试报告
```

支持 `架构设计@all`、变更申请单、评审表、反馈单、功能分析表等原有交付项。提交号填写到对应自定义字段；选型、目录布局和 Excel 表头须符合脚本的处理约定。

## 执行

```powershell
# 只检查配置和凭据是否齐全，不连接远程服务
python deliver.py --config delivery.local.json --check-config

# 输出目录必须为空；可用 --series 限定项目
New-Item -ItemType Directory -Force build | Out-Null
python -X utf8 deliver.py --config delivery.local.json --output build/workspace --series MCU_DEMO *> build/deliver_output.txt

# 即使有缺失项，也可以归档已生成的交付物并查看报告
python package_delivery.py --source build/workspace --output build/package --log build/deliver_output.txt
```

PowerShell 7 的重定向日志使用 UTF-8；Windows PowerShell 5.1 请改用 `python -X utf8 deliver.py ... 2>&1 | Out-File -Encoding utf8 build/deliver_output.txt`，并单独检查 Python 的退出码。

主流程有缺失项时返回非零。归档在副本中清理隐藏文件和普通 Markdown（保留 FreeRTOS/LWIP 中的 Markdown），将 `.c/.h/.xml` 换成 CRLF，原始收集目录可保留排查。`Common Files` 目录单独复制，其余交付目录压缩；缺失项或错误日志会使归档返回非零。每次交付应使用新的工作目录和归档目录。

默认生成：

```text
build/
  workspace/              # 收集的仓库、交付目录和 Excel 清单
  package/                # ZIP、Excel、Common Files 和 missing_file.txt（UTF-8）
  html_report/index.html   # 产物、缺失项和日志摘要
  deliver_output.txt
```

单独生成 HTML 报告：

```bash
python delivery_report_packager.py html-report --log build/deliver_output.txt --package-dir build/package --output build/html_report/index.html
```

## Jenkins

将本仓库配置为 Pipeline from SCM，在具有 Python、Git 的 Linux agent 上执行。需要 Git、Credentials Binding、HTML Publisher 插件。流水线使用 Jenkins 的 [Git 用户名/密码绑定](https://www.jenkins.io/blog/2021/07/27/git-credentials-binding-phase-1/)完成私有仓库克隆，因此配置中使用 `clone_protocol: "https"`。

参数中填写各凭据 ID：JSON 配置和 MCU 映射使用 Secret file，API key/token 使用 Secret text，Git 使用 Username with password。JSON 的 `selection_mapping` 应填写 `mcu_selection_mapping.local.ini`。启用钉钉时再填写对应凭据 ID；需要加密时再配置 ZIP 密码和 7-Zip。

流水线按构建号保存产物。运行结果包含错误时仍会尝试归档并发布报告。`*.local.json`、`*.local.ini`、环境文件、构建产物和本地验证文件均已加入 `.gitignore`。

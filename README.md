# audit-flow

**审计底稿明细表/审定表自动刷数工具 | The audit workpaper auto-fill terminator**

把科目余额表（TB）和序时账（GL）导入底稿、填明细、接审定表公式，再逐项核对金额，是审计中反复发生的手工工作。audit-flow 把这段流程交给可复用的本地 Python 工具。明细表按数据行数自动扩展，审定表保留 Excel 公式，并附勾稽报告和准则异常建议。首个随附模板覆盖应收账款。

固定核心引擎负责标准字段契约、模板、Excel 重算与审计勾稽。各客户通过 Adapter 映射不同账套的导出格式。运行无需启用 VBA 宏。

## 本次更新

第二轮代码审查整改（2026-11-25）

- `--cas-routed` 可按余额表科目选择 CAS 规则，常规运行继续使用通用筛查。凭证完整性、损益结转和外币结转由数据持有人逐项确认。可选的 `--account-names` 接收已核实的科目名称映射，并校验 JSON 格式和科目编码。
- 准则工作表新增「待核规则」区块，列出因缺少确认、配对或一级汇总行而延后的规则编号。命中数量和四道勾稽门禁独立统计。
- 补充 CAS 实务卡片、专题参考及本地原文的命名说明。新增 4 项 CLI 回归测试，全套共 36 项。使用合成数据在本机 Excel 验证了默认筛查、按科目筛查和已确认上下文的按科目筛查。

## 数据安全与使用边界

| 环节 | 处理方式与边界 |
| --- | --- |
| 原始账表 | 真实 TB、GL 和真实表头清单放在本地 `data/raw/`。原始表由本地 Python 处理。Agent 只可调用脱敏探针或已审核的执行入口，严禁直接打开、预览、全文读取、截图、上传或粘贴原始账表。 |
| 脱敏画像 | 探针读取表头和随后最多十行。摘要抹除，客商替换为占位符，金额以保留格式的合成值替换。通用或经批准的表头、借贷方向和格式统计会保留。共享 JSON 前由数据持有人确认这些信息可公开。 |
| 本地底稿 | `data/output/` 中的工作簿及运行暂存文件包含标准化 TB、GL、摘要、客商或真实金额，应按原始账表的保密级别管理。真实账套的控制台默认仅输出门禁状态。仅限数据持有人在本地使用 `--local-amounts`，完整输出不得进入 Agent 上下文或公共日志。 |
| Git 与发布 | `.gitignore` 排除 `data/raw/` 和 `data/output/` 的内容，目录占位文件除外。仅提交代码、通用模板和合成样例。上传、强制添加文件、复制到其他目录以及已经写入 Git 历史的内容均需要单独管控。发布前核对暂存清单。 |
| 适配器和运行环境 | 只运行已审核的 Python Adapter 和可信模板。模板会检查外部关系及活动内容，重算在本地 Excel 或 LibreOffice 上执行。文件权限、磁盘加密、同步、Office 插件及网络访问由使用方管理。 |

## 使用 Skill 前

1. 从合成数据开始。为真实账套建立单独的本地目录。样例、测试和公开问题单均用合成材料。
2. 按 [SKILL.md](SKILL.md) 先跑本地脱敏探针，再用可公开画像生成 Adapter。列名从探针输出逐字复制。字段含义、期间、币种口径和层级规则由数据持有人确认。
3. 保持勾稽门禁和金额容差 `0.00`。用 Python 或 Excel 计算财务金额。失败时核对来源与 Adapter，并保留失败证据。
4. 由审计人员复核数据口径、未覆盖科目和准则建议。四道门禁通过仅说明本次自动勾稽通过。CAS 建议不写回未审数或 AJE，最终调整及审计结论由复核人确认。

## 快速运行

需要 Python 3.10+。Windows 可用 Microsoft Excel 和 xlwings，也可安装 LibreOffice 并使 `soffice` 可用。

~~~sh
python -m venv .venv
# 激活虚拟环境后执行
python -m pip install -r requirements.txt
python core/masker.py --input data/mock/demo_tb.xlsx
python core/masker.py --input data/mock/demo_gl.csv
python run.py --adapter demo_kingdee --tb data/mock/demo_tb.xlsx --gl data/mock/demo_gl.csv
~~~

Windows 也可用 `.\.venv\Scripts\python.exe` 代替 `python`。仓库随附模板和模拟数据，`python -m scripts.generate_demo` 可重建这些文件。真实账表按项目放在 `data/raw/`，同次运行的 TB 和 GL 应来自同一个目录、同一账套、期间和货币单位。

- 探针支持 `.xlsx`、`.xlsm`、`.xltx`、`.xltm`、`.xls`、`.xlsb`、`.ods` 和 `.csv`。工作表号与表头行号从 1 开始。CSV 默认 UTF-8，其他编码或分隔符用 `--encoding`、`--delimiter`。探针只采样表头和随后最多十行。旧版或二进制格式使用文件中已保存的单元格值。
- 探针原样公开通用表头和数据持有人批准的表头。其他表头返回 `null`。`--approved-headers` 接收 `data/raw/` 或 `data/mock/` 下经本地审核的 JSON 字符串数组，或列名到标准字段的映射。后者可为非标准摘要或客商列指定 `summary` 或 `aux_name`。摘要抹除，客商以本次画像内稳定的占位符替换，金额保持符号、括号、千分位习惯和小数位数。
- `--ar-account` 指定应收一级科目编码，默认采用 Adapter 设置。`--recalc` 取 `auto`、`xlwings` 或 `libreoffice`。`--out` 接收 `data/output/` 下的 ASCII 文件名，省略时生成唯一名称。
- 退出码 0 表示 `PASS`。退出码 1 表示勾稽或重算 `FAILED`，生成 `*.review.xlsx`。退出码 2 表示参数、输入、契约或执行异常。
- 模拟数据的终端输出可用于验收。真实数据默认只输出门禁状态，详细金额留在本地报告。数据持有人可在本地添加 `--local-amounts` 查看完整终端证据。

## 数据口径和 Excel 模板

| 契约 | 字段 |
| --- | --- |
| Standard_TB | `account_code`、`account_name`、`level`、`direction`、`open_debit`、`open_credit`、`period_debit`、`period_credit`、`close_debit`、`close_credit`、`aux_name`（可选） |
| Standard_GL | `date`、`voucher_no`、`summary`、`account_code`、`debit`、`credit`、`aux_name` |

金额采用 Decimal 两位小数。示例 Adapter 支持千分位和括号负数，GL 空白借贷侧按导出口径转为零，TB 六列须显式提供。科目、凭证号以文本录入以保留前导零。示例接受同一账套、会计期间和货币单位内带一级与明细层级的 TB。科目层级依据编码前缀识别。目标一级科目需要一条无辅助核算的余额行。明细优先采用末级科目与客商行，源表层级汇总行全部保留在 `Raw_TB`，以源文件同一粒度核对清洗守恒。

审定表 `SUMIFS` 预置于模板，生成时保持公式文本原样，按一级科目从 `Raw_TB` 取值并从 `AJE` 汇总调整。明细从 `{{DATA_START}}` 插行，复制模板样式，写入该行期初、本期借贷和期末滚转公式。增加 AJE 调整金额后，Excel 动态更新审定数。底稿的勾稽报告为当次运行记录。

| 门禁 | 验收内容 |
| --- | --- |
| 1. 清洗前后守恒 | 从原始 TB 独立读取六列借贷金额，与标准化合计及行数比对，金额差异必须为 0.00 |
| 2. 未映射兜底 | 原始范围内的有效记录全部写入底层页，未纳入应收底稿的科目逐一列出告警 |
| 3. 跨表勾稽 | 审定表未审数、明细期末合计、Raw_TB 对应一级科目期末净额须一致 |
| 4. 公式健康度 | 检查预期公式存在、公式缓存存在且无 Excel 错误 |

新增客户见 `SKILL.md`。运行 `python -m unittest discover -s tests` 可验证脱敏、严格差额、未映射兜底、模板公式保留和缓存检查。

## CAS 准则异常筛查

四道勾稽结束后，`core/cas_checker.py` 默认使用 `rules/cas_general.yaml` 离线扫描标准 TB，在第二张工作表 `01_准则异常与调整建议` 列出异常、依据及待复核的 RJE 或 AJE。审阅重分类时核对合同和往来明细，审阅损益余额时先核对结转时点。四道门禁独立判定，分录由复核人确认后另行录入。

YAML 中的编码前缀采用一般企业常见科目编码。坏账准备科目可能覆盖多类应收资产。经核对后，在客户 Adapter 中设置 `cas_allowance_pairs = {"准备科目编码": ("应收原值科目编码",)}`。缺少配对或原值时，报告会提示补充资料。对经确认的配对，程序按对应范围内应收明细的借方净额合计筛查准备上限，调整金额由复核人根据 ECL 重算确定。

数据持有人可在运行时添加 `--cas-routed`，按余额表科目选择 `rules/cas_router.yaml` 匹配的规则模块。按科目筛查的下列参数须按本次数据逐项确认。

- `--complete-vouchers` 确认凭证已核对完整。
- `--period-closed` 确认相关损益结转已完成。
- `--currency-closed` 确认外币折算和汇兑结转已完成，且 TB 已统一为记账本位币。
- `--account-names FILE` 接收已核实的科目编码与名称的 UTF-8 JSON 对象，例如 `{"1122": "应收账款"}`。真实账套的映射保留在本地。

三个上下文确认默认关闭。规则因缺少确认、配对或汇总行而延后时，工作表底部的「待核规则」区块列出规则编号，命中项数量单独计算。

### 准则来源与致谢

> 本项目的 CAS 实务卡片和筛查规则参考毛燕庆老师整理的**审计文库（MaoDocs）**所收录的《企业会计准则》及《企业会计准则应用指南汇编2024》。感谢毛燕庆老师对准则资料的整理与维护。 [企业会计准则](https://docs.maoyanqing.com/accounting/ent/cas/) · [2024 应用指南](https://docs.maoyanqing.com/accounting/ent/casg/)

## 开源许可

项目采用 [MIT 许可证](LICENSE)。客户输入、底稿、真实映射和运行证据由数据持有人按其保密要求管理。

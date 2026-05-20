# 2026-05-20 工作日志：审计底稿 Roll Forward 工具 - CRA等级表动态填写

## 工作内容

实现CRA（控制风险评估）等级表的动态读取和填写功能。

## 修改内容

### 1. `roll_forward_core.py`

#### 新增函数：`load_cra_data(cra_path, subject_code)`
- 从CRA等级表（`输入；参考资料.xlsx` 的 "CRA" Sheet）读取指定科目的认定等级和比例
- 解析逻辑：
  - 查找以 `subject_code.` 开头的行（如 `C.`）
  - 提取认定名称（存在性、完整性、计价、权利义务、列报）
  - 读取C3列判断是否适用（X=不适用）
  - 读取C6列获取风险等级（低→Low，中→Moderate，高→High等）
  - 读取C9列获取比例值
- **重要修复**：同时处理全角括号 `（` 和半角括号 `(`，避免中文Excel中的括号导致解析失败

#### 修改 `process_single_subject` 函数
- 新增 `cra_path` 可选参数
- 在处理流程中调用 `load_cra_data` 加载CRA数据
- 将CRA数据存入 `company_info["cra_data"]` 传递给 `process_lead_sheet`

#### 修改 `process_lead_sheet` 函数
- **CRA等级和比例自动填写**：
  - 遍历CRA区域（R15-R19）
  - 根据模板B列的认定名称匹配CRA数据中的认定
  - 将等级填入C列（如 Low, Moderate, High）
  - 比较CRA等级表中的比例与底稿公式默认比例：
    - Minimal → 100%
    - Low → 75%
    - Moderate → 50%
    - High → 25%
  - **如果用户CRA等级表中的比例与默认不同，覆盖D列的公式，写入计算后的实际值**

## 测试结果

以科目 C（货币资金）测试：

| 行号 | 认定 | 原始C列 | 处理后C列 | 原始D列 | 处理后D列 | 原因 |
|------|------|--------|---------|--------|----------|------|
| R15 | 存在性 | 空 | 空 | 公式 | 公式 | CRA等级表中标记为X（不适用） |
| R16 | 完整性 | 空 | **Low** | 公式 | **1** | ✅ 从CRA等级表读取并写入 |
| R17 | 计价 | 空 | **Low** | 公式 | **1** | ✅ 从CRA等级表读取并写入 |
| R18 | 权利义务 | 空 | **Low** | 公式 | **1** | ✅ 从CRA等级表读取并写入 |
| R19 | 列报 | 空 | 空 | 公式 | 公式 | CRA等级表中标记为X（不适用） |

## 使用方式

在调用 `process_single_subject` 时，传入 `cra_path` 参数：

```python
success, message, output_path, warnings = process_single_subject(
    "C",  # 科目代码
    template_path,
    prior_path,
    pmte_path,
    "公司名称",
    "2026-12-31",  # 日期
    output_dir,
    subject_config,
    cra_path=r"path/to/输入；参考资料.xlsx"  # 新增参数
)
```

如果不传 `cra_path`，会尝试从 `pmte_path` 所在的文件中读取CRA数据。

## 已知问题

1. PMTE信息表路径（`pmte_test.xlsx`）目前需要单独指定，建议与CRA等级表统一路径
2. 如果CRA等级表中某科目不存在，会静默跳过，不影响其他功能

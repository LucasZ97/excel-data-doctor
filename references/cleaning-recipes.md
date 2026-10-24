# 脏数据清洗菜谱

执行清洗时按以下标准写法处理。安全修复优先直接用 `scripts/clean.py`（确定性脚本，比手写 pandas 更可靠）；下面的菜谱用于处理"需确认修复"和脚本覆盖不到的 edge case。所有操作作用于副本 DataFrame，最终写入 `<原文件名>_cleaned.xlsx`。

## 避坑（实测沉淀）

1. **openpyxl 只读模式读不到合并单元格**：`load_workbook(read_only=True)` 的 worksheet 没有 `merged_cells` 属性。诊断合并单元格必须用普通模式加载。
2. **合并单元格经 pandas 读取后变成 NaN**：只有左上角有值，其余为 NaN。如需保留合并语义，用 `df[col].ffill()` 向前填充，但必须先让用户确认。
3. **中文日期 pandas 直接解析不了**："2024年1月7日" 必须先把 年/月/日 替换成 `-` 再 `pd.to_datetime(..., format='mixed')`。
4. **中文 CSV 九成是 GBK 编码**：直接 `pd.read_csv` 默认 utf-8 会报 UnicodeDecodeError。按 utf-8-sig → gbk → gb18030 顺序尝试。
5. **金额列的"看起来像数字"陷阱**："1,200元" 剥离符号后能转数字，可以安全清洗；"abc"剥离后仍转不了，必须保留原值并标出，绝不静默丢弃或填 0。
6. **全角空格 U+3000**：不在 `[！-～]` 范围内，正则要单独加上；Python 的 `str.strip()` 能去掉它，可用于检测。
7. **近似重复行**：`df.duplicated()` 抓不到 "张三" vs " 张三 "。先 strip+lower 归一化后再判重，报"疑似重复行"让用户确认。
8. **`data_only=True` 的 workbook 不要 save**：公式会变成字面量，永久丢失。本技能不涉及公式，但处理他人文件时要记住这条。

## 1. 基础规范化（安全修复，默认执行）

```python
# 去首尾空格（字符串列）
for c in df.select_dtypes(include='object'):
    df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 全角转半角
def full_to_half(s):
    if not isinstance(s, str): return s
    return ''.join(chr(ord(ch) - 0xFEE0) if '！' <= ch <= '～' else ch for ch in s).replace('　', ' ')
for c in df.select_dtypes(include='object'):
    df[c] = df[c].apply(full_to_half)

# 删除全空行、全空列
df = df.dropna(how='all').dropna(axis=1, how='all')

# 规范列名：去空格、去换行
df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
```

## 2. 金额/数字列去杂质（安全修复）

```python
# 去掉"元/￥/¥/$/,"后转数值；百分号单独处理（除以100）
def clean_number(s):
    if pd.isna(s): return s
    t = str(s).strip().replace(',', '').replace('￥', '').replace('¥', '').replace('$', '').replace('元', '')
    if t.endswith('%'):
        try: return float(t[:-1]) / 100
        except: return s
    try: return float(t)
    except: return s

df[col] = df[col].apply(clean_number)
# 转成功后统一保留两位小数
```

注意：清洗前后各记录一列的 `describe()`，数值分布发生剧烈变化时要告警（可能是单位混杂，如"万"和"元"混用）。

## 3. 日期列统一（安全修复）

```python
df[col] = pd.to_datetime(df[col], errors='coerce', format='mixed')
fail = df[col].isna().sum()
# fail>0 时把解析失败的原值单独列出给用户确认，不要静默丢弃
df[col] = df[col].dt.strftime('%Y-%m-%d')
```

## 4. 重复行（需用户确认）

先展示重复行的内容和数量，用户确认后再 `df.drop_duplicates()`。注意：业务上"看起来重复"的行可能是有效重复（如同一客户多笔订单），必须让用户拍板。

## 5. 空值处理（需用户确认）

默认策略按列类型建议：
- 数值列 → 填 0 或列均值（注明选择了哪种）
- 文本列 → 填"未知"或保留空
- 日期列 → 保留空，不编造日期

执行前向用户说明每列的空值数和拟采用的策略。

## 6. 合并单元格（需用户确认）

```python
# openpyxl 读取时向前填充
# 或 pandas: df[col] = df[col].ffill()
```

先说明合并区域的位置和填充后的效果，用户确认后执行。

## 7. 只报告、不修的项

- 同一列出现两种单位（"万元"和"元"混用）：列出实例，让用户定口径。
- 数值异常值（用 IQR 法识别，超出 1.5 倍四分位距）：列出 top 10，不擅自删除。
- 疑似多行表头：给出重构表头的建议方案，让用户确认列名映射。

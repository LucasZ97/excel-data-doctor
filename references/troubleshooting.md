# 排障手册

## 依赖问题

**现象**：`ModuleNotFoundError: No module named 'pandas'`
**处理**：执行 `pip install pandas openpyxl` 后重试。不要预先安装——先 import，失败再装。

**现象**：读取 `.xls` 老格式报错 `Missing optional dependency 'xlrd'`
**处理**：`pip install xlrd`。或者让用户把文件另存为 .xlsx。

## 文件问题

**现象**：`PermissionError` / 文件被占用
**处理**：让用户关闭正在打开该文件的 Excel/WPS 后重试。不要强行写入。

**现象**：xlsx 有打开密码
**处理**：本技能无法处理密码保护的文件。请用户先去除密码，或提供密码后用 `msoffcrypto-tool` 解密（`pip install msoffcrypto-tool`），解密后再走正常流程。

**现象**：CSV 打开是乱码
**处理**：脚本已自动尝试 utf-8-sig → gbk → gb18030。仍乱码的可能是 UTF-16 或 Big5，手动用 `pd.read_csv(path, encoding='utf-16')` 等尝试，并在报告中注明实际使用的编码。
**边界**：极少数 GBK 文件的字节恰好也是合法 UTF-8（短文件、无生僻字时），脚本会按 utf-8 解出形似亚美尼亚文的乱码且不报错。判别方法：列名出现非中文字符时，改用 `encoding='gbk'` 重读一次。

## 规模问题

**现象**：文件超过 50 万行，诊断很慢
**处理**：用 `--sample 5000` 先抽样体检，给出"抽样体检"的结论并注明局限。如用户坚持全量，分 sheet 逐个处理。

**现象**：内存不足（MemoryError）
**处理**：改用 `pd.read_excel(..., chunksize=...)` 分块，或先让用户拆分文件。

## 结果问题

**现象**：清洗后健康评分没有提升
**处理**：按 SKILL.md 要求回滚检查——很可能是某列的"杂质"判断阈值误伤，或日期列被误判。对比清洗前后的列级问题清单，定位是哪一类修复没生效。

**现象**：用户说"清洗结果不对"
**处理**：原文件未动，直接从原文件重新跑。先问清是哪一列、期望是什么，再针对性修，不要全盘重来。

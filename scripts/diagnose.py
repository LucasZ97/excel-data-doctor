#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格体检医生 - 自动化诊断脚本
扫描 Excel/CSV 的数据质量问题，输出 JSON 体检报告。

用法:
    python3 diagnose.py <文件路径> [--sheet 表名] [--out 报告.json] [--sample N]

CSV 编码自动识别 utf-8-sig -> gbk -> gb18030，无需手动转码。
退出码: 0=正常, 1=文件读取失败, 2=缺少依赖。

依赖: pandas, openpyxl
"""
import argparse
import json
import re
import sys
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("缺少依赖 pandas，请先执行: pip install pandas openpyxl", file=sys.stderr)
    sys.exit(2)

FULLWIDTH_RE = re.compile(r'[！-～　]')
CURRENCY_CHARS = ['元', '￥', '¥', '$', '%', ',']
DATE_HINT_RE = re.compile(r'日期|时间|date|time', re.I)
ID_HINT_RE = re.compile(r'编号|编码|单号|\bid\b', re.I)


def normalize_date_str(s: str) -> str:
    """把中文日期 2024年1月7日 -> 2024-1-7，便于解析"""
    t = str(s).strip()
    t = re.sub(r'(\d)\s*年\s*', r'\1-', t)
    t = re.sub(r'(\d)\s*月\s*', r'\1-', t)
    t = re.sub(r'(\d)\s*日\s*', r'\1', t)
    return t.strip('- ')

PENALTY = {
    'null_rate': 25,        # 空值率惩罚系数
    'dup_rows': 15,
    'mixed_type': 10,
    'date_fail': 12,
    'dirty_numeric': 12,
    'whitespace': 5,
    'fullwidth': 5,
    'merged_cells': 8,
    'multi_header': 8,
    'constant_col': 4,
}


def is_fullwidth(s: str) -> bool:
    return bool(FULLWIDTH_RE.search(s))


def looks_like_date(values) -> bool:
    """抽样判断一列是否疑似日期列"""
    sample = [v for v in values if v not in (None, '')][:20]
    if len(sample) < 3:
        return False
    hit = 0
    for v in sample:
        s = str(v)
        if re.search(r'\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}', s):
            hit += 1
    return hit >= max(2, len(sample) * 0.5)


def diagnose_column(name, series: pd.Series) -> dict:
    s = series.copy()
    n = len(s)
    col = {'column': str(name), 'dtype': str(s.dtype), 'total_rows': n, 'issues': []}

    null_mask = s.isna() | (s.astype(str).str.strip() == '')
    null_count = int(null_mask.sum())
    null_rate = round(null_count / n, 4) if n else 0
    col['null_count'] = null_count
    col['null_rate'] = null_rate
    # 空值分级：少量(low) vs 偏多(medium)
    if 0 < null_count < 3 and null_rate <= 0.2:
        col['issues'].append({'type': '少量空值', 'detail': f'空值 {null_count} 个（{null_rate:.1%}）', 'severity': 'low'})
    elif null_rate > 0.2 or null_count >= 3:
        col['issues'].append({'type': '空值偏多', 'detail': f'空值 {null_count} 个（{null_rate:.1%}）', 'severity': 'medium'})

    vals = s[~null_mask].astype(str)
    if len(vals):
        # 首尾空格
        ws = int((vals != vals.str.strip()).sum())
        if ws:
            col['issues'].append({'type': '首尾空格', 'detail': f'{ws} 个值含首尾空格', 'severity': 'low', 'count': ws})
        # 全角字符
        fw = int(vals.apply(is_fullwidth).sum())
        if fw:
            col['issues'].append({'type': '全角字符', 'detail': f'{fw} 个值含全角字符', 'severity': 'low', 'count': fw})

    # 数字列里混入非数字字符：先剥离货币符号/逗号再判断
    stripped = vals.str.replace(r'[元￥¥$,，\s]', '', regex=True).str.rstrip('%')
    numeric_after = pd.to_numeric(stripped, errors='coerce')
    convertible = int(numeric_after.notna().sum())
    total_nonnull = len(vals)
    if convertible >= 2 and convertible < total_nonnull:
        dirty = total_nonnull - convertible
        bad_examples = vals[numeric_after.isna()].head(3).tolist()
        col['issues'].append({'type': '数字列含杂质', 'detail': f'{dirty} 个值含"元"/"￥"/逗号等杂质，如 {bad_examples}',
                              'severity': 'medium', 'count': dirty, 'examples': bad_examples})

    # 疑似日期列但解析失败（先做中文日期归一化）
    if looks_like_date(s.tolist()):
        normalized = s.apply(lambda v: normalize_date_str(v) if pd.notna(v) and str(v).strip() != '' else v)
        parsed = pd.to_datetime(normalized, errors='coerce', format='mixed')
        fail = int(parsed.isna().sum() - null_count)
        if fail > 0:
            examples = s[parsed.isna() & ~null_mask].astype(str).head(3).tolist()
            col['issues'].append({'type': '日期格式混乱', 'detail': f'{fail} 个值无法解析为日期，如 {examples}',
                                  'severity': 'medium', 'count': fail, 'examples': examples})

    # 常量列（除空值外只有一个值）
    uniq = s[~null_mask].nunique()
    col['unique_count'] = int(uniq)
    if uniq == 1 and n - null_count > 1:
        col['issues'].append({'type': '常量列', 'detail': f'非空值全部为"{s[~null_mask].iloc[0]}"，可能是无用列', 'severity': 'low'})

    # 疑似应唯一的列（如编号/单号列）出现重复值
    if uniq > 1 and ID_HINT_RE.search(str(name)) and (n - null_count - uniq) > 0:
        col['issues'].append({'type': '编号列重复', 'detail': f'疑似唯一标识列出现重复值（{uniq} 个唯一值 / {n - null_count} 个非空值）',
                              'severity': 'medium'})

    return col


def diagnose_sheet(df: pd.DataFrame, sheet_name: str, merged_count: int) -> dict:
    report = {'sheet': sheet_name, 'rows': len(df), 'cols': len(df.columns),
              'merged_cells': merged_count, 'issues': [], 'columns': []}
    if merged_count:
        report['issues'].append({'type': '合并单元格', 'detail': f'存在 {merged_count} 处合并单元格，会影响筛选/透视',
                                 'severity': 'medium', 'count': merged_count})

    # 疑似多行表头：前两行空值率极高
    if len(df) >= 2:
        head_null = (df.head(2).isna().mean(axis=1) > 0.6).sum()
        if head_null >= 2 and len(df.columns) > 3:
            report['issues'].append({'type': '疑似多行表头', 'detail': '前两行大面积为空，可能表头占了多行',
                                     'severity': 'medium'})

    dup = int(df.duplicated().sum())
    report['duplicate_rows'] = dup
    if dup:
        report['issues'].append({'type': '重复行', 'detail': f'存在 {dup} 行完全重复的数据', 'severity': 'medium', 'count': dup})
    else:
        # 忽略首尾空格/大小写后的近似重复
        norm = df.apply(lambda col: col.astype(str).str.strip().str.lower() if col.dtype == object else col.astype(str))
        near_dup = int(norm.duplicated().sum())
        if near_dup:
            report['issues'].append({'type': '疑似重复行', 'detail': f'忽略首尾空格后有 {near_dup} 行疑似重复（如"张三"与" 张三 "）',
                                     'severity': 'medium', 'count': near_dup})

    for c in df.columns:
        report['columns'].append(diagnose_column(c, df[c]))
    return report


def health_score(sheets: list) -> int:
    score = 100
    for sh in sheets:
        for iss in sh['issues']:
            t = iss['type']
            if t == '合并单元格':
                score -= PENALTY['merged_cells']
            elif t == '重复行':
                score -= min(PENALTY['dup_rows'], iss.get('count', 0))
            elif t == '疑似多行表头':
                score -= PENALTY['multi_header']
        for col in sh['columns']:
            n = col['total_rows'] or 1
            for iss in col['issues']:
                t = iss['type']
                if t == '空值偏多':
                    score -= PENALTY['null_rate'] * col['null_rate']
                elif t == '数字列含杂质':
                    score -= min(PENALTY['dirty_numeric'], iss.get('count', 0) / n * 50)
                elif t == '日期格式混乱':
                    score -= min(PENALTY['date_fail'], iss.get('count', 0) / n * 50)
                elif t == '首尾空格':
                    score -= PENALTY['whitespace'] * min(1, iss.get('count', 0) / n * 5)
                elif t == '全角字符':
                    score -= PENALTY['fullwidth'] * min(1, iss.get('count', 0) / n * 5)
                elif t == '常量列':
                    score -= PENALTY['constant_col']
    return max(0, round(score))


def read_csv_smart(path: str) -> pd.DataFrame:
    """中文 CSV 编码自动识别：utf-8-sig -> gbk -> gb18030。实测中文 Excel 导出的 CSV 大多是 GBK，直接 utf-8 读会报 UnicodeDecodeError。"""
    last_err = None
    for enc in ('utf-8-sig', 'gbk', 'gb18030'):
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=True, encoding=enc)
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
    raise last_err


def load_workbook(path: str, sheet=None):
    if path.lower().endswith('.csv'):
        df = read_csv_smart(path)
        return [(sheet or 'Sheet1', df, 0)]
    try:
        xls = pd.ExcelFile(path)
    except ImportError as e:
        raise RuntimeError(f'缺少读取该格式所需的依赖: {e}（.xls 老格式需要 pip install xlrd）')
    names = [sheet] if sheet else xls.sheet_names
    out = []
    try:
        from openpyxl import load_workbook as olw
        # 注意：read_only=True 时读不到 merged_cells，必须用普通模式
        wb = olw(path, data_only=True)
        merged = {ws.title: len(list(ws.merged_cells.ranges)) for ws in wb.worksheets}
        wb.close()
    except Exception:
        merged = {}
    for name in names:
        df = pd.read_excel(xls, sheet_name=name, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        out.append((name, df, merged.get(name, 0)))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('file', help='Excel/CSV 文件路径')
    ap.add_argument('--sheet', default=None, help='只诊断指定工作表')
    ap.add_argument('--out', default=None, help='JSON 报告输出路径（默认打印到 stdout）')
    ap.add_argument('--sample', type=int, default=0, help='抽样行数（0=全量）')
    args = ap.parse_args()

    try:
        sheets_data = load_workbook(args.file, args.sheet)
    except Exception as e:
        print(json.dumps({'error': f'文件读取失败: {e}'}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)

    sheets = []
    for name, df, merged_count in sheets_data:
        if args.sample and len(df) > args.sample:
            df = df.sample(args.sample, random_state=42).sort_index()
        sheets.append(diagnose_sheet(df, name, merged_count))

    total_issues = sum(len(sh['issues']) + sum(len(c['issues']) for c in sh['columns']) for sh in sheets)
    severity_counts = {'high': 0, 'medium': 0, 'low': 0, 'info': 0}
    for sh in sheets:
        for iss in sh['issues'] + [i for c in sh['columns'] for i in c['issues']]:
            sev = iss.get('severity', 'info')
            severity_counts[sev] = severity_counts.get(sev, 0) + 1
    report = {
        'file': args.file,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'health_score': health_score(sheets),
        'total_issues': total_issues,
        'severity_counts': severity_counts,
        'sheets': sheets,
    }
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f'体检报告已写入 {args.out}，健康评分：{report["health_score"]} 分')
    else:
        print(text)


if __name__ == '__main__':
    main()

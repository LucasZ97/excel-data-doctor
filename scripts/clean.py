#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格体检医生 - 一键安全修复脚本
只执行"安全修复"（去空格、全角转半角、删全空行列、规范列名、数字去杂质、日期统一），
不碰需要人工确认的项（重复行删除、空值填充、合并单元格拆分等）。

用法:
    python3 clean.py <文件路径> [--sheet 表名] [--out 清洗后文件] [--report 动作日志.json]

输出: 清洗后的新文件 + JSON 动作日志（哪一列、修了什么、影响了多少行）。
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


def full_to_half(s):
    if not isinstance(s, str):
        return s
    return ''.join(chr(ord(ch) - 0xFEE0) if '！' <= ch <= '～' else ch for ch in s).replace('　', ' ')


def normalize_date_str(s: str) -> str:
    t = str(s).strip()
    t = re.sub(r'(\d)\s*年\s*', r'\1-', t)
    t = re.sub(r'(\d)\s*月\s*', r'\1-', t)
    t = re.sub(r'(\d)\s*日\s*', r'\1', t)
    return t.strip('- ')


def looks_like_date(values) -> bool:
    sample = [v for v in values if v not in (None, '')][:20]
    if len(sample) < 3:
        return False
    hit = sum(1 for v in sample if re.search(r'\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}', str(v)))
    return hit >= max(2, len(sample) * 0.5)


def looks_like_dirty_numeric(series: pd.Series) -> bool:
    vals = series[~(series.isna() | (series.astype(str).str.strip() == ''))].astype(str)
    if len(vals) < 2:
        return False
    stripped = vals.str.replace(r'[元￥¥$,，\s]', '', regex=True).str.rstrip('%')
    ok = int(pd.to_numeric(stripped, errors='coerce').notna().sum())
    return ok >= 2 and ok >= len(vals) * 0.5


def read_input(path: str, sheet=None):
    if path.lower().endswith('.csv'):
        last_err = None
        for enc in ('utf-8-sig', 'gbk', 'gb18030'):
            try:
                df = pd.read_csv(path, dtype=str, keep_default_na=True, encoding=enc)
                return {'Sheet1': df}, False
            except (UnicodeDecodeError, UnicodeError) as e:
                last_err = e
        raise last_err
    xls = pd.ExcelFile(path)
    names = [sheet] if sheet else xls.sheet_names
    return {n: pd.read_excel(xls, sheet_name=n, dtype=object) for n in names}, True


def clean_sheet(df: pd.DataFrame):
    actions = []
    df = df.copy()
    df.columns = [str(c).strip().replace('\n', '') for c in df.columns]

    # 1. 去首尾空格 + 全角转半角
    for c in df.select_dtypes(include='object').columns:
        before = df[c].astype(str)
        df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x).apply(full_to_half)
        n = int((before != df[c].astype(str)).sum())
        if n:
            actions.append({'action': '去首尾空格/全角转半角', 'column': str(c), 'rows': n})

    # 2. 删除全空行、全空列
    n_rows = int(df.isna().all(axis=1).sum())
    n_cols = int(df.isna().all(axis=0).sum())
    df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    if n_rows:
        actions.append({'action': '删除全空行', 'column': '-', 'rows': n_rows})
    if n_cols:
        actions.append({'action': '删除全空列', 'column': '-', 'rows': n_cols})

    # 3. 数字列去杂质
    for c in df.select_dtypes(include='object').columns:
        if looks_like_dirty_numeric(df[c]):
            def clean_num(s):
                if pd.isna(s) or str(s).strip() == '':
                    return float('nan')
                t = str(s).strip().replace(',', '').replace('，', '').replace('￥', '') \
                    .replace('¥', '').replace('$', '').replace('元', '')
                try:
                    return float(t[:-1]) / 100 if t.endswith('%') else float(t)
                except ValueError:
                    return s  # 实在转不动的保留原值，不静默丢弃
            df[c] = df[c].apply(clean_num)
            n = int(pd.to_numeric(df[c], errors='coerce').notna().sum())
            actions.append({'action': '去"元/￥/逗号/百分号"转数值', 'column': str(c), 'rows': n})

    # 4. 日期列统一
    for c in df.select_dtypes(include='object').columns:
        if looks_like_date(df[c].tolist()):
            parsed = df[c].apply(lambda v: pd.to_datetime(
                normalize_date_str(v), errors='coerce', format='mixed')
                if pd.notna(v) and str(v).strip() != '' else pd.NaT)
            n = int(parsed.notna().sum())
            df[c] = parsed.dt.strftime('%Y-%m-%d')
            actions.append({'action': '统一日期格式 YYYY-MM-DD', 'column': str(c), 'rows': n})

    return df, actions


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('file', help='Excel/CSV 文件路径')
    ap.add_argument('--sheet', default=None, help='只清洗指定工作表')
    ap.add_argument('--out', default=None, help='清洗后文件路径（默认 <原文件名>_cleaned.xlsx）')
    ap.add_argument('--report', default=None, help='动作日志 JSON 输出路径（默认打印到 stdout）')
    args = ap.parse_args()

    try:
        sheets, is_excel = read_input(args.file, args.sheet)
    except Exception as e:
        print(json.dumps({'error': f'文件读取失败: {e}'}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)

    out_path = args.out or re.sub(r'(\.[^.]+)$', r'_cleaned\1', args.file)
    if not out_path.lower().endswith(('.xlsx', '.csv')):
        out_path += '_cleaned.xlsx'

    all_actions = []
    cleaned = {}
    for name, df in sheets.items():
        new_df, actions = clean_sheet(df)
        cleaned[name] = new_df
        for a in actions:
            a['sheet'] = name
            all_actions.append(a)

    if out_path.lower().endswith('.csv'):
        # 多 sheet 时只写第一个
        first = next(iter(cleaned.values()))
        first.to_csv(out_path, index=False)
    else:
        with pd.ExcelWriter(out_path, engine='openpyxl') as w:
            for name, df in cleaned.items():
                df.to_excel(w, sheet_name=name[:31], index=False)

    log = {
        'file': args.file,
        'cleaned_file': out_path,
        'cleaned_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'actions': all_actions,
        'note': '仅执行安全修复；重复行删除、空值填充、合并单元格拆分等需人工确认，未执行。',
    }
    text = json.dumps(log, ensure_ascii=False, indent=2)
    if args.report:
        with open(args.report, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f'清洗完成：{out_path}，共 {len(all_actions)} 个修复动作，日志已写入 {args.report}')
    else:
        print(text)


if __name__ == '__main__':
    main()

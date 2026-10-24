#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""表格体检医生 - 烟雾测试：一键验证技能包完整可用。用法: python3 tests/smoke_test.py"""
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL_DIR / 'scripts'
passed, failed = [], []

def check(name, cond, detail=''):
    (passed if cond else failed).append(name)
    print(('PASS' if cond else 'FAIL'), '-', name, detail)


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=SCRIPTS)
    return r


# 1. SKILL.md 元数据校验（对标官方 skill-creator 的 quick_validate 规则）
front = (SKILL_DIR / 'SKILL.md').read_text(encoding='utf-8')
m = re.search(r'^---\n(.*?)\n---', front, re.S)
check('SKILL.md 存在 frontmatter', bool(m))
meta = m.group(1) if m else ''
name = re.search(r'^name:\s*(\S+)', meta, re.M)
check('name 为 hyphen-case', bool(name) and re.fullmatch(r'[a-z0-9-]+', name.group(1)) is not None,
      f'got: {name.group(1) if name else None}')
desc = re.search(r'^description:\s*"(.*)"', meta, re.M | re.S)
check('description 存在', bool(desc))
check('description 无尖括号', bool(desc) and '<' not in desc.group(1) and '>' not in desc.group(1))
check('description 提到触发场景', bool(desc) and '当用户' in desc.group(1))

# 2. 造脏表 -> 诊断必须抓到关键问题
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    import openpyxl
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = '销售表'
    ws.append(['客户名称', '销售日期', '金额', '备注'])
    for r in [[' 张三 ', '2024/1/5', '1,200元', '正常'],
              ['李四', '2024年1月7日', '￥3500', None],
              ['王五', '不详', 'abc', '　全角']]:
        ws.append(r)
    dirty = tmp / 'dirty.xlsx'; wb.save(dirty)

    r = run([sys.executable, 'diagnose.py', str(dirty), '--out', str(tmp / 'r.json')])
    check('diagnose.py 正常退出', r.returncode == 0, r.stderr[:200])
    rep = json.loads((tmp / 'r.json').read_text(encoding='utf-8'))
    types = {i['type'] for sh in rep['sheets']
             for i in sh['issues'] + [x for c in sh['columns'] for x in c['issues']]}
    for t in ['首尾空格', '日期格式混乱', '数字列含杂质', '全角字符']:
        check(f'诊断检出「{t}」', t in types)
    check('健康评分 < 100', rep['health_score'] < 100, f"score={rep['health_score']}")
    before = rep['health_score']

    # 3. 清洗 -> 评分必须提升，原文件不动
    r = run([sys.executable, 'clean.py', str(dirty), '--out', str(tmp / 'cleaned.xlsx'),
             '--report', str(tmp / 'fixes.json')])
    check('clean.py 正常退出', r.returncode == 0, r.stderr[:200])
    check('清洗文件已生成', (tmp / 'cleaned.xlsx').exists())
    check('原文件未被改动', dirty.exists())
    r = run([sys.executable, 'diagnose.py', str(tmp / 'cleaned.xlsx'), '--out', str(tmp / 'r2.json')])
    after = json.loads((tmp / 'r2.json').read_text(encoding='utf-8'))['health_score']
    check('清洗后评分提升', after > before, f'{before} -> {after}')
    fixes = json.loads((tmp / 'fixes.json').read_text(encoding='utf-8'))
    check('动作日志非空', len(fixes['actions']) > 0)

    # 4. CSV 中文编码（GBK）能读
    (tmp / 'gbk.csv').write_bytes('姓名,金额\n张三,"1,200元"\n'.encode('gbk'))
    r = run([sys.executable, 'diagnose.py', str(tmp / 'gbk.csv'), '--out', str(tmp / 'r3.json')])
    rep3 = json.loads((tmp / 'r3.json').read_text(encoding='utf-8'))
    check('GBK 编码 CSV 正常诊断', r.returncode == 0 and rep3['sheets'][0]['rows'] == 1 and '姓名' in rep3['sheets'][0]['columns'][0]['column'], r.stderr[:200])

print(f'\n{len(passed)} passed, {len(failed)} failed')
sys.exit(1 if failed else 0)

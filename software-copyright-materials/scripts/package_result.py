#!/usr/bin/env python3
"""Package exactly three submission drafts after successful physical page QA."""
import argparse,json,zipfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--report',required=True);p.add_argument('--validation',required=True);p.add_argument('--output',required=True);a=p.parse_args()
meta=json.loads(Path(a.report).read_text(encoding='utf8'));validation=json.loads(Path(a.validation).read_text(encoding='utf8'))
if not validation.get('passed'):raise SystemExit('未通过渲染逐页核对，不允许打包为完成版。')
files=[Path(x) for x in meta['files'].values()]
if len(files)!=3 or not all(x.exists() for x in files):raise SystemExit('必须恰好存在两份 Word 和一份 Markdown')
if sorted(x.suffix for x in files)!=['.docx','.docx','.md']:raise SystemExit('输出格式错误')
with zipfile.ZipFile(a.output,'w',zipfile.ZIP_DEFLATED) as z:
 for f in files:z.write(f,arcname=f.name)
print('已打包：',a.output)

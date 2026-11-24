#!/usr/bin/env python3
"""Inventory an uploaded source directory/ZIP without executing its code.

Usage: python source_inventory.py --source project.zip --out work/inventory.json
"""
import argparse, hashlib, json, os, re, shutil, stat, tempfile, zipfile
from collections import Counter
from pathlib import Path

EXTS = {'.py':'Python','.js':'JavaScript','.jsx':'JSX','.ts':'TypeScript','.tsx':'TSX','.go':'Go','.java':'Java','.c':'C','.h':'C/C++','.hpp':'C++','.cc':'C++','.cpp':'C++','.cs':'C#','.rs':'Rust','.rb':'Ruby','.php':'PHP','.html':'HTML','.htm':'HTML','.css':'CSS','.scss':'SCSS','.sql':'SQL','.sh':'Shell','.ps1':'PowerShell','.kt':'Kotlin','.kts':'Kotlin','.swift':'Swift','.dart':'Dart','.vue':'Vue','.svelte':'Svelte','.m':'Objective-C','.mm':'Objective-C++','.ino':'Arduino','.lua':'Lua','.pl':'Perl','.r':'R','.jl':'Julia','.bat':'Batch','.cmd':'Batch','.xml':'XML'}
EXCLUDE_DIRS = {'.git','.hg','.svn','node_modules','venv','.venv','__pycache__','dist','build','target','.next','.nuxt','.turbo','coverage','.idea','.vscode','vendor','third_party','third-party','packages_cache','site-packages','.gradle','.mvn','bin','obj','out','.tox'}
EXCLUDE_FILE = re.compile(r'(^package-lock\.json$|^yarn\.lock$|^pnpm-lock\.yaml$|^poetry\.lock$|\.min\.(js|css)$|\.map$|\.(snap|generated|g)\.(js|ts)$)',re.I)
MAX_UNCOMPRESSED = 250 * 1024 * 1024
MAX_FILE = 4 * 1024 * 1024

def safe_unzip(path, dest):
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as z:
        total = 0
        for item in z.infolist():
            if item.is_dir(): continue
            rel = Path(item.filename.replace('\\','/'))
            if not item.filename or rel.is_absolute() or '..' in rel.parts or '\x00' in item.filename:
                raise ValueError(f'压缩包存在不安全路径：{item.filename}')
            if (item.external_attr >> 16) & 0o170000 == stat.S_IFLNK:
                raise ValueError(f'不允许压缩包符号链接：{item.filename}')
            total += item.file_size
            if total > MAX_UNCOMPRESSED or item.file_size > MAX_FILE * 3:
                raise ValueError('ZIP 过大；请清理依赖包与编译产物后再试')
            out = dest / rel
            out.parent.mkdir(parents=True, exist_ok=True)
            with z.open(item) as src, out.open('wb') as dst: shutil.copyfileobj(src,dst)

def decode(b):
    for enc in ('utf-8-sig','utf-16','gb18030'):
        try: return b.decode(enc), enc
        except UnicodeError: pass
    return b.decode('utf-8','replace'), 'lossy-utf8'

def inventory(source, out, extract_to=None):
    source=Path(source).expanduser().resolve(); out=Path(out).resolve()
    if source.is_file() and source.suffix.lower()=='.zip':
        extract_dir=Path(extract_to or out.parent / '_source_extract').resolve()
        safe_unzip(source,extract_dir); root=extract_dir
    elif source.is_dir(): root=source
    else: raise ValueError('输入必须是目录或 ZIP 源码压缩包')
    files=[]; excluded=[]; counter=Counter()
    for p in sorted(root.rglob('*')):
        if not p.is_file() or p.is_symlink():continue
        rel=p.relative_to(root).as_posix()
        if any(part in EXCLUDE_DIRS or part.startswith('.') and part!='.github' for part in p.relative_to(root).parts[:-1]):
            excluded.append({'path':rel,'reason':'依赖、缓存、构建或隐藏目录'});continue
        if p.suffix.lower() not in EXTS or EXCLUDE_FILE.search(p.name):
            excluded.append({'path':rel,'reason':'非源码或生成文件'});continue
        if p.stat().st_size>MAX_FILE:
            excluded.append({'path':rel,'reason':'单文件过大，需人工核对'});continue
        data=p.read_bytes()
        if b'\x00' in data[:2048]:excluded.append({'path':rel,'reason':'疑似二进制文件'});continue
        content,enc=decode(data)
        lines=content.splitlines()
        effective=sum(1 for line in lines if line.strip())
        record={'path':rel,'absolute_path':str(p),'language':EXTS[p.suffix.lower()],'encoding':enc,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),'total_lines':len(lines),'nonempty_lines':effective,'max_line_chars':max(map(len,lines),default=0),'possible_entrypoint':bool(re.search(r'(^|/)(main|app|server|index|run|start|application)\.',rel,re.I))}
        files.append(record);counter[record['language']]+=effective
    if not files: raise ValueError('源码包中未发现可识别源码；请人工核对扩展名和目录')
    result={'source':str(source),'extracted_root':str(root),'file_count':len(files),'total_nonempty_lines':sum(f['nonempty_lines'] for f in files),'languages':dict(counter),'files':files,'excluded':excluded,'notes':['只统计可识别源文件中的非空行；不把依赖、锁文件、编译产物当作自研代码。','统计值仅供核对，申请官网“源程序量”字段应与最终申报口径一致。','项目扫描不执行任何上传源码。']}
    out.parent.mkdir(parents=True,exist_ok=True);out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'inventory':str(out),'source_files':len(files),'nonempty_source_lines':result['total_nonempty_lines'],'languages':result['languages']},ensure_ascii=False))
    return result

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--source',required=True);p.add_argument('--out',required=True);p.add_argument('--extract-to');a=p.parse_args();inventory(a.source,a.out,a.extract_to)

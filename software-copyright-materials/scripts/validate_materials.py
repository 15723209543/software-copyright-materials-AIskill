#!/usr/bin/env python3
"""Render actual DOCX pages and validate code lines and screenshot placeholders."""
import argparse,json,re,shutil,subprocess,sys, tempfile
from pathlib import Path
import fitz
from docx import Document

def docx_to_pdf(path,workdir):
    if not shutil.which('libreoffice'):raise RuntimeError('缺少 LibreOffice，不能确认 Word 实际分页。请在有 Office/LibreOffice 环境再次校验。')
    profile=workdir / ('profile_'+path.stem[:10]);profile.mkdir(exist_ok=True)
    cmd=['libreoffice','-env:UserInstallation=file://'+str(profile.resolve()),'--headless','--convert-to','pdf','--outdir',str(workdir),str(path)]
    done=subprocess.run(cmd,capture_output=True,text=True,timeout=180)
    out=workdir/(path.stem+'.pdf')
    if done.returncode or not out.exists():raise RuntimeError('Word 渲染失败：'+done.stderr+done.stdout)
    return out

def validate(meta,keep_preview=None):
    paths={k:Path(v) for k,v in meta['files'].items()};issues=[];warnings=[];output={}
    for k,path in paths.items():
        if not path.is_file():issues.append('缺少文件：'+str(path))
    if issues:return {'passed':False,'errors':issues}
    manual=Document(paths['manual']);text='\n'.join(p.text for p in manual.paragraphs)
    if '4 软件界面与用户操作流程' not in text:issues.append('说明书缺少第4章功能操作流程')
    slots=re.findall(r'【截图提示：.*?】',text)
    expected_screens=len(re.findall(r'^4\.\d+ ',text,re.M))
    if len(slots)!=expected_screens:issues.append(f'截图提示与第4章操作项数量不一致：{len(slots)}/{expected_screens}')
    if expected_screens<10:warnings.append(f'第四章仅有 {expected_screens} 种真实操作；目标为至少10种，但不能编造不存在的操作')
    if '【截图提示' not in text:issues.append('说明书没有明确的截图提示占位符')
    for op in re.findall(r'^4\.\d+ .*',text,re.M):
        if op.count('【')>0:warnings.append('操作标题中包含占位文字：'+op[:45])
    md=paths['form'].read_text(encoding='utf-8')
    if '【待申请人确认' not in md:warnings.append('官网填报信息无人工核实标记；请确认法律身份信息已由申请人提供')
    if len(re.findall(r'^\|[^|]+\|',md,re.M))<65:issues.append('官网填报清单字段不足，可能被截断')
    with tempfile.TemporaryDirectory(prefix='scskill_qa_') as td:
        work=Path(td)
        for key in ('code','manual'):
            pdf=docx_to_pdf(paths[key],work)
            document=fitz.open(pdf)
            if key=='code':
                expected=meta['code']['expected_pages']
                if document.page_count !=expected:issues.append(f'程序鉴别文档实际页数 {document.page_count}，应为 {expected} 页（存在自动溢出或空白页）')
                actual=[];endpoints=[]
                for i,page in enumerate(document):
                    # Original code line numbers are in the left margin of each line.
                    content=page.get_text('text')
                    numbered=[ln for ln in content.splitlines() if re.match(r'^\s*\d{5}\s+',ln)]
                    actual.append(len(numbered))
                    blocks=[b for b in page.get_text('blocks') if (b[4] or '').strip()]
                    if blocks:
                        lowest=max(b[3] for b in blocks);endpoints.append(round(lowest/page.rect.height,3))
                    else:endpoints.append(0)
                    if len(numbered)<50:issues.append(f'程序文档实际第 {i+1} 页仅提取到 {len(numbered)} 行源码（至少50行；检查折行与溢出）')
                    if blocks and endpoints[-1]<0.58:warnings.append(f'第 {i+1} 页下半部分可能存在较多空白（内容截止约 {endpoints[-1]*100:.0f}% 页高）')
                    for line in numbered:
                        if '【待填写】' in line:issues.append('程序中存在填充占位文本')
                output['code_pages_actual']=len(actual);output['code_lines_each_page']=actual;output['code_page_content_end_ratio']=endpoints
                if sum(actual)!=len(actual)*0+sum(p['source_nonempty_lines'] for p in meta['code']['pages']):
                    issues.append(f'实际识别的源码行数 {sum(actual)} 与所选源代码行数 {sum(p["source_nonempty_lines"] for p in meta["code"]["pages"])} 不一致')
            else:
                output['manual_pages_actual']=document.page_count
                if document.page_count<3:warnings.append('说明书较短，请确认第四章操作说明足够详细')
            if keep_preview:
                preview=Path(keep_preview);preview.mkdir(parents=True,exist_ok=True)
                # previews are QA artifacts, not one of the three user deliverables
                for i in range(min(document.page_count,3)):
                    p=document.load_page(i)
                    pix=p.get_pixmap(matrix=fitz.Matrix(1.2,1.2),alpha=False)
                    pix.save(str(preview/f'{key}_page_{i+1}.png'))
                if key=='code' and document.page_count>=4:
                    p=document.load_page(document.page_count-1)
                    p.get_pixmap(matrix=fitz.Matrix(1.2,1.2),alpha=False).save(str(preview/'code_last_page.png'))
    report={'passed':not issues,'errors':issues,'warnings':warnings,'results':output}
    return report

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--report',required=True,help='build_materials.py 生成的 _build_report.json');p.add_argument('--out',required=True);p.add_argument('--preview-dir');a=p.parse_args()
    meta=json.loads(Path(a.report).read_text(encoding='utf-8'))
    result=validate(meta,a.preview_dir)
    Path(a.out).parent.mkdir(parents=True,exist_ok=True)
    Path(a.out).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False,indent=2))
    sys.exit(0 if result['passed'] else 2)

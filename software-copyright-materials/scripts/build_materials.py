#!/usr/bin/env python3
"""Build three software copyright application deliverables from audited project evidence.

Input project_analysis.json MUST be authored after inspecting the actual code. This
script does not invent product features, rights ownership, dates, or screenshots.
"""
import argparse, json, math, re, sys
from pathlib import Path
from collections import defaultdict
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SOURCE_EXT_ALLOW={'.py','.js','.jsx','.ts','.tsx','.go','.java','.c','.h','.hpp','.cc','.cpp','.cs','.rs','.rb','.php','.html','.htm','.css','.scss','.sql','.sh','.ps1','.kt','.kts','.swift','.dart','.vue','.svelte','.m','.mm','.ino','.lua','.pl','.r','.jl','.bat','.cmd','.xml'}

def font(run, chinese='宋体', pt=10.5, bold=False):
    run.font.name=chinese;run.font.size=Pt(pt);run.bold=bold;run.font.color.rgb=RGBColor(0,0,0)
    rpr=run._element.get_or_add_rPr();rfonts=rpr.rFonts
    if rfonts is None:
        rfonts=OxmlElement('w:rFonts');rpr.insert(0,rfonts)
    rfonts.set(qn('w:eastAsia'),chinese);rfonts.set(qn('w:ascii'),chinese);rfonts.set(qn('w:hAnsi'),chinese)

def base_doc():
    d=Document();sec=d.sections[0];sec.page_height=Cm(29.7);sec.page_width=Cm(21)
    sec.top_margin=Cm(1.5);sec.bottom_margin=Cm(1.5);sec.left_margin=Cm(1.6);sec.right_margin=Cm(1.6)
    sec.header_distance=Cm(0);sec.footer_distance=Cm(0)
    normal=d.styles['Normal'];normal.font.name='宋体';normal.font.size=Pt(10.5)
    normal.paragraph_format.space_before=Pt(0);normal.paragraph_format.space_after=Pt(6)
    for nm in ['Heading 1','Heading 2','Heading 3']:
        st=d.styles[nm];st.font.name='黑体';st.font.color.rgb=RGBColor(0,0,0)
        st.font.size=Pt({'Heading 1':16,'Heading 2':14,'Heading 3':12}[nm])
        rpr=st._element.get_or_add_rPr();ff=rpr.rFonts
        if ff is None:ff=OxmlElement('w:rFonts');rpr.insert(0,ff)
        ff.set(qn('w:eastAsia'),'黑体')
    return d

def ptext(doc,txt,style=None):
    p=doc.add_paragraph(style=style)
    for i,line in enumerate(str(txt).split('\n')):
        if i:p.add_run().add_break()
        font(p.add_run(line),'宋体',10.5)
    return p

def read_source_lines(inv, order):
    table={f['path']:f for f in inv['files']}
    paths=order or [f['path'] for f in inv['files']]
    missing=[x for x in paths if x not in table]
    if missing:raise ValueError('analysis.source_order 引用了不存在的源码文件：'+','.join(missing))
    stream=[]
    for rel in paths:
        entry=table[rel];path=Path(entry['absolute_path'])
        if not path.is_file():raise ValueError('无法读取源码：'+str(path))
        enc=entry['encoding'];enc='utf-8' if enc=='lossy-utf8' else enc
        text=path.read_text(encoding=enc,errors='replace')
        for lineno,line in enumerate(text.splitlines(),1):
            if not line.strip():continue
            stream.append({'path':rel,'lineno':lineno,'text':line.expandtabs(4)})
    return stream

def select_stream(stream):
    if len(stream)<50:raise ValueError(f'源程序仅有 {len(stream)} 行非空代码，无法合法满足“每页至少50行”，不生成伪造代码；请补充真实源码。')
    if len(stream)<=3000:return stream,'全部代码（不超过3000非空行）'
    return stream[:1500]+stream[-1500:],'大于3000行：按固定源码顺序选取前1500行和后1500行（各30页）'

def code_style_by_lines(n):
    if n<=55:return 9.0,13.0
    if n<=65:return 8.3,11.3
    if n<=80:return 7.5,9.0
    return 6.45,7.05

def code_header(doc,path,lineno):
    p=doc.add_paragraph();p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(5)
    p.paragraph_format.keep_with_next=True
    p.paragraph_format.line_spacing=1.0
    font(p.add_run(f'代码文件：{path}（源文件第{lineno}行起）'),'黑体',22,True)
    return p

def make_code_doc(inv, analysis, output):
    stream=read_source_lines(inv,analysis.get('source_order'))
    stream,policy=select_stream(stream)
    pages_n=min(60,max(1,len(stream)//50));base=len(stream)//pages_n;extra=len(stream)%pages_n
    sizes=[base+(1 if i<extra else 0) for i in range(pages_n)]
    d=base_doc();d.core_properties.title=f'{analysis["software_name"]} {analysis["version"]} 程序鉴别材料'
    # No cover, no header/footer. Every physical page is intended to contain code.
    cursor=0;prev=None;page_info=[]
    for pi,num in enumerate(sizes,1):
        chunk=stream[cursor:cursor+num];cursor+=num
        fsize,spacing=code_style_by_lines(num)
        # Source files can be short: avoid overflow when many 22pt path headings appear on one page.
        transitions=sum(1 for j,x in enumerate(chunk) if j==0 or x['path']!=chunk[j-1]['path'] or x['lineno']!=chunk[j-1]['lineno']+1)
        if transitions>2:spacing=max(7.7,spacing-0.55*(transitions-2))
        titles=0
        for x in chunk:
            k=(x['path'],x['lineno'])
            if prev is None or x['path']!=prev['path'] or x['lineno']!=prev['lineno']+1:
                code_header(d,x['path'],x['lineno']);titles+=1
            para=d.add_paragraph();para.paragraph_format.space_before=Pt(0);para.paragraph_format.space_after=Pt(0)
            para.paragraph_format.line_spacing_rule=WD_LINE_SPACING.EXACTLY
            para.paragraph_format.line_spacing=Pt(spacing)
            para.paragraph_format.keep_together=True;para.paragraph_format.keep_with_next=False
            txt=f'{x["lineno"]:05d}  {x["text"]}'
            font(para.add_run(txt),'宋体',fsize)
            prev=x
        page_info.append({'page':pi,'source_nonempty_lines':num,'headings':titles,'first':f'{chunk[0]["path"]}:{chunk[0]["lineno"]}','last':f'{chunk[-1]["path"]}:{chunk[-1]["lineno"]}','max_rendered_characters':max(len(x['text']) for x in chunk)})
        if pi < pages_n:d.add_page_break()
    d.save(output)
    return {'source_lines_total':len(stream),'selection_policy':policy,'expected_pages':pages_n,'pages':page_info,'source_order':analysis.get('source_order') or [x['path'] for x in inv['files']]}

def add_kv_table(d,items):
    t=d.add_table(rows=0,cols=2);t.style='Table Grid';t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for key,val in items:
        c=t.add_row().cells;c[0].text=str(key);c[1].text=str(val)
        for cell in c:
            for p in cell.paragraphs:
                for run in p.runs:font(run,'宋体',10)

def add_section(d,section_num,title,paragraphs):
    if not paragraphs: return False
    d.add_heading(f'{section_num} {title}',level=1)
    for item in paragraphs:
        if isinstance(item,dict):
            if item.get('title'):d.add_heading(item['title'],level=2)
            ptext(d,item.get('text',''))
        elif str(item).strip():ptext(d,item)
    return True

def manual_document(analysis,inv,out):
    d=base_doc();name=analysis['software_name'];version=analysis['version']
    p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(90)
    font(p.add_run(name),'黑体',24,True)
    p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;font(p.add_run(version),'黑体',16)
    p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(30)
    font(p.add_run('软件说明书'),'黑体',26,True)
    p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(18)
    font(p.add_run('计算机软件著作权登记用软件说明文档'),'宋体',11)
    d.add_page_break()
    d.add_heading('1 软件概述与基本信息',1)
    d.add_heading('1.1 文档用途与软件概况',2)
    ptext(d,analysis['overview'])
    d.add_heading('1.2 软件基本信息',2)
    add_kv_table(d,[('软件名称',name),('软件版本',version),('主要技术语言',', '.join(inv.get('languages',{}))),('实际源码文件数',inv.get('file_count','待核实')),('源码非空行数',inv.get('total_nonempty_lines','待核实')),('适用对象',analysis.get('users','【待项目分析确认】')),('软件用途',analysis.get('purpose','【待项目分析确认】'))])
    d.add_heading('1.3 软件主要功能',2)
    for i,f in enumerate(analysis['features'],1):ptext(d,f'{i}. {f["name"]}：{f["description"]}（对应源码：{", ".join(f["evidence_paths"]) }）')
    d.add_heading('2 需求分析与总体设计',1)
    add_section(d,'2.1','目标与使用场景',analysis.get('requirements',[]))
    add_section(d,'2.2','系统总体架构与业务流程',analysis.get('architecture',[]))
    d.add_heading('3 项目组成与运行环境',1)
    add_section(d,'3.1','项目源码模块组织',analysis.get('modules',[]))
    add_section(d,'3.2','开发与运行环境',analysis.get('environment',[]))
    add_section(d,'3.3','软件启动、访问与退出',analysis.get('startup',[]))
    d.add_heading('4 软件界面与用户操作流程',1)
    ptext(d,'本章只介绍源码和实际界面能够支持的操作。以下【截图提示】均为申请人使用本次申请的软件版本实际运行后补充的截图位置，不代表已经取得截图。')
    for i,op in enumerate(analysis['operations'],1):
        d.add_heading(f'4.{i} {op["title"]}',2)
        ptext(d,op['description'])
        ptext(d,'操作过程与页面反馈：')
        for j,step in enumerate(op['steps'],1):ptext(d,f'（{j}）{step}')
        ptext(d,'执行结果与注意事项：'+op['feedback'])
        ptext(d,'对应源码：'+', '.join(op['evidence_paths']))
        ptext(d,f'【截图提示：图4-{i}，{op["screenshot_prompt"]}；请从 {name} {version} 实际运行界面截取，截图中应清晰展示对应控件和结果，不要使用虚构或其他版本截图。】')
    for n,key,title in [(5,'implementation','核心程序结构与交互实现'),(6,'interfaces','前后端服务、数据接口与通信'),(7,'storage','数据结构、存储与处理流程'),(8,'error_handling','异常处理、输入校验与安全边界'),(9,'testing','实际运行、功能验证与故障处理'),(10,'maintenance','部署、维护、版本限制与总结')]:
        add_section(d,n,title,analysis.get(key,[]))
    d.core_properties.title=f'{name} {version} 软件说明书';d.save(out)

def val(x):
    if x is None or str(x).strip() in ('','null','unknown','待确认','未提供','不确定'):
        return '【待申请人确认：请在调用指令提供或官网核实】'
    if isinstance(x,list):return '；'.join(map(str,x)) or '【待申请人确认】'
    return str(x)

FIELD_GROUPS=[
 ('A. 申请账号、实名认证与业务选择',[
 ('平台与登记项目','platform','中国版权保护中心版权登记系统 / R11 计算机软件著作权登记申请'),('登录账号实名状态','account_verified',None),('申请人/登记办理人姓名或单位','applicant_name',None),('申请人类别','applicant_type',None),('申请人与著作权人的关系','applicant_relation_to_holder',None),('申请人联系电话','applicant_phone',None),('申请人联系邮箱','applicant_email',None),('代理申请状态','agent_mode',None),('代理人姓名或机构','agent_name',None),('当前系统/页面版本','form_system_version',None)]),
 ('B. 软件基本信息',[
 ('软件全称','software_name',None),('软件简称','short_name',None),('版本号','version',None),('软件分类号（按实时分类表选）','classification_code',None),('适用行业代码（按实时选项）','industry_code',None),('软件作品说明（原创/修改/合成/翻译等按实际）','work_type',None),('前一登记号（适用于升级修改等）','prior_registration_number',None),('修改软件说明/新增内容','modification_description',None),('项目用途与应用领域','application_domain',None)]),
 ('C. 开发与发表情况',[
 ('开发方式（独立/合作/委托/下达任务等）','development_mode',None),('开发完成日期','completion_date',None),('发表状态（已发表/未发表）','published',None),('首次发表日期','publication_date',None),('首次发表国家','publication_country',None),('首次发表城市/地区','publication_city',None),('首次发表方式及证据','publication_method',None),('开发起始日期（若系统要求）','start_date',None),('是否已升级/改版','upgrade_status',None)]),
 ('D. 著作权与权属',[
 ('权利取得方式（原始/继受）','right_acquisition',None),('继受类别（转让/继承/承受等）','successor_type',None),('权利范围（全部/部分）','right_scope',None),('部分权利具体内容','partial_rights',None),('共有著作权人名单及顺序','copyright_holders',None),('所有权归属依据','ownership_basis',None),('职务开发/非职务开发说明','employment_status',None),('学校/单位内部审批状态（适用时）','organization_approval',None),('合作、委托、任务来源相关协议','right_documents',None)]),
 ('E. 著作权人身份与联系方式（所有共有人均应核对）',[
 ('自然人姓名/机构全称','holder_name',None),('著作权人英文名称（若要求）','holder_en_name',None),('主体类型','holder_entity_type',None),('证件类型','holder_id_type',None),('证件号码/统一社会信用代码','holder_id_number',None),('国籍/注册地','holder_country',None),('省、市、区','holder_region',None),('通讯地址','holder_address',None),('邮政编码','holder_postcode',None),('著作权人联系电话','holder_phone',None),('著作权人联系邮箱','holder_email',None),('联系人姓名','contact_name',None),('联系人手机号','contact_phone',None),('联系人邮箱','contact_email',None),('联系人证件及联系方式（若要求）','contact_identity',None),('联系人通讯地址','contact_address',None),('联系人邮编','contact_postcode',None),('多个权利人的同意及授权文件','multiple_holder_consent',None)]),
 ('F. 软件技术特点与源程序量',[
 ('开发硬件环境','dev_hardware',None),('运行硬件环境','runtime_hardware',None),('开发操作系统及版本','dev_os',None),('运行操作系统及版本','runtime_os',None),('开发软件环境/IDE及版本','dev_software',None),('运行支撑环境/依赖','runtime_software',None),('开发工具/框架及版本','frameworks',None),('编程语言及版本','programming_languages',None),('数据库/存储组件','database',None),('部署形态（Web/桌面/移动/嵌入式）','deployment_type',None),('源程序总量（非空行数，需确认统计口径）','source_lines',None),('主要功能说明（可直接复制）','main_functions',None),('技术特点说明（可直接复制）','technical_features',None),('软件设计特点/优势','technical_advantages',None),('所属技术领域','technical_field',None),('可运行版本信息/入口','runtime_entry',None)]),
 ('G. 申请材料、签章及上传自检',[
 ('程序鉴别材料 Word 文件','code_docx',None),('软件说明文档 Word 文件','manual_docx',None),('PDF 转换是否必要（按官网当前页面）','pdf_required',None),('申请人主体身份证明','identity_proof',None),('联系人身份证明（如要求）','contact_proof',None),('代理委托书（代理时）','power_of_attorney',None),('合作开发合同（合作时）','collaboration_contract',None),('委托开发合同（委托时）','commission_contract',None),('任务书（下达任务时）','task_document',None),('权利转让/继承/承受证明（继受时）','successor_proof',None),('原权利人授权许可（改编他人作品时）','permission_proof',None),('软件登记历史材料（版本升级时）','previous_certificate',None),('学校或单位科研归属证明（适用时）','institutional_proof',None),('签章/实名认证/申请人承诺','signature_commitment',None),('其他实际页面要求的声明或附件','other_attachments',None)]),
 ('H. 提交与跟踪信息',[
 ('申请渠道 / 业务编号','application_channel',None),('申请表核对时间','check_date',None),('提交日期','submission_date',None),('申请受理号（提交后填写）','acceptance_number',None),('补正通知与补正截止日','correction_deadline',None),('审核状态','review_status',None),('证书下载及登记号（核准后）','certificate_number',None)]),
]

def make_official_md(analysis,inv,code_name,manual_name,output):
    official=analysis.get('official',{}).copy();official.setdefault('software_name',analysis['software_name']);official.setdefault('version',analysis['version'])
    official.setdefault('programming_languages','、'.join(inv.get('languages',{})))
    official.setdefault('source_lines',inv.get('total_nonempty_lines',None))
    official.setdefault('code_docx',code_name);official.setdefault('manual_docx',manual_name)
    if 'main_functions' not in official:official['main_functions']='；'.join(f['name']+'：'+f['description'] for f in analysis['features'])
    if 'technical_features' not in official: official['technical_features']='；'.join(analysis.get('technical_highlights',[])) or None
    lines=[f'# {analysis["software_name"]} {analysis["version"]} 软件著作权官网填报信息','',
           '> 【申报信息来源】申请人、著作权人、联系人、开发及发表日期来自本次调用消息提供的字段；技术信息来自项目源码。',
           '> 【隐私保护】证件号码、个人电话、邮箱、详细通信地址默认仅显示“请直接在官网填写”，除非调用消息明确同意写入此MD。',
           '> 本文用于对照当前版权登记系统 R11 表单逐项填报，不是官网在线申请表；部分字段仅在特定权属情形或页面版本中出现。',
           '> 【待申请人确认】项目不得猜测、不得以作者邮箱替代著作权人身份信息。敏感证件与联系信息可留空，直接在官网填写。',
           '> 技术字段依据源码分析；著作权人与发表、开发完成时间由申请人核实。最终以申请时官网实际页面为准。','']
    pending=[]
    # By default never expose identity numbers / phone / personal email / full contact address
    # in a shareable Markdown artifact, even if user supplied those in the chat prompt.
    sensitive={'holder_id_number','contact_identity','applicant_phone','applicant_email',
               'holder_phone','holder_email','contact_phone','contact_email',
               'holder_address','contact_address','holder_postcode','contact_postcode'}
    allow_sensitive=official.get('include_sensitive_in_md') is True
    for title,fields in FIELD_GROUPS:
        lines+=['## '+title,'','| 项目 | 建议填写内容 | 状态/依据 |','|---|---|---|']
        for label,key,default in fields:
            value=official.get(key,default)
            s=val(value) if (allow_sensitive or key not in sensitive) else '【请直接在官网填写：敏感信息默认不进入MD】'
            status='待人工确认' if s.startswith('【待') else ('源码分析/项目材料' if key in ('source_lines','main_functions','technical_features','programming_languages','code_docx','manual_docx','software_name','version') else '调用消息提供/需复核')
            if s.startswith('【待'):pending.append(label)
            s=s.replace('|','\\|').replace('\n','；')
            lines.append(f'| {label} | {s} | {status} |')
        lines.append('')
    lines += ['## I. 推荐可粘贴的说明文本','',
        '### 主要功能','',str(official.get('main_functions','【待完善】')),'',
        '### 技术特点','',val(official.get('technical_features')),'',
        '### 开发与运行环境说明','',f'开发环境：{val(official.get("dev_os"))}；{val(official.get("dev_software"))}。运行环境：{val(official.get("runtime_os"))}；{val(official.get("runtime_software"))}。','',
        '## J. 最终交叉检查清单','',
        '- [ ] 软件全称、简称、版本号在申请表、代码 Word、说明书 Word 中一致',
        '- [ ] 源码完整性、代码筛选连续性与源程序量填写口径已经复核',
        '- [ ] 程序鉴别 Word 每页不少于50行真实源码，且未插入伪代码或大块空白',
        '- [ ] 说明书第4章各操作均有真实功能依据，截图占位符已用同版本真实截图替换',
        '- [ ] 著作权人、开发方式、权利范围、开发完成日与发表状态经权利人确认',
        '- [ ] 所有合作开发、单位/学校开发、委托开发、历史升级等特殊权属材料按情况准备',
        '- [ ] 核对官网当日页面字段、身份认证、授权签章、上传文件类型与大小限制',
        '- [ ] 身份证、手机号、邮箱等隐私信息不上传公共仓库，不写入项目源码','',
        '## K. 无法仅凭源码确定、必须本人核实的信息','']
    lines += [f'- [ ] {x}' for x in pending]
    lines += ['','## L. 法规与填报参考（使用时再核实现行页面）','',
        '- 国家版权局《计算机软件著作权登记办法》：https://www.ncac.gov.cn/xxfb/flfg/bmgz/202410/t20241015_869486.html',
        '- 中国版权保护中心：https://www.ccopyright.com.cn/',
        '- 北京市政务服务网计算机软件著作权登记指南：https://banshi.beijing.gov.cn/pubtask/task/1/110000000000/3e283672-76be-4c8c-98e8-0bebe9bd06bf.html',
        '- 2026年版权登记诚信要求（行业协会转发）：https://www.ssia.org.cn/page65?article_id=2588',
        '- 字段名称、必填与限制可能随官网系统更新；本清单为“宁多勿漏”的人工核对清单，非对当前在线界面的实时抓取。','']
    output.write_text('\n'.join(lines),encoding='utf-8')
    return {'fields_total':sum(len(f) for _,f in FIELD_GROUPS),'pending_manual_confirm':pending}

def require_analysis(a,inv):
    req=['software_name','version','overview','features','operations']
    for k in req:
        if not a.get(k):raise ValueError(f'项目分析缺少必要字段 {k}；先基于实际源码完成核查，不得填造假内容。')
    paths={x['path'] for x in inv['files']}
    for group in ('features','operations'):
        for i,entry in enumerate(a[group],1):
            required=('name','description','evidence_paths') if group=='features' else ('title','description','steps','feedback','screenshot_prompt','evidence_paths')
            for k in required:
                if not entry.get(k):raise ValueError(f'{group}[{i}] 缺少 {k}')
            for path in entry['evidence_paths']:
                if path not in paths:raise ValueError(f'{group}[{i}] 引用了不存在的源码路径：{path}')
    if len(a['operations'])<10:print('警告：可证实的操作功能少于10项，不能为了数量编造；请核对是否遗漏源码功能。',file=sys.stderr)

def build(inv_path,analysis_path,out_dir):
    inv=json.loads(Path(inv_path).read_text(encoding='utf-8'))
    analysis=json.loads(Path(analysis_path).read_text(encoding='utf-8'))
    require_analysis(analysis,inv)
    out=Path(out_dir);out.mkdir(parents=True,exist_ok=True)
    slug=re.sub(r'[\\/:*?"<>|\n]+','_',f'{analysis["software_name"]}_{analysis["version"]}')
    code=out/(slug+'_程序鉴别材料.docx');manual=out/(slug+'_软件说明书.docx');form=out/(slug+'_官网填报信息.md')
    code_meta=make_code_doc(inv,analysis,code)
    manual_document(analysis,inv,manual)
    form_meta=make_official_md(analysis,inv,code.name,manual.name,form)
    meta={'files':{'code':str(code),'manual':str(manual),'form':str(form)},'code':code_meta,'form':form_meta,'software_name':analysis['software_name'],'version':analysis['version']}
    (out.parent/'_build_report.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'deliverables':[str(code),str(manual),str(form)],'expected_code_pages':code_meta['expected_pages'],'official_form_fields':form_meta['fields_total'],'unconfirmed_items':len(form_meta['pending_manual_confirm'])},ensure_ascii=False,indent=2))
    return meta

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--inventory',required=True);p.add_argument('--analysis',required=True);p.add_argument('--output',required=True);a=p.parse_args()
    build(a.inventory,a.analysis,a.output)

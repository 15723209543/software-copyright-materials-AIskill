# software-copyright-materials / 软件著作权申请资料生成 AI Skill

> 版本：**v1.0** · 作者：**15723209543@163.com** · 许可：**MIT**  
> GitHub 仓库规划地址：**https://github.com/15723209543/software-copyright-materials-AIskill**（须以实际创建结果为准）

## 1. Skill 简介

读取用户上传的**整个程序源码目录或 ZIP 压缩包**，结合真实功能、源码文件和目录结构，自动整理中国计算机软件著作权申请常用的三份材料。可用于前后端网站、桌面应用、嵌入式软件、算法应用等项目。必须依据当前项目本身的代码生成材料，而不是复制示例项目的功能。

## 2. 三份核心输出

| 文件 | 格式 | 重点 |
|---|---|---|
| 程序鉴别材料 | Word `.docx` | **真实源码**、所有代码宋体、每一篇代码开始时以**黑体二号**写清文件路径/文件名、每个实际页面不少于50行、没有大空白 |
| 软件说明书 | Word `.docx` | 详细写软件背景、目录、功能与实现，重点写**第4章至少10种真实操作（有多少写多少，不能编造）**，每项使用`【截图提示：……】`说明应补哪张真实图 |
| 官网填报信息 | Markdown `.md` | 面向中国版权保护中心 R11 登记的**详细字段清单**，尽量多列而不漏，自动填写可核实的技术事实；身份、权属、日期待申请人确认 |

生成过程还会建立源码清单和审查报告用于内部检查，但最终输出压缩包只有三份核心资料。

## 3. 核心功能

1. 读完整源码 ZIP、识别源码文件和主要编程语言，排除常见第三方依赖与构建目录，记录哈希、真实源码行数与相对路径。
2. 阅读源代码判断实际实现的功能，生成可追溯的代码证据目录，不虚构用户页面、API 或数据库。
3. 自动将源码编排成符合用户版式偏好的 Word，并提供实际渲染后**逐页50行及空白检查**。
4. 借鉴用户先前软件说明书的写法，重点编写第4章“软件界面与用户操作流程”，逐项放置真实截图指导占位符。
5. 官网材料覆盖软件名称、版本、分类、日期、发表情况、开发方式、权属、联系人、开发运行环境、编程语言、主要功能、技术特点、代码量、证明材料、提交后的补正信息。
6. 遇到真实资料无法确认的法律事实，明确标记【待申请人确认】，不会自动填写作者身份或虚构开发时间。

## 4. 项目目录

```text
software-copyright-materials-AIskill/
├── software-copyright-materials/
│   ├── SKILL.md
│   ├── agents/openai.yaml
│   ├── assets/
│   │   ├── output-tree.example.txt
│   │   └── project_analysis.example.json
│   ├── references/
│   │   ├── example-style.md
│   │   ├── source-identification.md
│   │   ├── manual-style.md
│   │   ├── official-form-checklist.md
│   │   └── prompt-templates.md
│   └── scripts/
│       ├── source_inventory.py
│       ├── build_materials.py
│       ├── validate_materials.py
│       ├── package_result.py
│       └── requirements.txt
├── software-copyright-materials.zip  # 网页版上传的同名内层 Skill 文件夹压缩包
├── README.md
├── LICENSE.md
├── 调用模板_Codex.md
├── 调用模板_网页版GPT.md
└── install.ps1
```

## 5. 安装方法

### 5.1 Codex（Windows PowerShell）

先克隆仓库（仓库创建后可使用）或直接解压下载的仓库压缩包：

```powershell
git clone https://github.com/15723209543/software-copyright-materials-AIskill.git
cd software-copyright-materials-AIskill
.\install.ps1
```

也可以手动将内层 `software-copyright-materials/` 文件夹复制到：

```text
%USERPROFILE%\.codex\skills\software-copyright-materials\
```

确保文件位置为 `.codex/skills/software-copyright-materials/SKILL.md`，重启或重新打开 Codex 会话。

### 5.2 Codex（macOS/Linux）

```bash
git clone https://github.com/15723209543/software-copyright-materials-AIskill.git
mkdir -p ~/.codex/skills
cp -R software-copyright-materials-AIskill/software-copyright-materials ~/.codex/skills/
python3 -m pip install -r ~/.codex/skills/software-copyright-materials/scripts/requirements.txt
```

若没有系统管理员权限，可使用 Python 虚拟环境安装依赖。

### 5.3 网页版 ChatGPT

上传外层目录中的 `software-copyright-materials.zip`，同时上传完整项目 ZIP、软件说明书参考文件（如果希望匹配既有样式）。然后复制 `调用模板_网页版GPT.md` 中的调用消息。**此方式是在当前会话内按文件内规则操作，不是永久安装 GPT Skill。**

## 6. 调用方式

**先填写调用消息里的申请信息，再交由 AI 生成。** 本 Skill 遵循与 `make-readme-and-license` 一样的填报体验：Codex 和网页版模板分别给出可复制的完整表单，包括申请人、著作权人、联系人、软件名与版本、开发方式、发表日期、权属、协议等，而不是要求用户编辑 Skill 的 `SKILL.md`。不确定的可以留空并在官网 MD 标记待确认。敏感证件号建议直接在官网填写。

Codex：

```text
$software-copyright-materials 请从本项目完整源码生成两份Word（程序鉴别材料、软件说明书）以及一份官网填报信息MD；注意每页至少50行真实代码，第4章逐项添加【截图提示】；执行分页校验后交付。
```

网页版：见 `调用模板_网页版GPT.md`（模板中包含与 Codex 一致的申请人资料输入区）。

## 7. Python 辅助脚本用法

```bash
pip install -r software-copyright-materials/scripts/requirements.txt
python software-copyright-materials/scripts/source_inventory.py --source project.zip --out work/inventory.json
# 由 AI 依据实际项目生成、核实 work/project_analysis.json
python software-copyright-materials/scripts/build_materials.py --inventory work/inventory.json --analysis work/project_analysis.json --output out/正式资料
python software-copyright-materials/scripts/validate_materials.py --report out/_build_report.json --out work/qa_report.json
python software-copyright-materials/scripts/package_result.py --report out/_build_report.json --validation work/qa_report.json --output out/软著申请资料.zip
```

需要 **Python 3.10+**、`python-docx`、`PyMuPDF`。物理页 QA 要求可以使用的 **LibreOffice**；Word 手动验证可由用户使用 MS Word 替代，但脚本目前调用 LibreOffice。`project_analysis.json` 是 AI 分析源码之后的工作事实文件，不能直接拿 `.example.json` 填充正式文书。

## 8. 注意事项与责任边界

- 不是法律代理服务，也不保证版权登记审核通过。申报前请以中国版权保护中心当次要求核对电子材料、页眉页码和签章。
- 历史书写偏好是不加页眉页脚，但部分办事指南要求页眉软件全称、版本号和页码；如冲突，优先核对和遵循当前提交渠道要求。
- 代码行数以实际自研源码非空行统计，最终官网字段的统计口径应由申请人确认。
- 无法获取的身份、首次发表、开发完成时间、合作权属信息不作虚假填写。
- 不要公开上传含有密钥、证件号码、数据库密码或私人联系方式的完整源码。

**官方参考：** [中国版权保护中心](https://www.ccopyright.com.cn/) · [《计算机软件著作权登记办法》](https://www.ncac.gov.cn/xxfb/flfg/bmgz/202410/t20241015_869486.html)

## 9. 版权许可

Copyright © 2026 15723209543@163.com. Licensed under the MIT License. See [LICENSE.md](LICENSE.md).

$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'software-copyright-materials'
$dstRoot = Join-Path $HOME '.codex\skills'
$dst = Join-Path $dstRoot 'software-copyright-materials'
if (-not (Test-Path (Join-Path $src 'SKILL.md'))) { throw '未找到内层 SKILL.md' }
New-Item -ItemType Directory -Force -Path $dstRoot | Out-Null
if (Test-Path $dst) { Write-Warning "目标 $dst 已存在，以下操作将更新同名 Skill 文件" }
Copy-Item -Path $src -Destination $dstRoot -Recurse -Force
Write-Host "已安装 $dst。重新打开 Codex 会话后，使用 `"`$software-copyright-materials`" 调用。"
Write-Host '运行脚本需要 python-docx、PyMuPDF，排版分页验证还需要 LibreOffice。'

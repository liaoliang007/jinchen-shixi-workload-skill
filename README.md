# 锦城毕业实习工作量统计

> 成都锦城学院教务处 · 按指导教师批量生成《XX指导学生实习工作量统计表》(Word + PDF) 的可复用 AI 技能

**作者：廖哥（廖亮）** ｜ 创建日期：2026-09-30 ｜ 版本 v1.0

## 功能
- 从实习明细 Excel（名单教师实习明细 sheet）按指导教师分组
- 参照教务处样例样式逐教师生成统计表：9列固定字段、每生一行折算学时、教师合计纵向合并
- 布局定稿：删除"附件4"、落款紧邻标题、页边距2/3、学年学期单行、合计列缩1/3、表头两行、实习单位单行截断
- 一键转 PDF（Word COM），支持文件占用三级兜底
- 支持批量（无参数）与单教师（传姓名参数）两种模式

## 用法
```bash
python scripts/generate_workload_tables.py            # 批量全部教师
python scripts/generate_workload_tables.py 何贤江     # 单教师
```

依赖：python-docx、openpyxl、pywin32、pypdf、pymupdf（Windows + MS Word）

## 文件结构
- `SKILL.md` — 技能说明（含踩坑记录：Word COM 只读陷阱、noWrap 失效、输出文件占用兜底）
- `scripts/generate_workload_tables.py` — 一键生成脚本

---
*署名：廖哥（廖亮）· 成都锦城学院教务处*

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据结构化候选人数据生成招聘交付物。

输入 JSON 结构（CANDIDATE_SCHEMA）：
{
  "job": {"title": "岗位名称", "date": "2026-11-17"},
  "candidates": [
    {
      "name": "候选人A",
      "boss_id": "BOSS标识",
      "current_title": "当前职位",
      "company": "当前公司",
      "years": 3,
      "education": "本科",
      "skills": ["Python", "SQL"],
      "expected_salary": "20-30K",
      "status": "已收简历",
      "scores": {
        "skill": 4, "experience": 4, "stability": 3,
        "salary": 4, "activity": 5, "culture": 4
      },
      "total_score": 4.0,
      "grade": "A",
      "strengths": "核心优势描述",
      "risks": "风险点描述",
      "recommendation": "推荐面试",
      "recommend_reason": "推荐理由",
      "resume_summary": "简历摘要",
      "pending_questions": "待核实问题",
      "notes": "备注",
      "last_contact": "2026-11-17 10:00",
      "next_followup": "2026-11-18"
    }
  ]
}

字段说明：
- scores 六个维度均为 1-5 分；total_score 与 grade 可省略，脚本会自动计算。
- 其余文本字段可省略，脚本会填「-」。
- 「status」建议使用：已打招呼 / 已回复 / 已收简历 / 待跟进 / 拒绝 / 无响应。

输出文件：
1. <岗位>_候选人评分表与台账.xlsx（三个工作表：推荐名单、评分明细、沟通台账）
2. <岗位>_简历筛选报告.md
3. <岗位>_沟通话术.md
4. <岗位>_面试问题清单.md（通用题库 + 每位推荐面试候选人的个性化追问 + 面试评分表）

若 openpyxl 不可用，Excel 自动降级为三个 CSV 文件，不影响报告生成。
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

WEIGHTS = {
    "skill": 0.25,
    "experience": 0.25,
    "stability": 0.15,
    "salary": 0.10,
    "activity": 0.15,
    "culture": 0.10,
}

SCORE_LABELS = {
    "skill": "技能匹配度",
    "experience": "经验相关性",
    "stability": "稳定性",
    "salary": "薪资匹配度",
    "activity": "活跃度与沟通意愿",
    "culture": "文化契合度",
}

GRADE_RULES = [
    (4.2, "A"),
    (3.5, "B"),
    (2.8, "C"),
    (0.0, "D"),
]


def get(c, key, default="-"):
    value = c.get(key)
    if value is None or value == "" or value == []:
        return default
    if isinstance(value, list):
        return "、".join(str(v) for v in value)
    return value


def calc_total(scores):
    total = 0.0
    for key, weight in WEIGHTS.items():
        try:
            value = float(scores.get(key, 0))
        except (TypeError, ValueError):
            value = 0.0
        total += value * weight
    return round(total, 2)


def calc_grade(total):
    for threshold, grade in GRADE_RULES:
        if total >= threshold:
            return grade
    return "D"


def prepare(data):
    candidates = data.get("candidates", [])
    for c in candidates:
        scores = c.get("scores") or {}
        if c.get("total_score") in (None, ""):
            c["total_score"] = calc_total(scores)
        else:
            try:
                c["total_score"] = round(float(c["total_score"]), 2)
            except (TypeError, ValueError):
                c["total_score"] = calc_total(scores)
        if not c.get("grade"):
            c["grade"] = calc_grade(c["total_score"])
    candidates.sort(key=lambda x: x["total_score"], reverse=True)
    return candidates


def write_excel(output_dir, job_title, candidates):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        write_csv_fallback(output_dir, job_title, candidates)
        return None, True

    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def style_sheet(ws, widths):
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = width
        ws.freeze_panes = "A2"

    # 推荐名单
    ws1 = wb.active
    ws1.title = "推荐名单"
    ws1.append(["排名", "姓名", "当前职位", "当前公司", "年限", "学历",
                "技能标签", "期望薪资", "总分", "等级", "推荐结论", "推荐理由"])
    for i, c in enumerate(candidates, start=1):
        ws1.append([
            i, get(c, "name"), get(c, "current_title"), get(c, "company"),
            get(c, "years"), get(c, "education"), get(c, "skills"),
            get(c, "expected_salary"), c["total_score"], c["grade"],
            get(c, "recommendation"), get(c, "recommend_reason"),
        ])
    style_sheet(ws1, [6, 14, 20, 20, 8, 10, 24, 14, 8, 8, 14, 40])

    # 评分明细
    ws2 = wb.create_sheet("评分明细")
    ws2.append(["姓名"] + [SCORE_LABELS[k] for k in WEIGHTS] +
               ["总分", "等级", "核心优势", "风险点", "待核实问题"])
    for c in candidates:
        scores = c.get("scores") or {}
        ws2.append([
            get(c, "name"),
            *[scores.get(k, "-") for k in WEIGHTS],
            c["total_score"], c["grade"],
            get(c, "strengths"), get(c, "risks"), get(c, "pending_questions"),
        ])
    style_sheet(ws2, [14] + [14] * len(WEIGHTS) + [8, 8, 36, 36, 36])

    # 沟通台账
    ws3 = wb.create_sheet("沟通台账")
    ws3.append(["姓名", "BOSS标识", "沟通状态", "最后沟通时间",
                "下次跟进", "备注"])
    for c in candidates:
        ws3.append([
            get(c, "name"), get(c, "boss_id"), get(c, "status"),
            get(c, "last_contact"), get(c, "next_followup"), get(c, "notes"),
        ])
    style_sheet(ws3, [14, 18, 14, 20, 14, 40])

    path = output_dir / f"{job_title}_候选人评分表与台账.xlsx"
    wb.save(path)
    return path, False


def write_csv_fallback(output_dir, job_title, candidates):
    import csv

    def dump(filename, header, rows):
        path = output_dir / filename
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerows(rows)
        return path

    dump(f"{job_title}_推荐名单.csv",
         ["排名", "姓名", "当前职位", "当前公司", "年限", "学历", "技能标签",
          "期望薪资", "总分", "等级", "推荐结论", "推荐理由"],
         [[i, get(c, "name"), get(c, "current_title"), get(c, "company"),
           get(c, "years"), get(c, "education"), get(c, "skills"),
           get(c, "expected_salary"), c["total_score"], c["grade"],
           get(c, "recommendation"), get(c, "recommend_reason")]
          for i, c in enumerate(candidates, start=1)])

    dump(f"{job_title}_评分明细.csv",
         ["姓名"] + [SCORE_LABELS[k] for k in WEIGHTS] +
         ["总分", "等级", "核心优势", "风险点", "待核实问题"],
         [[get(c, "name"), *[(c.get("scores") or {}).get(k, "-") for k in WEIGHTS],
           c["total_score"], c["grade"], get(c, "strengths"),
           get(c, "risks"), get(c, "pending_questions")] for c in candidates])

    dump(f"{job_title}_沟通台账.csv",
         ["姓名", "BOSS标识", "沟通状态", "最后沟通时间", "下次跟进", "备注"],
         [[get(c, "name"), get(c, "boss_id"), get(c, "status"),
           get(c, "last_contact"), get(c, "next_followup"), get(c, "notes")]
          for c in candidates])


def write_report(output_dir, job_title, candidates, date_str):
    counts = {g: sum(1 for c in candidates if c["grade"] == g) for g in ("A", "B", "C", "D")}
    lines = [
        f"# {job_title} — 简历筛选报告",
        "",
        f"- 生成日期：{date_str}",
        f"- 候选人总数：{len(candidates)}",
        f"- 等级分布：A 级 {counts['A']} 人 / B 级 {counts['B']} 人 / "
        f"C 级 {counts['C']} 人 / D 级 {counts['D']} 人",
        f"- 已收到简历：{sum(1 for c in candidates if c.get('status') == '已收简历')} 人",
        "",
        "## 推荐名单（按匹配分排序）",
        "",
        "| 排名 | 姓名 | 总分 | 等级 | 推荐结论 | 推荐理由 |",
        "|------|------|------|------|----------|----------|",
    ]
    for i, c in enumerate(candidates, start=1):
        lines.append(
            f"| {i} | {get(c, 'name')} | {c['total_score']} | {c['grade']} | "
            f"{get(c, 'recommendation')} | {get(c, 'recommend_reason')} |"
        )

    lines.extend(["", "## 候选人详情", ""])
    for c in candidates:
        lines.extend([
            f"### {get(c, 'name')}（{c['grade']} 级 / {c['total_score']} 分）",
            "",
            f"- 当前职位：{get(c, 'current_title')} @ {get(c, 'company')}",
            f"- 工作年限：{get(c, 'years')} 年｜学历：{get(c, 'education')}",
            f"- 技能标签：{get(c, 'skills')}",
            f"- 期望薪资：{get(c, 'expected_salary')}｜沟通状态：{get(c, 'status')}",
            f"- 核心优势：{get(c, 'strengths')}",
            f"- 风险点：{get(c, 'risks')}",
            f"- 待核实问题：{get(c, 'pending_questions')}",
            f"- 简历摘要：{get(c, 'resume_summary')}",
            "",
        ])

    path = output_dir / f"{job_title}_简历筛选报告.md"
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def write_interview_guide(output_dir, job_title, skill_root, candidates):
    """生成面试问题清单：通用题库 + 每个推荐面试候选人的个性化追问。"""
    source = skill_root / "references" / "interview-questions.md"
    if source.exists():
        content = source.read_text(encoding="utf-8")
    else:
        content = f"# {job_title} 面试问题清单\n\n（题库文件缺失，请检查 skill 安装完整性。）\n"

    interviewees = [c for c in candidates
                    if c.get("grade") in ("A", "B")
                    and "推荐" in str(c.get("recommendation", ""))]
    if not interviewees:
        interviewees = [c for c in candidates if c.get("grade") in ("A", "B")]

    content += f"\n\n---\n\n# {job_title} — 候选人专属面试追问\n\n"
    content += f"本批次进入面试的候选人共 {len(interviewees)} 人。以下问题由脚本根据每位候选人的简历疑点与优势自动生成，面试时优先提问。\n"

    for i, c in enumerate(interviewees, start=1):
        content += f"\n## {i}. {get(c, 'name')}（{c['grade']} 级 / {c['total_score']} 分）\n\n"
        content += f"- 当前职位：{get(c, 'current_title')} @ {get(c, 'company')}\n"
        content += f"- 推荐理由：{get(c, 'recommend_reason')}\n\n"
        content += "**必问的个性化追问：**\n\n"

        pending = c.get("pending_questions")
        risks = c.get("risks")
        strengths = c.get("strengths")
        count = 0
        if pending not in (None, "", "-"):
            count += 1
            content += f"{count}. 【核实简历疑点】{pending}——请候选人当场说明，回答与上一步沟通或简历不一致的，标记为存疑。\n"
        if risks not in (None, "", "-"):
            count += 1
            content += f"{count}. 【验证风险点】针对「{risks}」，请举例说明并追问细节，确认风险是否属实、是否可接受。\n"
        if strengths not in (None, "", "-"):
            count += 1
            content += f"{count}. 【深挖优势】简历显示「{strengths}」，请候选人展开讲具体案例：本人角色、做法、可量化的结果。\n"
        # 销售岗通用必问，结合候选人情况
        count += 1
        content += f"{count}. 【业绩验证】请给出上一份销售工作近三个月的实际业绩数字与团队排名，并提供可核实的口径。\n"
        count += 1
        content += f"{count}. 【稳定性】请说明近几段工作经历的起止时间与离职原因，重点解释与简历不一致或时间重叠的部分。\n"
        count += 1
        content += f"{count}. 【行业经验（硬性门槛）】本岗要求同行业经验 1 年以上，请说明你在财税/代账/税务代理行业的具体从业时长与岗位。\n"
        count += 1
        content += f"{count}. 【薪资与到岗】期望薪资是多少？对本岗「无责任底薪 + 高提成 + 五险 + 奖金、综合月薪 8-13K+」是否接受？最快何时到岗？\n"

        content += "\n**面试记录**（现场填写）：\n\n"
        content += "- 回答要点：__________\n- 判断结论：□ 推荐录用　□ 待定/复核　□ 不录用\n- 理由：__________\n"

    path = output_dir / f"{job_title}_面试问题清单.md"
    path.write_text(content, encoding="utf-8")
    return path


def write_templates(output_dir, job_title, skill_root, data):
    source = skill_root / "references" / "message-templates.md"
    if source.exists():
        content = source.read_text(encoding="utf-8")
    else:
        content = "# 沟通话术\n\n（话术库文件缺失，请检查 skill 安装完整性。）\n"

    custom = data.get("templates")
    if custom:
        content += "\n\n## 本次定制话术\n\n"
        for item in custom:
            content += f"- {item}\n"

    path = output_dir / f"{job_title}_沟通话术.md"
    path.write_text(content, encoding="utf-8")
    return path


def main():
    parser = argparse.ArgumentParser(description="生成候选人评分表、筛选报告与沟通话术")
    parser.add_argument("--input", required=True, help="候选人 JSON 文件路径")
    parser.add_argument("--output-dir", required=True, help="输出目录")
    parser.add_argument("--job", default="", help="岗位名称，覆盖 JSON 中的 job.title")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：输入文件不存在：{input_path}", file=sys.stderr)
        sys.exit(1)

    try:
        data = json.loads(input_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        print(f"错误：JSON 解析失败：{e}", file=sys.stderr)
        sys.exit(1)

    job_title = args.job or (data.get("job", {}).get("title") or "招聘")
    job_title = "".join(ch for ch in job_title if ch not in '\\/:*?"<>|').strip() or "招聘"
    date_str = data.get("job", {}).get("date") or datetime.now().strftime("%Y-%m-%d")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    candidates = prepare(data)
    excel_path, fell_back = write_excel(output_dir, job_title, candidates)
    report_path = write_report(output_dir, job_title, candidates, date_str)
    skill_root = Path(__file__).resolve().parent.parent
    template_path = write_templates(output_dir, job_title, skill_root, data)
    interview_path = write_interview_guide(output_dir, job_title, skill_root, candidates)

    print(f"岗位：{job_title}")
    print(f"候选人数量：{len(candidates)}")
    if fell_back:
        print("提示：openpyxl 不可用，已降级输出 CSV 文件")
    else:
        print(f"评分表与台账：{excel_path}")
    print(f"筛选报告：{report_path}")
    print(f"沟通话术：{template_path}")
    print(f"面试问题清单：{interview_path}")


if __name__ == "__main__":
    main()

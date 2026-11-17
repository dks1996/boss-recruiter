# boss-recruiter（BOSS直聘招聘助手）

一个 WorkBuddy Skill，把「岗位要求 → 候选人搜索 → 匹配评分 → 自动沟通 → 简历筛选 → 面试问题 → 交付结果」串成完整招聘流水线。支持在 BOSS直聘上按岗位要求找人、沟通索简历、筛选评估，并输出五类交付物。

> **隐私提示**：`references/company-profile.md` 中包含企业真实信息（公司名、地址、联系人、薪资结构等）。公开仓库前请先检查并脱敏，或保持仓库为 Private。

## 目录结构

```
boss-recruiter/
├── SKILL.md                          # 技能说明与完整工作流
├── LESSONS.md                        # ⚡ 实战经验与提速速查（下次先读这个）
├── references/
│   ├── company-profile.md            # 公司与岗位固定配置（含真实企业信息）
│   ├── screening-framework.md        # 候选人评分模型与 A/B/C/D 分级
│   ├── browser-workflow.md           # BOSS直聘浏览器操作 SOP 与风控红线（含实战提速要点）
│   ├── message-templates.md          # 沟通话术库
│   └── interview-questions.md        # 面试题库与评分表
└── scripts/
    └── build_shortlist.py            # 生成 Excel 台账 + 各类 Markdown 交付物
```

## 工作流概览

1. 加载公司/岗位配置，确认浏览器控制能力可用
2. 在 BOSS直聘搜索候选人并采集信息
3. 按加权评分模型匹配分级（A/B/C/D）
4. 对 A/B 级候选人自动打招呼、追问、索要简历
5. 收到简历后深度评估，回填评分与推荐结论
6. 生成五类交付物：评分表与台账、筛选报告、沟通话术、沟通台账、面试问题清单

## ⚡ 实战提速

实跑复盘的经验与坑点集中在 [`LESSONS.md`](LESSONS.md)，**下次执行前先读**。最关键两条：

- **浏览器只用官方内置**（`mcp__browser__*`）：第三方 Playwright 引擎会被 BOSS 识别反爬并强制跳转。
- **主动联系只走「推荐牛人」页**（免费）：「搜索」页的「联系Ta」是 ¥388 付费墙，仅用于采集/评分。

## 使用方式

在 WorkBuddy 中直接描述招聘需求即可触发，例如：

> 用 boss-recruiter 帮我招财税销售顾问，招 5 人

## 合规说明

- 遵守 BOSS直聘平台规则，控制自动化频率，遇验证码/风控立即停止
- 不索取或存储账号密码，候选人隐私数据仅用于招聘
- 最终录用决定由用户负责
- 依赖浏览器控制能力（Browser Use 插件 / workbuddy-browser-use）
- 交付脚本基于 Python 3，Excel 输出依赖 `openpyxl`（缺失时自动降级为 CSV）

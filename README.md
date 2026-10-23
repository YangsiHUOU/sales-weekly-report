# 销售周报 Skill

将每周 Excel、员工分享原稿与照片、会议或活动材料，整理成可发公司群的红色销售周报长图。默认适配亚太财险深圳分公司的已确认版式。

## 安装

把下面这句话复制到 Codex：

```text
请使用 skill-installer 安装这个销售周报技能：https://github.com/YangsiHUOU/sales-weekly-report/tree/main/skills/sales-weekly-report
```

安装后，在下一轮消息中使用 `$sales-weekly-report`，或直接说“我要制作本周销售周报”。

[下载技能安装包](https://github.com/YangsiHUOU/sales-weekly-report/releases/latest/download/sales-weekly-report.zip) · [查看技能文件](https://github.com/YangsiHUOU/sales-weekly-report/tree/main/skills/sales-weekly-report) · [下载空白 Excel 模板](https://github.com/YangsiHUOU/sales-weekly-report/releases/latest/download/weekly-input-template.xlsx)

如果手动安装，将 ZIP 中的 `sales-weekly-report` 文件夹放入 Codex 的个人技能目录：默认 macOS/Linux 为 `~/.codex/skills/`，Windows 为 `%USERPROFILE%\.codex\skills\`；设置过 `CODEX_HOME` 时使用其 `skills` 子目录。已安装同名技能时，先备份再替换。

## 每周使用

1. 填写空白 Excel，金额单位为万元，统计日期填写本周实际起止日期。
2. 把 Excel 发给 Codex，附上员工分享及人物照片（没有时说明取消该栏目）。
3. 提供公司动态的通知、纪要或事件说明，以及需要使用的现场照片。
4. Codex 检索统计周内发布的三条深圳保险资讯，润色文字并生成 PNG 长图。

可直接使用这段提示：

```text
使用 $sales-weekly-report 制作本周销售周报。Excel 是本周数据，附件是公司动态材料和现场照片。本周没有一线有话说，请取消该栏目。行业动态请自行检索统计周内正式发布的深圳保险信息，检查标题换行、正文留白和照片比例后，交付 PNG 长图。
```

## 包含的内容

- 行业动态：三条本周深圳保险资讯，核验发布日期与来源。
- 业绩通报：本周业绩走势，以及团队、机构、个人 TOP3。
- 一线有话说：原稿润色、人物照片叠加卡，可按周取消。
- 公司动态：会议、培训、公益活动及荣誉，区分主办与参加。
- 红色长图生成脚本、空白 Excel、内容和排版规范。

宽度为 1080 像素，高度随栏目及内容调整。正文首行缩进两字，段落间不空行；现场照片按比例展示，排版超过可用行数时报错，避免文字被静默截断。

## 运行环境

建议在带工作区依赖和联网检索能力的 Codex 环境中使用。生成脚本需要 Node.js、`@oai/artifact-tool` 与 `sharp`，由当前 Codex 的工作区运行时提供。这个 ZIP 是技能资源包，不是可双击运行的独立程序。没有相应依赖的环境需要先提供兼容运行时。

仓库只包含技能和空白填报模板，不包含历史销售数据、员工照片、现场照片或已生成的周报。

---
name: exchange-ews
description: OpenClaw 的 Exchange EWS 双层技能中枢：由 agent 负责意图理解、策略决策与流程编排，由 EWS 原语层负责可审计的邮箱动作执行（收发、检索、移动、标记等），并通过明确边界实现安全、可回放、可扩展的自动化。
---

# Exchange EWS

此技能采用两层结构：

- 决策层：OpenClaw agent（意图理解、策略、流程编排）
- 执行层：`scripts/ews_cli.py`（只提供基础邮箱原语）

## 核心原则

- OpenClaw agent 是唯一决策者：负责意图理解、策略选择、流程编排、风险控制。
- Exchange 能力层是被动执行面：只接受 agent 下发的结构化动作。
- 所有动作必须可审计、可回放、可回滚（由上层 agent 统一控制）。

## 准备环境

```bash
pip3 install -r requirements.txt
```

推荐使用本地 `.env`（不会提交到 Git）：

```bash
cp .env.example .env
# 然后编辑 .env，填入真实 EWS_* 配置
```

`scripts/ews_cli.py` 会在启动时自动读取当前目录或技能目录下的 `.env`（仅在对应变量尚未存在时注入，不覆盖系统环境变量）。

也可以继续使用 shell 环境变量：

```bash
export EWS_ENDPOINT="https://mail.example.com/EWS/Exchange.asmx"
export EWS_EMAIL="user@example.com"
export EWS_USERNAME="user_or_DOMAIN\\user"
export EWS_AUTH_TYPE="NTLM"
export EWS_PASSWORD_CMD="your-secret-command"
```

说明：
- 非敏感配置（端点、邮箱、用户名、认证类型）放系统环境变量。
- 密码优先从 `EWS_PASSWORD` 读取；若未设置，则执行 `EWS_PASSWORD_CMD` 获取。
- `EWS_PASSWORD_CMD` 需要输出纯密码到 stdout（尾部换行可有可无）。
- 可按平台接入系统密钥管理：
  - macOS Keychain: `security find-generic-password -a "$USER" -s openclaw_ews_password -w`
  - Linux Secret Service: `secret-tool lookup service openclaw account "$USER"`
  - Windows SecretManagement: `pwsh -NoProfile -Command "(Get-Secret -Name openclaw_ews_password -AsPlainText)"`

分平台落地步骤：

```bash
# macOS: 首次写入/更新密码
security add-generic-password -a "$USER" -s openclaw_ews_password -w 'your_ews_password' -U

# macOS: 在 shell 配置中设置（~/.zshrc / ~/.bashrc）
export EWS_PASSWORD_CMD='security find-generic-password -a "$USER" -s openclaw_ews_password -w'
```

```bash
# Linux (Secret Service): 写入密码
printf '%s' 'your_ews_password' | secret-tool store --label='OpenClaw EWS Password' service openclaw account "$USER"

# Linux: 在 shell 配置中设置（~/.bashrc / ~/.zshrc）
export EWS_PASSWORD_CMD='secret-tool lookup service openclaw account "$USER"'
```

```powershell
# Windows (PowerShell SecretManagement): 首次安装模块
Install-Module Microsoft.PowerShell.SecretManagement,Microsoft.PowerShell.SecretStore -Scope CurrentUser

# Windows: 写入密码
Set-Secret -Name openclaw_ews_password -Secret "your_ews_password"

# Windows: 为当前会话设置命令（持久化请写入 PowerShell profile）
$env:EWS_PASSWORD_CMD='pwsh -NoProfile -Command "(Get-Secret -Name openclaw_ews_password -AsPlainText)"'
```

## 基础检查

```bash
python3 scripts/ews_cli.py healthcheck
```

查看未读邮件（快捷命令）：

```bash
python3 scripts/ews_cli.py unread --folder "Inbox"
```

性能建议：所有需要先拉取邮件再本地筛选的操作默认都使用 `10`（含 `limit/scan-limit`），并优先走服务端过滤（如未读/附件）后本地兜底；只有用户明确要求查看更多时再增大参数。

可选包装器（直接使用系统环境变量，如 `~/.zshrc` 中的 `EWS_*`）：

```bash
cd ~/.openclaw/workspace
./scripts/run_ews_command.sh healthcheck
```

## 原语能力清单

```bash
python3 scripts/ews_cli.py agent capabilities
```

## 统一执行入口

```bash
python3 scripts/ews_cli.py agent op --name "<primitive_op>" [args...] [--dry-run|--apply]
```

支持能力：

- 写与发：
  `compose.new` `compose.get` `compose.list` `compose.delete`
  `compose.subject.set` `compose.body.set`
  `compose.recipients.add_to` `compose.recipients.add_cc` `compose.recipients.add_bcc`
  `compose.recipients.clear_to` `compose.recipients.clear_cc` `compose.recipients.clear_bcc`
  `compose.attachments.add` `compose.attachments.remove` `compose.attachments.list`
  `send.now` `send.later` `send.schedule.list` `send.schedule.cancel` `send.schedule.flush_due`
- 收与看：
  `mailbox.inbox.list` `mailbox.message.list_unread` `mailbox.message.get` `conversation.view`
  `attachment.list` `attachment.meta` `attachment.open` `attachment.preview` `attachment.download`
- 回与转：
  `message.quote` `message.reply` `message.reply_with_quote` `message.reply_all` `message.forward` `message.template_reply`
- 管与整：
  `message.mark_read` `message.mark_unread`
  `message.star` `message.unstar`
  `message.move` `message.archive` `message.delete`
  `message.tag` `message.tags.add` `message.tags.remove` `message.tags.clear`
- 找与筛：
  `mail.search` `mail.filter` `mail.sort`
- 防打扰：
  `sender.block` `sender.blocked.list` `sender.block.remove`
  `message.mark_spam`
  `message.unsubscribe` `message.unsubscribe.list` `message.unsubscribe.remove`
  `contact.priority_alert` `contact.priority_alert.get` `contact.priority_alert.remove`
- 自动化：
  `rule.set` `rule.from_text` `rule.get` `rule.list` `rule.delete` `automation.run`
  `autoreply.set` `autoreply.get` `autoreply.disable`

## 执行安全

- 变更类操作默认 `--dry-run`
- 只有显式传入 `--apply` 才会执行写操作
- 任意 channel 场景中，发送带附件邮件时默认只传附件路径（`--attach /abs/path/to/file`），不要把附件全文贴入会话。

## 多 channel 大附件防 token 超限（必读）

- 目标：避免 `Total tokens of image and text exceed max message tokens`。
- 默认策略：`发送附件 != 解析附件内容`，优先直接走 `send.now` + `--attach`。
- 对于 `.wsdl`、`.xml`、日志、代码等大文本附件：禁止在同一条消息里粘贴全文。
- 只有用户明确要求“分析附件内容”时，才做分块读取（建议每块 2k-4k tokens）并分轮处理。
- 若仅用于邮件转发，推荐最短指令模板：  
  `用 exchange-ews 发邮件到 <收件人>，主题 <主题>，正文 <正文>，附件 <绝对路径>，不要解析附件内容`

发送示例（直接发，含附件）：

```bash
python3 scripts/ews_cli.py agent op --name send.now \
  --to "recipient@example.com" \
  --subject "测试邮件（含附件）" \
  --body "附件已附上。" \
  --attach "/abs/path/to/file" \
  --apply
```

## 常见排查

- 连接失败：先跑 `healthcheck`；若失败通常是端点不可达、代理/防火墙限制或证书链问题。
- 权限失败：检查 `EWS_USERNAME` 格式（如 `DOMAIN\\username`）和密码类型（普通密码/应用密码）。
- 结果为空：确认邮箱文件夹是否正确，可先执行 `mailbox.inbox.list` 做最小验证。

## 调度与本地状态说明

- `send.later` 只写入本地 `scheduled_outbox` 队列，不会自动后台发送。
- 需要显式执行 `send.schedule.flush_due` 才会把到期任务真正发送到 Exchange。
- `sender.block` / `message.unsubscribe` / `contact.priority_alert` / `rule.*` / `autoreply.*` 为本地状态管理能力，默认写入 `EWS_AGENT_STATE_FILE` 指向的状态文件。

## 本地自动化（Agent 读信后动作）

- `automation.run` 会读取本地 `rules`，扫描邮箱消息后按规则执行动作（可 `--dry-run` 预演）。
- 可用 `rule.from_text` 用自然语言创建规则（无需手写 JSON）。
- 示例：
  `python3 scripts/ews_cli.py agent op --name rule.from_text --rule-id urgent_cn --rule-text "发件人包含 boss@example.com 且主题包含 紧急 时，标记已读并添加标签 urgent，然后移动到 Important" --apply`
- 规则示例（`rule.set --rule-id urgent --rule-json '<json>'`）：
  - `enabled`: 是否启用
  - `conditions`: `sender_contains` / `subject_contains` / `body_contains` / `has_attachments` / `is_unread`
  - `actions`: 支持 `mark_read` `mark_unread` `star` `unstar` `move` `archive` `mark_spam` `delete` `tag_set` `tag_add` `tag_remove` `reply_template` `forward`
- `automation.run` 复用 `agent op` 公共筛选参数（如 `--folder` `--scan-limit` `--limit` `--keyword` `--sender-filter` `--has-attachments`）。

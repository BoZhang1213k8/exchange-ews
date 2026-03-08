# exchange-ews

OpenClaw 的 Exchange EWS 双层技能：

- 决策层：由 OpenClaw agent 负责意图理解、策略选择、流程编排与风险控制
- 执行层：由 `scripts/ews_cli.py` 提供可审计的邮箱原语执行能力

设计目标：安全、可回放、可扩展、职责边界清晰。

## 这个 Skill 能做什么

`exchange-ews` 是一个面向 Exchange 邮箱自动化的执行技能，核心价值是把“邮件操作能力”标准化为可编排、可审计的原语接口，供上层 agent 组合成完整业务流程。它可以：

- 提供统一的邮箱操作能力：收发邮件、草稿处理、检索过滤、状态变更与目录操作
- 将能力封装为稳定执行层：上层只关注“做什么”，执行层负责“如何调用 EWS”
- 输出结构化 JSON 结果：便于在多步骤任务中做状态传递、错误处理与审计留痕
- 支持安全优先的凭据接入：通过环境变量和 `EWS_PASSWORD_CMD` 对接系统密钥管理
- 作为 OpenClaw 决策层的基础设施：适合构建通知分发、邮件巡检、规则化处理等自动化流程

## 功能概览

- 邮件收发与草稿处理
- 邮件检索、过滤、排序
- 已读/未读、移动等状态与目录操作
- 统一 JSON 输出，便于上层 agent 编排与审计

## 快速开始

### 1) 安装依赖

```bash
pip3 install -r requirements.txt
```

### 2) 配置环境变量

推荐使用本地 `.env`（不会提交到仓库）：

```bash
cp .env.example .env
# 编辑 .env，填入真实 EWS_* 参数
```

`scripts/ews_cli.py` 会在启动时自动读取 `.env`（当前目录优先；仅注入未设置变量，不覆盖已有系统环境变量）。

也可以直接使用 export：

```bash
export EWS_ENDPOINT="https://mail.example.com/EWS/Exchange.asmx"
export EWS_EMAIL="user@example.com"
export EWS_USERNAME="user_or_DOMAIN\\user"
export EWS_AUTH_TYPE="NTLM"
export EWS_PASSWORD_CMD="your-secret-command"
```

### 3) 健康检查

```bash
python3 scripts/ews_cli.py healthcheck
```

### 4) 查看未读邮件

```bash
python3 scripts/ews_cli.py unread --folder "Inbox"
```

性能建议：凡是需要先拉取邮件再本地筛选的操作，默认都按 `10` 封执行（包括 `unread` 及 `agent op` 下的收件箱列表/检索类操作）。只有在确实要查看更多时再调大 `--limit` / `--scan-limit`。

支持常用筛选与分页参数：`--keyword`、`--sender-filter`、`--has-attachments`、`--offset`、`--scan-limit`、`--sort-by`、`--sort-asc/--sort-desc`。

## 凭据与安全

请勿将真实凭据写入仓库。

- 不要提交真实 `EWS_ENDPOINT`、`EWS_EMAIL`、`EWS_USERNAME`、`EWS_PASSWORD`
- 优先使用 `EWS_PASSWORD_CMD` + 系统密钥管理，不要常态化使用 `--password` 明文参数
- 不要提交导出的邮件内容、附件、日志和临时文件
- 默认把邮件数据视为敏感信息处理

推荐发布前扫描：

```bash
gitleaks detect --no-git -s .
trufflehog filesystem .
```

## 使用边界

- 本项目仅用于你有合法授权的 Exchange 邮箱与组织环境
- 使用者需自行遵守所在组织安全策略与法律法规
- 项目不对误用、越权访问或数据泄露承担责任

## 项目结构

```text
exchange-ews/
├── SKILL.md
├── requirements.txt
├── _meta.json
└── scripts/
    └── ews_cli.py
```

## 许可证

MIT

# exchange-ews

OpenClaw 的 Exchange EWS 双层技能：

- 决策层：由 OpenClaw agent 负责意图理解、策略选择、流程编排与风险控制
- 执行层：由 `scripts/ews_cli.py` 提供可审计的邮箱原语执行能力

设计目标：安全、可回放、可扩展、职责边界清晰。

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

#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import re
import subprocess
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from exchangelib import Account, BASIC, DELEGATE, NTLM, Configuration, Credentials, FileAttachment, Mailbox, Message

AUTH_MAP = {"NTLM": NTLM, "BASIC": BASIC, "AUTO": None}
MUTATING_OPS = {
    "compose.new",
    "compose.create",
    "compose.delete",
    "compose.subject.set",
    "compose.body.set",
    "compose.add_recipients",
    "compose.recipients.add_to",
    "compose.recipients.add_cc",
    "compose.recipients.add_bcc",
    "compose.recipients.clear_to",
    "compose.recipients.clear_cc",
    "compose.recipients.clear_bcc",
    "compose.add_attachment",
    "compose.attachments.add",
    "compose.attachments.remove",
    "message.reply_with_quote",
    "send.now",
    "send.later",
    "send.schedule.cancel",
    "send.schedule.flush_due",
    "message.reply",
    "message.reply_all",
    "message.forward",
    "message.template_reply",
    "message.mark_read",
    "message.mark_unread",
    "message.star",
    "message.move",
    "message.tag",
    "message.tags.add",
    "message.tags.remove",
    "message.tags.clear",
    "message.archive",
    "message.delete",
    "message.mark_spam",
    "sender.block",
    "message.unsubscribe",
    "contact.priority_alert",
    "rule.set",
    "rule.from_text",
    "rule.delete",
    "autoreply.set",
    "autoreply.disable",
    "automation.run",
}
LOCAL_ONLY_OPS = {
    "compose.new",
    "compose.create",
    "compose.delete",
    "compose.subject.set",
    "compose.body.set",
    "compose.add_recipients",
    "compose.recipients.add_to",
    "compose.recipients.add_cc",
    "compose.recipients.add_bcc",
    "compose.recipients.clear_to",
    "compose.recipients.clear_cc",
    "compose.recipients.clear_bcc",
    "compose.add_attachment",
    "compose.attachments.add",
    "compose.attachments.remove",
    "compose.attachments.list",
    "compose.get",
    "compose.list",
    "send.later",
    "send.schedule.list",
    "send.schedule.cancel",
    "send.schedule.flush_due",
    "sender.block",
    "sender.blocked.list",
    "sender.block.remove",
    "message.unsubscribe",
    "message.unsubscribe.list",
    "message.unsubscribe.remove",
    "contact.priority_alert",
    "contact.priority_alert.get",
    "contact.priority_alert.remove",
    "rule.set",
    "rule.from_text",
    "rule.get",
    "rule.list",
    "rule.delete",
    "autoreply.set",
    "autoreply.get",
    "autoreply.disable",
}


def _strip_wrapping_quotes(value: str) -> str:
    if len(value) >= 2 and ((value[0] == value[-1] == '"') or (value[0] == value[-1] == "'")):
        return value[1:-1]
    return value


def load_dotenv_files() -> None:
    candidates = [
        Path.cwd() / ".env",
        Path(__file__).resolve().parents[1] / ".env",
    ]
    seen: set[str] = set()
    for env_path in candidates:
        resolved = str(env_path.resolve())
        if resolved in seen or not env_path.exists() or not env_path.is_file():
            continue
        seen.add(resolved)
        try:
            for raw_line in env_path.read_text(encoding="utf-8").splitlines():
                line = raw_line.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("export "):
                    line = line[len("export ") :].strip()
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                if not key or key in os.environ:
                    continue
                os.environ[key] = _strip_wrapping_quotes(value.strip())
        except Exception:
            # .env is best-effort local convenience; ignore parse/read failures.
            continue


@dataclass
class Envelope:
    status: str
    data: Dict[str, Any]
    error: Optional[Dict[str, Any]]
    meta: Dict[str, Any]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def parse_csv(raw: str) -> List[str]:
    return [x.strip() for x in (raw or "").split(",") if x.strip()]


def parse_recipients(raw: str, required: bool = True) -> List[Mailbox]:
    values = [Mailbox(email_address=x) for x in parse_csv(raw)]
    if required and not values:
        raise ValueError("No valid recipients")
    return values


def parse_datetime_utc(raw: str, field_name: str) -> datetime:
    text = (raw or "").strip()
    if not text:
        raise ValueError(f"--{field_name} required")
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError as exc:
        raise ValueError(f"Invalid --{field_name}. Use ISO8601, e.g. 2026-03-08T09:30:00+08:00") from exc
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    else:
        parsed = parsed.astimezone(timezone.utc)
    return parsed


def emit(args, status: str, data: Dict[str, Any], error: Optional[Dict[str, Any]], command: str, started: float) -> int:
    payload = Envelope(
        status=status,
        data=data,
        error=error,
        meta={
            "command": command,
            "timestamp": now_iso(),
            "duration_ms": int((time.time() - started) * 1000),
            "dry_run": bool(getattr(args, "dry_run", False)),
        },
    )
    print(json.dumps(payload.__dict__, ensure_ascii=False, indent=2))
    return 0 if status != "error" else 1


def resolve_password(args) -> str:
    if args.password:
        return args.password
    env_password = os.getenv("EWS_PASSWORD")
    if env_password:
        return env_password
    cmd = args.password_cmd or os.getenv("EWS_PASSWORD_CMD")
    if not cmd:
        return ""
    proc = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if proc.returncode != 0:
        message = (proc.stderr or proc.stdout or "").strip() or f"exit code {proc.returncode}"
        raise ValueError(f"EWS_PASSWORD_CMD failed: {message}")
    return (proc.stdout or "").strip()


def build_account(args) -> Account:
    endpoint = args.endpoint or os.getenv("EWS_ENDPOINT")
    email = args.email or os.getenv("EWS_EMAIL")
    username = args.username or os.getenv("EWS_USERNAME")
    password = resolve_password(args)
    auth_name = (args.auth or os.getenv("EWS_AUTH_TYPE") or "NTLM").upper()
    if not all([endpoint, email, username, password]):
        raise ValueError("Missing EWS credentials/env. Require endpoint/email/username/password (or EWS_PASSWORD_CMD)")
    config = Configuration(
        service_endpoint=endpoint,
        credentials=Credentials(username=username, password=password),
        auth_type=AUTH_MAP.get(auth_name, NTLM),
    )
    return Account(primary_smtp_address=email, config=config, autodiscover=False, access_type=DELEGATE)


def resolve_folder(account: Account, name: str):
    if not name or name.lower() == "inbox":
        return account.inbox
    for folder in account.root.walk():
        if folder.name == name:
            return folder
    raise ValueError(f"Folder not found: {name}")


def serialize_message(item, include_body: bool = False, include_attachments: bool = False) -> Dict[str, Any]:
    data = {
        "id": str(getattr(item, "id", "")),
        "subject": str(getattr(item, "subject", "") or ""),
        "from": (item.sender.email_address if getattr(item, "sender", None) else ""),
        "to": [x.email_address for x in (getattr(item, "to_recipients", None) or [])],
        "datetime_received": item.datetime_received.isoformat() if getattr(item, "datetime_received", None) else "",
        "is_read": bool(getattr(item, "is_read", False)),
        "has_attachments": bool(getattr(item, "has_attachments", False)),
        "conversation_id": str(getattr(getattr(item, "conversation_id", None), "id", "") or ""),
    }
    if include_body:
        data["text_body"] = str(getattr(item, "text_body", "") or "")
        data["body"] = str(getattr(item, "body", "") or "")
    if include_attachments:
        data["attachments"] = [
            {"name": getattr(a, "name", ""), "content_type": getattr(a, "content_type", ""), "size": getattr(a, "size", 0)}
            for a in (getattr(item, "attachments", None) or [])
            if isinstance(a, FileAttachment)
        ]
    return data


def state_path(args) -> Path:
    raw = args.agent_state_file or os.getenv("EWS_AGENT_STATE_FILE") or "outputs/personal/agent_mailbox_state.json"
    return Path(raw)


def load_state(args) -> Dict[str, Any]:
    default = {
        "drafts": {},
        "scheduled_outbox": [],
        "blocked_senders": [],
        "unsubscribe_list": [],
        "priority_contact_alerts": {},
        "rules": {},
        "autoreply": {"enabled": False, "start": "", "end": "", "message": ""},
        "automation_history": [],
    }
    p = state_path(args)
    if not p.exists():
        return default
    try:
        payload = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(payload, dict):
            default.update(payload)
    except Exception:
        pass
    return default


def save_state(args, payload: Dict[str, Any]) -> str:
    p = state_path(args)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(p)


def send_message(args, account: Account, subject: str, body: str, to_raw: str, cc_raw: str = "", bcc_raw: str = "", attach_raw: str = "") -> Dict[str, Any]:
    to = parse_recipients(to_raw, required=True)
    cc = parse_recipients(cc_raw, required=False)
    bcc = parse_recipients(bcc_raw, required=False)
    attachments = parse_csv(attach_raw)
    if args.dry_run:
        return {"status": "dry_run", "subject": subject, "to": [x.email_address for x in to], "cc": [x.email_address for x in cc], "bcc": [x.email_address for x in bcc], "attachments": attachments}
    msg = Message(account=account, folder=account.sent, subject=subject, body=body, to_recipients=to, cc_recipients=cc, bcc_recipients=bcc)
    for path in attachments:
        p = Path(path)
        msg.attach(FileAttachment(name=p.name, content=p.read_bytes()))
    msg.send_and_save()
    return {"status": "sent", "subject": subject, "to": [x.email_address for x in to], "cc": [x.email_address for x in cc], "bcc": [x.email_address for x in bcc], "attachments": attachments}


def render_quote(item) -> str:
    return (
        "\n\n--- Original Message ---\n"
        f"From: {(item.sender.email_address if getattr(item, 'sender', None) else '')}\n"
        f"Date: {(item.datetime_received.isoformat() if getattr(item, 'datetime_received', None) else '')}\n"
        f"Subject: {getattr(item, 'subject', '')}\n\n"
        f"{str(getattr(item, 'text_body', '') or getattr(item, 'body', '') or '')}"
    )


def get_draft_or_raise(state: Dict[str, Any], draft_id: str) -> Dict[str, Any]:
    draft = state["drafts"].get(draft_id)
    if not isinstance(draft, dict):
        raise ValueError("Draft not found")
    return draft


def find_attachment(files: List[FileAttachment], name: str, index: int) -> FileAttachment:
    if name:
        for a in files:
            if getattr(a, "name", "") == name:
                return a
        raise ValueError("Attachment not found by name")
    if index < 0 or index >= len(files):
        raise ValueError("Attachment index out of range")
    return files[index]


def parse_tags(raw: Any) -> List[str]:
    if isinstance(raw, list):
        return [str(x).strip() for x in raw if str(x).strip()]
    return parse_csv(str(raw or ""))


def parse_nl_rule_text(rule_text: str) -> Dict[str, Any]:
    text = (rule_text or "").strip()
    if not text:
        raise ValueError("--rule-text required")

    conditions: Dict[str, Any] = {}
    actions: List[Dict[str, Any]] = []

    boundary = r"(?=(?:\s*(?:且|并且|并|and|then|然后|，|,|。|;|；|时|则|就))|$)"
    sender_patterns = [
        rf"发件人(?:包含|是|为)\s*[\"“]?(.+?)[\"”]?{boundary}",
        rf"来自\s*[\"“]?(.+?)[\"”]?{boundary}",
        rf"from\s+[\"“]?(.+?)[\"”]?{boundary}",
    ]
    for pattern in sender_patterns:
        hit = re.search(pattern, text, flags=re.IGNORECASE)
        if hit:
            conditions["sender_contains"] = hit.group(1).strip()
            break

    subject_patterns = [
        rf"(?:主题|标题)(?:包含|含有|是|为)\s*[\"“]?(.+?)[\"”]?{boundary}",
        rf"subject\s+contains\s+[\"“]?(.+?)[\"”]?{boundary}",
    ]
    for pattern in subject_patterns:
        hit = re.search(pattern, text, flags=re.IGNORECASE)
        if hit:
            conditions["subject_contains"] = hit.group(1).strip()
            break

    body_patterns = [
        rf"(?:正文|内容)(?:包含|含有)\s*[\"“]?(.+?)[\"”]?{boundary}",
        rf"body\s+contains\s+[\"“]?(.+?)[\"”]?{boundary}",
    ]
    for pattern in body_patterns:
        hit = re.search(pattern, text, flags=re.IGNORECASE)
        if hit:
            conditions["body_contains"] = hit.group(1).strip()
            break

    if re.search(r"(未读|unread)", text, flags=re.IGNORECASE):
        conditions["is_unread"] = True
    if re.search(r"(有附件|含附件|带附件|has attachment)", text, flags=re.IGNORECASE):
        conditions["has_attachments"] = True

    if re.search(r"(标记已读|设为已读|mark read)", text, flags=re.IGNORECASE):
        actions.append({"type": "mark_read"})
    if re.search(r"(标记未读|设为未读|mark unread)", text, flags=re.IGNORECASE):
        actions.append({"type": "mark_unread"})
    if re.search(r"(取消星标|取消加星|unstar)", text, flags=re.IGNORECASE):
        actions.append({"type": "unstar"})
    elif re.search(r"(星标|加星|标星|star)", text, flags=re.IGNORECASE):
        actions.append({"type": "star"})
    if re.search(r"(归档|archive)", text, flags=re.IGNORECASE):
        actions.append({"type": "archive"})
    if re.search(r"(标记垃圾|垃圾邮件|spam)", text, flags=re.IGNORECASE):
        actions.append({"type": "mark_spam"})
    if re.search(r"(删除|delete)", text, flags=re.IGNORECASE):
        actions.append({"type": "delete"})

    move_hit = re.search(rf"(?:移动到|移到|move to)\s*[\"“]?(.+?)[\"”]?{boundary}", text, flags=re.IGNORECASE)
    if move_hit:
        actions.append({"type": "move", "target_folder": move_hit.group(1).strip()})

    tag_add_hit = re.search(rf"(?:添加标签|加标签|tag add)\s*[\"“]?(.+?)[\"”]?{boundary}", text, flags=re.IGNORECASE)
    if tag_add_hit:
        tags = [x.strip() for x in re.split(r"[,\s，、]+", tag_add_hit.group(1)) if x.strip()]
        if tags:
            actions.append({"type": "tag_add", "tags": tags})

    tag_set_hit = re.search(rf"(?:设置标签|tag set)\s*[\"“]?(.+?)[\"”]?{boundary}", text, flags=re.IGNORECASE)
    if tag_set_hit:
        tags = [x.strip() for x in re.split(r"[,\s，、]+", tag_set_hit.group(1)) if x.strip()]
        if tags:
            actions.append({"type": "tag_set", "tags": tags})

    reply_hit = re.search(r"(?:自动回复|回复)\s*[：: ]\s*[\"“]([^\"”]+)[\"”]", text, flags=re.IGNORECASE)
    if reply_hit:
        actions.append({"type": "reply_template", "body": reply_hit.group(1).strip(), "quote_original": False})

    fwd_hit = re.search(rf"(?:转发给|forward to)\s*[\"“]?(.+?)[\"”]?{boundary}", text, flags=re.IGNORECASE)
    if fwd_hit:
        to_values = [x.strip() for x in re.split(r"[,\s，、]+", fwd_hit.group(1)) if x.strip()]
        if to_values:
            actions.append({"type": "forward", "to": ",".join(to_values), "quote_original": True})

    if not conditions:
        conditions["is_unread"] = True
    if not actions:
        actions.append({"type": "mark_read"})

    return {"enabled": True, "conditions": conditions, "actions": actions, "source_text": text}


def rule_matches_message(rule: Dict[str, Any], item) -> bool:
    if not bool(rule.get("enabled", True)):
        return False
    conditions = rule.get("conditions", {}) or {}
    if not isinstance(conditions, dict):
        return False
    sender = (item.sender.email_address if getattr(item, "sender", None) else "").lower()
    subject = str(getattr(item, "subject", "") or "")
    body = str(getattr(item, "text_body", "") or getattr(item, "body", "") or "")
    if conditions.get("sender_contains"):
        if str(conditions.get("sender_contains", "")).lower() not in sender:
            return False
    if conditions.get("subject_contains"):
        if str(conditions.get("subject_contains", "")).lower() not in subject.lower():
            return False
    if conditions.get("body_contains"):
        if str(conditions.get("body_contains", "")).lower() not in body.lower():
            return False
    if "has_attachments" in conditions:
        if bool(getattr(item, "has_attachments", False)) != bool(conditions.get("has_attachments")):
            return False
    if "is_unread" in conditions:
        if (not bool(getattr(item, "is_read", False))) != bool(conditions.get("is_unread")):
            return False
    return True


def normalize_action(action: Any) -> Dict[str, Any]:
    if isinstance(action, str):
        return {"type": action}
    if isinstance(action, dict):
        return dict(action)
    return {"type": ""}


def apply_automation_action(args, account: Account, item, action: Dict[str, Any]) -> Dict[str, Any]:
    kind = str(action.get("type", "")).strip().lower()
    if not kind:
        return {"type": kind, "status": "skipped", "reason": "missing action type"}
    if kind == "mark_read":
        if args.dry_run:
            return {"type": kind, "status": "dry_run"}
        item.is_read = True
        item.save(update_fields=["is_read"])
        return {"type": kind, "status": "ok"}
    if kind == "mark_unread":
        if args.dry_run:
            return {"type": kind, "status": "dry_run"}
        item.is_read = False
        item.save(update_fields=["is_read"])
        return {"type": kind, "status": "ok"}
    if kind in {"star", "unstar"}:
        value = bool(action.get("value", True))
        if kind == "unstar":
            value = False
        if args.dry_run:
            return {"type": kind, "status": "dry_run", "value": value}
        item.is_flagged = value
        item.save(update_fields=["is_flagged"])
        return {"type": kind, "status": "ok", "value": value}
    if kind in {"tag_set", "tag_add", "tag_remove"}:
        tags = parse_tags(action.get("tags", []))
        current = list(getattr(item, "categories", []) or [])
        if kind == "tag_set":
            next_tags = tags
        elif kind == "tag_add":
            next_tags = list(dict.fromkeys(current + tags))
        else:
            remove_set = set(tags)
            next_tags = [x for x in current if x not in remove_set]
        if args.dry_run:
            return {"type": kind, "status": "dry_run", "tags": next_tags}
        item.categories = next_tags
        item.save(update_fields=["categories"])
        return {"type": kind, "status": "ok", "tags": next_tags}
    if kind == "move":
        target_folder = str(action.get("target_folder", "")).strip()
        if not target_folder:
            return {"type": kind, "status": "skipped", "reason": "target_folder required"}
        if args.dry_run:
            return {"type": kind, "status": "dry_run", "target_folder": target_folder}
        item.move(resolve_folder(account, target_folder))
        return {"type": kind, "status": "ok", "target_folder": target_folder}
    if kind == "archive":
        target_folder = str(action.get("target_folder", "Archive")).strip() or "Archive"
        if args.dry_run:
            return {"type": kind, "status": "dry_run", "target_folder": target_folder}
        item.move(resolve_folder(account, target_folder))
        return {"type": kind, "status": "ok", "target_folder": target_folder}
    if kind == "mark_spam":
        target_folder = str(action.get("target_folder", "Junk Email")).strip() or "Junk Email"
        if args.dry_run:
            return {"type": kind, "status": "dry_run", "target_folder": target_folder}
        item.move(resolve_folder(account, target_folder))
        return {"type": kind, "status": "ok", "target_folder": target_folder}
    if kind == "delete":
        if args.dry_run:
            return {"type": kind, "status": "dry_run"}
        item.delete()
        return {"type": kind, "status": "ok"}
    if kind == "reply_template":
        body = str(action.get("body", "已收到，稍后回复。"))
        quote_original = bool(action.get("quote_original", False))
        recipient = item.sender.email_address if getattr(item, "sender", None) else ""
        if not recipient:
            return {"type": kind, "status": "skipped", "reason": "sender missing"}
        result = send_message(
            args,
            account,
            f"RE: {getattr(item, 'subject', '')}",
            body + (render_quote(item) if quote_original else ""),
            recipient,
        )
        return {"type": kind, "status": result.get("status", "ok"), "to": recipient}
    if kind == "forward":
        to_raw = str(action.get("to", "")).strip()
        if not to_raw:
            return {"type": kind, "status": "skipped", "reason": "to required"}
        body = str(action.get("body", ""))
        quote_original = bool(action.get("quote_original", True))
        result = send_message(
            args,
            account,
            f"FWD: {getattr(item, 'subject', '')}",
            body + (render_quote(item) if quote_original else ""),
            to_raw,
            str(action.get("cc", "") or ""),
            str(action.get("bcc", "") or ""),
        )
        return {"type": kind, "status": result.get("status", "ok"), "to": parse_csv(to_raw)}
    return {"type": kind, "status": "skipped", "reason": "unsupported action type"}


def agent_op(args, account: Optional[Account]) -> Dict[str, Any]:
    op = args.name.strip().lower()
    alias_map = {
        "compose.create": "compose.new",
        "compose.recipients.add_to": "compose.add_recipients",
        "compose.recipients.add_cc": "compose.add_recipients",
        "compose.recipients.add_bcc": "compose.add_recipients",
        "compose.attachments.add": "compose.add_attachment",
        "conversation.get": "conversation.view",
        "message.unstar": "message.star",
    }
    op = alias_map.get(op, op)
    state = load_state(args)
    changed = False

    def finalize(data: Dict[str, Any]) -> Dict[str, Any]:
        if changed and not args.dry_run:
            data["state_file"] = save_state(args, state)
        return data

    if op == "compose.new":
        draft_id = args.draft_id or f"draft-{int(time.time())}"
        draft = {"id": draft_id, "subject": args.subject, "body": args.body, "to": parse_csv(args.to), "cc": parse_csv(args.cc), "bcc": parse_csv(args.bcc), "attachments": parse_csv(args.attach), "updated_at": now_iso()}
        if not args.dry_run:
            state["drafts"][draft_id] = draft
            changed = True
        return finalize({"status": "dry_run" if args.dry_run else "draft_saved", "draft": draft})
    if op == "compose.get":
        draft = get_draft_or_raise(state, args.draft_id)
        return {"status": "ok", "draft": draft}
    if op == "compose.list":
        rows = list(state["drafts"].values())
        rows.sort(key=lambda x: x.get("updated_at", ""), reverse=True)
        return {"status": "ok", "count": len(rows), "items": rows[: args.limit]}
    if op == "compose.delete":
        _ = get_draft_or_raise(state, args.draft_id)
        if args.dry_run:
            return {"status": "dry_run", "draft_id": args.draft_id}
        state["drafts"].pop(args.draft_id, None)
        changed = True
        return finalize({"status": "ok", "draft_id": args.draft_id})
    if op == "compose.subject.set":
        draft = get_draft_or_raise(state, args.draft_id)
        if args.dry_run:
            return {"status": "dry_run", "draft_id": args.draft_id, "subject": args.subject}
        draft["subject"] = args.subject
        draft["updated_at"] = now_iso()
        changed = True
        return finalize({"status": "ok", "draft_id": args.draft_id, "subject": draft["subject"]})
    if op == "compose.body.set":
        draft = get_draft_or_raise(state, args.draft_id)
        if args.dry_run:
            return {"status": "dry_run", "draft_id": args.draft_id, "body": args.body}
        draft["body"] = args.body
        draft["updated_at"] = now_iso()
        changed = True
        return finalize({"status": "ok", "draft_id": args.draft_id, "body": draft["body"]})
    if op == "compose.add_recipients":
        draft = get_draft_or_raise(state, args.draft_id)
        to_add = parse_csv(args.to)
        cc_add = parse_csv(args.cc)
        bcc_add = parse_csv(args.bcc)
        if args.name.strip().lower() == "compose.recipients.add_to":
            cc_add = []
            bcc_add = []
        if args.name.strip().lower() == "compose.recipients.add_cc":
            to_add = []
            bcc_add = []
        if args.name.strip().lower() == "compose.recipients.add_bcc":
            to_add = []
            cc_add = []
        merged = {
            "to": list(dict.fromkeys(draft.get("to", []) + to_add)),
            "cc": list(dict.fromkeys(draft.get("cc", []) + cc_add)),
            "bcc": list(dict.fromkeys(draft.get("bcc", []) + bcc_add)),
        }
        if not args.dry_run:
            draft.update(merged)
            draft["updated_at"] = now_iso()
            changed = True
        return finalize({"status": "dry_run" if args.dry_run else "updated", "draft_id": args.draft_id, **merged})
    if op in {"compose.recipients.clear_to", "compose.recipients.clear_cc", "compose.recipients.clear_bcc"}:
        draft = get_draft_or_raise(state, args.draft_id)
        key = "to" if op.endswith("_to") else ("cc" if op.endswith("_cc") else "bcc")
        if args.dry_run:
            return {"status": "dry_run", "draft_id": args.draft_id, "cleared": key}
        draft[key] = []
        draft["updated_at"] = now_iso()
        changed = True
        return finalize({"status": "ok", "draft_id": args.draft_id, "cleared": key})
    if op == "compose.add_attachment":
        draft = get_draft_or_raise(state, args.draft_id)
        attachments = list(dict.fromkeys(draft.get("attachments", []) + parse_csv(args.attach)))
        if not args.dry_run:
            draft["attachments"] = attachments
            draft["updated_at"] = now_iso()
            changed = True
        return finalize({"status": "dry_run" if args.dry_run else "updated", "draft_id": args.draft_id, "attachments": attachments})
    if op == "compose.attachments.remove":
        draft = get_draft_or_raise(state, args.draft_id)
        current = draft.get("attachments", [])
        values = set(parse_csv(args.attach))
        kept = [x for x in current if x not in values]
        removed = [x for x in current if x in values]
        if args.dry_run:
            return {"status": "dry_run", "draft_id": args.draft_id, "removed": removed, "attachments": kept}
        draft["attachments"] = kept
        draft["updated_at"] = now_iso()
        changed = True
        return finalize({"status": "ok", "draft_id": args.draft_id, "removed": removed, "attachments": kept})
    if op == "compose.attachments.list":
        draft = get_draft_or_raise(state, args.draft_id)
        values = draft.get("attachments", [])
        return {"status": "ok", "draft_id": args.draft_id, "count": len(values), "items": values}
    if op == "send.now":
        if account is None:
            raise ValueError("send.now requires EWS account")
        if args.draft_id:
            draft = state["drafts"].get(args.draft_id)
            if not isinstance(draft, dict):
                raise ValueError("Draft not found")
            result = send_message(args, account, draft.get("subject", ""), draft.get("body", ""), ",".join(draft.get("to", [])), ",".join(draft.get("cc", [])), ",".join(draft.get("bcc", [])), ",".join(draft.get("attachments", [])))
            if not args.dry_run:
                state["drafts"].pop(args.draft_id, None)
                changed = True
            return finalize(result)
        return send_message(args, account, args.subject, args.body, args.to, args.cc, args.bcc, args.attach)
    if op == "send.later":
        when_utc = parse_datetime_utc(args.schedule_at, "schedule-at")
        # Validate recipients early so invalid jobs are not queued.
        parse_recipients(args.to, required=True)
        job = {
            "id": f"job-{int(time.time())}",
            "schedule_at": when_utc.isoformat(),
            "subject": args.subject,
            "body": args.body,
            "to": parse_csv(args.to),
            "cc": parse_csv(args.cc),
            "bcc": parse_csv(args.bcc),
            "attachments": parse_csv(args.attach),
        }
        if args.dry_run:
            return {"status": "dry_run", "job": job}
        state["scheduled_outbox"].append(job)
        changed = True
        return finalize({"status": "scheduled", "job": job})
    if op == "send.schedule.list":
        rows = state.get("scheduled_outbox", [])
        rows = sorted(rows, key=lambda x: x.get("schedule_at", ""))
        return {"status": "ok", "count": len(rows), "items": rows}
    if op == "send.schedule.cancel":
        job_id = (args.job_id or "").strip()
        if not job_id:
            raise ValueError("--job-id required")
        rows = state.get("scheduled_outbox", [])
        hit = [x for x in rows if str(x.get("id", "")) == job_id]
        if not hit:
            raise ValueError("Scheduled job not found")
        if args.dry_run:
            return {"status": "dry_run", "job_id": job_id}
        state["scheduled_outbox"] = [x for x in rows if str(x.get("id", "")) != job_id]
        changed = True
        return finalize({"status": "ok", "job_id": job_id, "cancelled": True})
    if op == "send.schedule.flush_due":
        now_dt = datetime.now(timezone.utc)
        due: List[Dict[str, Any]] = []
        pending: List[Dict[str, Any]] = []
        invalid: List[Dict[str, Any]] = []
        for job in state.get("scheduled_outbox", []):
            raw_time = str(job.get("schedule_at", "")).strip()
            try:
                scheduled_at = parse_datetime_utc(raw_time, "schedule-at")
            except ValueError as exc:
                invalid.append({"id": job.get("id"), "schedule_at": raw_time, "error": str(exc)})
                pending.append(job)
                continue
            normalized = scheduled_at.isoformat()
            if raw_time != normalized and not args.dry_run:
                job["schedule_at"] = normalized
                changed = True
            if scheduled_at <= now_dt:
                due.append(job)
            else:
                pending.append(job)
        if args.dry_run:
            return {
                "status": "dry_run",
                "due_count": len(due),
                "due_items": due,
                "invalid_count": len(invalid),
                "invalid_items": invalid,
            }
        if account is None and due:
            raise ValueError("send.schedule.flush_due requires EWS account for due items")
        sent = []
        failed = []
        for job in due:
            try:
                result = send_message(
                    args,
                    account,
                    job.get("subject", ""),
                    job.get("body", ""),
                    ",".join(job.get("to", [])),
                    ",".join(job.get("cc", [])),
                    ",".join(job.get("bcc", [])),
                    ",".join(job.get("attachments", [])),
                )
                sent.append({"id": job.get("id"), "result": result.get("status", "sent")})
            except Exception as exc:
                failed.append({"id": job.get("id"), "error": str(exc)})
                pending.append(job)
        old_count = len(state.get("scheduled_outbox", []))
        state["scheduled_outbox"] = pending
        if old_count != len(pending) or sent:
            changed = True
        return finalize(
            {
                "status": "ok",
                "flushed_count": len(sent),
                "items": sent,
                "failed_count": len(failed),
                "failed_items": failed,
                "invalid_count": len(invalid),
                "invalid_items": invalid,
                "pending_count": len(pending),
            }
        )
    if op == "sender.block":
        sender = (args.sender or "").strip().lower()
        if not sender:
            raise ValueError("--sender required")
        if sender not in state["blocked_senders"] and not args.dry_run:
            state["blocked_senders"].append(sender)
            changed = True
        return finalize({"status": "dry_run" if args.dry_run else "ok", "blocked_senders": state["blocked_senders"] + ([sender] if args.dry_run else [])})
    if op == "sender.blocked.list":
        rows = state.get("blocked_senders", [])
        return {"status": "ok", "count": len(rows), "items": rows}
    if op == "sender.block.remove":
        sender = (args.sender or "").strip().lower()
        rows = state.get("blocked_senders", [])
        if sender not in rows:
            return {"status": "ok", "removed": False, "sender": sender}
        if args.dry_run:
            return {"status": "dry_run", "removed": True, "sender": sender}
        state["blocked_senders"] = [x for x in rows if x != sender]
        changed = True
        return finalize({"status": "ok", "removed": True, "sender": sender})
    if op == "message.unsubscribe":
        sender = (args.sender or "").strip().lower()
        if not sender:
            raise ValueError("--sender required")
        if sender not in state["unsubscribe_list"] and not args.dry_run:
            state["unsubscribe_list"].append(sender)
            changed = True
        return finalize({"status": "dry_run" if args.dry_run else "ok", "unsubscribe_list": state["unsubscribe_list"] + ([sender] if args.dry_run else [])})
    if op == "message.unsubscribe.list":
        rows = state.get("unsubscribe_list", [])
        return {"status": "ok", "count": len(rows), "items": rows}
    if op == "message.unsubscribe.remove":
        sender = (args.sender or "").strip().lower()
        rows = state.get("unsubscribe_list", [])
        if sender not in rows:
            return {"status": "ok", "removed": False, "sender": sender}
        if args.dry_run:
            return {"status": "dry_run", "removed": True, "sender": sender}
        state["unsubscribe_list"] = [x for x in rows if x != sender]
        changed = True
        return finalize({"status": "ok", "removed": True, "sender": sender})
    if op == "contact.priority_alert":
        sender = (args.sender or "").strip().lower()
        if not sender:
            raise ValueError("--sender required")
        payload = {"level": args.alert_level or "high", "updated_at": now_iso()}
        if args.dry_run:
            return {"status": "dry_run", "contact": sender, "alert": payload}
        state["priority_contact_alerts"][sender] = payload
        changed = True
        return finalize({"status": "ok", "contact": sender, "alert": payload})
    if op == "contact.priority_alert.get":
        sender = (args.sender or "").strip().lower()
        alerts = state.get("priority_contact_alerts", {})
        if sender:
            return {"status": "ok", "contact": sender, "alert": alerts.get(sender)}
        return {"status": "ok", "count": len(alerts), "items": alerts}
    if op == "contact.priority_alert.remove":
        sender = (args.sender or "").strip().lower()
        if not sender:
            raise ValueError("--sender required")
        if sender not in state.get("priority_contact_alerts", {}):
            return {"status": "ok", "removed": False, "contact": sender}
        if args.dry_run:
            return {"status": "dry_run", "removed": True, "contact": sender}
        state["priority_contact_alerts"].pop(sender, None)
        changed = True
        return finalize({"status": "ok", "removed": True, "contact": sender})
    if op == "rule.set":
        if not args.rule_id:
            raise ValueError("--rule-id required")
        payload = json.loads(args.rule_json) if args.rule_json else {"enabled": True}
        if args.dry_run:
            return {"status": "dry_run", "rule_id": args.rule_id, "rule": payload}
        state["rules"][args.rule_id] = payload
        changed = True
        return finalize({"status": "ok", "rule_id": args.rule_id, "rule": payload})
    if op == "rule.from_text":
        if not args.rule_id:
            raise ValueError("--rule-id required")
        payload = parse_nl_rule_text(args.rule_text)
        if args.dry_run:
            return {"status": "dry_run", "rule_id": args.rule_id, "rule": payload}
        state["rules"][args.rule_id] = payload
        changed = True
        return finalize({"status": "ok", "rule_id": args.rule_id, "rule": payload})
    if op == "rule.get":
        if not args.rule_id:
            raise ValueError("--rule-id required")
        return {"status": "ok", "rule_id": args.rule_id, "rule": state.get("rules", {}).get(args.rule_id)}
    if op == "rule.list":
        rows = state.get("rules", {})
        return {"status": "ok", "count": len(rows), "items": rows}
    if op == "rule.delete":
        if not args.rule_id:
            raise ValueError("--rule-id required")
        if args.rule_id not in state.get("rules", {}):
            return {"status": "ok", "removed": False, "rule_id": args.rule_id}
        if args.dry_run:
            return {"status": "dry_run", "removed": True, "rule_id": args.rule_id}
        state["rules"].pop(args.rule_id, None)
        changed = True
        return finalize({"status": "ok", "removed": True, "rule_id": args.rule_id})
    if op == "autoreply.set":
        payload = {"enabled": True, "start": args.start or "", "end": args.end or "", "message": args.message or "", "updated_at": now_iso()}
        if args.dry_run:
            return {"status": "dry_run", "autoreply": payload}
        state["autoreply"] = payload
        changed = True
        return finalize({"status": "ok", "autoreply": payload})
    if op == "autoreply.get":
        return {"status": "ok", "autoreply": state.get("autoreply", {})}
    if op == "autoreply.disable":
        payload = {"enabled": False, "start": "", "end": "", "message": "", "updated_at": now_iso()}
        if args.dry_run:
            return {"status": "dry_run", "autoreply": payload}
        state["autoreply"] = payload
        changed = True
        return finalize({"status": "ok", "autoreply": payload})

    if account is None:
        raise ValueError(f"{op} requires EWS account")

    folder = resolve_folder(account, args.folder or "Inbox")
    if op == "automation.run":
        rules = state.get("rules", {})
        if not isinstance(rules, dict) or not rules:
            return {"status": "ok", "scanned": 0, "matched_messages": 0, "applied_actions": 0, "runs": []}
        rows = list(folder.all().order_by("-datetime_received")[: args.scan_limit])
        if args.sender_filter:
            rows = [x for x in rows if args.sender_filter.lower() in ((x.sender.email_address if getattr(x, "sender", None) else "").lower())]
        if args.keyword:
            rows = [x for x in rows if args.keyword.lower() in ((str(getattr(x, "subject", "") or "") + " " + str(getattr(x, "text_body", "") or "")).lower())]
        if args.has_attachments:
            rows = [x for x in rows if bool(getattr(x, "has_attachments", False))]
        scanned = rows[: args.limit]
        runs = []
        matched_messages = 0
        applied_actions = 0
        for item in scanned:
            message_actions = []
            matched_rule_ids = []
            for rule_id, raw_rule in rules.items():
                if not isinstance(raw_rule, dict):
                    continue
                if not rule_matches_message(raw_rule, item):
                    continue
                matched_rule_ids.append(str(rule_id))
                actions = raw_rule.get("actions", [])
                if not isinstance(actions, list):
                    actions = [actions]
                if not actions:
                    actions = [{"type": "mark_read"}]
                for action in actions:
                    normalized = normalize_action(action)
                    result = apply_automation_action(args, account, item, normalized)
                    message_actions.append({"rule_id": str(rule_id), "action": normalized, "result": result})
                    if result.get("status") in {"ok", "sent"}:
                        applied_actions += 1
            if matched_rule_ids:
                matched_messages += 1
                runs.append(
                    {
                        "message": serialize_message(item),
                        "matched_rules": matched_rule_ids,
                        "actions": message_actions,
                    }
                )
        report = {
            "status": "dry_run" if args.dry_run else "ok",
            "scanned": len(scanned),
            "matched_messages": matched_messages,
            "applied_actions": applied_actions,
            "runs": runs,
        }
        if not args.dry_run:
            history = list(state.get("automation_history", []))
            history.append(
                {
                    "timestamp": now_iso(),
                    "folder": args.folder or "Inbox",
                    "scanned": len(scanned),
                    "matched_messages": matched_messages,
                    "applied_actions": applied_actions,
                }
            )
            state["automation_history"] = history[-200:]
            changed = True
        return finalize(report)
    if op in {"mailbox.inbox.list", "mail.search", "mail.filter", "mail.sort", "mailbox.message.list_unread"}:
        rows = list(folder.all().order_by("-datetime_received")[: args.scan_limit])
        if op == "mailbox.message.list_unread":
            rows = [x for x in rows if not bool(getattr(x, "is_read", False))]
        keyword = (args.keyword or "").lower()
        sender_filter = (args.sender_filter or "").lower()
        if keyword:
            rows = [x for x in rows if keyword in ((str(getattr(x, "subject", "") or "") + " " + str(getattr(x, "text_body", "") or "")).lower())]
        if sender_filter:
            rows = [x for x in rows if sender_filter in ((x.sender.email_address if getattr(x, "sender", None) else "").lower())]
        if args.has_attachments:
            rows = [x for x in rows if bool(getattr(x, "has_attachments", False))]
        serialized = [serialize_message(x) for x in rows]
        if args.sort_by == "subject":
            serialized.sort(key=lambda x: x.get("subject", "").lower(), reverse=args.sort_desc)
        elif args.sort_by == "from":
            serialized.sort(key=lambda x: x.get("from", "").lower(), reverse=args.sort_desc)
        else:
            serialized.sort(key=lambda x: x.get("datetime_received", ""), reverse=args.sort_desc)
        paged = serialized[args.offset : args.offset + args.limit]
        return {"status": "ok", "count": len(paged), "total": len(serialized), "items": paged}
    if op == "mailbox.message.get":
        item = folder.get(id=args.id)
        return {"status": "ok", "item": serialize_message(item, include_body=args.include_body, include_attachments=True)}
    if op == "conversation.view":
        rows = []
        for item in list(folder.all().order_by("-datetime_received")[: args.scan_limit]):
            cid = str(getattr(getattr(item, "conversation_id", None), "id", "") or "")
            if args.conversation_id and cid != args.conversation_id:
                continue
            rows.append(serialize_message(item, include_body=args.include_body, include_attachments=True))
        return {"status": "ok", "count": len(rows), "items": rows[: args.limit]}
    if op in {"attachment.list", "attachment.meta", "attachment.open", "attachment.preview", "attachment.download"}:
        item = folder.get(id=args.id)
        files = [a for a in (item.attachments or []) if isinstance(a, FileAttachment)]
        if not files:
            raise ValueError("No file attachments found")
        if op == "attachment.list":
            rows = [{"name": a.name, "content_type": a.content_type, "size": a.size} for a in files]
            return {"status": "ok", "count": len(rows), "items": rows}
        picked = find_attachment(files, args.attachment_name, args.attachment_index)
        if op == "attachment.meta":
            return {"status": "ok", "name": picked.name, "content_type": picked.content_type, "size": picked.size}
        out = Path(args.out_dir or "outputs/attachments")
        out.mkdir(parents=True, exist_ok=True)
        target = out / picked.name
        if not args.dry_run:
            target.write_bytes(picked.content)
        ctype = mimetypes.guess_type(str(target))[0] or picked.content_type
        data = {"status": "dry_run" if args.dry_run else "ok", "name": picked.name, "content_type": ctype, "path": str(target)}
        if op == "attachment.preview" and ctype and ctype.startswith("image/") and not args.dry_run:
            data["preview_base64"] = base64.b64encode(picked.content[:4096]).decode("ascii")
        return data
    item = folder.get(id=args.id) if args.id else None
    if op == "message.quote":
        return {"status": "ok", "quote": render_quote(item), "id": args.id}
    if op == "message.reply_with_quote":
        return send_message(args, account, f"RE: {item.subject}", (args.body or "") + render_quote(item), item.sender.email_address if item.sender else "")
    if op == "message.reply":
        return send_message(args, account, f"RE: {item.subject}", (args.body or "") + (render_quote(item) if args.quote_original else ""), item.sender.email_address if item.sender else "")
    if op == "message.reply_all":
        targets = [item.sender.email_address] if item.sender else []
        targets.extend([x.email_address for x in (item.to_recipients or []) if x.email_address != account.primary_smtp_address and x.email_address not in targets])
        return send_message(args, account, f"RE: {item.subject}", (args.body or "") + (render_quote(item) if args.quote_original else ""), ",".join(targets))
    if op == "message.forward":
        return send_message(args, account, f"FWD: {item.subject}", (args.body or "") + (render_quote(item) if args.quote_original else ""), args.to, args.cc, args.bcc, args.attach)
    if op == "message.template_reply":
        text = args.body or "已收到，稍后回复。"
        return send_message(args, account, f"RE: {item.subject}", text + (render_quote(item) if args.quote_original else ""), item.sender.email_address if item.sender else "")
    if args.dry_run and op in MUTATING_OPS:
        return {"status": "dry_run", "op": op, "id": args.id}
    if op == "message.mark_read":
        item.is_read = True
        item.save(update_fields=["is_read"])
        return {"status": "ok", "op": op, "id": args.id}
    if op == "message.mark_unread":
        item.is_read = False
        item.save(update_fields=["is_read"])
        return {"status": "ok", "op": op, "id": args.id}
    if op == "message.star":
        item.is_flagged = bool(args.star_value)
        item.save(update_fields=["is_flagged"])
        return {"status": "ok", "op": op, "id": args.id}
    if op == "message.move":
        item.move(resolve_folder(account, args.target_folder))
        return {"status": "ok", "op": op, "id": args.id}
    if op == "message.archive":
        item.move(resolve_folder(account, args.target_folder or "Archive"))
        return {"status": "ok", "op": op, "id": args.id}
    if op == "message.delete":
        item.delete()
        return {"status": "ok", "op": op, "id": args.id}
    if op == "message.tag":
        item.categories = parse_csv(args.tags)
        item.save(update_fields=["categories"])
        return {"status": "ok", "op": op, "id": args.id, "tags": item.categories}
    if op == "message.tags.add":
        merged = list(dict.fromkeys((list(getattr(item, "categories", []) or [])) + parse_csv(args.tags)))
        item.categories = merged
        item.save(update_fields=["categories"])
        return {"status": "ok", "op": op, "id": args.id, "tags": item.categories}
    if op == "message.tags.remove":
        to_remove = set(parse_csv(args.tags))
        item.categories = [x for x in (list(getattr(item, "categories", []) or [])) if x not in to_remove]
        item.save(update_fields=["categories"])
        return {"status": "ok", "op": op, "id": args.id, "tags": item.categories}
    if op == "message.tags.clear":
        item.categories = []
        item.save(update_fields=["categories"])
        return {"status": "ok", "op": op, "id": args.id, "tags": []}
    if op == "message.mark_spam":
        item.move(resolve_folder(account, args.target_folder or "Junk Email"))
        return {"status": "ok", "op": op, "id": args.id}
    raise ValueError(f"Unsupported op: {op}")


def cmd_healthcheck(args) -> int:
    started = time.time()
    account = build_account(args)
    return emit(args, "ok", {"connected": True, "inbox_total_count": account.inbox.total_count}, None, "healthcheck", started)


def cmd_agent_capabilities(args) -> int:
    started = time.time()
    data = {
        "decision_center": "openclaw_agent",
        "execution_layer": "ews_cli_primitives",
        "capabilities": {
            "write_send": [
                "compose.new",
                "compose.get",
                "compose.list",
                "compose.delete",
                "compose.subject.set",
                "compose.body.set",
                "compose.recipients.add_to",
                "compose.recipients.add_cc",
                "compose.recipients.add_bcc",
                "compose.recipients.clear_to",
                "compose.recipients.clear_cc",
                "compose.recipients.clear_bcc",
                "compose.attachments.add",
                "compose.attachments.remove",
                "compose.attachments.list",
                "send.now",
                "send.later",
                "send.schedule.list",
                "send.schedule.cancel",
                "send.schedule.flush_due",
            ],
            "receive_view": [
                "mailbox.inbox.list",
                "mailbox.message.list_unread",
                "mailbox.message.get",
                "conversation.view",
                "attachment.list",
                "attachment.meta",
                "attachment.open",
                "attachment.preview",
                "attachment.download",
            ],
            "reply_forward": [
                "message.quote",
                "message.reply",
                "message.reply_with_quote",
                "message.reply_all",
                "message.forward",
                "message.template_reply",
            ],
            "manage_organize": [
                "message.mark_read",
                "message.mark_unread",
                "message.star",
                "message.unstar",
                "message.move",
                "message.archive",
                "message.delete",
                "message.tag",
                "message.tags.add",
                "message.tags.remove",
                "message.tags.clear",
            ],
            "search_filter": ["mail.search", "mail.filter", "mail.sort"],
            "quiet_mode": [
                "sender.block",
                "sender.blocked.list",
                "sender.block.remove",
                "message.mark_spam",
                "message.unsubscribe",
                "message.unsubscribe.list",
                "message.unsubscribe.remove",
                "contact.priority_alert",
                "contact.priority_alert.get",
                "contact.priority_alert.remove",
            ],
            "automation": ["rule.set", "rule.from_text", "rule.get", "rule.list", "rule.delete", "automation.run", "autoreply.set", "autoreply.get", "autoreply.disable"],
        },
    }
    return emit(args, "ok", data, None, "agent capabilities", started)


def cmd_agent_op(args) -> int:
    started = time.time()
    op = args.name.strip().lower()
    account = None if op in LOCAL_ONLY_OPS else build_account(args)
    data = agent_op(args, account)
    return emit(args, "ok", data, None, f"agent op {op}", started)


def cmd_unread(args) -> int:
    started = time.time()
    account = build_account(args)
    op_args = argparse.Namespace(**vars(args))
    op_args.name = "mailbox.message.list_unread"
    # Keep unread lookup lightweight by default while ensuring requested page size can be fulfilled.
    op_args.scan_limit = max(int(getattr(op_args, "scan_limit", 0) or 0), int(getattr(op_args, "limit", 0) or 0))
    data = agent_op(op_args, account)
    return emit(args, "ok", data, None, "unread", started)


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Exchange EWS primitive execution for OpenClaw agent")
    p.add_argument("--endpoint", default="", help="EWS endpoint")
    p.add_argument("--email", default="", help="Primary SMTP email")
    p.add_argument("--username", default="", help="Exchange username")
    p.add_argument("--password", default="", help="Exchange password")
    p.add_argument("--password-cmd", default="", help="Command to fetch Exchange password (fallback: EWS_PASSWORD_CMD env)")
    p.add_argument("--auth", choices=["NTLM", "BASIC", "AUTO"], default="NTLM", help="Auth type")
    p.add_argument("--agent-state-file", default="", help="Local state file path")
    sub = p.add_subparsers(dest="command", required=True)

    h = sub.add_parser("healthcheck")
    h.set_defaults(func=cmd_healthcheck)

    unread = sub.add_parser("unread", help="List unread messages")
    unread.add_argument("--folder", default="Inbox")
    unread.add_argument("--keyword", default="")
    unread.add_argument("--sender-filter", default="")
    unread.add_argument("--has-attachments", action="store_true")
    unread.add_argument("--sort-by", choices=["received", "subject", "from"], default="received")
    unread.add_argument("--sort-desc", dest="sort_desc", action="store_true")
    unread.add_argument("--sort-asc", dest="sort_desc", action="store_false")
    unread.add_argument("--offset", type=int, default=0)
    unread.add_argument("--scan-limit", type=int, default=10)
    unread.add_argument("--limit", type=int, default=10)
    unread.set_defaults(func=cmd_unread, sort_desc=True)

    a = sub.add_parser("agent", help="Agent primitives")
    a_sub = a.add_subparsers(dest="agent_command", required=True)
    cap = a_sub.add_parser("capabilities")
    cap.set_defaults(func=cmd_agent_capabilities)

    op = a_sub.add_parser("op")
    op.add_argument("--name", required=True)
    op.add_argument("--folder", default="Inbox")
    op.add_argument("--id", default="")
    op.add_argument("--conversation-id", default="")
    op.add_argument("--draft-id", default="")
    op.add_argument("--subject", default="")
    op.add_argument("--body", default="")
    op.add_argument("--to", default="")
    op.add_argument("--cc", default="")
    op.add_argument("--bcc", default="")
    op.add_argument("--attach", default="")
    op.add_argument("--schedule-at", default="")
    op.add_argument("--job-id", default="")
    op.add_argument("--keyword", default="")
    op.add_argument("--sender-filter", default="")
    op.add_argument("--has-attachments", action="store_true")
    op.add_argument("--sort-by", choices=["received", "subject", "from"], default="received")
    op.add_argument("--sort-desc", dest="sort_desc", action="store_true")
    op.add_argument("--sort-asc", dest="sort_desc", action="store_false")
    op.add_argument("--offset", type=int, default=0)
    op.add_argument("--scan-limit", type=int, default=10)
    op.add_argument("--limit", type=int, default=10)
    op.add_argument("--attachment-name", default="")
    op.add_argument("--attachment-index", type=int, default=0)
    op.add_argument("--out-dir", default="outputs/attachments")
    op.add_argument("--include-body", action="store_true")
    op.add_argument("--quote-original", action="store_true")
    op.add_argument("--target-folder", default="")
    op.add_argument("--tags", default="")
    op.add_argument("--star-value", action="store_true", default=True)
    op.add_argument("--unstar", dest="star_value", action="store_false")
    op.add_argument("--sender", default="")
    op.add_argument("--alert-level", default="high")
    op.add_argument("--rule-id", default="")
    op.add_argument("--rule-json", default="")
    op.add_argument("--rule-text", default="")
    op.add_argument("--start", default="")
    op.add_argument("--end", default="")
    op.add_argument("--message", default="")
    op.add_argument("--dry-run", action="store_true", default=True)
    op.add_argument("--apply", action="store_true")
    op.set_defaults(func=cmd_agent_op, sort_desc=True)
    return p


def normalize_flags(args):
    if hasattr(args, "apply") and args.apply and hasattr(args, "dry_run"):
        args.dry_run = False
    if hasattr(args, "offset") and args.offset < 0:
        raise ValueError("--offset must be >= 0")
    if hasattr(args, "limit") and args.limit <= 0:
        raise ValueError("--limit must be > 0")
    if hasattr(args, "scan_limit") and args.scan_limit <= 0:
        raise ValueError("--scan-limit must be > 0")


def main() -> int:
    load_dotenv_files()
    p = parser()
    args = p.parse_args()
    normalize_flags(args)
    started = time.time()
    try:
        return args.func(args)
    except Exception as exc:
        return emit(args, "error", {}, {"type": exc.__class__.__name__, "message": str(exc)}, " ".join(os.sys.argv[1:]), started)


if __name__ == "__main__":
    raise SystemExit(main())

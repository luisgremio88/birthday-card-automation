from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
import unicodedata
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pythoncom
import win32com.client


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def load_email_defaults() -> dict[str, str]:
    defaults = {
        "senderEmail": "aniversarios@exemplo.com",
        "bccEmail": "auditoria@exemplo.com",
    }
    config_path = PROJECT_ROOT / "config" / "email_defaults.json"
    if not config_path.exists():
        return defaults
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return defaults
    return {
        "senderEmail": str(config.get("senderEmail") or defaults["senderEmail"]).strip(),
        "bccEmail": str(config.get("bccEmail") or defaults["bccEmail"]).strip(),
    }


EMAIL_DEFAULTS = load_email_defaults()
FONT_PATH = Path(r"C:\Windows\Fonts\tahomabd.ttf")
EMAIL_IMAGE_PIXELS = 567
DEFAULT_SENDER = EMAIL_DEFAULTS["senderEmail"]
DEFAULT_BCC = EMAIL_DEFAULTS["bccEmail"]
DEFAULT_INLINE_CID = "cartao_aniversario"
REQUIRED_COLUMNS = ["nome", "tabelionato", "email"]
MAPI_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
MAPI_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
VALID_EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
LOCK_TIMEOUT_SECONDS = 15 * 60
SENT_STATUSES = {"enviado", "preparando_envio"}
DRAFT_STATUSES = {"rascunho", *SENT_STATUSES}

PROFILE_CONFIG = {
    "associado": {
        "label": "Associado",
        "excel_path": PROJECT_ROOT / "uploads" / "aniversariantes_associado.xlsx",
        "template_path": PROJECT_ROOT / "templates" / "cartao_base_limpo_associado.png",
        "output_dir": PROJECT_ROOT / "gerados" / "associado",
        "log_path": PROJECT_ROOT / "logs" / "envios_associado.csv",
        "subject": "Feliz aniversario Associado",
        "body_intro": "Feliz aniversario",
        "body_highlight_fallback": "Associado CNB/RS",
        "body_message": "Desejamos um excelente dia e um novo ciclo repleto de realizacoes.",
        "name_box": {"x": 179, "y": 1538, "width": 1043, "height": 80, "align": "left"},
    },
    "diretoria": {
        "label": "Diretoria",
        "excel_path": PROJECT_ROOT / "uploads" / "aniversariantes_diretoria.xlsx",
        "template_path": PROJECT_ROOT / "templates" / "cartao_base_limpo_diretoria.png",
        "output_dir": PROJECT_ROOT / "gerados" / "diretoria",
        "log_path": PROJECT_ROOT / "logs" / "envios_diretoria.csv",
        "subject": "Parabens ao membro da Diretoria",
        "body_intro": "Parabens ao membro da Diretoria",
        "body_highlight_fallback": "Diretoria CNB/RS",
        "body_message": "Reconhecemos sua dedicacao e desejamos um novo ciclo de muito sucesso, saude e realizacoes.",
        "name_box": {"x": 320, "y": 710, "width": 1360, "height": 115, "align": "center"},
    },
}


@dataclass
class BirthdayEntry:
    nome: str
    tabelionato: str
    email: str
    data_aniversario: str


@dataclass
class LogState:
    sent_keys: set[tuple[str, str, str]]
    draft_keys: set[tuple[str, str, str]]


def normalize_header(value: object) -> str:
    normalized = str(value or "").strip().lower()
    normalized = unicodedata.normalize("NFD", normalized)
    normalized = "".join(char for char in normalized if unicodedata.category(char) != "Mn")
    return re.sub(r"\s+", "_", normalized)


def slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFD", value.strip().lower())
    normalized = "".join(char for char in normalized if unicodedata.category(char) != "Mn")
    normalized = re.sub(r"[^a-z0-9]+", "-", normalized)
    return normalized.strip("-") or "sem-nome"


def normalize_name_key(value: str) -> str:
    normalized = unicodedata.normalize("NFD", value.strip().lower())
    normalized = "".join(char for char in normalized if unicodedata.category(char) != "Mn")
    return re.sub(r"\s+", " ", normalized)


def is_valid_email(value: str) -> bool:
    return bool(VALID_EMAIL_PATTERN.fullmatch(value.strip()))


@contextmanager
def run_lock(profile: str, run_date: date):
    lock_dir = PROJECT_ROOT / "logs"
    lock_dir.mkdir(parents=True, exist_ok=True)
    lock_path = lock_dir / f"automation_{profile}_{run_date.isoformat()}.lock"
    now = time.time()

    if lock_path.exists():
        try:
            lock_age = now - lock_path.stat().st_mtime
        except OSError:
            lock_age = 0
        if lock_age < LOCK_TIMEOUT_SECONDS:
            raise RuntimeError(
                "Ja existe uma execucao em andamento para este perfil/data. "
                "Aguarde terminar para evitar envio duplicado."
            )

    lock_path.write_text(
        f"pid={os.getpid()}\nprofile={profile}\ndate={run_date.isoformat()}\nstarted={datetime.now().isoformat(timespec='seconds')}\n",
        encoding="utf-8",
    )
    try:
        yield
    finally:
        try:
            lock_path.unlink()
        except FileNotFoundError:
            pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera cartoes de aniversario e prepara/envia e-mails pelo Outlook.",
    )
    parser.add_argument(
        "--profile",
        choices=sorted(PROFILE_CONFIG.keys()),
        default="associado",
        help="Perfil de envio. Padrao: associado",
    )
    parser.add_argument("--excel", type=Path, help="Sobrescreve o caminho da planilha.")
    parser.add_argument("--template", type=Path, help="Sobrescreve o caminho do template base limpo.")
    parser.add_argument("--output-dir", type=Path, help="Sobrescreve a pasta de cartoes gerados.")
    parser.add_argument("--log-path", type=Path, help="Sobrescreve o arquivo CSV de log.")
    parser.add_argument("--sender-email", default=DEFAULT_SENDER, help=f"E-mail remetente. Padrao: {DEFAULT_SENDER}")
    parser.add_argument("--bcc-email", default=DEFAULT_BCC, help=f"E-mail fixo em Cco. Padrao: {DEFAULT_BCC}")
    parser.add_argument("--subject", help="Sobrescreve o assunto do e-mail.")
    parser.add_argument("--date", help="Data para teste no formato YYYY-MM-DD ou DD/MM.")
    parser.add_argument("--send", action="store_true", help="Envia de verdade. Sem esta flag, abre rascunho.")
    parser.add_argument("--force", action="store_true", help="Ignora o log e permite reenviar na mesma data.")
    return parser.parse_args()


def profile_settings(args: argparse.Namespace) -> dict:
    base = PROFILE_CONFIG[args.profile].copy()
    base["excel_path"] = args.excel or base["excel_path"]
    base["template_path"] = args.template or base["template_path"]
    base["output_dir"] = args.output_dir or base["output_dir"]
    base["log_path"] = args.log_path or base["log_path"]
    base["subject"] = args.subject or base["subject"]
    return base


def resolve_run_date(raw_value: str | None) -> date:
    if not raw_value:
        return date.today()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw_value):
        return datetime.strptime(raw_value, "%Y-%m-%d").date()
    if re.fullmatch(r"\d{2}/\d{2}", raw_value):
        today = date.today()
        return datetime.strptime(f"{raw_value}/{today.year}", "%d/%m/%Y").date()
    raise ValueError("Use --date no formato YYYY-MM-DD ou DD/MM.")


def format_birthday_cell(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%d/%m")

    text = str(value).strip()
    if not text:
        return ""
    match = re.match(r"^(\d{1,2})[/-](\d{1,2})(?:[/-]\d{2,4})?$", text)
    if match:
        return f"{int(match.group(1)):02d}/{int(match.group(2)):02d}"
    for pattern in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(text, pattern).strftime("%d/%m")
        except ValueError:
            continue
    return text[:5]


def is_blocked_birth_date(value: object) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, (pd.Timestamp, datetime, date)):
        parsed = value.date() if isinstance(value, datetime) else value
    else:
        text = str(value).strip()
        if re.fullmatch(r"\d{1,2}[/-]\d{1,2}$", text):
            return False
        parsed = None
        for pattern in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y"):
            try:
                parsed = datetime.strptime(text, pattern).date()
                break
            except ValueError:
                continue
        if parsed is None:
            return False

    return parsed == date(1900, 1, 1) or parsed > date(2000, 1, 1)


def load_birthdays(excel_path: Path) -> list[BirthdayEntry]:
    if not excel_path.exists():
        raise FileNotFoundError(f"Planilha nao encontrada: {excel_path}")

    frame = pd.read_excel(excel_path)
    frame = frame.rename(columns={column: normalize_header(column) for column in frame.columns})
    date_column = "data_aniversario" if "data_aniversario" in frame.columns else "data_de_nascimento"
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if date_column not in frame.columns:
        missing.append("data_aniversario ou data_de_nascimento")
    if missing:
        raise ValueError(f"Colunas obrigatorias ausentes na planilha: {', '.join(missing)}")

    frame = frame[[*REQUIRED_COLUMNS, date_column]].copy()
    frame = frame.rename(columns={date_column: "data_aniversario"})
    for column in ("nome", "tabelionato", "email"):
        frame[column] = frame[column].fillna("").astype(str).str.strip()
    frame = frame[~frame["data_aniversario"].map(is_blocked_birth_date)]
    frame["data_aniversario"] = frame["data_aniversario"].map(format_birthday_cell)
    frame = frame[(frame["nome"] != "") & (frame["email"] != "") & (frame["data_aniversario"] != "")]
    frame = frame[frame["email"].map(is_valid_email)]
    frame["_nome_key"] = frame["nome"].map(normalize_name_key)
    frame["_email_key"] = frame["email"].str.lower()
    frame = frame.drop_duplicates(subset="_email_key", keep="first")
    frame = frame.drop_duplicates(subset="_nome_key", keep="first")

    return [
        BirthdayEntry(
            nome=row["nome"],
            tabelionato=row["tabelionato"],
            email=row["email"],
            data_aniversario=row["data_aniversario"],
        )
        for _, row in frame.iterrows()
    ]


def birthdays_for_date(entries: Iterable[BirthdayEntry], run_date: date) -> list[BirthdayEntry]:
    today_key = run_date.strftime("%d/%m")
    return [entry for entry in entries if entry.data_aniversario[:5] == today_key]


def fit_single_line_name(draw: ImageDraw.ImageDraw, text: str, max_width: int) -> ImageFont.FreeTypeFont:
    for font_size in range(88, 23, -1):
        font = ImageFont.truetype(str(FONT_PATH), font_size)
        bbox = draw.textbbox((0, 0), text, font=font)
        if (bbox[2] - bbox[0]) <= max_width:
            return font
    return ImageFont.truetype(str(FONT_PATH), 24)


def generate_card(template_path: Path, output_dir: Path, person_name: str, run_date: date, name_box: dict) -> Path:
    if not template_path.exists():
        raise FileNotFoundError(f"Template nao encontrado: {template_path}")
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"Fonte nao encontrada: {FONT_PATH}")

    output_dir.mkdir(parents=True, exist_ok=True)
    image = Image.open(template_path).convert("RGBA")
    draw = ImageDraw.Draw(image)
    safe_name = person_name.strip().upper()

    font = fit_single_line_name(draw, safe_name, name_box["width"])
    bbox = draw.textbbox((0, 0), safe_name, font=font)
    text_height = bbox[3] - bbox[1]
    y_position = name_box["y"] + ((name_box["height"] - text_height) / 2) - bbox[1]

    if name_box.get("align") == "center":
        text_width = bbox[2] - bbox[0]
        text_x = name_box["x"] + (name_box["width"] - text_width) / 2
    else:
        text_x = name_box["x"]

    draw.text((text_x, y_position), safe_name, font=font, fill="white")

    filename = f"{run_date.strftime('%Y-%m-%d')}-{slugify(person_name)}.png"
    output_path = output_dir / filename
    image.save(output_path)
    return output_path


def generate_email_card_image(card_path: Path) -> Path:
    email_path = card_path.with_name(f"{card_path.stem}-email.png")
    image = Image.open(card_path).convert("RGBA")
    resized = image.resize((EMAIL_IMAGE_PIXELS, EMAIL_IMAGE_PIXELS), Image.Resampling.LANCZOS)
    resized.save(email_path, dpi=(96, 96))
    return email_path


def read_log_state(log_path: Path, profile: str) -> LogState:
    if not log_path.exists():
        return LogState(sent_keys=set(), draft_keys=set())

    latest_status_by_key = {}
    any_draft_keys = set()
    any_sent_keys = set()
    sent_keys = set()
    draft_keys = set()
    with log_path.open("r", encoding="utf-8", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            if (row.get("perfil") or profile) != profile:
                continue
            key = (
                row.get("perfil", ""),
                row.get("data_referencia", ""),
                row.get("email", "").strip().lower(),
            )
            status = row.get("status", "")
            latest_status_by_key[key] = status
            if status == "rascunho":
                any_draft_keys.add(key)
            if status == "enviado":
                any_sent_keys.add(key)

    for key, status in latest_status_by_key.items():
        if key in any_sent_keys or status == "preparando_envio":
            sent_keys.add(key)
            draft_keys.add(key)
        elif key in any_draft_keys:
            draft_keys.add(key)

    return LogState(sent_keys=sent_keys, draft_keys=draft_keys)


def append_log(
    log_path: Path,
    profile: str,
    run_date: date,
    entry: BirthdayEntry,
    card_path: Path,
    status: str,
    details: str,
) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    file_exists = log_path.exists()
    with log_path.open("a", encoding="utf-8", newline="") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=[
                "timestamp",
                "perfil",
                "data_referencia",
                "nome",
                "email",
                "tabelionato",
                "arquivo_cartao",
                "status",
                "detalhes",
            ],
        )
        if not file_exists:
            writer.writeheader()
        writer.writerow(
            {
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "perfil": profile,
                "data_referencia": run_date.isoformat(),
                "nome": entry.nome,
                "email": entry.email,
                "tabelionato": entry.tabelionato,
                "arquivo_cartao": str(card_path),
                "status": status,
                "detalhes": details,
            }
        )


def build_email_html(entry: BirthdayEntry, run_date: date, settings: dict) -> str:
    reference_label = run_date.strftime("%d/%m/%Y")
    unit_line = entry.tabelionato or settings["body_highlight_fallback"]
    return f"""
<html>
  <body style="font-family: Arial, sans-serif; color: #143040;">
    <p>{settings["body_intro"]}, <strong>{entry.nome}</strong>!</p>
    <p>Data de referencia: {reference_label}</p>
    <p><strong>{unit_line}</strong></p>
    <p>{settings["body_message"]}</p>
    <img src="cid:{DEFAULT_INLINE_CID}" alt="Cartao de aniversario" width="{EMAIL_IMAGE_PIXELS}" height="{EMAIL_IMAGE_PIXELS}" style="display: block; width: 15cm; height: 15cm;" />
  </body>
</html>
""".strip()


def find_outlook_account(outlook, sender_email: str):
    session = outlook.Session
    fallback_account = None
    for account in session.Accounts:
        if fallback_account is None:
            fallback_account = account
        if str(account.SmtpAddress).strip().lower() == sender_email.strip().lower():
            return account
    if fallback_account is not None:
        fallback_email = str(fallback_account.SmtpAddress).strip()
        print(
            f"[AVISO] A conta {sender_email} nao foi encontrada no Outlook. "
            f"Usando {fallback_email} como remetente."
        )
        return fallback_account
    raise RuntimeError("Nenhuma conta configurada foi encontrada no Outlook.")


def outlook_application():
    pythoncom.CoInitialize()
    try:
        return win32com.client.gencache.EnsureDispatch("Outlook.Application")
    except Exception:  # noqa: BLE001
        return win32com.client.Dispatch("Outlook.Application")


def check_outlook(sender_email: str) -> str:
    outlook = outlook_application()
    account = find_outlook_account(outlook, sender_email)
    return str(account.SmtpAddress).strip()


def create_outlook_mail(
    sender_email: str,
    recipient_email: str,
    fixed_bcc_email: str,
    subject: str,
    html_body: str,
    image_path: Path,
    should_send: bool,
) -> None:
    outlook = outlook_application()
    mail = outlook.CreateItem(0)
    account = find_outlook_account(outlook, sender_email)
    bcc_targets = []
    for email in (recipient_email, fixed_bcc_email):
        normalized = str(email).strip()
        if normalized and normalized.lower() not in {target.lower() for target in bcc_targets}:
            bcc_targets.append(normalized)

    mail.SendUsingAccount = account
    # Mantem o envio tecnico no campo "Para" e protege os destinatarios reais no Cco.
    mail.To = sender_email
    mail.BCC = "; ".join(bcc_targets)
    mail.Subject = subject
    mail.HTMLBody = html_body

    attachment = mail.Attachments.Add(str(image_path))
    accessor = attachment.PropertyAccessor
    accessor.SetProperty(MAPI_CONTENT_ID, DEFAULT_INLINE_CID)
    accessor.SetProperty(MAPI_ATTACHMENT_HIDDEN, True)

    if should_send:
        mail.Send()
    else:
        mail.Display()


def process_entry(
    entry: BirthdayEntry,
    args: argparse.Namespace,
    run_date: date,
    log_state: LogState,
    settings: dict,
) -> str:
    log_key = (args.profile, run_date.isoformat(), entry.email.strip().lower())
    if args.send and not args.force and log_key in log_state.sent_keys:
        return "ignorado"
    if not args.send and not args.force and log_key in log_state.draft_keys:
        return "ignorado"
    if not is_valid_email(entry.email):
        append_log(settings["log_path"], args.profile, run_date, entry, Path(""), "erro", "E-mail invalido na planilha.")
        raise ValueError(f"E-mail invalido na planilha: {entry.email}")

    card_path = generate_card(settings["template_path"], settings["output_dir"], entry.nome, run_date, settings["name_box"])
    email_card_path = generate_email_card_image(card_path)
    html_body = build_email_html(entry, run_date, settings)

    try:
        if args.send:
            append_log(
                settings["log_path"],
                args.profile,
                run_date,
                entry,
                card_path,
                "preparando_envio",
                "Envio iniciado. Este registro bloqueia repeticao automatica em caso de interrupcao.",
            )
        create_outlook_mail(
            sender_email=args.sender_email,
            recipient_email=entry.email,
            fixed_bcc_email=args.bcc_email,
            subject=settings["subject"],
            html_body=html_body,
            image_path=email_card_path,
            should_send=args.send,
        )
        status = "enviado" if args.send else "rascunho"
        details = (
            "Email enviado pelo Outlook com destinatarios no Cco."
            if args.send
            else "Email aberto no Outlook para conferencia com destinatarios no Cco."
        )
        append_log(settings["log_path"], args.profile, run_date, entry, card_path, status, details)
        return status
    except Exception as exc:  # noqa: BLE001
        append_log(settings["log_path"], args.profile, run_date, entry, card_path, "erro", str(exc))
        raise


def main() -> int:
    args = parse_args()
    settings = profile_settings(args)

    try:
        run_date = resolve_run_date(args.date)
        with run_lock(args.profile, run_date):
            entries = load_birthdays(settings["excel_path"])
            selected_entries = birthdays_for_date(entries, run_date)
            log_state = read_log_state(settings["log_path"], args.profile)

            if not selected_entries:
                print(f"Nenhum aniversariante encontrado para {run_date.strftime('%d/%m/%Y')} no perfil {args.profile}.")
                return 0

            print(
                f"Aniversariantes encontrados para {run_date.strftime('%d/%m/%Y')} "
                f"no perfil {args.profile}: {len(selected_entries)}"
            )

            error_count = 0
            skipped_count = 0
            for entry in selected_entries:
                try:
                    status = process_entry(entry, args, run_date, log_state, settings)
                    if status == "ignorado":
                        skipped_count += 1
                    print(f"- {entry.nome} <{entry.email}>: {status}")
                except Exception as exc:  # noqa: BLE001
                    error_count += 1
                    print(f"- {entry.nome} <{entry.email}>: erro -> {exc}")

            if skipped_count:
                print(f"[INFO] {skipped_count} aniversariante(s) ignorado(s) por ja constarem no historico do dia.")
            if error_count:
                print(f"[ERRO] {error_count} aniversariante(s) nao foram processados.")
                return 1

            return 0
    except Exception as exc:  # noqa: BLE001
        print(f"[ERRO] {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

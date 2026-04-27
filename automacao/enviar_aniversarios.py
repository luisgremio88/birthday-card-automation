from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import win32com.client


PROJECT_ROOT = Path(__file__).resolve().parents[1]
FONT_PATH = Path(r"C:\Windows\Fonts\tahomabd.ttf")
EMAIL_IMAGE_PIXELS = 567
DEFAULT_SENDER = "luis.dias@cnbrs.org.br"
DEFAULT_INLINE_CID = "cartao_aniversario"
REQUIRED_COLUMNS = ["nome", "tabelionato", "email", "data_aniversario"]
MAPI_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
MAPI_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

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
        "name_box": {"x": 430, "y": 665, "width": 1115, "height": 145, "align": "center"},
    },
}


@dataclass
class BirthdayEntry:
    nome: str
    tabelionato: str
    email: str
    data_aniversario: str


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
    if re.fullmatch(r"\d{2}/\d{2}", text):
        return text
    for pattern in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, pattern).strftime("%d/%m")
        except ValueError:
            continue
    return text[:5]


def load_birthdays(excel_path: Path) -> list[BirthdayEntry]:
    if not excel_path.exists():
        raise FileNotFoundError(f"Planilha nao encontrada: {excel_path}")

    frame = pd.read_excel(excel_path)
    frame = frame.rename(columns={column: normalize_header(column) for column in frame.columns})
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError(f"Colunas obrigatorias ausentes na planilha: {', '.join(missing)}")

    frame = frame[REQUIRED_COLUMNS].copy()
    for column in ("nome", "tabelionato", "email"):
        frame[column] = frame[column].fillna("").astype(str).str.strip()
    frame["data_aniversario"] = frame["data_aniversario"].map(format_birthday_cell)
    frame = frame[(frame["nome"] != "") & (frame["email"] != "") & (frame["data_aniversario"] != "")]

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


def read_log(log_path: Path, profile: str) -> set[tuple[str, str, str]]:
    if not log_path.exists():
        return set()
    with log_path.open("r", encoding="utf-8", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        return {
            (row.get("perfil", ""), row.get("data_referencia", ""), row.get("email", "").strip().lower())
            for row in reader
            if row.get("status") == "enviado"
        }


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
    for account in session.Accounts:
        if str(account.SmtpAddress).strip().lower() == sender_email.strip().lower():
            return account
    raise RuntimeError(f"A conta {sender_email} nao foi encontrada no Outlook.")


def create_outlook_mail(
    sender_email: str,
    recipient_email: str,
    subject: str,
    html_body: str,
    image_path: Path,
    should_send: bool,
) -> None:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    account = find_outlook_account(outlook, sender_email)

    mail.SendUsingAccount = account
    mail.To = recipient_email
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
    existing_sent: set[tuple[str, str, str]],
    settings: dict,
) -> str:
    log_key = (args.profile, run_date.isoformat(), entry.email.strip().lower())
    if not args.force and log_key in existing_sent:
        return "ignorado"

    card_path = generate_card(settings["template_path"], settings["output_dir"], entry.nome, run_date, settings["name_box"])
    email_card_path = generate_email_card_image(card_path)
    html_body = build_email_html(entry, run_date, settings)

    try:
        create_outlook_mail(
            sender_email=args.sender_email,
            recipient_email=entry.email,
            subject=settings["subject"],
            html_body=html_body,
            image_path=email_card_path,
            should_send=args.send,
        )
        status = "enviado" if args.send else "rascunho"
        details = "Email enviado pelo Outlook." if args.send else "Email aberto no Outlook para conferencia."
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
        entries = load_birthdays(settings["excel_path"])
        selected_entries = birthdays_for_date(entries, run_date)
        existing_sent = read_log(settings["log_path"], args.profile)
    except Exception as exc:  # noqa: BLE001
        print(f"[ERRO] {exc}")
        return 1

    if not selected_entries:
        print(f"Nenhum aniversariante encontrado para {run_date.strftime('%d/%m/%Y')} no perfil {args.profile}.")
        return 0

    print(
        f"Aniversariantes encontrados para {run_date.strftime('%d/%m/%Y')} "
        f"no perfil {args.profile}: {len(selected_entries)}"
    )

    for entry in selected_entries:
        try:
            status = process_entry(entry, args, run_date, existing_sent, settings)
            print(f"- {entry.nome} <{entry.email}>: {status}")
        except Exception as exc:  # noqa: BLE001
            print(f"- {entry.nome} <{entry.email}>: erro -> {exc}")

    return 0


if __name__ == "__main__":
    sys.exit(main())

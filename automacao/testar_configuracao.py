from __future__ import annotations

import argparse
import sys
from datetime import date

from enviar_aniversarios import (
    DEFAULT_SENDER,
    PROFILE_CONFIG,
    birthdays_for_date,
    check_outlook,
    load_birthdays,
    resolve_run_date,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Testa configuracao sem abrir rascunho e sem enviar e-mail.")
    parser.add_argument("--profile", choices=sorted(PROFILE_CONFIG.keys()), default="associado")
    parser.add_argument("--date", help="Data de teste no formato YYYY-MM-DD ou DD/MM.")
    parser.add_argument("--sender-email", default=DEFAULT_SENDER)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    settings = PROFILE_CONFIG[args.profile]
    run_date = resolve_run_date(args.date) if args.date else date.today()
    errors = []

    print(f"Perfil: {args.profile}")
    print(f"Data de referencia: {run_date.strftime('%d/%m/%Y')}")

    for label, path in (
        ("Planilha", settings["excel_path"]),
        ("Template", settings["template_path"]),
    ):
        if path.exists():
            print(f"[OK] {label}: {path}")
        else:
            errors.append(f"{label} nao encontrado: {path}")
            print(f"[ERRO] {label} nao encontrado: {path}")

    try:
        entries = load_birthdays(settings["excel_path"])
        selected_entries = birthdays_for_date(entries, run_date)
        print(f"[OK] Registros validos na planilha: {len(entries)}")
        print(f"[OK] Aniversariantes na data: {len(selected_entries)}")
        for entry in selected_entries[:10]:
            print(f" - {entry.nome} <{entry.email}>")
    except Exception as exc:  # noqa: BLE001
        errors.append(str(exc))
        print(f"[ERRO] Leitura da planilha: {exc}")

    try:
        account_email = check_outlook(args.sender_email)
        print(f"[OK] Outlook acessivel. Conta usada: {account_email}")
    except Exception as exc:  # noqa: BLE001
        errors.append(str(exc))
        print(f"[ERRO] Outlook: {exc}")

    if errors:
        print("\nResultado: existem pendencias para corrigir.")
        return 1

    print("\nResultado: configuracao pronta.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

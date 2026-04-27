from __future__ import annotations

import argparse
from pathlib import Path

from psd_tools import PSDImage


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PSD_ROOT = Path(r"C:\Users\Suporte\OneDrive - CNB RS")

PROFILE_OUTPUTS = {
    "associado": {
        "template_png": PROJECT_ROOT / "templates" / "cartao_base_limpo_associado.png",
        "public_png": PROJECT_ROOT / "frontend" / "public" / "templates" / "cartao-associado.png",
    },
    "diretoria": {
        "template_png": PROJECT_ROOT / "templates" / "cartao_base_limpo_diretoria.png",
        "public_png": PROJECT_ROOT / "frontend" / "public" / "templates" / "cartao-diretoria.png",
    },
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrai uma base limpa do PSD de cartao de aniversario.")
    parser.add_argument("--profile", choices=sorted(PROFILE_OUTPUTS.keys()), default="associado")
    parser.add_argument("--psd", type=Path, help="Caminho do PSD. Se omitido, procura automaticamente.")
    parser.add_argument("--hidden-layer", default="Web", help="Nome da camada que sera ocultada.")
    return parser.parse_args()


def resolve_psd_path(profile: str, custom_path: Path | None) -> Path:
    if custom_path:
        if not custom_path.exists():
            raise FileNotFoundError(f"PSD nao encontrado: {custom_path}")
        return custom_path

    marker = "1_" if profile == "associado" else "2_"
    return next(path for path in PSD_ROOT.rglob("*.psd") if path.name.startswith(marker))


def main() -> None:
    args = parse_args()
    psd_path = resolve_psd_path(args.profile, args.psd)
    outputs = PROFILE_OUTPUTS[args.profile]

    psd = PSDImage.open(psd_path)
    for layer in psd.descendants():
        if layer.name == args.hidden_layer:
            layer.visible = False

    image = psd.composite(force=True)
    for output_path in outputs.values():
        output_path.parent.mkdir(parents=True, exist_ok=True)
        image.save(output_path)
        print(output_path)


if __name__ == "__main__":
    main()

from __future__ import annotations

import argparse
from pathlib import Path

from PIL import ImageDraw
from psd_tools import PSDImage


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PSD_ROOT = Path(r"C:\Users\Suporte\OneDrive - CNB RS")

PROFILE_OUTPUTS = {
    "associado": {
        "template_png": PROJECT_ROOT / "templates" / "cartao_base_limpo_associado.png",
        "public_png": PROJECT_ROOT / "frontend" / "public" / "templates" / "cartao-associado.png",
        "hidden_layers": ["Web"],
    },
    "diretoria": {
        "template_png": PROJECT_ROOT / "templates" / "cartao_base_limpo_diretoria.png",
        "public_png": PROJECT_ROOT / "frontend" / "public" / "templates" / "cartao-diretoria.png",
        "hidden_layers": ["Rettangolo 1", "Web"],
    },
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrai uma base limpa do PSD de cartao de aniversario.")
    parser.add_argument("--profile", choices=sorted(PROFILE_OUTPUTS.keys()), default="associado")
    parser.add_argument("--psd", type=Path, help="Caminho do PSD. Se omitido, procura automaticamente.")
    parser.add_argument("--hidden-layer", action="append", help="Nome de camada extra que sera ocultada.")
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
    hidden_layers = set(outputs["hidden_layers"])
    hidden_layers.update(args.hidden_layer or [])

    psd = PSDImage.open(psd_path)
    for layer in psd.descendants():
        if layer.name in hidden_layers:
            layer.visible = False

    image = psd.composite(force=True)
    if args.profile == "diretoria":
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle(
            (70, 134, 1930, 2140),
            radius=150,
            outline="white",
            width=18,
        )
    for output_key in ("template_png", "public_png"):
        output_path = outputs[output_key]
        output_path.parent.mkdir(parents=True, exist_ok=True)
        image.save(output_path)
        print(output_path)


if __name__ == "__main__":
    main()

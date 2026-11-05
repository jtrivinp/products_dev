"""Adaptador de consola provisional; la GUI lo sustituirá para el usuario final."""

import argparse
from pathlib import Path
from .engine import process


def main() -> None:
    parser = argparse.ArgumentParser(description="Motor de propuestas de compensacion SAP")
    parser.add_argument("partidas", type=Path); parser.add_argument("retenciones_legacy", nargs="?", type=Path)
    parser.add_argument("--modo", choices=["abiertas", "domiciliados", "masmenos"], default="abiertas")
    parser.add_argument("--retenciones", type=Path); parser.add_argument("--clases", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=Path("PARTIDAS_PROCESADAS.xlsx"))
    parser.add_argument("--ventana-dias", type=int, default=31); parser.add_argument("--clase-pago-domiciliado")
    parser.add_argument("--tolerancia-centimos", type=int, default=3); parser.add_argument("--max-facturas-domiciliado", type=int, default=4)
    args = parser.parse_args(); retention = args.retenciones or args.retenciones_legacy
    stats = process(args.modo, args.partidas, args.output, retention, args.clases, args.ventana_dias, args.clase_pago_domiciliado, args.tolerancia_centimos, args.max_facturas_domiciliado)
    print(f"Modo: {args.modo}\nSalida: {args.output}\nPartidas propuestas: {stats['proposed_rows']}\nGrupos: {stats['groups']}")
    for kind, count in sorted(stats["by_type"].items()): print(f"  {kind}: {count}")


if __name__ == "__main__": main()

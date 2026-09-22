"""CLI: populate a fixed PPTX template from JSON data or a UUID workspace.

Direct data mode:
    python main.py --template templates/sales_report_template.pptx \
                   --data data/example.json \
                   --output output/direct_mode.pptx

UUID workspace mode (content.json + image files resolved by stem):
    python main.py --template templates/sales_report_template.pptx \
                   --workspace workspace/<uuid4> \
                   --output output/workspace_mode.pptx
"""
import argparse
from pathlib import Path

from src.renderer import load_data, render
from src.workspace_loader import load_workspace, render_workspace


def main():
    parser = argparse.ArgumentParser(description="Render a PPTX template with JSON data.")
    parser.add_argument("--template", required=True, help="Path to the .pptx template")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--data", help="Path to the JSON data file (direct mode)")
    source.add_argument("--workspace",
                        help="Path to a UUID4 workspace folder (workspace mode)")
    parser.add_argument("--output", required=True, help="Path for the generated .pptx")
    args = parser.parse_args()

    if args.workspace:
        ws = load_workspace(args.workspace)
        result = render_workspace(args.template, ws, args.output)
        loaded_from = f"workspace {Path(ws['path']).name}"
    else:
        data, data_dir = load_data(args.data)
        result = render(args.template, data, args.output,
                        base_dirs=(str(data_dir), "."))
        loaded_from = str(args.data)

    print(f"Output written: {args.output}  (from {loaded_from})")
    print(f"  text keys replaced : {', '.join(result['text_keys']) or '-'}")
    print(f"  image keys replaced: {', '.join(result['image_keys']) or '-'}")


if __name__ == "__main__":
    main()

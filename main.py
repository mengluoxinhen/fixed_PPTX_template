"""CLI: populate a fixed PPTX template with JSON data.

    python main.py --template templates/sales_report_template.pptx \
                   --data data/example.json \
                   --output output/sales_report.pptx
"""
import argparse
from pathlib import Path

from src.renderer import load_data, render


def main():
    parser = argparse.ArgumentParser(description="Render a PPTX template with JSON data.")
    parser.add_argument("--template", required=True, help="Path to the .pptx template")
    parser.add_argument("--data", required=True, help="Path to the JSON data file")
    parser.add_argument("--output", required=True, help="Path for the generated .pptx")
    args = parser.parse_args()

    data, data_dir = load_data(args.data)
    # Relative image paths resolve against the data dir, then the CWD.
    result = render(args.template, data, args.output, base_dirs=(str(data_dir), "."))

    print(f"Output written: {args.output}")
    print(f"  text keys replaced : {', '.join(result['text_keys']) or '-'}")
    print(f"  image keys replaced: {', '.join(result['image_keys']) or '-'}")


if __name__ == "__main__":
    main()

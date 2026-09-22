"""Create (or refresh) the test UUID workspace.

    python build_test_workspace.py

Creates:
    workspace/550e8400-e29b-41d4-a716-446655440000/
        content.json   text values ONLY (title, sales) -> proves other
                       placeholders stay unchanged
        product.png    resolved by {{image:product}}
        logo.jpg       resolved by {{image:logo}} (.jpg -> extension-free key)
        notes_readme.txt  non-image file, must be ignored by discovery

There is deliberately NO product_image.* and NO missing_image.* file.
"""
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WORKSPACE_ROOT = ROOT / "workspace"
TEST_WORKSPACE_ID = "550e8400-e29b-41d4-a716-446655440000"  # a real UUID4
WORKSPACE_PATH = WORKSPACE_ROOT / TEST_WORKSPACE_ID

CONTENT = {
    "title": "2026年销售报告",
    "sales": "1258万元",
}


def main():
    WORKSPACE_PATH.mkdir(parents=True, exist_ok=True)

    with open(WORKSPACE_PATH / "content.json", "w", encoding="utf-8") as f:
        json.dump(CONTENT, f, ensure_ascii=False, indent=2)

    shutil.copyfile(ROOT / "assets" / "product.png", WORKSPACE_PATH / "product.png")
    shutil.copyfile(ROOT / "assets" / "logo.jpg", WORKSPACE_PATH / "logo.jpg")
    (WORKSPACE_PATH / "notes_readme.txt").write_text(
        "non-image file in workspace; must be ignored by asset discovery\n",
        encoding="utf-8",
    )

    print(f"Workspace written: {WORKSPACE_PATH}")


if __name__ == "__main__":
    main()

import argparse
import csv
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import DefaultDict, Iterable, List

import pandas as pd
import pytesseract
from pdf2image import convert_from_path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_POPPLER_PATH = PROJECT_ROOT / "poppler" / "poppler-25.12.0" / "Library" / "bin"
DEFAULT_TESSERACT_PATH = PROJECT_ROOT / "tesseract" / "tesseract.exe"
PART_NUMBER_PATTERN = re.compile(r"^[A-Z0-9]+(?:-[A-Z0-9]+)+[A-Z0-9]*$")
PART_NUMBER_SEARCH = re.compile(r"[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+[A-Z0-9]*")


@dataclass
class ShipmentRow:
    page_number: int
    quantity: int
    po_number: str
    packing_slip: str
    part_number: str
    raw_line: str


def configure_tesseract(tesseract_path: Path) -> None:
    pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)


def normalize_qty(token: str) -> int:
    normalized = token.upper().replace("I", "1").replace("L", "1").replace("|", "1")
    return int(re.sub(r"[^0-9]", "", normalized))


def normalize_part_number(remainder: str) -> str:
    tokens = remainder.split()
    if not tokens:
        return ""

    def clean_candidate(value: str) -> str:
        cleaned = re.sub(r"[^A-Z0-9\-]", "", value.upper())
        cleaned = re.sub(r"(?<=-)[IL](?=\d|$)", "1", cleaned)
        cleaned = re.sub(r"(?<=\d)[IL](?=-)", "1", cleaned)
        return cleaned

    candidates = [
        clean_candidate(tokens[0]),
        clean_candidate("".join(tokens[:2])),
        clean_candidate("".join(tokens[:3])),
        clean_candidate("".join(tokens)),
    ]

    for candidate in candidates:
        if PART_NUMBER_PATTERN.fullmatch(candidate):
            return candidate
        match = PART_NUMBER_SEARCH.search(candidate)
        if match:
            return match.group(0)

    return clean_candidate(tokens[0])


def parse_shipment_line(cleaned: str, page_number: int) -> ShipmentRow | None:
    tokens = cleaned.split()
    if len(tokens) < 5:
        return None

    qty_token = tokens[0]
    if not re.fullmatch(r"[0-9Il|]+", qty_token):
        return None

    try:
        quantity = normalize_qty(qty_token)
    except ValueError:
        return None

    index = 1
    po_parts = []
    while index < len(tokens) and tokens[index].isdigit():
        po_parts.append(tokens[index])
        if "".join(po_parts).startswith("700") and len("".join(po_parts)) >= 8:
            break
        index += 1

    po_number = "".join(po_parts)
    if not po_number.startswith("700") or len(po_number) < 8:
        return None

    index += 1
    if index < len(tokens) and tokens[index].upper() == "SO":
        index += 1

    if index >= len(tokens) or not tokens[index].isdigit():
        return None
    packing_slip = tokens[index]
    index += 1

    remainder = " ".join(tokens[index:])
    part_number = normalize_part_number(remainder)
    if not part_number:
        return None

    return ShipmentRow(
        page_number=page_number,
        quantity=quantity,
        po_number=po_number,
        packing_slip=packing_slip,
        part_number=part_number,
        raw_line=cleaned,
    )


def extract_rows_from_text(text: str, page_number: int) -> List[ShipmentRow]:
    rows: List[ShipmentRow] = []
    for raw_line in text.splitlines():
        cleaned = " ".join(raw_line.split())
        if not cleaned or "Qty Shipped" in cleaned or "P.O. Number" in cleaned:
            continue
        row = parse_shipment_line(cleaned, page_number)
        if row:
            rows.append(row)
    return rows


def ocr_pdf(pdf_path: Path, poppler_path: Path, dpi: int) -> Iterable[str]:
    images = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        poppler_path=str(poppler_path),
    )
    for image in images:
        yield pytesseract.image_to_string(image, lang="eng", config="--psm 6")


def aggregate_quantities(rows: Iterable[ShipmentRow]) -> DefaultDict[str, int]:
    totals: DefaultDict[str, int] = defaultdict(int)
    for row in rows:
        totals[row.part_number] += row.quantity
    return totals


def aggregate_quantities_by_po(rows: Iterable[ShipmentRow]) -> DefaultDict[tuple[str, str], int]:
    totals: DefaultDict[tuple[str, str], int] = defaultdict(int)
    for row in rows:
        totals[(row.po_number, row.part_number)] += row.quantity
    return totals


def write_csv(output_path: Path, totals: DefaultDict[str, int]) -> None:
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["part_number", "total_qty_shipped"])
        for part_number in sorted(totals):
            writer.writerow([part_number, totals[part_number]])


def write_excel(
    output_path: Path,
    rows: List[ShipmentRow],
    part_totals: DefaultDict[str, int],
    po_part_totals: DefaultDict[tuple[str, str], int],
) -> None:
    part_totals_frame = pd.DataFrame(
        [
            {"part_number": part_number, "total_qty_shipped": total}
            for part_number, total in sorted(part_totals.items())
        ]
    )
    po_part_totals_frame = pd.DataFrame(
        [
            {"po_number": po_number, "part_number": part_number, "total_qty_shipped": total}
            for (po_number, part_number), total in sorted(po_part_totals.items())
        ]
    )
    detail_frame = pd.DataFrame(
        [
            {
                "page_number": row.page_number,
                "po_number": row.po_number,
                "packing_slip": row.packing_slip,
                "part_number": row.part_number,
                "quantity": row.quantity,
                "raw_line": row.raw_line,
            }
            for row in rows
        ]
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        part_totals_frame.to_excel(writer, sheet_name="Part Totals", index=False)
        po_part_totals_frame.to_excel(writer, sheet_name="PO Part Totals", index=False)
        detail_frame.to_excel(writer, sheet_name="Detail", index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read a shipment PDF and total quantity shipped by part number."
    )
    parser.add_argument("pdf_path", help="Path to the PDF report to process.")
    parser.add_argument(
        "--output",
        help="Optional CSV output path. Defaults to '<pdf name>_part_totals.csv'.",
    )
    parser.add_argument(
        "--excel-output",
        help="Optional Excel output path. Defaults to '<pdf name>_part_totals.xlsx'.",
    )
    parser.add_argument("--dpi", type=int, default=220, help="OCR render DPI. Default: 220.")
    parser.add_argument(
        "--poppler-path",
        default=str(DEFAULT_POPPLER_PATH),
        help="Path to the Poppler bin folder.",
    )
    parser.add_argument(
        "--tesseract-path",
        default=str(DEFAULT_TESSERACT_PATH),
        help="Path to tesseract.exe.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    pdf_path = Path(args.pdf_path).expanduser().resolve()
    if not pdf_path.exists():
        print(f"PDF not found: {pdf_path}")
        return 1

    poppler_path = Path(args.poppler_path)
    tesseract_path = Path(args.tesseract_path)
    if not poppler_path.exists():
        print(f"Poppler path not found: {poppler_path}")
        return 1
    if not tesseract_path.exists():
        print(f"Tesseract path not found: {tesseract_path}")
        return 1

    configure_tesseract(tesseract_path)

    rows: List[ShipmentRow] = []
    for page_number, text in enumerate(ocr_pdf(pdf_path, poppler_path, args.dpi), start=1):
        rows.extend(extract_rows_from_text(text, page_number))

    if not rows:
        print("No shipment rows were detected in the PDF.")
        return 1

    totals = aggregate_quantities(rows)
    po_part_totals = aggregate_quantities_by_po(rows)
    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else pdf_path.with_name(f"{pdf_path.stem}_part_totals.csv")
    )
    excel_output_path = (
        Path(args.excel_output).expanduser().resolve()
        if args.excel_output
        else pdf_path.with_name(f"{pdf_path.stem}_part_totals.xlsx")
    )
    write_csv(output_path, totals)
    write_excel(excel_output_path, rows, totals, po_part_totals)

    print(f"Rows detected: {len(rows)}")
    print(f"Unique part numbers: {len(totals)}")
    print(f"Saved totals to: {output_path}")
    print(f"Saved Excel totals to: {excel_output_path}")
    print("\nTop totals:")
    for part_number, total in sorted(totals.items(), key=lambda item: (-item[1], item[0]))[:15]:
        print(f"  {part_number}: {total}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

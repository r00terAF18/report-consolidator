import re
import shutil
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False


def write_cell_safe(ws, coord: str, val: Any) -> bool:
    """Safely write to a cell coordinate, resolving merged cells to their top-left anchor."""
    try:
        cell = ws[coord]
        if type(cell).__name__ == "MergedCell":
            for rng in ws.merged_cells.ranges:
                if coord in rng:
                    ws.cell(row=rng.min_row, column=rng.min_col, value=val)
                    return True
        cell.value = val
        return True
    except Exception:
        return False


def find_target_cell(ws, label_keywords: list[str], max_row: int = 50, col_offset: int = 1) -> str | None:
    """Scan the sheet to locate the cell adjacent to a matching label."""
    for row in range(1, max_row + 1):
        for col in range(1, 20):
            cell = ws.cell(row=row, column=col)
            val = str(cell.value or "").strip().lower()
            if any(k in val for k in label_keywords):
                target_col = col + col_offset
                return f"{get_column_letter(target_col)}{row}"
    return None


class Consolidator:
    def __init__(self, folder_path: Path):
        self.folder_path = folder_path

    def locate_files(self) -> dict[str, Path | None]:
        files: dict[str, Path | None] = {
            "target": None,
            "company_ddr": None,
            "pob": None,
            "mdr": None,        # Mechanical Daily Report
            "edr": None,        # Electrical Daily Report (.xlsx or .pdf)
            "mud_report": None, # Mud Daily Report
        }

        # Check Excel files
        for path in self.folder_path.glob("*.xlsx"):
            if path.name.startswith("~$"):
                continue
            name_lower = path.name.lower()
            if "tool poosher" in name_lower or "toolpusher" in name_lower:
                files["target"] = path
            elif "pob" in name_lower:
                files["pob"] = path
            elif "mud report" in name_lower or "mud_report" in name_lower:
                files["mud_report"] = path
            elif "mdr" in name_lower:
                files["mdr"] = path
            elif "edr" in name_lower:
                files["edr"] = path
            elif "ddr" in name_lower:
                files["company_ddr"] = path

        # Check PDF files if EDR / Mud Report were not in Excel
        for path in self.folder_path.glob("*.pdf"):
            name_lower = path.name.lower()
            if not files["edr"] and "edr" in name_lower:
                files["edr"] = path
            elif not files["mud_report"] and "mud report" in name_lower:
                files["mud_report"] = path
            elif not files["mdr"] and "mdr" in name_lower:
                files["mdr"] = path

        return files

    def extract_pob(self, path: Path) -> dict[str, Any]:
        result: dict[str, Any] = {"total": None}
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                for idx, cell in enumerate(row):
                    if isinstance(cell, str) and any(w in cell.lower() for w in ["total pob", "total person", "total"]):
                        for val in row[idx + 1:]:
                            if isinstance(val, (int, float)) and 0 < val < 600:
                                result["total"] = int(val)
                                wb.close()
                                return result
            wb.close()
        except Exception:
            pass
        return result

    def extract_mechanical_report(self, path: Path) -> dict[str, Any]:
        """Extract Generator hours, fuel, and mechanical operations from MDR."""
        result: dict[str, Any] = {
            "gen_hours": {},
            "fuel_stock": None,
            "fuel_consumed": None,
            "remarks": [],
        }
        if path.suffix.lower() == ".xlsx":
            try:
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        row_vals = [c for c in row if c is not None]
                        row_str = " ".join(str(c).lower() for c in row_vals)

                        # Fuel data
                        if "fuel" in row_str or "diesel" in row_str:
                            for idx, val in enumerate(row):
                                if isinstance(val, str):
                                    if "consumed" in val.lower() or "daily" in val.lower():
                                        for v in row[idx + 1: idx + 4]:
                                            if isinstance(v, (int, float)) and v > 0:
                                                result["fuel_consumed"] = float(v)
                                                break
                                    elif "stock" in val.lower() or "remain" in val.lower():
                                        for v in row[idx + 1: idx + 4]:
                                            if isinstance(v, (int, float)) and v > 0:
                                                result["fuel_stock"] = float(v)
                                                break

                        # Generator / Engine Hours (e.g. Gen 1, Gen 2, Engine #1)
                        gen_match = re.search(r"(gen(?:erator)?|engine)\s*#?\s*([1-4])", row_str)
                        if gen_match:
                            gen_num = gen_match.group(2)
                            for val in row_vals:
                                if isinstance(val, (int, float)) and 0 < val < 100000:
                                    result["gen_hours"][f"gen_{gen_num}"] = float(val)
                                    break

                        # Maintenance remarks / operations
                        if any(k in row_str for k in ["repair", "service", "maintenance", "overhaul", "check", "inspect"]):
                            for c in row_vals:
                                if isinstance(c, str) and len(c.strip()) > 15:
                                    result["remarks"].append(c.strip())
                wb.close()
            except Exception:
                pass

        elif path.suffix.lower() == ".pdf" and HAS_PYPDF:
            try:
                reader = PdfReader(str(path))
                text = "\n".join(page.extract_text() or "" for page in reader.pages)
                for line in text.splitlines():
                    if any(k in line.lower() for k in ["repair", "service", "maintenance", "inspection"]):
                        result["remarks"].append(line.strip())
            except Exception:
                pass

        return result

    def extract_electrical_report(self, path: Path) -> dict[str, Any]:
        """Extract Electrical equipment status, load, and remarks from EDR."""
        result: dict[str, Any] = {
            "gen_loads": {},
            "remarks": [],
        }

        if path.suffix.lower() == ".xlsx":
            try:
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        row_vals = [c for c in row if c is not None]
                        row_str = " ".join(str(c).lower() for c in row_vals)

                        # Electrical loads (kW or Amps)
                        gen_match = re.search(r"gen(?:erator)?\s*#?\s*([1-4])", row_str)
                        if gen_match:
                            gen_num = gen_match.group(1)
                            for val in row_vals:
                                if isinstance(val, (int, float)) and 0 < val < 3000:
                                    result["gen_loads"][f"gen_{gen_num}"] = float(val)
                                    break

                        # Electrical activities
                        if any(k in row_str for k in ["electrical", "motor", "mcc", "lighting", "sensor", "breaker", "cable"]):
                            for c in row_vals:
                                if isinstance(c, str) and len(c.strip()) > 15:
                                    result["remarks"].append(c.strip())
                wb.close()
            except Exception:
                pass

        elif path.suffix.lower() == ".pdf" and HAS_PYPDF:
            try:
                reader = PdfReader(str(path))
                text = "\n".join(page.extract_text() or "" for page in reader.pages)
                for line in text.splitlines():
                    line_clean = line.strip()
                    if any(k in line_clean.lower() for k in ["mcc", "generator", "motor", "lighting", "repaired", "checked", "troubleshoot"]):
                        if len(line_clean) > 15:
                            result["remarks"].append(line_clean)
            except Exception:
                pass

        return result

    def extract_company_ddr(self, path: Path) -> dict[str, Any]:
        result: dict[str, Any] = {"depth": None}
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                for idx, cell in enumerate(row):
                    if isinstance(cell, str) and any(k in cell.lower() for k in ["current depth", "present depth", "hole depth", "depth (m)"]):
                        for val in row[idx + 1: idx + 5]:
                            if isinstance(val, (int, float)) and val > 0:
                                result["depth"] = round(float(val), 2)
                                wb.close()
                                return result
            wb.close()
        except Exception:
            pass
        return result

    def run(self) -> tuple[bool, list[str]]:
        logs: list[str] = []

        if not self.folder_path.is_dir():
            return False, [f"Error: Path '{self.folder_path}' is not a valid directory."]

        files = self.locate_files()
        target = files["target"]

        if not target:
            return False, [f"Target Error: No Toolpusher DDR (*TOOL POOSHER*.xlsx) found in '{self.folder_path.name}'."]

        logs.append(f"Master target located: {target.name}")

        # Extract data from available sources
        extracted_data: dict[str, Any] = {}

        for source_key, source_name in [
            ("pob", "POB"),
            ("mdr", "Mechanical Report (MDR)"),
            ("edr", "Electrical Report (EDR)"),
            ("company_ddr", "Company DDR"),
        ]:
            file_path = files[source_key]
            if file_path:
                logs.append(f"Reading {source_name}: {file_path.name}")
                if source_key == "pob":
                    extracted_data["pob"] = self.extract_pob(file_path)
                elif source_key == "mdr":
                    extracted_data["mdr"] = self.extract_mechanical_report(file_path)
                elif source_key == "edr":
                    extracted_data["edr"] = self.extract_electrical_report(file_path)
                elif source_key == "company_ddr":
                    extracted_data["ddr"] = self.extract_company_ddr(file_path)
            else:
                logs.append(f"Warning: No {source_name} file found. Skipping.")

        # Safety backup
        backup_path = target.with_suffix(".xlsx.bak")
        try:
            shutil.copy2(target, backup_path)
            logs.append(f"Safety backup saved: {backup_path.name}")
        except Exception as e:
            return False, [f"File Error: Could not generate backup file: {e}"]

        # Write to Toolpusher DDR
        try:
            wb = openpyxl.load_workbook(target)
            ws = wb.active

            updated_fields = 0

            # 1. POB
            pob_val = extracted_data.get("pob", {}).get("total")
            if pob_val is not None:
                coord = find_target_cell(ws, ["pob", "total pob", "total person"]) or "C10"
                if write_cell_safe(ws, coord, pob_val):
                    logs.append(f"Updated POB Total -> {pob_val} (at {coord})")
                    updated_fields += 1

            # 2. Drilling Depth
            depth_val = extracted_data.get("ddr", {}).get("depth")
            if depth_val is not None:
                coord = find_target_cell(ws, ["current depth", "present depth", "depth"]) or "H5"
                if write_cell_safe(ws, coord, depth_val):
                    logs.append(f"Updated Current Depth -> {depth_val} (at {coord})")
                    updated_fields += 1

            # 3. Mechanical Report Data (Fuel & Generator Hours)
            mdr_data = extracted_data.get("mdr", {})
            if mdr_data.get("fuel_consumed") is not None:
                coord = find_target_cell(ws, ["fuel consumed", "daily fuel", "fuel daily"])
                if coord and write_cell_safe(ws, coord, mdr_data["fuel_consumed"]):
                    logs.append(f"Updated Daily Fuel Consumed -> {mdr_data['fuel_consumed']} (at {coord})")
                    updated_fields += 1

            if mdr_data.get("fuel_stock") is not None:
                coord = find_target_cell(ws, ["fuel stock", "fuel remaining", "fuel balance"])
                if coord and write_cell_safe(ws, coord, mdr_data["fuel_stock"]):
                    logs.append(f"Updated Fuel Remaining -> {mdr_data['fuel_stock']} (at {coord})")
                    updated_fields += 1

            # 4. Remarks (Mechanical & Electrical updates)
            all_remarks = []
            if mdr_data.get("remarks"):
                all_remarks.append("--- MECHANICAL UPDATES ---")
                all_remarks.extend(mdr_data["remarks"][:5])
            if extracted_data.get("edr", {}).get("remarks"):
                all_remarks.append("--- ELECTRICAL UPDATES ---")
                all_remarks.extend(extracted_data["edr"]["remarks"][:5])

            if all_remarks:
                combined_text = "\n".join(all_remarks)
                coord = find_target_cell(ws, ["maintenance", "repair", "remarks", "mechanical/electrical"])
                if coord:
                    write_cell_safe(ws, coord, combined_text)
                    logs.append(f"Updated Maintenance & Electrical Remarks (at {coord})")
                    updated_fields += 1

            out_file = target.with_name(f"{target.stem}_CONSOLIDATED.xlsx")
            wb.save(out_file)
            wb.close()
            logs.append(f"Success: Consolidated report saved as '{out_file.name}' ({updated_fields} fields written).")
            return True, logs
        except Exception as e:
            if "wb" in locals():
                wb.close()
            return False, logs + [f"Execution Error while writing workbook: {e}"]

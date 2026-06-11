
import sys
import re
from pathlib import Path
from datetime import datetime


OLD_TO_NEW = {
    "Properties.Resources.新竹_科目成績單":
        "Properties.Resources.國中評量成績單樣板_科目成績單",

    "Properties.Resources.新竹_領域成績單":
        "Properties.Resources.國中評量成績單樣板_領域成績單",

    "Properties.Resources.新竹_科目及領域成績單_科目組距":
        "Properties.Resources.國中評量成績單樣板_科目及領域成績單_科目組距",

    "Properties.Resources.新竹_科目及領域成績單_領域組距":
        "Properties.Resources.國中評量成績單樣板_科目及領域成績單_領域組距",
}

NEW_RESOURCE_NAMES = [
    "國中評量成績單樣板_科目成績單",
    "國中評量成績單樣板_領域成績單",
    "國中評量成績單樣板_科目及領域成績單_科目組距",
    "國中評量成績單樣板_科目及領域成績單_領域組距",
]

OLD_RESOURCE_NAMES = [
    "新竹_科目成績單",
    "新竹_領域成績單",
    "新竹_科目及領域成績單_科目組距",
    "新竹_科目及領域成績單_領域組距",
]

EXPECTED_WORD_FILES = [
    "國中評量成績單樣板(科目成績單).doc",
    "國中評量成績單樣板(領域成績單).doc",
    "國中評量成績單樣板(科目及領域成績單_科目組距).doc",
    "國中評量成績單樣板(科目及領域成績單_領域組距).doc",
]


def read_text(path: Path) -> str:
    if not path.exists():
        return ""

    for encoding in ("utf-8-sig", "utf-8", "cp950", "big5"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue

    return path.read_text(errors="ignore")


def find_file(project_dir: Path, file_name: str):
    matches = list(project_dir.rglob(file_name))
    return matches[0] if matches else None


def check_printform(project_dir: Path):
    result = []

    print_form = find_file(project_dir, "PrintForm.cs")
    if print_form is None:
        result.append(("ERROR", "PrintForm.cs not found."))
        return result, None

    text = read_text(print_form)

    result.append(("INFO", f"Checked file: {print_form}"))

    for old_ref, new_ref in OLD_TO_NEW.items():
        if new_ref in text:
            result.append(("PASS", f"New reference exists: `{new_ref}`"))
        else:
            result.append(("FAIL", f"New reference missing: `{new_ref}`"))

        if old_ref in text:
            result.append(("FAIL", f"Old reference still exists: `{old_ref}`"))
        else:
            result.append(("PASS", f"Old reference removed: `{old_ref}`"))

    # Check switch-case names remain unchanged
    expected_cases = [
        'case "領域成績單":',
        'case "科目成績單":',
        'case "科目及領域成績單_領域組距":',
        'case "科目及領域成績單_科目組距":',
    ]

    for case_text in expected_cases:
        if case_text in text:
            result.append(("PASS", f"Switch case remains: `{case_text}`"))
        else:
            result.append(("WARN", f"Switch case not found or changed: `{case_text}`"))

    return result, print_form


def check_resources_resx(project_dir: Path):
    result = []

    resx = find_file(project_dir, "Resources.resx")
    if resx is None:
        result.append(("ERROR", "Properties/Resources.resx not found."))
        return result, None

    text = read_text(resx)
    result.append(("INFO", f"Checked file: {resx}"))

    for name in NEW_RESOURCE_NAMES:
        pattern = f'name="{name}"'
        if pattern in text:
            result.append(("PASS", f"New resource exists in Resources.resx: `{name}`"))
        else:
            result.append(("FAIL", f"New resource missing in Resources.resx: `{name}`"))

    for name in OLD_RESOURCE_NAMES:
        pattern = f'name="{name}"'
        if pattern in text:
            result.append(("WARN", f"Old resource name still exists in Resources.resx: `{name}`"))
        else:
            result.append(("PASS", f"Old resource name removed from Resources.resx: `{name}`"))

    for name in NEW_RESOURCE_NAMES:
        duplicated_pattern = f'name="{name}1"'
        if duplicated_pattern in text:
            result.append(("FAIL", f"Duplicated resource name found: `{name}1`"))
        else:
            result.append(("PASS", f"No duplicated resource name: `{name}1`"))

    # Check whether expected doc file names are referenced in resx.
    # This mainly works when resources are stored as file references.
    for doc_name in EXPECTED_WORD_FILES:
        if doc_name in text:
            result.append(("PASS", f"Expected Word file referenced in Resources.resx: `{doc_name}`"))
        else:
            result.append(("WARN", f"Expected Word file name not found in Resources.resx: `{doc_name}`. This may be OK if resources are embedded as base64."))

    return result, resx


def check_resources_designer(project_dir: Path):
    result = []

    designer = find_file(project_dir, "Resources.Designer.cs")
    if designer is None:
        result.append(("ERROR", "Properties/Resources.Designer.cs not found."))
        return result, None

    text = read_text(designer)
    result.append(("INFO", f"Checked file: {designer}"))

    for name in NEW_RESOURCE_NAMES:
        # Usually generated as: internal static byte[] xxx
        if re.search(rf"\bbyte\[\]\s+{re.escape(name)}\b", text):
            result.append(("PASS", f"New byte[] property exists: `{name}`"))
        elif name in text:
            result.append(("WARN", f"New resource name exists but byte[] property pattern not confirmed: `{name}`"))
        else:
            result.append(("FAIL", f"New byte[] property missing: `{name}`"))

    for name in OLD_RESOURCE_NAMES:
        if re.search(rf"\bbyte\[\]\s+{re.escape(name)}\b", text):
            result.append(("WARN", f"Old byte[] property still exists: `{name}`"))
        else:
            result.append(("PASS", f"Old byte[] property not found: `{name}`"))

    for name in NEW_RESOURCE_NAMES:
        if f"{name}1" in text:
            result.append(("FAIL", f"Duplicated designer property found: `{name}1`"))
        else:
            result.append(("PASS", f"No duplicated designer property: `{name}1`"))

    return result, designer


def search_all_old_references(project_dir: Path):
    result = []
    ignored_dirs = {
        ".git", ".vs", "bin", "obj", "packages", "node_modules",
    }

    files_to_scan = []
    for path in project_dir.rglob("*"):
        if not path.is_file():
            continue

        if any(part in ignored_dirs for part in path.parts):
            continue

        if path.suffix.lower() not in {
            ".cs", ".resx", ".config", ".xml", ".txt", ".md"
        }:
            continue

        files_to_scan.append(path)

    for old_ref in OLD_TO_NEW.keys():
        found_locations = []

        for path in files_to_scan:
            text = read_text(path)
            if old_ref in text:
                found_locations.append(path)

        if found_locations:
            for path in found_locations:
                result.append(("FAIL", f"Old reference `{old_ref}` found in: {path}"))
        else:
            result.append(("PASS", f"No old reference found in project: `{old_ref}`"))

    return result


def write_report(project_dir: Path, all_results):
    lines = []
    lines.append("# Word Template Resource Check Report")
    lines.append("")
    lines.append(f"Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Project path: `{project_dir}`")
    lines.append("")

    status_order = ["ERROR", "FAIL", "WARN", "PASS", "INFO"]

    summary = {key: 0 for key in status_order}
    for section_name, results in all_results:
        for status, _ in results:
            summary[status] = summary.get(status, 0) + 1

    lines.append("## Summary")
    lines.append("")
    for key in status_order:
        lines.append(f"- {key}: {summary.get(key, 0)}")
    lines.append("")

    final_pass = summary.get("ERROR", 0) == 0 and summary.get("FAIL", 0) == 0

    if final_pass:
        lines.append("## Final Result")
        lines.append("")
        lines.append("PASS: No blocking issues found.")
    else:
        lines.append("## Final Result")
        lines.append("")
        lines.append("FAIL: Blocking issues found. Please fix ERROR/FAIL items.")
    lines.append("")

    for section_name, results in all_results:
        lines.append(f"## {section_name}")
        lines.append("")
        for status, message in results:
            lines.append(f"- **{status}**: {message}")
        lines.append("")

    report_path = project_dir / "check_word_template_resources_report.md"
    report_path.write_text("\n".join(lines), encoding="utf-8-sig")

    return report_path, final_pass


def main():
    if len(sys.argv) >= 2:
        project_dir = Path(sys.argv[1]).resolve()
    else:
        project_dir = Path.cwd().resolve()

    if not project_dir.exists():
        print(f"Project directory does not exist: {project_dir}")
        sys.exit(2)

    all_results = []

    printform_results, _ = check_printform(project_dir)
    all_results.append(("PrintForm.cs check", printform_results))

    resx_results, _ = check_resources_resx(project_dir)
    all_results.append(("Resources.resx check", resx_results))

    designer_results, _ = check_resources_designer(project_dir)
    all_results.append(("Resources.Designer.cs check", designer_results))

    old_ref_results = search_all_old_references(project_dir)
    all_results.append(("Project-wide old reference check", old_ref_results))

    report_path, final_pass = write_report(project_dir, all_results)

    print(f"Report generated: {report_path}")

    if final_pass:
        print("PASS: No blocking issues found.")
        sys.exit(0)
    else:
        print("FAIL: Blocking issues found. Please check the report.")
        sys.exit(1)


if __name__ == "__main__":
    main()
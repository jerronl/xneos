import argparse
from pathlib import Path
import importlib.resources as pkg_resources
import shutil
import os, sys
import xneos.templates


def copy_template_file(filename, target_dir):
    target_path = Path(target_dir) / filename
    with pkg_resources.path(xneos.templates, filename) as src_path:
        shutil.copy(src_path, target_path)
    return target_path


def ensure_stable_xlwings_dll():
    dll_folder = Path(sys.prefix)

    target = dll_folder / "xlwings64.dll"
    if target.exists():
        return True

    matches = list(dll_folder.glob("xlwings64-*.dll"))
    if not matches:
        print("No matching xlwings64-*.dll found. Did you install xlwings?")
        return False

    shutil.copy(matches[0], target)
    return True


def quickstart_xneos(project_name):
    path = Path(project_name)
    if path.exists():
        print(f"[!] Folder '{project_name}' already exists.")
        return
    os.makedirs(path)

    copy_template_file("xneos_template.xlsm", path)
    copy_template_file("xneos_main.py", path)
    config_dict = {
        "USE UDF SERVER": "True",
        "SHOW CONSOLE": "True",
        "CONDA PATH": str(Path(os.environ.get("CONDA_EXE")).parent.parent.resolve()),
        "UDF MODULES": "xneos_main",
        "CONDA ENV": os.environ.get("CONDA_DEFAULT_ENV"),
        "DEBUG UDFS": "True",
    }
    with open(path / "xlwings.conf", "w", encoding="utf-8") as f:
        for key, val in config_dict.items():
            f.write(f'"{key}","{val}"\n')

    with open(path / "manualstart.bat", "w", encoding="utf-8") as f:
        f.write(f"""
echo start
call {str(Path(os.environ.get("CONDA_EXE")).parent.parent.resolve())}/Scripts/activate.bat
call conda activate {os.environ.get("CONDA_DEFAULT_ENV")}
python xneos_main.py
echo running...                """
        )


    if ensure_stable_xlwings_dll():
        print(f"Project created at: {path.resolve()}")
        print(
            """
    To enable Excel integration, please run:

        xlwings addin install

    This will install the Excel add-in to allow RunPython to work.
            """
        )


def main():
    parser = argparse.ArgumentParser(
        prog="xneos", description="XNEOS Command Line Tools"
    )
    subparsers = parser.add_subparsers(dest="command")

    quick_parser = subparsers.add_parser(
        "quickstart", help="Generate a new xneos Excel+Python project"
    )
    quick_parser.add_argument("project_name", help="Target project folder name")

    args = parser.parse_args()

    if args.command == "quickstart":
        quickstart_xneos(args.project_name)
    else:
        parser.print_help()

# if __name__ == "__main__":
#     main()
# This script provides a command line interface to quickly set up an xneos project.
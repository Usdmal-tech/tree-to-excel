import os
import sys

import pytest

sys.path.insert(
    0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
)
from tree_to_excel import generate_output_path, parse_tree_file, save_to_excel


# ------ Fixture for format without /A (standard tree) -------
@pytest.fixture
def tree_file(tmp_path):
    content = r"""Folder PATH listing for volume xx-xxx xxx
Volume serial number is xxxxxxxx xxxx:xxxx
C:\PROJECTS
├───src
│   ├───core
│   │   └───utils
│   │       └───helpers.py
│   └───tree_to_excel.py
├───docs
│   └───readme.txt
└───tests
    └───test_core.py
"""
    f = tmp_path / "tree_std.txt"
    f.write_text(content, encoding="utf-8")
    return str(f)


# ---------------- Fixture for format with /A ----------------
@pytest.fixture
def tree_file_a(tmp_path):
    content = r"""Folder PATH listing for volume xx-xxx xxx
Volume serial number is xxxxxxxx xxxx:xxxx
C:\PROJECTS
+---src
│   +---core
│   │   \---utils
│   │       \---helpers.py
│   \---tree_to_excel.py
+---docs
│   \---readme.txt
\---tests
    \---test_core.py
"""
    f = tmp_path / "tree_a.txt"
    f.write_text(content, encoding="utf-8")
    return str(f)


# ----------------- Test for standard format -----------------
def test_parse_standard(tree_file):
    items, folders, files = parse_tree_file(tree_file)
    assert folders == 5, f"Expected 5 folders, got {folders}"
    assert files == 4, f"Expected 4 files, got {files}"
    # Check that helpers.py is defined as a file
    helpers = next((it for it in items if it["name"] == "helpers.py"), None)
    assert helpers is not None
    assert helpers["is_dir"] is False
    assert helpers["ancestors"] == ["src", "core", "utils"]


# ------------------- Test for /A format ---------------------
def test_parse_a(tree_file_a):
    items, folders, files = parse_tree_file(tree_file_a)
    assert folders == 5, f"Expected 5 folders, got {folders}"
    assert files == 4, f"Expected 4 files, got {files}"
    helpers = next((it for it in items if it["name"] == "helpers.py"), None)
    assert helpers is not None
    assert helpers["is_dir"] is False


# ----------------- Test for folder with dot -----------------
def test_folder_with_dot_standard(tmp_path):
    # Empty folder with dot in name — should be misidentified as a file
    content = r"""Folder PATH listing for volume xx-xxx xxx
Volume serial number is xxxxxxxx xxxx:xxxx
C:\PROJECTS
├───src
│   └───project.v1
│
└───docs
"""
    f = tmp_path / "tree_dot_folder.txt"
    f.write_text(content, encoding="utf-8")
    items, folders, files = parse_tree_file(str(f))
    # Папка project.v1 не имеет потомков, есть маркер и точка -> будет файлом
    dot_folder = next((it for it in items if it["name"] == "project.v1"), None)
    assert dot_folder is not None
    assert (
        dot_folder["is_dir"] is False
    ), "Empty folder with dot should be misidentified as a file"
    # Check that `docs` is defined as a folder (has children)
    docs = next((it for it in items if it["name"] == "docs"), None)
    assert docs is not None
    assert docs["is_dir"] is True


# --------- Test for empty folder with /A key ----------
def test_empty_folder_a(tmp_path):
    content = r"""Folder PATH listing for volume xx-xxx xxx
Volume serial number is xxxxxxxx xxxx:xxxx
C:\PROJECTS
+---src
│   \---empty_folder
\---docs
"""
    f = tmp_path / "tree_empty_a.txt"
    f.write_text(content, encoding="utf-8")
    items, folders, files = parse_tree_file(str(f))
    empty = next((it for it in items if it["name"] == "empty_folder"), None)
    assert empty is not None
    assert empty["is_dir"] is False, "Empty folder in /A should be treated as a file"


# ------------- Test for file without extension --------------
def test_file_without_extension_standard(tmp_path):
    content = r"""Folder PATH listing for volume xx-xxx xxx
Volume serial number is xxxxxxxx xxxx:xxxx
C:\PROJECTS
├───src
│   └───README
└───docs
"""
    f = tmp_path / "tree_no_ext.txt"
    f.write_text(content, encoding="utf-8")
    items, folders, files = parse_tree_file(str(f))
    readme = next((it for it in items if it["name"] == "README"), None)
    assert readme is not None
    assert (
        readme["is_dir"] is True
    ), "File without extension should be misidentified as a folder"


# --------------- Test for missing root path -----------------
def test_parse_no_root(tmp_path):
    f = tmp_path / "bad.txt"
    f.write_text("Just text with no path", encoding="utf-8")
    with pytest.raises(ValueError, match="Root path not found"):
        parse_tree_file(str(f))


# -------------- Test for generate_output_path ---------------
def test_generate_output_path():
    input_path = os.path.join("C:", "temp", "my_tree.txt")
    out = generate_output_path(input_path, "flat")
    expected = os.path.join(
        os.path.dirname(os.path.abspath(input_path)), "my_tree_flat.xlsx"
    )
    assert out == expected


# ----------- Parametrized test for all save modes -----------
@pytest.mark.parametrize("mode", ["flat", "merged", "merged_full"])
def test_save_to_excel_modes(tree_file, tmp_path, mode):
    items, folders, files = parse_tree_file(tree_file)
    output_file = tmp_path / f"result_{mode}.xlsx"
    save_to_excel(items, str(output_file), mode, folders, files)
    assert output_file.exists()
    assert output_file.stat().st_size > 0


# ---- Parametrized test for all save modes (with /A key) ----
@pytest.mark.parametrize("mode", ["flat", "merged", "merged_full"])
def test_save_to_excel_modes_a(tree_file_a, tmp_path, mode):
    """Проверяет сохранение в Excel для дерева, сгенерированного с ключом /A."""
    items, folders, files = parse_tree_file(tree_file_a)
    output_file = tmp_path / f"result_a_{mode}.xlsx"
    save_to_excel(items, str(output_file), mode, folders, files)
    assert output_file.exists()
    assert output_file.stat().st_size > 0


# ------------------- Test for empty file --------------------
def test_parse_tree_file_empty(tmp_path):
    empty = tmp_path / "empty.txt"
    empty.write_text("", encoding="utf-8")
    with pytest.raises(ValueError):
        parse_tree_file(str(empty))


# ---------- Parametrized test for various encodings ----------
@pytest.mark.parametrize(
    "encoding, filename",
    [
        ("utf-8", "файл.txt"),
        ("cp866", "файл.txt"),
        ("cp1251", "файл.txt"),
        ("euc-kr", "파일.txt"),
        ("gb2312", "文件.txt"),
        ("iso-8859-1", "fiché.txt"),
        ("shift_jis", "ファイル.txt"),
        ("windows-1251", "файл.txt"),
        ("windows-1252", "fiché.txt"),
    ],
)
def test_encoding_various(tmp_path, encoding, filename):
    """
    Test that parser correctly reads files in various encodings and
    preserves the actual filename (not just structure).
    """
    content = f"""Folder PATH listing for volume xx-xxx xxx
Volume serial number is xxxxxxxx xxxx:xxxx
C:\\PROJECTS
+---src
    \\---{filename}
"""
    f = tmp_path / f"tree_{encoding}.txt"
    f.write_text(content, encoding=encoding)

    # Pass encoding explicitly to bypass auto-detection in this test
    items, folders, files = parse_tree_file(str(f), encoding=encoding)

    assert folders == 1, f"Expected 1 folder, got {folders}"
    assert files == 1, f"Expected 1 file, got {files}"

    # Check that filename is preserved exactly
    file_items = [it for it in items if not it["is_dir"]]
    assert len(file_items) == 1, "Expected exactly one file item"
    assert (
        file_items[0]["name"] == filename
    ), f"Expected filename '{filename}', got '{file_items[0]['name']}'"

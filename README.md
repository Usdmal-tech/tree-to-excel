<a id="readme-top"></a>

![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![GitHub repo size](https://img.shields.io/github/repo-size/Usdmal-tech/tree-to-excel)
![GitHub last commit](https://img.shields.io/github/last-commit/Usdmal-tech/tree-to-excel)
![Version](https://img.shields.io/badge/version-1.0.3-blueviolet)
![CI](https://github.com/Usdmal-tech/tree-to-excel/actions/workflows/ci.yml/badge.svg)

# Tree to Excel Parser

**Brief description:**

Converts the text output of the `tree` command (standard or with the `/A` switch) into a structured Excel spreadsheet, automatically detecting the input file encoding and [correctly](#6-parsing-features-and-limitations) identifying folders and files even for complex cases (empty folders, files without extensions, folders with a dot in the name).

The program supports three display modes – flat (`flat`), merged (`merged`), and full merged (`merged_full`) – with color-coded levels and bold font for folders.

Useful for visualizing directory hierarchies, preserving and/or auditing file structures, documenting projects, preparing nesting reports, and for subsequent analysis in Excel without manual copying and formatting.

> **IMPORTANT:** The program was primarily written for Cyrillic scripts, so it accurately detects Cyrillic fonts in UTF-8 and CP866 encodings. For other encodings, the result is not guaranteed.

<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#1-features">Features</a>
    </li>
    <li>
      <a href="#2-installation">Installation</a>
      <ul>
        <li><a href="#requirements">Requirements</a></li>
        <li><a href="#steps">Steps</a></li>
      </ul>
    </li>
    <li>
      <a href="#3-usage">Usage</a>
      <ul>
        <li><a href="#syntax">Syntax</a></li>
        <li><a href="#options">Options</a></li>
        <li><a href="#examples">Examples</a></li>
        <li><a href="#notes">Notes</a></li>
      </ul>
    </li>
    <li>
      <a href="#4-output-modes">Output modes</a>
      <ul>
        <li><a href="#41-flat-mode">`flat` mode</a></li>
        <li><a href="#42-merged-mode">`merged` mode</a></li>
        <li><a href="#43-merged_full-mode">`merged_full` mode</a></li>
      </ul>
    </li>
    <li>
      <a href="#5-example-output">Example output</a>
      <ul>
        <li><a href="#51-input-without-a">Input (without `/A`)</a></li>
        <li><a href="#52-input-with-a">Input (with `/A`)</a></li>
      </ul>
    </li>
    <li>
      <a href="#6-parsing-features-and-limitations">Parsing features and limitations</a>
      <ul>
        <li><a href="#61-standard-tree-format-without-a">Standard `tree` format (without `/A`)</a></li>
        <li><a href="#62-tree-a-ascii-format">`tree /A` (ASCII) format</a></li>
        <li><a href="#63-workarounds">Workarounds</a></li>
        <li><a href="#64-encodings">Encodings</a></li>
        <li><a href="#65-general-note">General note</a></li>
      </ul>
    </li>
    <li>
      <a href="#7-testing">Testing</a>
      <ul>
        <li><a href="#running-tests">Running tests</a></li>
      </ul>
    </li>
    <li><a href="#8-license">License</a></li>
    <li><a href="#9-contributing">Contributing</a></li>
  </ol>
</details>

<!-- FEATURES -->
## 1. Features

- **Support for two `tree` formats**  
  Works with both the standard `tree` output (using `├───`, `└───`) and the ASCII format (`tree /A`), which uses `+---`, `\---`.

- **Three data presentation modes**
  - `flat` – folders only, each on a separate row;
  - `merged` – folders only, with merged cells for repeated ancestor names (useful for visualizing hierarchy);
  - `merged_full` – all items (folders and files), with cell merging.

Nesting levels are placed in corresponding columns. Folders are highlighted in **bold**.

- **Automatic encoding detection**  
  The input file is analysed using the `chardet` library. If the detection confidence is below 0.85, the program falls back to **UTF-8** and displays a warning, recommending that you explicitly specify the encoding via the `--encoding` parameter if needed. This ensures that even with uncertain detection, data is not lost, and the user can easily correct the situation.

- **color-coded levels**  
  Each nesting level gets its own background color in Excel. The path to the item is highlighted in yellow.

- **Statistics at the end of the table**  
  A summary row is automatically added with the total number of folders and files.

- **Flexible output**  
  If no output file is specified, a name is generated automatically based on the input file name and the selected mode.

- **Simplicity**  
  The tool is designed to be used from the command line and does not require a graphical interface.

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- INSTALLATION -->
## 2. Installation

The easiest way is to install from PyPI:

```bash
pip install tree-to-excel
```

### Requirements

- Python version **3.8** or higher.
- Installed `pip` package manager.

### Steps

1. Clone the repository or download the project files.
2. Navigate to the project root directory (where `requirements.txt` is located).
3. Install the required dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. To run tests, you will also need `pytest` – it is already included in `requirements.txt`, but if you do not plan to test, you can skip installing it (it is not used by the main script).

After the dependencies are installed, the `tree_to_excel.py` script is ready to use.

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- USAGE -->
## 3. Usage

### Syntax

```bash
python src/tree_to_excel/tree_to_excel.py -i <input_file> [-m <mode>] [-o <output_file>] [-e <encoding>]
```

> After installation, the script is available as the tree-to-excel command.

### Options

| Option           | Description                                                                                                                                             |
|------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------|
| `-i, --input`    | **Required.** Path to the text file containing the `tree` command output.                                                                               |
| `-m, --mode`     | Table generation mode. Allowed values: `flat`, `merged`, `merged_full`. Default: `flat`.                                                                |
| `-o, --output`   | Path to save the Excel file. If not specified, a name is generated automatically: `<input_filename>_<mode>.xlsx` in the same folder as the input file.  |
| `-e, --encoding` | Force the encoding of the input file (e.g., UTF-8, cp1251). Overrides automatic detection. Useful when auto-detection fails (e.g., for rare encodings). |
| `--version`      | Show the program version and exit.                                                                                                                      |

### Examples

1. **Basic run** (mode `flat`, output file generated automatically):

   ```bash
   python src/tree_to_excel/tree_to_excel.py -i tree_output.txt
   ```

2. **Specifying mode and output file**:

   ```bash
   python src/tree_to_excel/tree_to_excel.py -i tree.txt -m merged_full -o result.xlsx
   ```

3. **Specifying encoding**:

   ```bash
   python src/tree_to_excel/tree_to_excel.py -i tree.txt -m merged_full -e CP866
   ```

4. **Getting help**:

   ```bash
   python src/tree_to_excel/tree_to_excel.py -h
   ```

### Notes

- The input file must contain the output of the `tree` command in one of the supported formats (standard or with the `/A` switch). Examples of such files can be found in the `examples/` folder.
- If the input file does not contain a valid root path (in the format `X:\...`), the program will exit with an error.
- On successful execution, the path to the created Excel file is printed to the console.
- In case of errors (missing file, parsing issues), messages are printed to `stderr`.

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- OUTPUT MODES -->
## 4. Output modes

The program supports three modes that determine which items appear in the Excel sheet and how they are grouped. Your choice depends on your needs: whether you want to see only folders or all items, and whether you need visual hierarchy with merged cells.

### 4.1. `flat` mode

- **What goes into the table:** folders only (files are excluded).
- **How it looks:** each item (folder) occupies a separate row. The folder name is placed in the column corresponding to its nesting level (L1, L2, …). Thus, the level is visible by the column position.
- **Peculiarities:** cells are not merged; the structure is read across columns. This is convenient for further filtering or sorting in Excel.

**Example command:**

```bash
python src/tree_to_excel/tree_to_excel.py tree.txt -m flat
```

### 4.2. `merged` mode

- **What goes into the table:** folders only.
- **How it looks:** for each level, columns are filled with ancestor folder names. If an ancestor name repeats for several child items, the cells with that name are merged vertically. This creates a visual tree, similar to the folder view in a file explorer.
- **Peculiarities:** merging is performed only for folders; files are ignored. The result is a compact representation of the hierarchy.

**Example command:**

```bash
python src/tree_to_excel/tree_to_excel.py -i tree.txt -m merged
```

### 4.3. `merged_full` mode

- All items – both folders and files – are included in the table.
- Visually, the result is similar to the `merged` mode, but now files are also included in the rows. Folders are highlighted in **bold**, files in regular font. This makes it easy to distinguish directories from files at a glance.
- Cell merging works for all repeated ancestor names.

**Example command:**

```bash
python src/tree_to_excel/tree_to_excel.py -i tree.txt -m merged_full
```

<br>

|     Mode      |      Folders       |       Files        |    Cell merging    |
|:-------------:|:------------------:|:------------------:|:------------------:|
|    `flat`     | :white_check_mark: |        :x:         |        :x:         |
|   `merged`    | :white_check_mark: |        :x:         | :white_check_mark: |
| `merged_full` | :white_check_mark: | :white_check_mark: | :white_check_mark: |

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- EXAMPLE OUTPUT -->
## 5. Example output

### 5.1. Input (without `/A`)

Example text file `project_tree.txt` with the following content (output of the `tree` command without the `/A` switch):

```text
Folder PATH listing for volume System
Volume serial number is 00000001 1234:5678
C:\MY_PROJECT
├───src
│   ├───core
│   │   ├───utils
│   │   │   └───helpers.py
│   │   └───main.py
│   └───tests
│       └───test_core.py
├───docs
│   ├───readme.md
│   └───guide.pdf
└───scripts
    └───deploy.bat
```

After execution, the file `project_tree_[mode].xlsx` is created:

- Example output in `flat` mode:

<div align="center">

![flat mode](examples/project_tree_flat.jpg)

</div>

- Example output in `merged` mode:

<div align="center">

![merged mode](examples/project_tree_merged.jpg)

</div>

- Example output in `merged_full` mode:

<div align="center">

![merged_full mode](examples/project_tree_merged_full.jpg)

</div>

### 5.2. Input (with `/A`)

Example text file `project_tree.txt` with the following content (output of the `tree` command with the `/A` switch):

```text
Folder PATH listing for volume System
Volume serial number is 00000001 1234:5678
C:\MY_PROJECT
+---src
│   +---core
│   │   +---utils
│   │   │   \---helpers.py
│   │   \---main.py
│   \---tests
│       \---test_core.py
+---docs
│   +---readme.md
│   \---guide.pdf
\---scripts
    \---deploy.bat
```

After execution, the file `project_tree_[mode].xlsx` is created:

- Example output in `flat` mode:

<div align="center">

![flat mode](examples/project_tree_a_flat.jpg)

</div>

- Example output in `merged` mode:

<div align="center">

![merged mode](examples/project_tree_a_merged.jpg)

</div>

- Example output in `merged_full` mode:

<div align="center">

![merged_full mode](examples/project_tree_a_merged_full.jpg)

</div>

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- PARSING FEATURES AND LIMITATIONS -->
## 6. Parsing features and limitations

When converting the tree to Excel, it is important to be aware of the parser's behavior and its limitations, especially regarding how it distinguishes folders from files.

### 6.1. Standard `tree` format (without `/A`)

The parser determines the item type (folder or file) as follows:

1. **If the item has children** – it is a **folder**.
2. **If the item has no children, but has a marker** (`+---`, `\---`, `├───`, `└───`):
   - **and the name contains a dot (`.`)**: the item is treated as a **file**.
   - **and the name does not contain a dot**: the item is treated as an **empty folder**.

This logic may fail in the following cases:

- **Folders with a dot in the name**  
  For example, `project.v1`, `.config`, `my.folder`.  
  Because the name contains a dot, the parser will consider them files, even if they are empty folders (or folders with no visible children in the `tree` output).  
  *Result:* such folders will not appear in `flat` and `merged` modes (which only output folders) and will be displayed as files in `merged_full`.

- **Files without an extension**  
  For example, `README`, `Makefile`, `LICENSE`, `Dockerfile`.  
  *Result:* they will be incorrectly included in folder lists in `flat` and `merged` modes, and in `merged_full` they will be highlighted in bold as folders.

### 6.2. `tree /A` (ASCII) format

In this format, all markers use only the characters `+`, `\`, `|`, and spaces. The parser **does not use markers** to determine the item type – only items that have children are considered folders.

**Empty folders are not recognized** and are always identified as files. The `tree /A` command on Windows does not output empty folders with branch markers as clearly as the standard output, or it outputs them without continuation indicators, making it impossible for the parser to distinguish them from extension-less files.

### 6.3. Workarounds

If your structure contains **folders with dots** or **files without extensions**, and you want to avoid misclassification, use one of the following methods:

- Use the `tree /A` (ASCII) output for such structures – then folders with dots will be correctly identified (provided they have children). However, **empty folders** will then be incorrectly classified as files (and consequently will not appear in `flat` and `merged` modes).
- Keep the standard `tree` output and **manually check** the resulting Excel file, adjusting item types if necessary (this can be done in the spreadsheet if you know which items should be folders).

### 6.4. Encodings

The program automatically detects the input file encoding using the `chardet` library.  
The confidence threshold is set to **0.85** – if the detector does not reach this confidence, it falls back to `UTF-8` as a default, and a warning is printed to the console recommending that you specify the encoding manually if the names appear incorrect.

For most common encodings (UTF-8, WINDOWS-1251, ISO-8859-1), auto-detection works reliably. For less common ones (CP866, EUC-KR, GB2312, SHIFT-JIS, etc.), errors are possible.

**If you encounter garbled characters in Excel**, run the program again with the `--encoding` parameter, specifying the correct encoding.

Example:

```bash
python src/tree_to_excel/tree_to_excel.py -i my_tree.txt --encoding CP866
```

### 6.5. General note

The tool is designed for **quick analysis and visualization** of folder structures, not for absolutely accurate reconstruction of the file system.

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- TESTING -->
## 7. Testing

The project includes a test suite to verify the correctness of parsing and the behavior of various modes. The tests cover:

- standard `tree` format and `/A` format;
- handling of files encoded in UTF-8, CP866, CP1251, EUC-KR, GB2312, ISO-8859-1, SHIFT_JIS, WINDOWS-1251, WINDOWS-1252 (you can also test with any other encoding yourself);
- edge cases (folders with dots, files without extensions, empty folders);
- all saving modes (`flat`, `merged`, `merged_full`);
- handling of erroneous input (missing root path, empty file).

### Running tests

1. Install development dependencies (if you haven't already):

   ```bash
   pip install -r requirements.txt
   ```

2. Run the following command from the project root:

   ```bash
   pytest -v
   ```

3. For a more detailed report, use:

   ```bash
   pytest -v --tb=short
   ```

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- LICENSE -->
## 8. License

This project is distributed under the **MIT** license. The full license text is available in the [`LICENSE`](LICENSE) file in the root of the repository.

<p align="right"><a href="#readme-top">back to top</a></p>

---

<!-- CONTRIBUTING -->
## 9. Contributing

If you find a bug or would like to suggest an improvement, please contact me at [usdmal@rambler.ru](mailto:usdmal@rambler.ru) or create an [Issue](https://github.com/Usdmal-tech/tree_to-excel/issues) with a detailed description.

<p align="right"><a href="#readme-top">back to top</a></p>

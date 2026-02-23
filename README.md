# VSME Template Migration Tool

The VSME Template Migration Tool automates the migration of data filled in older versions of the VSME Digital Template (Excel-based) into the latest template version.
It ensures that all data-filled Named Ranges (NRs) and data cells with missing NRs are transferred correctly and consistently between template versions — minimizing manual effort and the risk of errors during updates.

## Working overview

Everything starts from an old version filled with data and a new empty one, both imported as openpyxl objects. Thanks to data being organised with NRs in the VSME template, name ranges for both files are accessed through the `.definedNames` method. With NR attributes, sheet and range information for each NR are extracted and ranges are then transformed into shapes using the `.range_boundaries` method and stored into a specific Python class (see "classes.py"). With these info, pandas DataFrames are created, with cols of: "name_ranges", "sheets", "cell_ranges" and "cell_shapes". Values are subsequently copied from the old workbook and attached to the first DataFrame in the col "cell_values". With "name_ranges" col as key, values are joined from the first DataFrame (old wb) to the second (new wb). Lastly, the old values are pasted in the new empty workbook (openpyxl object) based on their new references, found in each row of the DataFrame.
These mechanics are similarly repeated for additional DataFrames of the missing NRs (with, in that case, no name ranges references and joining performed as concatenation based on the hard-coded cell references).

## Checklist before new release

- Update of the hard-coded tables (missingNR_df, dict_of_changes_NR), checking if cell references and NR denominations have changed;
- Check if reference of the Table of Contents have changed (see funct: create_table_of_contents);
- Check if new languages have been added (if so, add "INCOMPLETE" label in funct: check_status_incomplete);
- Think about which new issues after a successful migration can arise with a newer version (ex change in the list of NACE codes with no matches for certain old categories), and, if so, add an error string manually to list_migrationissues (`list[str]`)\*\*\*
- For every whole new functionality, check whether new data is required in latest version and, if so, create new mechanisms to handle how that new data will be displayed after migration;
- Be careful about pasting old data into cells in which validation or rules have been added (ex checkboxes or fixed enumeration lists)

## Requirements

1. Python environment (.venv, see Converter for creating and activating it) and dependencies:
   - openpyxl: Read/write Excel 2010 xlsx/xlsm/xltx/xltm files
   - <https://pandas.pydata.org/>: open source data analysis and manipulation tool
2. Filled-out Excel template with unmodified name ranges (see [digital templates](https://github.com/EFRAG-EU/Digital-Template-to-XBRL-Converter/tree/main/digital-templates))

## Notes & Disclaimers

The tool does not alter the old VSME version inputted — it creates a new copy (of the newest VSME version).
Some formatting (like checkboxes and arrows to navigate tabs) are not transferred because openpyxl cannot process it.

_Disclaimer_: the authors are sorry if the code is sometimes not written in the most Pythonic way (we did not know which fishes to fry at some points in the process, and prioritised results).

\*\*\* As openpyxl objects cannot be dynamically recomputed (you have to open the xlsx file itself for recomputation of formula cells), we could not infer the list of issues dyamically from the Table of Validation. That is why potential migration issues have to be discovered and registered in list_migrationissues every new release.

## Installation

### For Users

```bash
pip install git+https://github.com/EFRAG-EU/Digital-Template-Data-Migration-Tool.git
```

### For Development

1. Clone the repository:

   ```bash
   git clone https://github.com/EFRAG-EU/Digital-Template-Data-Migration-Tool.git
   cd Digital-Template-Data-Migration-Tool
   ```

2. Create and activate a virtual environment (recommended):

   ```bash
   python -m venv .venv
   # On Windows:
   .venv\Scripts\activate
   # On Linux/Mac:
   source .venv/bin/activate
   ```

3. Install the package in editable mode with development dependencies:

   ```bash
   pip install -e .[dev]
   ```

   **Note:** The `[dev]` extra installs optional development dependencies (pytest and testing tools) defined in `pyproject.toml`. For production use, you can omit `[dev]` and just use `pip install -e .`

## Usage

### Running the Migration Tool

The tool can be run as a Python module:

```bash
# Interactive mode (prompts for file selection)
python -m migration_tool

# With a specific file path
python -m migration_tool path/to/your/template.xlsx
```

**Note:** The `-m` flag is required for the module to work correctly with relative imports.

### Running Tests

Run all tests with pytest:

```bash
pytest tests/ -v
```

Or run tests with more detailed output:

```bash
pytest tests/ -v --tb=short
```

### Package Management

View installed package information:

```bash
pip show migration_tool
```

Uninstall the package:

```bash
pip uninstall migration_tool
```

Reinstall after making changes (when in editable mode, changes are automatically reflected):

```bash
pip install -e .[dev]
```

## Authors

Developed by EFRAG's Digital Team.
Maintained by EFRAG's Digital Team.
For inquiries or support, contact: <andrea.baschiera@efrag.org> / <thibault.magro@efrag.org>

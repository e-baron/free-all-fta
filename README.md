# free-all-fta

## Aim 

Python tool to create allocative FTA calculations in an Excel workbook and optionally generate SVG/HTML fault-tree diagrams.

The repository includes a ready-to-use sample workbook: [free-all-fta-example.xlsx](free-all-fta-example.xlsx).

This FTA tool provides an iterative approach to identify the failure rates of equipment and check whether the required safety target is met. It also allows to check this based on mean time to repair (or safe down time).

## What This Tool Does

The script reads an Excel workbook with three sheets:

- `events`
- `gates`
- `tree`

It computes gate-level values from the event/gate hierarchy, writes formulas and result columns back into the workbook, and can generate diagram files.

Two calculation modes are supported:

- **Markov allocative mode** (default)
- **Pure-probability allocative mode** (`--pure`)

## Requirements

- Python 3.9+
- `openpyxl`

Install dependency:

### On Windows

#### Python installation

- Open `PowerShell`
- Execute: 
```bash
winget install Python.Python.3.14
```
- Close `PowerShell`
- Reopen `PowerShell`


#### Library installation

```bash
pip install openpyxl
```

### On MacOS

#### Python installation

- Open a Terminal
- Execute:
```bash
brew update
brew install python
```


#### Library installation

```bash
python3 -m pip install --user --break-system-packages openpyxl
```

## Quick Start With The Provided Excel File

### 1. Run default (Markov) mode

```bash
python3 free-all-fta.py free-all-fta-example.xlsx
```

This creates:

- `free-all-fta-example/free-all-fta-example_alloc.xlsx`

### 2. Run pure-probability mode

```bash
python3 free-all-fta.py free-all-fta-example.xlsx --pure
```

This creates:

- `free-all-fta-example/free-all-fta-example_pure.xlsx`

### 3. Generate diagrams too

```bash
python3 free-all-fta.py free-all-fta-example.xlsx --diag
```

This creates SVG pages and an HTML index in the output directory.

## Command Syntax

```bash
python3 free-all-fta.py <excel-file> [options]
```

### Positional argument

- `<excel-file>`: path to the source workbook (for example `free-all-fta-example.xlsx`).

### Options

- `--alloc`
	- Backward-compatible no-op.
	- Allocative mode is already the default.

- `--markov`
	- Explicitly selects Markov allocative formulas.
	- This is also the default if `--pure` is not provided.

- `--pure`
	- Selects pure-probability allocative formulas.

- `--diag`
	- Generates FTA diagrams (SVG + HTML index) in addition to the output workbook.

- `--alloc-out <path>`
	- Output workbook path for Markov mode.
	- Kept for backward compatibility.

- `--markov-out <path>`
	- Alias for `--alloc-out`.

- `--pure-out <path>`
	- Output workbook path for pure mode.

- `--diag-prefix <text>`
	- Prefix used for generated diagram filenames.

- `--diag-dir <path>`
	- Output directory for generated workbook and diagrams.
	- If omitted, defaults to a folder based on input file name.

### Option compatibility

- `--pure` and `--markov` are mutually exclusive.

## Practical Examples

### Markov + diagrams with default naming

```bash
python free-all-fta.py free-all-fta-example.xlsx --diag
```

### Pure probability output + diagrams with default naming

```bash
python free-all-fta.py free-all-fta-example.xlsx --pure --diag
```

### Markov output to a custom file

```bash
python free-all-fta.py free-all-fta-example.xlsx --markov-out ./out/markov.xlsx
```

### Pure probability output to a custom file + diagrams in custom folder

```bash
python free-all-fta.py free-all-fta-example.xlsx --pure --pure-out ./out/pure.xlsx --diag --diag-dir ./out/diagrams --diag-prefix projectA
```

## Expected Workbook Structure

The tool expects these sheets to exist:

- `events`
- `gates`
- `tree`

### `events` required columns

- `id`
- `model_type`
- `detection_time` for `ConstantRate` or `Fixed` model_type
- `negation_time` for `ConstantRate` or `Fixed` model_type
- `Allocated value` for `ConstantRate` or `Fixed` model_type
- `mean_repair_time` for `ConstantRate` or `Fixed` model_type, it is automatically calculated based on `detection_time` and `negation_time`

Optional but used when present:

- `label` (used by diagram rendering)

### `gates` required columns

- `id`
- `safety target` (used for differences and target propagation): provide a safety target for the top gate (the one gate at the top of your tree)

The script creates/updates these columns as needed:

- `Formula`
- `calculated frequency`
- `calculated mean repair time`
- `differences`: this is the difference between the `safety target` and the `calculated frequency`. If the `safety target` is met by the calculations based on the provided values in the `events` worksheet, this cell will display a green background.
- `calculated probability` (pure mode)

Optional but used when present:

- `label` (diagram text)
- `is_paged` (diagram paging): when `True` is selected, a SVG is provided for this gate. The FTA is cut at this gate: the gate is represented as a triangle in the parent diagram, and a new diagram is provided for the Gate and its children.

### `tree` expected format

This is where you are constructing the hierarchy of gates down to events (the leafs of the tree).

Hierarchy is read from level columns named like:

- `L1_ID`, `L2_ID`, ...
- optional matching type columns: `L1_TYPE`, `L2_TYPE`, ...

Supported gate types are:

- `AND`
- `OR`

## Supported Event Model Types

Recognized `model_type` values include:

- `ConstantRate` (or `constant_rate`): to be used when we now that the system is repairable
- `Fixed`: to be used when we are not sure whether the system is repairable or we want to ignore the time to repair
- `Probability`: to be used when an event is not directly linked to time (e.g. for human errors...)
- `Multiplicity`: to be used when we want to represent a number of pieces of equipment
- `True`: this is equivalent to a probability = 1
- `False`: this is equivalent to a probability = 0

## Output Files

Depending on options, outputs include:

- Updated workbook (`*_alloc.xlsx` or `*_pure.xlsx`, unless custom path is provided)
- Diagram files (`*_fta_*.svg`)
- Diagram index page (`*_fta.html`)

## Troubleshooting

- If you get missing-column errors, verify sheet names and column headers exactly.
- If formulas look stale in Excel, open the output file and force recalculation (the script already requests full calc on load when possible).
- If diagrams are not needed, omit `--diag` for faster execution.

## License

See [LICENSE](LICENSE).

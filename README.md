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

equivalent of:

```bash
python free-all-fta.py free-all-fta-example.xlsx --diag --markov
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

## Columns Used Per Calculation Mode

Both modes use the same workbook sheets (`events`, `gates`, `tree`).

### Markov mode (`--markov`, default)

Used columns in `events`:

- `id`
- `model_type`
- `allocated value`
- `mean_repair_time`

Used columns in `gates`:

- `id`
- `formula`
- `safety target` (used for top-down target propagation and `differences`)
- `calculated frequency`
- `calculated mean repair time`
- `differences`

Used columns in `tree`:

- `L1_ID`, `L2_ID`, ... (hierarchy)
- Optional: `L1_TYPE`, `L2_TYPE`, ... with gate types `AND`/`OR`

### Pure mode (`--pure`)

Used columns in `events`:

- `id`
- `model_type`
- `allocated value`
- `mean_repair_time`
- Optional for `Probability` model: `probability` (if missing, `allocated value` is used)

Used columns in `gates`:

- `id`
- `formula`
- `calculated frequency`
- `calculated mean repair time`
- `calculated probability`
- Optional: `safety target` and `differences` (if present, `differences` is computed)

Used columns in `tree`:

- `L1_ID`, `L2_ID`, ... (hierarchy)
- Optional: `L1_TYPE`, `L2_TYPE`, ... with gate types `AND`/`OR`

## How Formulas Are Built Per Mode

This section explains, for each mode, which inputs are used to build formulas and how gate outputs are calculated.

### Markov mode (`--markov`, default)

Event-level interpretation:

- `ConstantRate` / `Fixed`:
	- `allocated value` is interpreted as rate `lambda`.
	- `mean_repair_time` is interpreted as `T`.
- `Probability` / `Multiplicity`:
	- `allocated value` is used as a multiplicative factor.
- `True` / `False`:
	- factor `1` / `0`.

Gate formulas:

- `AND` gate:
	- With multiple rate children:
		- `calculated frequency = product(lambda_i * T_i) * sum(1 / T_i) * product(factors)`
		- `calculated mean repair time = 1 / sum(1 / T_i)`
		The hazardous state exists only when all child states are simultaneously present. The overlap duration is governed by the shortest restoration process. The harmonic combination: `TEQ = 1 / Σ(1/Ti)` preserves the expected overlap duration.
	- With one rate child:
		- `calculated frequency = lambda_1 * product(factors)`
		- `calculated mean repair time = T_1`
	- With only factor children:
		- `calculated frequency = product(factors)`
		- `calculated mean repair time` is empty.

- `OR` gate:
	- `calculated frequency = sum(rate children lambdas) + sum(factor children)`
	- `calculated mean repair time = sum(lambda_i * T_i) / sum(lambda_i)`
		- Computed from rate children only. Any child failure can independently create the hazardous state. The gate occurrence rate is therefore the sum of child occurrence rates. The equivalent duration must preserve: `QEQ = λEQ · TEQ`. Therefore the duration is the rate-weighted average persistence time.

Safety-target propagation and difference:

- Along `AND` chains where a parent has exactly one child gate:
	- `child safety target = parent safety target / product(event child allocated value)`
- `differences = safety target - calculated frequency`

Which inputs each calculated gate column depends on:

- `calculated frequency`:
	- `events.model_type`
	- `events.allocated value`
	- `events.mean_repair_time` (for rate-based children)
	- `tree.Lx_ID`, `tree.Lx_TYPE`
- `calculated mean repair time`:
	- `events.mean_repair_time` (for rate-based children)
	- `events.allocated value` (for OR weighted average)
	- `tree.Lx_ID`, `tree.Lx_TYPE`
- `differences`:
	- `gates.safety target`
	- `gates.calculated frequency`

### Pure mode (`--pure`)

Event-level interpretation:

- `ConstantRate` / `Fixed`:
	- `lambda = allocated value`
	- `T = mean_repair_time`
	- event probability contribution `q_i = 1 - EXP(-(lambda * T))`
	- event frequency contribution `w_i = lambda`
- `Probability`:
	- `q_i = probability` when the `probability` column exists, otherwise `q_i = allocated value`.
- `Multiplicity`:
	- `q_i = allocated value` as a factor in gate probability expressions.
- `True` / `False`:
	- `q_i = 1` / `0`.

Gate formulas:

- `AND` gate:
	- `calculated probability = MIN(1, product(q_i))`
	- `calculated mean repair time = MAX(T_i)`.
	The probability of the hazardous condition is directly computed from the child probabilities. There is no physical rate-weighting mechanism as in Markov theory. A weighted average duration has no probabilistic meaning. The longest child duration is the most representative persistence horizon. Using: `TEQ = MAX(Ti)` is conservative and avoids artificially increasing: `wEQ`
	- `calculated frequency = raw_product(q_i) / MAX(T_i)`
		- Uses raw (uncapped) probability expression for frequency.

- `OR` gate:
	- `calculated probability = MIN(1, sum(q_i))`
	- `calculated mean repair time = MAX(T_i)`.
	The conjunction probability is obtained directly from the child probabilities.No Markov overlap-duration computation is involved. There is no unique mathematically derived equivalent duration.Using: `MAX(Ti)` provides a conservative exposure duration and keeps the interpretation identical to OR gates. It also avoids mixing probabilistic calculations with Markov-derived duration formulas.
	- `calculated frequency = sum(w_i)` for rate/gate branches
		- This avoids deriving frequency from a capped probability.

Difference in pure mode:

- If `safety target` and `differences` columns are present:
	- `differences = safety target - calculated frequency`

Which inputs each calculated gate column depends on:

- `calculated probability`:
	- `events.model_type`
	- `events.allocated value`
	- `events.probability` (optional, only for `Probability` model)
	- `events.mean_repair_time` (for rate-to-probability conversion)
	- `tree.Lx_ID`, `tree.Lx_TYPE`
- `calculated mean repair time`:
	- `events.mean_repair_time` (for rate/gate branches)
	- `tree.Lx_ID`, `tree.Lx_TYPE`
- `calculated frequency`:
	- `events.allocated value`
	- `events.mean_repair_time`
	- child gate `calculated frequency` and `calculated mean repair time`
	- `tree.Lx_ID`, `tree.Lx_TYPE`

### `events` required columns

- `id`
- `model_type`
- `allocated value`
- `mean_repair_time`

Additional notes:

- In pure mode, if `model_type` is `Probability`, the `probability` column is used when present (otherwise `allocated value` is used).
- `detection_time` / `negation_time` are not directly required by the calculation engine; they can still be part of your workbook if you use them to derive `mean_repair_time`.

Optional but used when present:

- `label` (used by diagram rendering)

### `gates` required columns

- `id`
- `formula`
- `calculated frequency`
- `calculated mean repair time`

Mode-specific notes:

- Markov mode also uses `safety target` and `differences`.
- Pure mode also uses `calculated probability`; `safety target`/`differences` are optional but used when present.

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

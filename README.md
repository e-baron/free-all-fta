# free-all-fta

## Aim

Python tool to create Markov allocative fault-tree calculations in an Excel workbook and optionally generate SVG/HTML fault-tree diagrams.

This tool provides an iterative approach to identify equipment failure rates and check whether the required safety target is met. It also supports calculations based on mean time to repair (or safe down time).

## What the tool does

The script reads an Excel workbook containing these worksheets:

- `events`
- `gates`
- `tree` (or another worksheet selected with `--tree-name`)

It calculates gate-level values from the event/gate hierarchy, writes formulas and result columns into the workbook, and can generate diagrams. Markov allocative calculation is the only supported calculation mode.

## Requirements

- Python 3.9+
- `openpyxl`

Install `openpyxl`:

### Windows

Install Python from PowerShell:

```bash
winget install Python.Python.3.14
```

Then reopen PowerShell and run:

```bash
pip install openpyxl
```

### macOS

Install Python with Homebrew:

```bash
brew update
brew install python
```

Then install the library:

```bash
python3 -m pip install --user --break-system-packages openpyxl
```

## Quick start

Run Markov calculations:

```bash
python3 free-all-fta.py free-all-fta-example.xlsx
```

This creates `free-all-fta-example/free-all-fta-example_alloc.xlsx`.

Generate Markov calculations and diagrams:

```bash
python3 free-all-fta.py free-all-fta-example.xlsx --diag
```

This creates the output workbook, SVG diagram pages, and an HTML index in the output directory.

## Command syntax

```bash
python3 free-all-fta.py <excel-file> [options]
```

### Options

- `--alloc`: Backward-compatible no-op; allocative Markov mode is the default.
- `--markov`: Explicitly selects the only supported calculation mode.
- `--diag`: Generates Markov FTA diagrams (SVG pages and an HTML index) alongside the workbook.
- `--tree-name <worksheet>`: Uses the named worksheet as the tree instead of `tree`. It must contain columns such as `L1_ID`. The name is included in automatically generated output filenames.
- `--alloc-out <path>`: Output workbook path; retained for backward compatibility.
- `--markov-out <path>`: Alias for `--alloc-out`.
- `--diag-prefix <text>`: Prefix for generated diagram filenames.
- `--diag-dir <path>`: Directory for the generated workbook and diagrams. Defaults to a directory based on the input filename.

Examples:

```bash
python3 free-all-fta.py free-all-fta-example.xlsx --markov-out ./out/markov.xlsx
python3 free-all-fta.py free-all-fta-example.xlsx --diag --diag-dir ./out/diagrams --diag-prefix projectA
```

## Workbook structure

### `events` columns

Required:

- `id`
- `model_type`
- `allocated value`
- `mean_repair_time`

Optional:

- `label` (used in diagrams)
- `probability` is not required by Markov calculations; Probability events use `allocated value` as their factor.
- `detection_time` and `negation_time` may be present if used by workbook formulas to derive `mean_repair_time`.

### `gates` columns

Required:

- `id`
- `formula`
- `safety target`
- `calculated frequency`
- `calculated mean repair time`
- `differences`

The script creates or updates the `Formula`, `calculated frequency`, `calculated mean repair time`, `differences`, and `calculated probability` columns as needed. `calculated probability` is used for Markov probability-only gates; it does not indicate a separate calculation mode.

Optional:

- `label` (used in diagrams)
- `is_paged` (or `is_page`): when set to `True`, the gate is shown as a transfer symbol in its parent diagram and receives a separate SVG page.

### `tree` worksheet

The tree describes the gate/event hierarchy, with one node per row. Hierarchy columns are named `L1_ID`, `L2_ID`, and so on. Optional matching columns `L1_TYPE`, `L2_TYPE`, and so on specify gate types `AND` or `OR`.

## Markov calculations

### Event interpretation

- `ConstantRate`: `allocated value` is the rate $\\lambda$; `mean_repair_time` is the repair time $T$.
- `Fixed`: `allocated value` is a fixed frequency $\\lambda$; `mean_repair_time` is ignored and no repair time is produced.
- `Probability` and `Multiplicity`: `allocated value` is a multiplicative factor in gate calculations.
- `True` and `False`: factors of $1$ and $0$, respectively.

### AND gates

Rate children with repair times are represented by $(\\lambda_i, T_i)$. Rate children without repair times contribute fixed frequencies $\\lambda_F$; factors multiply the result.

With multiple repairable rate children:

$$
\\lambda_{EQ} = \\left(\\prod_i \\lambda_i T_i\\right)\\left(\\sum_i \\frac{1}{T_i}\\right)\\prod_F \\lambda_F \\prod_j f_j,
\\qquad
T_{EQ} = \\frac{1}{\\sum_i 1/T_i}.
$$

With one repairable rate child, the frequency is its rate multiplied by fixed-frequency children and factors, and the gate repair time is that child's $T$. If there are no repair-time children, the frequency is the product of fixed frequencies and factors; the calculated repair time is left empty.

### OR gates

The calculated frequency is the sum of rate-child frequencies and factor children. The calculated repair time is computed from rate children with a repair time only:

$$
T_{EQ} = \\frac{\\sum_i \\lambda_i T_i}{\\sum_i \\lambda_i}.
$$

It is left empty when no child has a repair time.

### Safety targets

Along an AND chain where a parent has exactly one child gate, the child target is propagated as:

$$
\\text{child safety target} = \\frac{\\text{parent safety target}}{\\prod \\text{event-child allocated values}}.
$$

The `differences` value is `safety target - calculated frequency`. Nonnegative differences are highlighted green; negative differences are highlighted red.

## Probability-only gates

A gate whose descendants are only `Probability`, `True`, or `False` events (or other probability-only gates) has no frequency or repair time. An AND gate may also include `Multiplicity` events; an OR gate containing a `Multiplicity` event is not considered probability-only.

For a probability-only gate, `calculated probability` is the child-probability product for AND or sum for OR, capped at 1. The `differences` value compares `safety target` with this probability. Diagrams show the probability value without a frequency or repair time.

An OR gate must not directly mix probability children (probability-only gates and `Probability`/`True`/`False` events) with frequency children. The script reports the offending gate and children when dimensions are mixed.

## Cut sets worksheet

The generated `_alloc.xlsx` workbook includes a `cut-sets` worksheet formatted as an Excel table named `mcs_table`. It lists the minimal cut sets of the tree's top gate(s), sorted by frequency from highest to lowest; cut sets without a computable frequency appear last. `True` events are ignored and `False` events remove the cut sets they belong to. Values are Excel formulas based on the `events` worksheet.

Columns include:

- `order`: number of events in the cut set.
- `MCS`: cut-set event IDs and labels, one per line.
- `w (frequency /h)`: cut-set frequency calculated with the Markov AND formulas.
- `w (%)`: cut-set frequency divided by the overall frequency of the top gate(s); blank when unavailable or zero.

## Supported event models

- `ConstantRate` (or `constant_rate`): repairable system represented by a rate and repair time.
- `Fixed`: fixed frequency with no repair time.
- `Probability`: event not directly linked to time, such as a human error; used as a factor.
- `Multiplicity`: number of equipment items, used as a factor.
- `True`: equivalent to a factor of 1.
- `False`: equivalent to a factor of 0.

## Output files

Depending on the options, output includes:

- Updated workbook (`*_alloc.xlsx`, unless a custom output path is supplied), including the `cut-sets` worksheet.
- Diagram files (`*_fta_*.svg`) and an HTML index (`*_fta.html`) when `--diag` is specified.

## Troubleshooting

- For missing-column errors, verify worksheet names and column headers exactly.
- If formulas look stale in Excel, open the output workbook and force recalculation; the script requests full recalculation on load when possible.
- Omit `--diag` when diagrams are not needed.

## License

See [LICENSE](LICENSE).

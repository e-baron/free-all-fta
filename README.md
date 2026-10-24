# free-all-fta

## Calculation model and scope

The application implements one restricted, state-based Markov fault-tree calculation model only. It is not a general-purpose dynamic fault-tree solver.

Every `ConstantRate` and `Fixed` event used by the selected tree must have a frequency $w$ in $1/h$ and a strictly positive `mean_repair_time` (displayed as MTTR) $T$ in hours. The two event types use the same Markov formulas today, but remain distinct: `ConstantRate` asserts a constant-rate stochastic failure process; `Fixed` identifies an externally supplied or externally calculated frequency. Retaining the distinction allows future models to treat them differently.

In this application, MTTR is the mean duration for which the event condition remains active after occurrence. For a technical failure it normally represents detection plus repair, restoration, or negation time. For a human or operational `Fixed` event, it represents the mean duration of the resulting erroneous or unsafe condition—not repair of the person. The software does not infer or assign MTTR, a global SDT, mission time, or a default duration.

Frequency operands can be combined under OR and AND when every frequency operand has valid frequency and MTTR values. Frequency and Probability cannot be mixed under OR. Probability may scale frequency under AND. Multiplicity is a non-negative, dimensionless AND-only scaling factor. `True` and `False` follow Boolean identities and cannot be OR-combined with a frequency.

The calculations assume statistical independence unless dependence is explicitly represented, constant frequency/rate values, constant effective restoration/state-ending rates, exponentially represented state durations, and applicability of the implemented rare-event Markov approximation. Validation rejects unsupported combinations before output generation, but passing validation does not prove correct fault-tree logic, independence, valid data or MTTR assumptions, completeness, absence of common causes, suitability for a safety case, or compliance with a railway standard.

> **Important:** The Markov model treats every frequency-producing event as a state with a frequency and MTTR. For human and operational Fixed events, analysts must justify the duration of the resulting condition. The software does not assign or infer this duration automatically.

## Aim and usage

Python tool to create Markov allocative fault-tree calculations in an Excel workbook and optionally generate SVG/HTML fault-tree diagrams.

### Requirements

- Python 3.9 or newer
- `openpyxl`

### Install on Windows

Open PowerShell and install Python if it is not already available:

```powershell
winget install Python.Python.3.14
```

After installation, reopen PowerShell, go to the project folder, and install `openpyxl`:

```powershell
python --version
python -m pip install openpyxl
```

If `python` is not recognized, reopen PowerShell or modify/reinstall Python with the option to add Python to `PATH` enabled.

### Quick start on Windows

Run these commands from the project folder:

```powershell
python free-all-fta.py free-all-fta-example.xlsx
python free-all-fta.py free-all-fta-example.xlsx --diag
python free-all-fta.py free-all-fta-example.xlsx --diag --diag-dir .\out --diag-prefix projectA
```

The first command creates `free-all-fta-example\free-all-fta-example_alloc.xlsx`. The diagram command also creates SVG pages and an HTML index in the selected output directory.

### macOS / Linux

Install Python 3.9+ and `openpyxl` using your platform's package manager, then run the same script with `python3` if that is the Python 3 command on your system:

```bash
python3 -m pip install openpyxl
python3 free-all-fta.py free-all-fta-example.xlsx --diag
```

### Command syntax

```powershell
python free-all-fta.py <excel-file> [options]
```

Options:

- `--alloc`: backward-compatible no-op; allocative Markov mode is the default.
- `--markov`: explicitly selects the only supported calculation model.
- `--diag`: generates diagrams in addition to the workbook.
- `--tree-name <worksheet>`: selects a tree worksheet instead of `tree`.
- `--alloc-out <path>`: output workbook path.
- `--markov-out <path>`: alias for `--alloc-out`.
- `--diag-prefix <text>`: prefix for diagram files.
- `--diag-dir <path>`: output directory; defaults to a directory based on the input filename.

The workbook contains `events`, `gates`, and the selected tree worksheet. Events require `id`, `model_type`, `allocated value`, and `mean_repair_time`. Gates require `id`, `formula`, and `safety target`; result columns are created/updated as needed. Tree hierarchy headers use `L1_ID`, `L2_ID`, etc., with optional matching `L1_TYPE`, `L2_TYPE` values `AND` or `OR`.

## Markov calculations

### Event model semantics

- `ConstantRate`: `allocated value` is constant rate $w$ [1/h]; `mean_repair_time` is mandatory positive MTTR $T$ [h]. Typical for technical repairable/restorable failure.
- `Fixed`: `allocated value` is externally supplied/calculated frequency $w$ [1/h]; `mean_repair_time` is also mandatory positive MTTR $T$ [h]. It uses the same current calculations as `ConstantRate`, while retaining its separate input meaning.
- `Probability`: `allocated value` is a dimensionless factor $q$ with $0\le q\le1$.
- `Multiplicity`: `allocated value` is a non-negative dimensionless allocation/scaling factor $m$; it is valid under AND only.
- `True` and `False`: dimensionless Boolean constants 1 and 0.
- Unknown, blank, misspelled, or unsupported event models are rejected.

### AND gates

Frequency events are represented as $(w_i,T_i)$. A single frequency may be multiplied by Probability and Multiplicity factors. Multiple frequency operands are allowed when all have valid positive MTTRs:

$$
w_{AND}=\left(\prod_i w_iT_i\right)\left(\sum_i\frac{1}{T_i}\right)\prod_jq_j\prod_km_k,
\qquad T_{AND}=\frac{1}{\sum_i1/T_i}.
$$

For one frequency operand:

$$w_{result}=w_{frequency}\prod_jq_j\prod_km_k,\qquad T_{result}=T_{frequency}.$$

For dimensionless-only AND gates, factors multiply and Boolean identities apply: $A\land True=A$ and $A\land False=False$.

### OR gates

Frequency-only OR gates sum frequencies and calculate the frequency-weighted MTTR:

$$w_{OR}=\sum_iw_i,\qquad T_{OR}=\frac{\sum_iw_iT_i}{\sum_iw_i}.$$

Every frequency branch must have a valid MTTR. If the total frequency is zero, the frequency remains zero and the guarded MTTR formula returns blank to avoid division by zero; that result has no available equivalent MTTR for further multi-frequency calculations until the total becomes positive.

Probability/Boolean-only OR gates use the existing bounded-sum approximation $q_{OR}=\min(1,\sum_iq_i)$. Boolean identities are applied consistently: $A\lor True=True$ and $A\lor False=A$. Frequency OR Probability/Boolean and Multiplicity under OR are rejected.

### Minimal cut sets and outputs

The output workbook includes a `cut-sets` worksheet formatted as an Excel table named `mcs_table`. Minimal cut sets are Boolean-expanded in memory and validated before any output is generated; each cut set uses the same AND semantics as gates. Output can include the updated workbook and, with `--diag`, SVG pages and an HTML index.

### Conservative treatment of repeated events

The `calculated frequency` reported for a gate is calculated recursively from the displayed fault-tree structure. At each gate, the application combines the equivalent frequency and MTTR of its direct children using the supported Markov AND or OR equations. It does not globally simplify repeated basic events across the complete tree before calculating gate frequencies.

Consequently, if the same basic event contributes through several branches, overlapping contributions may be counted more than once. The recursive gate frequency may therefore overestimate the frequency that an exact Boolean- or dependency-aware quantification would produce. This behavior is intentionally retained as a conservative (pessimistic) approximation. In particular, the implemented treatment tends to be conservative for repeated-event contributions combined through OR paths because overlapping contributions are not explicitly subtracted.

> **Important qualification:** “Conservative” is not a mathematical guarantee that the recursive result is an upper bound for every possible fault-tree topology or dependency structure.

For example, consider `TOP = A OR (A AND B)`. Boolean simplification gives `TOP = A`, since $A\lor(A\land B)=A$. The recursive gate calculation instead evaluates the two displayed branches independently and may approximate $w_{TOP}=w_A+w_{(A\land B)}$. The `A AND B` contribution is retained even though it is Boolean-absorbed by `A`; when that contribution is positive, this produces a conservative result for this example. The displayed tree need not be incorrect: repeated events may intentionally occur on different functional or causal paths.

Minimal-cut-set analysis is separate. Minimal cut sets are generated independently from the recursive gate calculation. During their generation, duplicate cut sets and non-minimal supersets are removed using Boolean logic. For example, the sets `{A}` and `{A, B}` reduce to `{A}` because occurrence of `A` alone is sufficient. Thus the two calculations have complementary purposes:

- **Recursive gate calculation:** preserves the displayed fault-tree structure and applies Markov formulas gate by gate.
- **Minimal-cut-set calculation:** performs Boolean minimization and identifies minimal combinations of basic events sufficient to produce the Top Event.

Minimal cut sets may themselves overlap. Therefore, $\sum_i w_{MCS,i}$ is a rare-event approximation of the Top Event frequency when cut sets are not mutually exclusive; it is not asserted to be the exact Top Event frequency. A useful diagnostic is the MCS coverage:

$$\mathrm{Coverage}_{MCS}=\frac{\sum_i w_{MCS,i}}{w_{TOP,recursive}}.$$

Coverage close to 100% indicates numerical consistency between recursive gate quantification and minimal-cut-set quantification for the analyzed tree and assumptions. Deviations may result from repeated basic events, Boolean absorption, overlapping minimal cut sets, rare-event approximations, or different aggregation paths. Do not force or normalize this indicator to 100%: retaining comparison against the recursively calculated Top Event frequency makes it diagnostically useful. Nor should MCS contributions be redefined as $w_{MCS,i}/\sum_jw_{MCS,j}$ solely to force their sum to 100%.

## Appendix: Markov model compatibility and limitations

### Assumptions and unsupported behavior

This restricted approximation assumes independence, constant frequency/rate values, constant effective restoration/state-ending rates, exponential Markov behavior, meaningful positive MTTR for every frequency-producing event, and rare-event applicability. The application does not support frequency operands without MTTR, frequency OR probability/Boolean, Multiplicity under OR, automatic probability/frequency conversion, automatic MTTR assignment, global MTTR/SDT, arbitrary dynamic or sequence-dependent gates, priority AND, functional dependency gates, automatic standby behavior, non-exponential duration behavior, implicit dependency or common-cause modeling, or automatic initiator/enabler interpretation.

**Repeated basic events and Boolean dependencies:** The recursive gate calculation does not perform exact global Boolean quantification or dependency-aware reduction of repeated basic events. Repeated-event contributions appearing through multiple branches may therefore be counted more than once. The result is intentionally treated as a conservative approximation; minimal-cut-set analysis provides a complementary Boolean-reduced analysis.

### Recursive quantification versus Boolean reduction

The recursive gate-frequency calculation preserves the displayed fault-tree structure. It does not perform exact global Boolean quantification or dependency-aware reduction of repeated basic events. Repeated-event contributions appearing through multiple branches may therefore be counted more than once. This is intentionally treated as a conservative approximation, not as a formal upper-bound guarantee. Minimal-cut-set analysis is provided as a complementary Boolean-reduced analysis.

The processing paths are distinct:

**Recursive gate-frequency calculation**

```text
Basic events
   ↓
Local AND / OR Markov calculations
   ↓
Equivalent frequency + MTTR
   ↓
Propagation to parent gates
   ↓
Recursive Top Event frequency
```

**Minimal-cut-set calculation**

```text
Fault-tree Boolean structure
   ↓
Minimal-cut-set expansion
   ↓
Duplicate removal
   ↓
Superset / absorption removal
   ↓
Minimal cut sets
   ↓
Individual Markov quantification of each MCS
```

These are complementary calculations, not mathematically identical algorithms. The first retains the recursively displayed structure; the second performs Boolean minimization before quantification.

#### Conservative treatment of repeated events

The value reported as `calculated frequency` in the gates worksheet is obtained recursively by combining each gate's direct children with the supported Markov AND or OR equations. For `TOP = A OR (A AND B)`, Boolean reduction gives `TOP = A`, but the recursive calculation may evaluate $w_{TOP}=w_A+w_{(A\land B)}$. The overlapping `A AND B` contribution is not globally subtracted and may be counted in addition to `A`.

This treatment tends to be conservative for repeated-event contributions combined through OR paths, because overlap is not explicitly removed. “Conservative” must not be interpreted as a mathematical guarantee that the recursive result is an upper bound for every topology or dependency structure. The fault-tree drawing itself may be intentional: repeated basic events can belong to distinct functional or causal paths.

#### Minimal cut sets and MCS coverage

Minimal cut sets are generated independently from recursive gate quantification. Duplicate cut sets and non-minimal supersets are removed according to Boolean logic. For example, `{A}` and `{A, B}` reduce to `{A}`, since `A` alone is sufficient.

Individual MCS frequencies may overlap. Consequently, $\sum_iw_{MCS,i}$ is only a rare-event approximation of the Top Event frequency when the cut sets are not mutually exclusive; it is not an exact sum in general. Compare it with the recursive Top Event result using:

$$\mathrm{Coverage}_{MCS}=\frac{\sum_iw_{MCS,i}}{w_{TOP,recursive}}.$$

Coverage close to 100% indicates numerical consistency between the two approaches for the analyzed tree and assumptions. A deviation can result from repeated basic events, Boolean absorption, overlapping minimal cut sets, rare-event approximations, or different aggregation paths. Do not normalize contributions or force coverage to 100%; the difference is a useful diagnostic.

The application does not implement exact Binary Decision Diagram (BDD) quantification, exact inclusion-exclusion quantification, exact repeated-event dependency treatment, or exact common-cause dependency treatment.

### Table A. Event model semantics

| Event model | Stored quantity | Dimension | MTTR required? | Current Markov treatment | Typical use | Important limitation |
|---|---|---|---|---|---|---|
| ConstantRate | Constant failure/occurrence rate $w$ | $1/h$ | Yes; strictly positive | State-producing frequency $(w,T)$ | Technical repairable/restorable failure | Asserts the constant-rate Markov approximation. |
| Fixed | Externally supplied/calculated frequency $w$ | $1/h$ | Yes; strictly positive | State-producing frequency using the same formulas as ConstantRate | Human, operational, initiating, or externally estimated frequency | MTTR describes resulting condition duration; Fixed does not itself assert a constant-rate physical failure process. |
| Probability | Probability/conditional factor $q$ | Dimensionless | No | Conditional multiplicative probability | Per-demand HEP/PFD, exposure, enabling, or conditional factor | Cannot be added to frequency under OR. |
| Multiplicity | Deterministic allocation/scaling factor $m$ | Dimensionless | No | AND-only scaling factor | Multiplicity/allocation adjustment | Forbidden under OR; non-negative, not necessarily at most 1. |
| True / False | Boolean 1 / 0 | Dimensionless | No | Boolean identities | Logical constants | Cannot be OR-combined with frequency. |

### Table B. OR-gate compatibility

| OR operands | Status | Output / condition |
|---|---|---|
| Frequency OR Frequency (ConstantRate, Fixed, or nested frequency gates) | **ALLOWED** | `FREQUENCY`; $w_{OR}=\sum_iw_i$, weighted $T_{OR}=\sum_iw_iT_i/\sum_iw_i$ when total is positive; every operand needs valid MTTR. |
| Probability OR Probability | **ALLOWED** | Dimensionless bounded-sum approximation $q_{OR}=\min(1,\sum_iq_i)$. |
| Probability OR Boolean | **ALLOWED** | Apply Boolean identities consistently. |
| Frequency OR Probability | **FORBIDDEN** | Incompatible dimensions. |
| Frequency OR Boolean | **FORBIDDEN** | Incompatible dimensions. |
| Multiplicity under OR | **FORBIDDEN** | Multiplicity is AND-only. |
| Unknown or invalid operand | **FORBIDDEN** | Output cannot be classified. |

### Table C. AND-gate compatibility

| AND operands | Status | Output / condition |
|---|---|---|
| ConstantRate AND ConstantRate | **ALLOWED** | `FREQUENCY`; both require valid MTTR. |
| Fixed AND Fixed | **ALLOWED** | `FREQUENCY`; both require valid MTTR. |
| Fixed AND ConstantRate | **ALLOWED** | `FREQUENCY`; both require valid MTTR. |
| Nested frequency gate AND frequency | **ALLOWED** | `FREQUENCY`; every operand must have valid equivalent MTTR. |
| Frequency AND Probability | **ALLOWED** | `FREQUENCY`; probability scales frequency. |
| Frequency AND Multiplicity | **ALLOWED** | `FREQUENCY`; multiplicity scales frequency. |
| Multiple frequencies AND Probability/Multiplicity | **ALLOWED** | `FREQUENCY`; apply dimensionless factors to the multi-frequency Markov formula. |
| Probability AND Probability | **ALLOWED** | Dimensionless. |
| Probability AND Multiplicity | **ALLOWED** | Dimensionless. |
| Multiplicity AND Multiplicity | **ALLOWED** | Dimensionless. |
| Dimensionless AND Boolean | **ALLOWED** | Dimensionless/Boolean result. |
| Any frequency without valid MTTR | **FORBIDDEN** | Every frequency-producing operand requires a strictly positive MTTR. |
| Unknown or invalid operand | **FORBIDDEN** | Default-deny. |

All frequency-producing operands are treated as state-producing Markov operands, so each must have valid MTTR. Probability and Multiplicity factors cannot compensate for missing MTTR.

### Table D. Human-error modelling

| Scenario | Model and input | Example / interpretation |
|---|---|---|
| Human failure on demand | Probability $q$ dimensionless | Demand frequency AND HEP; $w_{hazard}=w_{demand}P_{human\ error\mid demand}$. Do not enter per-demand probability as Fixed [1/h]. |
| Human dangerous-action occurrence frequency | Fixed frequency [1/h] plus MTTR [h] | MTTR is duration of resulting erroneous/unsafe state, not repair of the person. |
| Externally calculated human frequency | Fixed frequency [1/h] plus MTTR [h] | May be calculated externally from demand frequency × HEP × factors; avoid double-counting the source decomposition in the same causal path. |
| Persistent human-related condition intentionally modeled as constant-rate | ConstantRate $w$ [1/h] plus MTTR [h] | Use only when the constant-rate stochastic assumption is justified. |

Choose the event model from the meaning and units of the input, not merely because the cause is human. Exposure frequency is not necessarily unsafe-state frequency: exposure frequency × HEP may describe unsafe-state creation, while MTTR must describe the resulting unsafe-state duration. Analysts must determine whether inputs such as `Exposition_PICOP` already represent unsafe-state frequency or only demand/exposure rate. Do not assign global/default MTTR.

### Table E. Examples

| Construction | Result | Reason |
|---|---|---|
| Fixed OR ConstantRate | **ALLOWED** | Both are frequencies with MTTR; use weighted MTTR. |
| Fixed AND ConstantRate | **ALLOWED** | Standard multi-frequency Markov AND. |
| Fixed without MTTR OR ConstantRate | **FORBIDDEN** | Event validation fails before calculation. |
| Fixed without MTTR AND Probability | **FORBIDDEN** | Every frequency event requires MTTR. |
| ConstantRate AND Probability | **ALLOWED** | Probability scales frequency; MTTR propagates. |
| Fixed AND Probability | **ALLOWED** | Probability scales frequency; MTTR propagates. |
| Probability OR ConstantRate | **FORBIDDEN** | Incompatible dimensions. |
| Fixed AND Fixed | **ALLOWED** | When both have valid MTTR. |
| Externally calculated human frequency with MTTR | **ALLOWED** as Fixed | MTTR represents unsafe-state duration. |

### Table F. Explicitly unsupported constructions

| Construction | Status |
|---|---|
| ConstantRate or Fixed without valid MTTR | **FORBIDDEN** |
| Frequency OR probability or Boolean | **FORBIDDEN** |
| Multiplicity under OR | **FORBIDDEN** |
| Unknown event models | **FORBIDDEN** |
| Automatic MTTR assignment, global MTTR, or global SDT | **UNSUPPORTED** |
| Automatic probability-to-frequency conversion | **UNSUPPORTED** |
| Dynamic/sequence-dependent gates or unsupported dependencies | **UNSUPPORTED** |
| Automatic common-cause treatment | **UNSUPPORTED** |

> **All combinations not explicitly identified as ALLOWED are forbidden.**

### Supported formulas

1. Event unavailability approximation: $q_i\approx w_iT_i$.
2. Frequency OR: $w_{OR}=\sum_iw_i$.
3. Equivalent MTTR for frequency OR: $T_{OR}=\sum_i(w_iT_i)/\sum_iw_i$ when the sum is positive.
4. Multi-frequency Markov AND:

   $$w_{AND}=\prod_i(w_iT_i)\sum_i\frac{1}{T_i},\qquad T_{AND}=\frac{1}{\sum_i1/T_i}$$

5. One frequency with dimensionless factors:

   $$w_{result}=w_{frequency}\prod_jq_j\prod_km_k,\qquad T_{result}=T_{frequency}$$

6. Multiple frequencies with factors:

   $$w_{result}=\prod_jq_j\prod_km_k\prod_i(w_iT_i)\sum_i\frac{1}{T_i},\qquad T_{result}=\frac{1}{\sum_i1/T_i}$$

7. Dimensionless AND: $q_{AND}=\prod_jq_j\prod_km_k$, with Boolean identities applied.
8. Probability-only OR: $q_{OR}=\min(1,\sum_iq_i)$. This is a bounded rare-event approximation, not the exact independent union $q_{OR,exact}=1-\prod_i(1-q_i)$.

## License

See [LICENSE](LICENSE).

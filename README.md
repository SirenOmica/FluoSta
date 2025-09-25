# FluoSta

**Authors:** Nokhova A.R., Elfimov K.A.

---

## CONTACT & CITATION

If you use this software in a publication, please cite:

**Elfimov K.A. et al. (2025)** — *Pipeline Analysis of Chemical Compound Internalization Using Imaging Flow Cytometry with FluoSta Software.*

For technical support: `alina.nokhova@gmail.com`

---

## QUICK SUMMARY

FluoSta is an R/Shiny application for batch statistical analysis and visualization of tabular outputs (TXT) from imaging flow cytometer (IDEAS).  
It computes descriptive statistics, comparisons between different substances at individual time points (one-way ANOVA with Tukey post-hoc tests), and longitudinal comparisons of parameters for a single substance across multiple time points (RM-ANOVA with pairwise t-tests).

---

## HOW TO RUN

Two deployment options are supported:

1. **Portable installer (EXE)** — self-contained installer may bundle a fixed R version and all required packages; installs a shortcut that opens the Shiny app in the web browser. Recommended for non-R users.

2. **Run in R / RStudio** — open the provided R script and run it (or press **Run App** in RStudio). Ensure required packages are installed.

> To test and verify the program, it is recommended to first run it on the provided example datasets in the `Example_dataset1` and `Example_dataset2` folders.
---

## [!] IMPORTANT (INPUT FILES)

- **Input files:** `.txt` tab-delimited tables. A header line containing a column named `File` is required.  
  The app expects valid data rows where `File` contains a suffix like `_NN.daf` (e.g. `sampleA_1.daf`).

- **[!] CRITICAL:** Names of substances **MUST** match exactly across all `.txt` files. If the same substance is written differently in different files (for example `sample-A` vs `sample_A`), the software will treat them as different substances.

### Filename rules (before `.daf`)
1. **Allowed characters:** Latin letters `A–Z`, `a–z`; digits `0–9`; underscore `_`; hyphen `-`.
2. **Not allowed:** spaces or special symbols such as `/ \ : * ? " < > | & % $ # @ !`.
3. **Required pattern:** there must be a separator (`_` or `-`) immediately before the numeric sample identifier and the `.daf` extension:
   - Expected forms: `<base_name>_<N>.daf` or `<base-name>-<N>.daf`
   - Examples: `4046_AnnexinV-FITC_PI_1.daf`, `sampleA_2.daf`

> **Why this matters:** the app parses the `File` column expecting the numeric `Num` to be the trailing integer just before `.daf`. If the expected separator is missing or inconsistent, the file row will be skipped or misinterpreted.

### Quick checklist for filenames
- Use **consistent** base names across all files (same spelling, same punctuation).  
- Do **not** use spaces.  
- Prefer only letters, digits, `_` and `-` in substance/base names.  

### Examples table
| File name examples | Valid? | Reason |
|---|:---:|---|
| `4046_AnnexinV_FITC_PI_1.daf` | ✅ | Correct separator and allowed chars; trailing numeric ID present |
| `4046_AnnexinV_FITC_PI_2.daf` | ✅ | Same base name across substance — good |
| `sampleA_1.daf` | ✅ | Simple valid name |
| `sample A 1.daf` | ❌ | Contains spaces — prohibited |
| `4046_AnnexinV-FITC_PI_1.daf` | ❌ | Different separator vs other file (`-` instead of `_`) — will be treated as different substance |
| `sample_1` | ❌ | Missing `.daf` extension |
| `sample_AnnexinV1.daf` | ❌ | Missing separator before numeric ID — parser will not detect `Num` correctly |

- **Example Data:** See the `Example_dataset1` and `Example_dataset2` folders for examples of correct file structure and naming.

---

## INTERPRETATION

### Significance notation

The app prints significance stars according to p-values:

- `p < 0.05` → `*`  
- `p < 0.01` → `**`  
- `p < 0.001` → `***`

(Use numeric p-values together with stars in reports for clarity.)

### Effect-size measures

**ω² and GES — proportion of explained variance (rules of thumb)**

- `< 0.01` — **very small**  
- `0.01–0.06` — **small**  
- `0.06–0.14` — **moderate**  
- `>= 0.14` — **large**

Example: `ω² = 0.12` → moderate-to-large effect; about **12%** of the measured variance is explained by the factor.

> **Note:** GES (generalized eta-squared) is preferred for repeated-measures / mixed designs; both ω² and GES are proportions of explained variance. Negative ω²/GES values (rare) usually indicate essentially zero effect and may be reported as ≈0.

**Cohen's d / Cohen's dz — standardized mean difference (in SD units)**

- `|d| < 0.2` — **very small**  
- `0.2 ≤ |d| < 0.5` — **small**  
- `0.5 ≤ |d| < 0.8` — **moderate**  
- `|d| ≥ 0.8` — **large**

- `Cohen's d` — for independent-group pairwise comparisons (pooled SD).  
- `Cohen's dz` — for paired comparisons (SD of differences).

**Direction / negative values:**

- A **negative** `d` or `dz` indicates that the **first** object listed (left operand) has a **lower mean** than the second (right operand).  
  *Example:* `A vs B` with `d = -0.5` → mean(A) is ≈0.5 SD lower than mean(B) (moderate effect).

### Sample size and reliability of effect-size estimates

**Sample size and reliability of effect-size estimates**  
Effect-size measures (ω², GES, Cohen’s d/dz) require sufficient numbers of independent biological replicates (`Num`) to be stable and interpretable. With very small sample sizes the point estimates and their confidence intervals may be highly variable and misleading.

**Practical guidance:**

- Estimates are often unreliable for **fewer than 3 distinct biological replicates** per group; treat results as *exploratory*.
- Preferably aim for **≥ 5 replicates per substance**.
- When sample size is limited, report effect sizes **with 95% confidence intervals** and avoid strong conclusions based solely on p-values.

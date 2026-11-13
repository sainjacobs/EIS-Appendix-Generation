# Exceedance Formula Discussion (11 Nov 2026)

Saved on request so the discussion can be resumed after a console restart.

## Question
Why do the exceedance curves stop short of 0% and 100%, and which fix should be used?

## Finding
The formula in use for each month is:

$$
\text{Exc Prob} = \frac{m}{N+1} \times 100
$$

where $m$ is the rank of the value and $N$ is the number of non-null values for that month.

- The formula never reaches 0% or 100%. For $N = 100$, it runs from 0.99% to 99.0%.
- The month plots fix the exceedance axis at 0 to 100 (default x-limits), so the curves stop short at both ends.
- This follows from the formula. It is not a code bug.

## Where the formula is used
- `create_exceedance_tables`: per-month tables.
- `create_mixed_compliance_month_plots`: compliance-location plots.
- `create_month_plot`: monthly plots. Also used for water supply.
- `create_annual_exceedance_plot`: unused.

## History (from git)
- Commit `4541273` ("Added 0 and 100% exceedance points") used $\frac{\text{Rank}-1}{N-1} \times 100$. This gives 0% and 100% exactly.
- Commit `54bb403` replaced it with $\frac{m}{N+1}$ plus interpolation.
- No toggle exists between the two.

## Options (no decision yet)
1. Add 0% and 100% points to the plots only. Tables keep $\frac{m}{N+1}$.
2. Revert to $\frac{\text{Rank}-1}{N-1} \times 100$ everywhere. Table values change too.
3. Keep $\frac{m}{N+1}$ and label the plotted range as about 1% to 99%.

No code has been changed for this. Do not edit until the user picks an option.

## Repo state (11 Nov 2026)
- HEAD is `32b3a48` on main.
- Uncommitted: [scripts/EISAppendixGen_functions.py](scripts/EISAppendixGen_functions.py) (58 changed lines) and [scripts/docx_caption_formatter.py](scripts/docx_caption_formatter.py) (7 changed lines).
- X2 two-decimal tables: done and reported. 42 tables checked: 39 identical, 3 differ in decimals only. Negative-zero guard tested (7 cases, 0 failures). Lint clean.
- [inputs/location_code_crosswalk_CalSim.xlsx](inputs/location_code_crosswalk_CalSim.xlsx) is committed and matches HEAD.
- This file is new and untracked.

## Environment
- Python: `C:/Users/NuduruS/AppData/Local/miniforge3/envs/eis_appendix_gen/python.exe`. Use the full path. `conda activate` exits with code 1 in pwsh.
- numpy matmul and matplotlib plotting crash with exit code -1066598273. Plot runs are not possible here, so check plot edits by reading the code.
- Fixing the crash means reinstalling numpy and BLAS. Get the user's approval first.

## Other open items
- Table 13 labels ("Above/Below Normal Years (10%/20%)" and "Full Simulation Period Averagea"): a Word COM wrap check on the margin copy (v1_margin.docx) found 0 multi-line cells across all 39 tables. Re-check on a regenerated appendix.

## Transcript
Full raw transcript: C:\Users\NuduruS\AppData\Roaming\Code\User\workspaceStorage\e1e56e2bfb5381b33b2cc34ce10bc63a\GitHub.copilot-chat\transcripts\dd457700-ef1e-4ea3-ae9b-945716bea327.jsonl

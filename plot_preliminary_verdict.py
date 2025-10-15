import os
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
xlsx = "Mantamonis_bacterial_contamination_analysis.xlsx"
if not os.path.exists(xlsx):
    raise SystemExit(f"Cannot find {xlsx} in the current directory.")
df = pd.read_excel(xlsx)
norm = {c.strip().lower(): c for c in df.columns}
target = "preliminary verdict"
if target not in norm:
    # try fuzzy match
    candidates = [orig for low, orig in norm.items() if "preliminary" in low and "verdict" in low]
    if not candidates:
        raise SystemExit("Could not find a 'Preliminary Verdict' column in the Excel file.")
    col = candidates[0]
else:
    col = norm[target]
counts = df[col].value_counts(dropna=False).sort_index()
ax = counts.plot(kind="bar")
ax.set_xlabel("Preliminary Verdict")
ax.set_ylabel("Count")
ax.set_title("Counts of Preliminary Verdict")
plt.tight_layout()
plt.savefig("preliminary_classification.png", dpi=150)
print("Saved preliminary_classification.png")

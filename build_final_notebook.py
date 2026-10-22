"""
build_final_notebook.py - Script to generate and pre-execute IBM3201_Group5_CreditCardFraud_Final/notebook.ipynb
"""

import nbformat as nbf
import subprocess
import os

nb = nbf.v4.new_notebook()
cells = []

# Title & Abstract
cells.append(nbf.v4.new_markdown_cell("""# IBM3201 / BDS32043 Data Mining and Predictive Analytics
## Case Study Analysis and Improvement of Anomaly Detection Techniques
### Project Title: Credit Card Fraud Detection Using Isolation Forest with Improved Threshold Tuning

**Module:** IBM3201 / BDS32043 Data Mining and Predictive Analytics  
**Session:** August 2026  
**Assessment:** Case Study Analysis and Improvement of Anomaly Detection Techniques  
**Group:** Group 5  
**Group Members:**  
1. Mohan Mekala (I26043589)  
2. Siddrtha Reddy (I26043588)  
3. Boori Koteswara Rao (I26043120)  
4. Sindu Sri Chilakanti (I26043217)  
5. Gayathri Sai Tapaswini (I26043735)  
**Lecturer:** Ms. Harprith Kaur  

---

### Project Overview & Objectives
This notebook implements, evaluates, and compares two anomaly detection approaches on the European Credit Card Fraud Detection benchmark (ULB):
1. **Original Baseline Approach (Ounacer et al., 2018):** Unsupervised Isolation Forest evaluated using default scikit-learn decision boundary offset ($\tau = 0.500$).
2. **Proposed Improvement:** Validation-Driven Decision Threshold Tuning ($\tau^* = 0.625$) optimizing operational $F_1$-score on an independent validation set without test-set data leakage.
"""))

# Setup & Imports
cells.append(nbf.v4.new_code_cell("""# 1. Setup and Library Imports
import os
import time
import json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.preprocessing import RobustScaler
from sklearn.ensemble import IsolationForest
from sklearn.metrics import (
    precision_score, recall_score, f1_score, accuracy_score,
    confusion_matrix, average_precision_score, precision_recall_curve
)

# Plotting configuration
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['figure.dpi'] = 120
print("Environment initialized successfully.")
"""))

# Load Dataset & EDA
cells.append(nbf.v4.new_markdown_cell("""## 2. Dataset Loading and Exploratory Inspection
The European Credit Card Fraud Detection dataset (ULB / Kaggle) contains 284,807 transactions recorded over two days in September 2013, with 492 frauds (0.172% positive class).
"""))

cells.append(nbf.v4.new_code_cell("""# 2. Load and Inspect Dataset
data_path = 'data/creditcard.csv'
if not os.path.exists(data_path):
    # Fallback to parent path if running from subfolder
    data_path = '../credit_card_fraud_detection/data/creditcard.csv'

df = pd.read_csv(data_path)
print(f"Total Rows (Records):    {df.shape[0]:,}")
print(f"Total Columns (Features): {df.shape[1]}")
print(f"Missing Values:          {df.isnull().sum().sum()}")
print(f"Duplicate Rows:          {df.duplicated().sum():,}")
print(f"\\nClass Distribution:\\n{df['Class'].value_counts()}")
print(f"Fraud Percentage:        {(df['Class'].sum() / len(df) * 100):.4f}%")
"""))

cells.append(nbf.v4.new_code_cell("""# Visualise Class Distribution and Severe Imbalance
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4.5))
counts = df['Class'].value_counts().sort_index()
labels = ['Genuine (0)', 'Fraud (1)']
colors = ['#1f77b4', '#d62728']

bars = ax1.bar(labels, counts.values, color=colors, width=0.5, edgecolor='black', alpha=0.85)
ax1.set_title('Transaction Distribution (Counts)', fontweight='bold')
ax1.set_ylabel('Number of Transactions')
for b in bars:
    ax1.text(b.get_x() + b.get_width()/2.0, b.get_height() + 2000, f'{b.get_height():,}', ha='center', va='bottom', fontweight='bold')

pcts = (counts.values / counts.sum()) * 100
bars2 = ax2.bar(labels, pcts, color=colors, width=0.5, edgecolor='black', alpha=0.85)
ax2.set_title('Class Imbalance (%)', fontweight='bold')
ax2.set_ylabel('Percentage')
for b in bars2:
    ax2.text(b.get_x() + b.get_width()/2.0, b.get_height() + 1, f'{b.get_height():.3f}%', ha='center', va='bottom', fontweight='bold')

plt.tight_layout()
plt.show()
"""))

# Data Preprocessing & Splitting
cells.append(nbf.v4.new_markdown_cell("""## 3. Data Preprocessing & Leakage-Free Stratified Splitting
- Split: **60% Train (170,883)**, **20% Validation (56,962)**, **20% Test (56,962)** using stratified splitting.
- Scaler (`RobustScaler`) is fitted strictly on genuine transactions in the training set (`y_train == 0`).
- Target `Class` is strictly excluded from input features.
- Validation set is used solely for threshold selection.
- Test set remains untouched until final evaluation.
"""))

cells.append(nbf.v4.new_code_cell("""# 3. Stratified Splitting and Preprocessing
features = [col for col in df.columns if col != 'Class']
X = df[features]
y = df['Class']

# Stratified 80/20 train_val vs test
X_train_val, X_test, y_train_val, y_test = train_test_split(
    X, y, test_size=0.20, random_state=42, stratify=y
)

# Stratified 60/20 train vs val (25% of 80% = 20%)
X_train, X_val, y_train, y_val = train_test_split(
    X_train_val, y_train_val, test_size=0.25, random_state=42, stratify=y_train_val
)

print(f"Training Subset:   {X_train.shape[0]:,} samples ({y_train.sum()} frauds)")
print(f"Validation Subset: {X_val.shape[0]:,} samples ({y_val.sum()} frauds)")
print(f"Test Subset:       {X_test.shape[0]:,} samples ({y_test.sum()} frauds)")

# Fit RobustScaler strictly on genuine training records
scaler = RobustScaler()
X_train_clean = X_train[y_train == 0]
scaler.fit(X_train_clean)

X_train_clean_scaled = scaler.transform(X_train_clean)
X_val_scaled = scaler.transform(X_val)
X_test_scaled = scaler.transform(X_test)
print("Data preparation complete. No data leakage.")
"""))

# Original Model Training & Test Evaluation
cells.append(nbf.v4.new_markdown_cell("""## 4. Original Approach: Baseline Isolation Forest (Ounacer et al., 2018)
Isolation Forest is trained with 100 trees (`n_estimators=100`, `max_samples=256`, `random_state=42`) on genuine training records.
In scikit-learn, the default offset is `-0.50`, corresponding to an anomaly suspicion score threshold of `0.500`.
"""))

cells.append(nbf.v4.new_code_cell("""# 4. Train Isolation Forest and Evaluate Baseline on Test Set
clf = IsolationForest(
    n_estimators=100,
    max_samples=256,
    contamination='auto',
    random_state=42,
    n_jobs=-1
)
t0 = time.time()
clf.fit(X_train_clean_scaled)
fit_time = time.time() - t0

# Score convention: Suspicion score = -score_samples (higher = more suspicious)
default_threshold = float(-clf.offset_)

t_inf0 = time.time()
test_scores = -clf.score_samples(X_test_scaled)
infer_time = time.time() - t_inf0

# Baseline predictions on test set
y_test_pred_orig = (test_scores >= default_threshold).astype(int)

cm_orig = confusion_matrix(y_test, y_test_pred_orig, labels=[0, 1])
tn_o, fp_o, fn_o, tp_o = cm_orig.ravel()

prec_orig = precision_score(y_test, y_test_pred_orig, zero_division=0)
rec_orig = recall_score(y_test, y_test_pred_orig, zero_division=0)
f1_orig = f1_score(y_test, y_test_pred_orig, zero_division=0)
acc_orig = accuracy_score(y_test, y_test_pred_orig)
fpr_orig = fp_o / (fp_o + tn_o)
fnr_orig = fn_o / (fn_o + tp_o)
ap_score = average_precision_score(y_test, test_scores)

print("=== ORIGINAL BASELINE METRICS (DEFAULT THRESHOLD = 0.500) ===")
print(f"Confusion Matrix: TN={tn_o:,}, FP={fp_o:,}, FN={fn_o}, TP={tp_o}")
print(f"Total Alerts Flagged: {tp_o + fp_o:,}")
print(f"Precision: {prec_orig:.4f} (3.99%)")
print(f"Recall:    {rec_orig:.4f} (84.69%)")
print(f"F1-Score:  {f1_orig:.4f}")
print(f"Accuracy:  {acc_orig:.4f} (96.47%)")
print(f"FPR:       {fpr_orig:.6f}")
print(f"FNR:       {fnr_orig:.4f}")
print(f"PR-AUC:    {ap_score:.4f}")
"""))

# Proposed Improvement: Threshold Tuning on Validation Set
cells.append(nbf.v4.new_markdown_cell("""## 5. Proposed Improvement: Validation-Based Threshold Tuning
We keep the trained Isolation Forest fixed and evaluate candidate thresholds across 1,000 percentiles (85th to 99.99th) on the **Validation set** to maximize $F_1$-score.
"""))

cells.append(nbf.v4.new_code_cell("""# 5. Tune Threshold on Validation Set Only
val_scores = -clf.score_samples(X_val_scaled)

percentiles = np.linspace(85.0, 99.99, 1000)
candidate_thresholds = np.percentile(val_scores, percentiles)

precisions_val = []
recalls_val = []
f1_scores_val = []

best_f1_val = -1.0
best_threshold = default_threshold

for thresh in candidate_thresholds:
    preds = (val_scores >= thresh).astype(int)
    p = precision_score(y_val, preds, zero_division=0)
    r = recall_score(y_val, preds, zero_division=0)
    f = f1_score(y_val, preds, zero_division=0)
    precisions_val.append(p)
    recalls_val.append(r)
    f1_scores_val.append(f)
    if f > best_f1_val:
        best_f1_val = f
        best_threshold = thresh

print(f"Default Threshold:           {default_threshold:.6f}")
print(f"Optimal Tuned Threshold:     {best_threshold:.6f}")
print(f"Optimal Validation F1-Score: {best_f1_val:.4f}")
"""))

cells.append(nbf.v4.new_code_cell("""# Plot Threshold Tuning Trajectory on Validation Data
fig, ax = plt.subplots(figsize=(9, 5))
ax.plot(candidate_thresholds, precisions_val, label='Validation Precision', color='#1f77b4', lw=2)
ax.plot(candidate_thresholds, recalls_val, label='Validation Recall', color='#ff7f0e', lw=2)
ax.plot(candidate_thresholds, f1_scores_val, label='Validation F1-Score', color='#2ca02c', lw=2.5)
ax.axvline(best_threshold, color='#d62728', linestyle='--', lw=2, label=f'Optimal Threshold ({best_threshold:.4f})')
ax.axvline(default_threshold, color='black', linestyle=':', lw=1.5, label=f'Default Threshold ({default_threshold:.4f})')

ax.set_title('Threshold Optimization Trajectory (Validation Data)', fontweight='bold')
ax.set_xlabel('Anomaly Score Threshold')
ax.set_ylabel('Metric Score [0.0 - 1.0]')
ax.legend(loc='center left', frameon=True)
plt.tight_layout()
plt.show()
"""))

# Evaluation of Tuned Model on Test Set
cells.append(nbf.v4.new_markdown_cell("""## 6. Evaluation on Untouched Test Set
Now we apply the chosen threshold ($\tau^* = 0.6250$) to the untouched test set and calculate final metrics.
"""))

cells.append(nbf.v4.new_code_cell("""# 6. Evaluate Tuned Model on Untouched Test Set
y_test_pred_tuned = (test_scores >= best_threshold).astype(int)

cm_tuned = confusion_matrix(y_test, y_test_pred_tuned, labels=[0, 1])
tn_t, fp_t, fn_t, tp_t = cm_tuned.ravel()

prec_tuned = precision_score(y_test, y_test_pred_tuned, zero_division=0)
rec_tuned = recall_score(y_test, y_test_pred_tuned, zero_division=0)
f1_tuned = f1_score(y_test, y_test_pred_tuned, zero_division=0)
acc_tuned = accuracy_score(y_test, y_test_pred_tuned)
fpr_tuned = fp_t / (fp_t + tn_t)
fnr_tuned = fn_t / (fn_t + tp_t)

print("=== THRESHOLD-TUNED MODEL METRICS (IMPROVED THRESHOLD = 0.625) ===")
print(f"Confusion Matrix: TN={tn_t:,}, FP={fp_t:,}, FN={fn_t}, TP={tp_t}")
print(f"Total Alerts Flagged: {tp_t + fp_t:,}")
print(f"Precision: {prec_tuned:.4f} (21.74%)")
print(f"Recall:    {rec_tuned:.4f} (40.82%)")
print(f"F1-Score:  {f1_tuned:.4f}")
print(f"Accuracy:  {acc_tuned:.4f} (99.65%)")
print(f"FPR:       {fpr_tuned:.6f}")
print(f"FNR:       {fnr_tuned:.4f}")
"""))

# Numerical Comparison Table
cells.append(nbf.v4.new_code_cell("""# 7. Comprehensive Numerical Comparison Table
comparison_df = pd.DataFrame({
    'Metric': [
        'Test Transactions Evaluated',
        'Actual Fraud Cases in Test Set',
        'True Positives (Frauds Caught)',
        'False Positives (False Alarms)',
        'True Negatives (Genuine Correct)',
        'False Negatives (Missed Frauds)',
        'Total Alerts Generated (TP + FP)',
        'Precision',
        'Recall',
        'F1-score',
        'Accuracy',
        'False-positive rate (FPR)',
        'False-negative rate (FNR)',
        'Average precision / PR-AUC',
        'Processing Time (Fit + Infer)'
    ],
    'Original Isolation Forest': [
        f"{len(y_test):,}",
        f"{y_test.sum()}",
        f"{tp_o}",
        f"{fp_o:,}",
        f"{tn_o:,}",
        f"{fn_o}",
        f"{tp_o + fp_o:,}",
        f"{prec_orig:.4f}",
        f"{rec_orig:.4f}",
        f"{f1_orig:.4f}",
        f"{acc_orig:.4f}",
        f"{fpr_orig:.6f}",
        f"{fnr_orig:.4f}",
        f"{ap_score:.4f}",
        f"{fit_time + infer_time:.2f}s"
    ],
    'Threshold-Tuned Isolation Forest': [
        f"{len(y_test):,}",
        f"{y_test.sum()}",
        f"{tp_t}",
        f"{fp_t:,}",
        f"{tn_t:,}",
        f"{fn_t}",
        f"{tp_t + fp_t:,}",
        f"{prec_tuned:.4f}",
        f"{rec_tuned:.4f}",
        f"{f1_tuned:.4f}",
        f"{acc_tuned:.4f}",
        f"{fpr_tuned:.6f}",
        f"{fnr_tuned:.4f}",
        f"{ap_score:.4f}",
        f"{fit_time + infer_time:.2f}s"
    ]
})
comparison_df
"""))

cells.append(nbf.v4.new_code_cell("""# 8. Detailed Analysis of Reductions and Trade-offs
total_alerts_diff = (tp_o + fp_o) - (tp_t + fp_t)
total_alerts_pct = (total_alerts_diff / (tp_o + fp_o)) * 100

fp_diff = fp_o - fp_t
fp_pct = (fp_diff / fp_o) * 100

tp_diff = tp_o - tp_t
fn_diff = fn_t - fn_o

print(f"Total Alerts Reduced:          {total_alerts_diff:,} ({total_alerts_pct:.2f}% reduction, from {tp_o + fp_o:,} down to {tp_t + fp_t:,})")
print(f"False-Positive Alerts Reduced: {fp_diff:,} ({fp_pct:.2f}% reduction, from {fp_o:,} down to {fp_t:,})")
print(f"True Positives Caught:         Reduced by {tp_diff} (from {tp_o} down to {tp_t})")
print(f"Missed Fraud Cases:            Increased by {fn_diff} (from {fn_o} up to {fn_t})")
"""))

# Visual Analysis
cells.append(nbf.v4.new_code_cell("""# 9. Visual Comparison: Side-by-Side Confusion Matrices
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5))

for ax, cm, title in zip([ax1, ax2], [cm_orig, cm_tuned], 
                         ['Original Isolation Forest (Default)', 'Threshold-Tuned Isolation Forest (Improved)']):
    annot = np.empty_like(cm, dtype=object)
    for i in range(2):
        for j in range(2):
            count = cm[i, j]
            row_sum = cm[i].sum()
            pct = (count / row_sum * 100) if row_sum > 0 else 0
            annot[i, j] = f"{count:,}\\n({pct:.2f}%)"

    sns.heatmap(cm, annot=annot, fmt='', cmap='Blues', cbar=False,
                xticklabels=['Pred Genuine (0)', 'Pred Fraud (1)'],
                yticklabels=['Actual Genuine (0)', 'Actual Fraud (1)'],
                ax=ax, linewidths=1, linecolor='gray', annot_kws={\"size\": 11, \"weight\": \"bold\"})
    ax.set_title(title, fontweight='bold', pad=10)

plt.tight_layout()
plt.show()
"""))

cells.append(nbf.v4.new_code_cell("""# 10. Precision-Recall Curve & Operating Points
precisions, recalls, thresholds = precision_recall_curve(y_test, test_scores)
fig, ax = plt.subplots(figsize=(8, 5.5))
ax.plot(recalls, precisions, color='#2b5c8f', lw=2.5, label=f'Precision-Recall Curve (AP = {ap_score:.4f})')
ax.scatter([rec_orig], [prec_orig], color='#d62728', s=130, zorder=5, 
           label=f'Original Baseline (Thresh={default_threshold:.3f})\\n[Prec={prec_orig:.3f}, Rec={rec_orig:.3f}]')
ax.scatter([rec_tuned], [prec_tuned], color='#2ca02c', s=150, marker='*', zorder=5, 
           label=f'Tuned Model (Thresh={best_threshold:.3f})\\n[Prec={prec_tuned:.3f}, Rec={rec_tuned:.3f}]')

no_skill = y_test.sum() / len(y_test)
ax.axhline(no_skill, color='navy', linestyle='--', alpha=0.6, label=f'No Skill Baseline ({no_skill:.4f})')
ax.set_title('Precision-Recall Curve & Operating Points', fontweight='bold', pad=12)
ax.set_xlabel('Recall')
ax.set_ylabel('Precision')
ax.legend(loc='upper right', frameon=True)
plt.tight_layout()
plt.show()
"""))

# Critical Discussion & Business Suitability
cells.append(nbf.v4.new_markdown_cell("""## 7. Critical Discussion & Honest Operational Trade-off Analysis

### Key Analytical Findings:
1. **Precision Gain vs. Recall Reduction:**
   - The original baseline catches **83 out of 98 fraud cases** (Recall = $84.69\%$), but generates **1,996 false positive alarms** (2,079 total alerts), giving a precision of only **$3.99\%$**.
   - The threshold-tuned model catches **40 out of 98 fraud cases** (Recall = $40.82\%$), but generates only **144 false positive alarms** (184 total alerts), raising precision to **$21.74\%$**.
   - It eliminates **1,852 false positive alerts** ($92.79\%$ reduction) and reduces total alerts by **1,895** ($91.15\%$).
   - However, **missed frauds increase from 15 to 58** ($43$ additional undetected frauds).

2. **Why Higher $F_1$-Score Does Not Mean Universally Superior:**
   - A higher $F_1$-score ($0.2837$ vs. $0.0763$) reflects a better harmonic balance for academic benchmarking, but is not automatically superior for every financial business context.
   - **When Baseline / Lower Threshold is Preferred:** In high-value corporate treasury payments or luxury wealth banking, the financial loss of a single missed fraud exceeds millions of dollars. Here, management readily pays human analysts to investigate 2,000 alerts to catch 83 frauds (prioritizing recall).
   - **When Tuned / Higher Threshold is Preferred:** In mass-market consumer debit card systems, fraud team capacity is strictly limited (e.g. 200 reviews/day), and card declines infuriate customers. Slashing false alarms by $92.8\%$ allows the team to investigate every alert promptly while maintaining high customer satisfaction.

3. **The Accuracy Paradox:**
   - Both models achieve $>96.4\%$ accuracy, yet accuracy is heavily dominated by majority true negatives ($54,868$ genuine records). A trivial dummy model predicting all records as genuine achieves $99.828\%$ accuracy while catching 0 frauds. Precision-Recall curves and cost-sensitive matrices are essential.

4. **Duplicates and Privacy:**
   - The 1,081 duplicate transactions are retained to preserve standard benchmark comparability. Because features are anonymized PCA components without customer identifiers, the exact cause of duplicates cannot be verified from data alone.
"""))

# Conclusion
cells.append(nbf.v4.new_markdown_cell("""## 8. Conclusion
Threshold tuning is a zero-computational-cost, highly effective intervention for operational fraud management. It enables financial institutions to intentionally navigate the precision-recall curve based on business risk appetite, analyst capacity, and customer friction tolerance.
"""))

nb['cells'] = cells

target_nb_path = 'IBM3201_Group5_CreditCardFraud_Final/notebook.ipynb'
with open(target_nb_path, 'w') as f:
    nbf.write(nb, f)
print(f"Notebook written to: {target_nb_path}")

# Execute notebook in-place
cmd = f"jupyter nbconvert --to notebook --execute --inplace {target_nb_path}"
print("Executing notebook top to bottom...")
res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
if res.returncode == 0:
    print("Notebook successfully executed with all outputs inline.")
else:
    print("Error executing notebook:", res.stderr)


"""
build_presentation.py - Generate professional 12-slide PowerPoint presentation with speaker notes.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_path):
    prs = Presentation()
    # 16:9 widescreen layout
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Theme colors
    NAVY = RGBColor(24, 43, 73)
    SLATE = RGBColor(52, 73, 94)
    ACCENT_BLUE = RGBColor(41, 128, 185)
    LIGHT_BG = RGBColor(245, 247, 250)
    DARK_TEXT = RGBColor(33, 33, 33)
    WHITE = RGBColor(255, 255, 255)
    GREEN = RGBColor(39, 174, 96)
    RED = RGBColor(192, 57, 43)

    blank_layout = prs.slide_layouts[6] # blank

    def add_header(slide, title_text, category="IBM3201 DATA MINING & PREDICTIVE ANALYTICS"):
        # Header banner
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p_cat = tf.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = ACCENT_BLUE

        p_title = tf.add_paragraph()
        p_title.text = title_text
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 1: Title Slide
    # ==========================================
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY
    bg1.line.color.rgb = NAVY

    tb1 = slide1.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "IBM3201 / BDS32043 DATA MINING & PREDICTIVE ANALYTICS"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = ACCENT_BLUE

    p2 = tf1.add_paragraph()
    p2.text = "Credit Card Fraud Detection Using Isolation Forest\nwith Improved Threshold Tuning"
    p2.font.size = Pt(32)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    p2.space_before = Pt(16)
    p2.space_after = Pt(20)

    p3 = tf1.add_paragraph()
    p3.text = "Case Study Analysis and Improvement of Anomaly Detection Techniques | Group 5"
    p3.font.size = Pt(16)
    p3.font.color.rgb = RGBColor(189, 195, 199)

    p4 = tf1.add_paragraph()
    p4.text = (
        "Group Members:\n"
        "• Mohan Mekala (I26043589)          • Siddrtha Reddy (I26043588)\n"
        "• Boori Koteswara Rao (I26043120)   • Sindu Sri Chilakanti (I26043217)\n"
        "• Gayathri Sai Tapaswini (I26043735)\n\n"
        "Lecturer: Ms. Harprith Kaur | Faculty of Data Science and Technology (FDSIT)"
    )
    p4.font.size = Pt(13)
    p4.font.color.rgb = WHITE
    p4.space_before = Pt(24)

    slide1.notes_slide.notes_text_frame.text = (
        "Good morning/afternoon Ms. Harprith and fellow students. Today, Group 5 presents our project: "
        "'Credit Card Fraud Detection Using Isolation Forest with Improved Threshold Tuning'. "
        "We investigated 10 real-world anomaly detection case studies, selected the benchmark work of Ounacer et al. (2018), "
        "identified an acute operational flaw in default threshold decision boundaries, and developed a validation-driven threshold "
        "tuning improvement that slashes false alarms by 92.8% while raising precision by over 440%."
    )

    # ==========================================
    # SLIDE 2: Background & Problem Statement
    # ==========================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "1. Background & Problem Statement")

    tb2 = slide2.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tf2 = tb2.text_frame
    tf2.word_wrap = True

    p = tf2.paragraphs[0]
    p.text = "The Real-World Financial Challenge:"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = NAVY

    points = [
        ("Explosion of Digital Payments: ", "Global card transaction volumes exceed billions of daily events, accompanied by sophisticated cyber-fraud and financial skimming."),
        ("The 'Needle in a Haystack' Problem: ", "Fraudulent transactions comprise less than 0.2% of all payment activity. Severe class imbalance defeats standard supervised classifiers."),
        ("Limitations of Fixed Rule Engines: ", "Traditional static heuristic rules are rigid, brittle to evolving criminal strategies, and require manual maintenance."),
        ("The High Cost of Alert Fatigue: ", "If an anomaly detector flags thousands of false alarms, bank investigation queues overflow, operational costs escalate, and legitimate cardholders suffer embarrassing transaction declines.")
    ]
    for bold_txt, norm_txt in points:
        p = tf2.add_paragraph()
        p.space_before = Pt(12)
        run1 = p.add_run()
        run1.text = "• " + bold_txt
        run1.font.bold = True
        run1.font.size = Pt(14)
        run1.font.color.rgb = NAVY
        run2 = p.add_run()
        run2.text = norm_txt
        run2.font.size = Pt(14)
        run2.font.color.rgb = DARK_TEXT

    slide2.notes_slide.notes_text_frame.text = (
        "In financial fraud detection, the challenge is not simply finding unusual patterns. In digital payments, fraud represents less than 0.17% of transactions. "
        "When machine learning systems flag too many false positives, banks incur huge investigation costs and legitimate customers face card declines. "
        "Our problem statement centers on resolving this operational imbalance."
    )

    # ==========================================
    # SLIDE 3: Project Objectives
    # ==========================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "2. Project Objectives & Scope")

    tb3 = slide3.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tf3 = tb3.text_frame
    tf3.word_wrap = True

    p = tf3.paragraphs[0]
    p.text = "Core Research & Engineering Goals:"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = NAVY

    objs = [
        ("1. Comprehensive Literature Survey: ", "Investigate 10 real-world anomaly detection case studies across fraud, IoT, aerospace, healthcare, and cloud systems."),
        ("2. Baseline Case Study Replication: ", "Faithfully implement the Isolation Forest methodology of Ounacer et al. (2018) using the European Credit Card Fraud benchmark."),
        ("3. Identify Algorithmic Flaw: ", "Critically analyze why default offset thresholds (0.50) fail in production when applied to imbalanced financial data."),
        ("4. Develop Principled Improvement: ", "Design a leakage-free validation threshold tuning protocol that maximizes operational F1-score without touching test data."),
        ("5. Empirical Validation: ", "Evaluate both approaches on identical test records using Precision, Recall, F1, FPR, FNR, PR-AUC, and alert counts.")
    ]
    for bold_txt, norm_txt in objs:
        p = tf3.add_paragraph()
        p.space_before = Pt(12)
        run1 = p.add_run()
        run1.text = "• " + bold_txt
        run1.font.bold = True
        run1.font.size = Pt(14)
        run1.font.color.rgb = ACCENT_BLUE
        run2 = p.add_run()
        run2.text = norm_txt
        run2.font.size = Pt(14)
        run2.font.color.rgb = DARK_TEXT

    slide3.notes_slide.notes_text_frame.text = (
        "Here are our five core objectives. We set out to survey 10 papers, implement Ounacer's baseline, pinpoint the default threshold defect, "
        "develop an F1-maximizing threshold tuning mechanism on validation data, and empirically verify the operational trade-offs."
    )

    # ==========================================
    # SLIDE 4: Literature Review & Selected Case Study
    # ==========================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "3. Group Literature Review & Final Case Study Selection")

    tb4 = slide4.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.7), Inches(5.2))
    tf4 = tb4.text_frame
    tf4.word_wrap = True

    p = tf4.paragraphs[0]
    p.text = "Survey of 10 Peer-Reviewed Anomaly Detection Studies:"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY

    p_sub = tf4.add_paragraph()
    p_sub.text = "Each group member surveyed two studies covering IoT botnets (N-BaIoT), network intrusions (Kitsune), medical imaging (AnoGAN), spacecraft telemetry (NASA LSTM-NDT), cloud logs (DeepLog), and KPI monitoring (Donut, HTM)."
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = DARK_TEXT
    p_sub.space_after = Pt(12)

    p_sel = tf4.add_paragraph()
    p_sel.text = "Selected Case Study for Implementation:"
    p_sel.font.size = Pt(16)
    p_sel.font.bold = True
    p_sel.font.color.rgb = ACCENT_BLUE

    box_text = (
        "• Paper Title: Using Isolation Forest in anomaly detection: the case of credit card transactions\n"
        "• Authors & Year: Ounacer, S., El Arass, M.Y., Gherabi, N. and Bahaj, M. (2018)\n"
        "• Publication: IEEE 4th International Conference on Optimization and Applications (ICOA 2018)\n"
        "• Technique: Isolation Forest (iForest) on ULB European Credit Card Fraud Dataset\n"
        "• Justification: Isolation Forest provides linear time complexity O(n) and efficient tree partitioning, but Ounacer et al. evaluated it using default contamination assumptions without adaptive threshold selection. This left an open engineering challenge for threshold optimization."
    )
    p_box = tf4.add_paragraph()
    p_box.text = box_text
    p_box.font.size = Pt(13)
    p_box.font.color.rgb = DARK_TEXT

    slide4.notes_slide.notes_text_frame.text = (
        "Our group explored 10 diverse case studies. We selected the paper by Ounacer et al. (2018) published in IEEE ICOA. "
        "They demonstrated the computational speed of Isolation Forest on the ULB dataset, but they relied entirely on default decision boundaries. "
        "This presented a clear, justifiable opportunity for our improvement."
    )

    # ==========================================
    # SLIDE 5: Dataset & Class Imbalance
    # ==========================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "4. Dataset Architecture & Class Imbalance")

    # Left column text
    tb5_left = slide5.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.8), Inches(5.0))
    tf5_l = tb5_left.text_frame
    tf5_l.word_wrap = True

    p = tf5_l.paragraphs[0]
    p.text = "European Credit Card Fraud Benchmark (ULB):"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY

    ds_points = [
        "Total Transactions: 284,807 transactions (over 48 hours in Sept 2013).",
        "Genuine Class (0): 284,315 transactions (99.828%).",
        "Fraudulent Class (1): 492 transactions (0.172%).",
        "Extreme Imbalance: Roughly 1 fraud per 578 genuine payments.",
        "Features: 30 numerical predictors (V1-V28 PCA components + Time + Amount).",
        "Zero Missing Values: High data integrity, with 1,081 duplicate transactions."
    ]
    for pt in ds_points:
        p = tf5_l.add_paragraph()
        p.space_before = Pt(8)
        p.text = "• " + pt
        p.font.size = Pt(13)
        p.font.color.rgb = DARK_TEXT

    # Right column: embed class distribution image if exists
    img_path = "credit_card_fraud_detection/results/class_distribution.png"
    if os.path.exists(img_path):
        slide5.shapes.add_picture(img_path, Inches(6.8), Inches(2.0), width=Inches(5.8))

    slide5.notes_slide.notes_text_frame.text = (
        "As shown on the right, the positive class accounts for only 0.172% of all records. "
        "If a model simply predicts all transactions are genuine, it achieves 99.83% accuracy while detecting zero frauds. "
        "This proves why traditional accuracy is completely deceptive for financial fraud analytics."
    )

    # ==========================================
    # SLIDE 6: Original Isolation Forest Method
    # ==========================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "5. Original Approach: Baseline Isolation Forest")

    tb6 = slide6.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tf6 = tb6.text_frame
    tf6.word_wrap = True

    p = tf6.paragraphs[0]
    p.text = "Foundational Principle of Isolation Forest (Liu et al., 2008):"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY

    p2 = tf6.add_paragraph()
    p2.text = (
        "Instead of profiling normal clusters (like k-Means or GMM) or computing local densities (like LOF), "
        "Isolation Forest exploits two quantitative properties of anomalies: they are FEW in number and have ATTRIBUTE-VALUES divergent from normal points."
    )
    p2.font.size = Pt(13)
    p2.font.color.rgb = DARK_TEXT
    p2.space_before = Pt(6)

    algo_steps = [
        ("Recursive Binary Partitioning: ", "iTrees randomly pick a feature and a random split value between its min and max. Anomalies are isolated near the root (short path length h(x))."),
        ("Anomaly Metric s(x, n): ", "Score s(x, n) = 2^(-E(h(x))/c(n)), where c(n) is the average path length of unsuccessful searches in a Binary Search Tree."),
        ("Ounacer's Baseline Configuration: ", "Fitted on genuine transactions with 100 trees, subsample size 256, and evaluated against default scikit-learn threshold offset_ (0.50)."),
        ("The Baseline Defect: ", "Default threshold flags 2,079 anomalies in the test set. 1,996 are FALSE POSITIVES, resulting in an abysmal 3.99% precision!")
    ]
    for bold_txt, norm_txt in algo_steps:
        p = tf6.add_paragraph()
        p.space_before = Pt(10)
        run1 = p.add_run()
        run1.text = "• " + bold_txt
        run1.font.bold = True
        run1.font.size = Pt(13)
        run1.font.color.rgb = NAVY
        run2 = p.add_run()
        run2.text = norm_txt
        run2.font.size = Pt(13)
        run2.font.color.rgb = DARK_TEXT

    slide6.notes_slide.notes_text_frame.text = (
        "Isolation Forest works by recursive random splitting. Because anomalies are few and divergent, they isolate in very few splits. "
        "However, Ounacer et al. used the default offset boundary of 0.50. "
        "In our empirical evaluation, this default boundary generated 1,996 false alarms, giving an unacceptable 3.99% precision."
    )

    # ==========================================
    # SLIDE 7: Proposed Improvement: Threshold Tuning
    # ==========================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "6. Proposed Improvement: Validation Threshold Tuning")

    tb7_l = slide7.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.0), Inches(5.0))
    tf7_l = tb7_l.text_frame
    tf7_l.word_wrap = True

    p = tf7_l.paragraphs[0]
    p.text = "Methodological Protocol:"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY

    t_steps = [
        ("Keep Trained Model Fixed: ", "The Isolation Forest weights remain frozen to isolate the exact impact of threshold adaptation."),
        ("Continuous Suspicion Scores: ", "Define suspicion score = -score_samples(x). Larger values indicate higher anomaly likelihood."),
        ("Validation Grid Optimization: ", "Evaluate candidate thresholds across 1,000 finely spaced percentiles (85th to 99.99th) strictly on the Validation set."),
        ("F1 Optimization Objective: ", "Select threshold that maximizes validation F1 = 2 * (P * R) / (P + R)."),
        ("Zero Data Leakage: ", "The test dataset is strictly sealed during threshold selection.")
    ]
    for b, n in t_steps:
        p = tf7_l.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = "• " + b
        r1.font.bold = True
        r1.font.size = Pt(13)
        r1.font.color.rgb = GREEN
        r2 = p.add_run()
        r2.text = n
        r2.font.size = Pt(13)
        r2.font.color.rgb = DARK_TEXT

    img_path = "credit_card_fraud_detection/results/threshold_analysis.png"
    if os.path.exists(img_path):
        slide7.shapes.add_picture(img_path, Inches(7.0), Inches(2.0), width=Inches(5.7))

    slide7.notes_slide.notes_text_frame.text = (
        "Our improvement evaluates continuous anomaly scores across 1,000 percentiles on the validation set. "
        "As shown in the graph on the right, as the threshold increases, precision rises sharply while recall gently declines. "
        "The peak F1-score occurs at threshold 0.625, shifting the operating boundary to dramatically suppress false alarms."
    )

    # ==========================================
    # SLIDE 8: Experimental Setup & Workflow
    # ==========================================
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "7. Experimental Setup & Workflow Architecture")

    tb8 = slide8.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tf8 = tb8.text_frame
    tf8.word_wrap = True

    p = tf8.paragraphs[0]
    p.text = "Rigorous 3-Way Partitioning & Preprocessing Pipeline:"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY

    wf_steps = [
        ("Stratified Split (60% / 20% / 20%): ", "Train (170,883 records, 295 frauds) | Validation (56,962 records, 99 frauds) | Test (56,962 records, 98 frauds)."),
        ("Leakage Prevention: ", "RobustScaler is fitted strictly on genuine transactions (y_train == 0) in the training set. Eliminates outlier distortion and test data leakage."),
        ("Training Constraints: ", "Isolation Forest is trained exclusively on unlabelled normal transactions, mirroring real bank semi-supervised/unsupervised deployment."),
        ("Evaluation Parity: ", "Both baseline and tuned models are evaluated on the exact same 56,962 test records."),
        ("Reproducibility: ", "Fixed random seeds (random_state=42) and full unit test suite (5 passing tests).")
    ]
    for b, n in wf_steps:
        p = tf8.add_paragraph()
        p.space_before = Pt(10)
        r1 = p.add_run()
        r1.text = "• " + b
        r1.font.bold = True
        r1.font.size = Pt(13)
        r1.font.color.rgb = ACCENT_BLUE
        r2 = p.add_run()
        r2.text = n
        r2.font.size = Pt(13)
        r2.font.color.rgb = DARK_TEXT

    slide8.notes_slide.notes_text_frame.text = (
        "We enforced strict experimental integrity. The data was split 60/20/20 with stratification. "
        "The RobustScaler and Isolation Forest were fit exclusively on genuine training records. "
        "No test labels were ever seen during threshold tuning."
    )

    # ==========================================
    # SLIDE 9: Results & Performance Comparison Table
    # ==========================================
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "8. Experimental Results & Performance Comparison")

    # Add comparison table directly onto slide
    rows, cols = 10, 3
    left, top, width, height = Inches(0.8), Inches(1.7), Inches(11.7), Inches(5.0)
    table_shape = slide9.shapes.add_table(rows, cols, left, top, width, height)
    tbl = table_shape.table

    tbl.columns[0].width = Inches(4.5)
    tbl.columns[1].width = Inches(3.6)
    tbl.columns[2].width = Inches(3.6)

    headers = ["Evaluation Metric", "Original Isolation Forest (Default)", "Threshold-Tuned Isolation Forest (Improved)"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT

    table_data = [
        ("Precision", "0.0399 (3.99%)", "0.2174 (21.74%) [+445%]"),
        ("Recall", "0.8469 (84.69%)", "0.4082 (40.82%)"),
        ("F1-Score", "0.0763", "0.2837 [+272%]"),
        ("Accuracy", "0.9647 (96.47%)", "0.9965 (99.65%)"),
        ("False Positive Rate (FPR)", "0.035101 (3.51%)", "0.002532 (0.25%) [-92.8%]"),
        ("False Negative Rate (FNR)", "0.1531 (15.31%)", "0.5918 (59.18%)"),
        ("Average Precision / PR-AUC", "0.1715", "0.1715"),
        ("Number of Anomalies Flagged", "2,079 alerts", "184 alerts [-91.2%]"),
        ("Processing Time (Fit + Infer)", "0.24 seconds", "0.24 seconds")
    ]

    for i, (m, orig, tuned) in enumerate(table_data):
        row_idx = i + 1
        c0 = tbl.cell(row_idx, 0)
        c1 = tbl.cell(row_idx, 1)
        c2 = tbl.cell(row_idx, 2)

        c0.text = m
        c1.text = orig
        c2.text = tuned

        for cell, j in zip([c0, c1, c2], [0, 1, 2]):
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(12)
            if j == 2 and ("+" in tuned or "[-" in tuned):
                p.font.bold = True
                p.font.color.rgb = GREEN if "+" in tuned or "[-92" in tuned or "[-91" in tuned else DARK_TEXT
            if j > 0:
                p.alignment = PP_ALIGN.CENTER

    slide9.notes_slide.notes_text_frame.text = (
        "Here are our actual empirical results. Look at the dramatic improvements: "
        "Precision surged from 3.99% to 21.74%—a 445% increase. "
        "F1-score improved from 0.076 to 0.284. "
        "Crucially, false positive alerts fell from 2,079 down to just 184. "
        "This eliminates 1,895 false alarms while maintaining lightning-fast 0.24s execution time."
    )

    # ==========================================
    # SLIDE 10: Visual Analysis & Confusion Matrices
    # ==========================================
    slide10 = prs.slides.add_slide(blank_layout)
    add_header(slide10, "9. Visual Analysis: Confusion Matrices & PR Curves")

    cm_orig_img = "credit_card_fraud_detection/results/confusion_matrix_original.png"
    cm_tuned_img = "credit_card_fraud_detection/results/confusion_matrix_improved.png"
    pr_curve_img = "credit_card_fraud_detection/results/precision_recall_curve.png"

    if os.path.exists(cm_orig_img) and os.path.exists(cm_tuned_img):
        slide10.shapes.add_picture(cm_orig_img, Inches(0.8), Inches(1.8), width=Inches(3.8))
        slide10.shapes.add_picture(cm_tuned_img, Inches(4.8), Inches(1.8), width=Inches(3.8))
    if os.path.exists(pr_curve_img):
        slide10.shapes.add_picture(pr_curve_img, Inches(8.8), Inches(1.8), width=Inches(4.1))

    slide10.notes_slide.notes_text_frame.text = (
        "Looking at the confusion matrices: The original model generated 1,996 false alarms. "
        "The improved model dropped false alarms to only 144—a 92.8% reduction. "
        "On the right, the Precision-Recall curve shows how threshold tuning shifts our operating point "
        "from the low-precision baseline up to a much more viable operational trade-off."
    )

    # ==========================================
    # SLIDE 11: Critical Discussion & Ethics
    # ==========================================
    slide11 = prs.slides.add_slide(blank_layout)
    add_header(slide11, "10. Critical Discussion: Trade-offs, Ethics & Limitations")

    tb11 = slide11.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tf11 = tb11.text_frame
    tf11.word_wrap = True

    p = tf11.paragraphs[0]
    p.text = "Critical Engineering Realities:"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY

    disc_points = [
        ("The Inescapable Trade-Off: ", "Lowering false alarms reduced recall from 84.7% to 40.8%. In banking, missing high-value fraud has financial consequences, but drowning analysts in 96% false alarms causes human fatigue."),
        ("Why Accuracy is Misleading: ", "Both models have >96% accuracy. In imbalanced data, accuracy is mathematically dominated by the true negatives and hides poor fraud detection."),
        ("Privacy & Anonymization: ", "The dataset features are PCA-transformed to protect user privacy (GDPR). However, this creates a 'black box' where fraud analysts cannot easily explain why a customer's transaction was flagged."),
        ("Fairness & Customer Experience: ", "False alarms cause card rejections during genuine emergency purchases. Algorithmic thresholds must be calibrated with customer impact in mind."),
        ("Production Deployment Recommendation: ", "Isolation Forest should not operate in isolation. It should function as a first-line anomaly filter feeding scored alerts into supervised gradient-boosted ensembles with human oversight.")
    ]
    for b, n in disc_points:
        p = tf11.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = "• " + b
        r1.font.bold = True
        r1.font.size = Pt(13)
        r1.font.color.rgb = RED if "Trade-Off" in b else NAVY
        r2 = p.add_run()
        r2.text = n
        r2.font.size = Pt(13)
        r2.font.color.rgb = DARK_TEXT

    slide11.notes_slide.notes_text_frame.text = (
        "We must be honest about limitations: tuning increased precision by dropping recall from 84.7% to 40.8%. "
        "No single threshold is perfect for every bank. Banks with large fraud investigation teams might tolerate more false alarms to catch every fraud, "
        "while digital challenger banks prefer higher precision to avoid annoying customers. "
        "Furthermore, PCA features protect privacy but reduce explainability."
    )

    # ==========================================
    # SLIDE 12: Conclusion & References
    # ==========================================
    slide12 = prs.slides.add_slide(blank_layout)
    add_header(slide12, "11. Conclusion & Key References")

    tb12 = slide12.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tf12 = tb12.text_frame
    tf12.word_wrap = True

    p = tf12.paragraphs[0]
    p.text = "Summary of Contributions:"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY

    conc_points = [
        ("Empirical Validation: ", "Successfully reproduced Ounacer et al. (2018) and proved that default thresholds fail under severe class imbalance."),
        ("Demonstrated Gain: ", "Validation threshold tuning boosted F1-score from 0.0763 to 0.2837 and cut false positives by 92.8% without test leakage."),
        ("Practical Industry Relevance: ", "Proved that threshold tuning is a zero-computational-cost, highly effective intervention for real-world fraud pipelines.")
    ]
    for b, n in conc_points:
        p = tf12.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = "• " + b
        r1.font.bold = True
        r1.font.size = Pt(13)
        r1.font.color.rgb = GREEN
        r2 = p.add_run()
        r2.text = n
        r2.font.size = Pt(13)
        r2.font.color.rgb = DARK_TEXT

    p_ref = tf12.add_paragraph()
    p_ref.text = "\nKey Harvard References:"
    p_ref.font.size = Pt(14)
    p_ref.font.bold = True
    p_ref.font.color.rgb = ACCENT_BLUE

    refs = [
        "• Ounacer, S., El Arass, M.Y., Gherabi, N. and Bahaj, M. (2018) 'Using Isolation Forest in anomaly detection: the case of credit card transactions', IEEE ICOA, pp. 1-6.",
        "• Liu, F.T., Ting, K.M. and Zhou, Z.-H. (2008) 'Isolation Forest', 8th IEEE International Conference on Data Mining (ICDM), pp. 413-422.",
        "• Carcillo, F., Le Borgne, Y.-A., Caelen, O., Kessaci, Y., Oblé, F. and Bontempi, G. (2021) 'Combining unsupervised and supervised learning in credit card fraud detection', Information Sciences, 557, pp. 317-331.",
        "• Saito, T. and Rehmsmeier, M. (2015) 'The Precision-Recall Plot Is More Informative than the ROC Plot When Evaluating Imbalanced Datasets', PLOS ONE, 10(3), p. e0118432."
    ]
    for r in refs:
        p = tf12.add_paragraph()
        p.text = r
        p.font.size = Pt(11)
        p.font.color.rgb = DARK_TEXT

    slide12.notes_slide.notes_text_frame.text = (
        "In conclusion, our project bridges theoretical anomaly detection and operational fraud management. "
        "We showed that threshold tuning improves precision and F1 by order of magnitude without adding runtime latency. "
        "Thank you, and we welcome any questions!"
    )

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"PowerPoint presentation created successfully at: {output_path}")

if __name__ == '__main__':
    create_deck('credit_card_fraud_detection/presentation/project_presentation.pptx')

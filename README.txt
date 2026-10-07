RELATIVE GRADE ANALYTICS - FINAL JUSTIFIED ANALYSIS HTML
========================================================
Developed by Dr. Vikas Panthi, PC CSE Core, SCOPE

OPEN
- Open index.html in Microsoft Edge or Google Chrome.
- Excel parsing and chart rendering use CDN libraries, so internet access is required when first loading those libraries.

ASSESSMENT IDENTIFICATION
- MARK_MODE = CAT1 / CAT-I variants -> CAT-I
- MARK_MODE = CAT2 / CAT-II variants -> CAT-II
- Combined mode matches REG_NO + COURSE_CODE + CLASS_ID.

DATA CLEANING PRINCIPLES
- Trims text and maps recognized column aliases.
- Standardizes known STUDENT_STATUS and MARK_MODE variants.
- Removes only exact duplicate rows automatically.
- Duplicate Student+Course+Class keys are flagged, not silently deleted.
- Invalid/missing Present marks are excluded from score statistics but remain visible for audit.
- Original academic meaning is not changed silently.

STATISTICAL OUTPUTS
- N, mean, normalized mean %, median, mode
- population SD, sample SD, variance, CV
- P10, Q1, Q3, P90, IQR, min, max, range
- adjusted skewness, bias-adjusted excess kurtosis
- Tukey 1.5*IQR outlier count
- status counts and data completeness
- student percentile and normalized Z-score
- course/faculty/class summaries
- CAT-I/CAT-II paired improvement, decline, Pearson correlation
- configurable relative-grade what-if scenario on normalized percentages

IMPORTANT
- Faculty/course/class comparisons are descriptive only.
- Statistical grade levels are not official grades unless they match the institution's approved policy.
- Outliers are review points, not automatic errors.

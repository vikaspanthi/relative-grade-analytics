RELATIVE GRADE ANALYTICS - FINAL MARK_MODE VERSION
=================================================

HOW IT WORKS
1. Open index.html in Microsoft Edge or Google Chrome.
2. Select one or more same-format Excel/CSV result files.
3. Click Analyze.
4. The application reads MARK_MODE automatically:
   - CAT1 / CAT-I / CAT 1 -> CAT-I analysis
   - CAT2 / CAT-II / CAT 2 -> CAT-II analysis
5. Choose Analysis Mode:
   - CAT-I Only
   - CAT-II Only
   - Combined CAT-I + CAT-II
6. Combined mode matches the same REG_NO + COURSE_CODE + CLASS_ID.
7. Default combined weighting is 50% CAT-I + 50% CAT-II and can be changed.

SUPPORTED WORKFLOWS
- One Excel containing both CAT1 and CAT2 rows
- Separate CAT1 and CAT2 Excel files selected together
- CAT1-only Excel
- CAT2-only Excel

ANALYSIS
- Dashboard statistics and charts
- Course-wise analysis
- Faculty-wise analysis using ERP_ID + FACULTY_NAME
- Class-wise analysis
- Student-level results
- Present / Absent / Debarred separation
- CAT-I vs CAT-II improvement/decline in Combined mode
- Relative-grade statistical scenario
- Data validation
- CSV export
- PDF via Print / Save as PDF

IMPORTANT
Relative-grade levels are statistical what-if outputs and do not overwrite official grades.
Excel parsing uses SheetJS from a CDN, so internet access is required when loading XLSX files.

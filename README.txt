Relative Grade Analytics - Modular HTML Edition

Files:
  index.html  - application structure
  styles.css  - layout, responsive design, show/hide and print rules
  app.js      - analysis logic, visualizations, visibility controls and exports

Open index.html in Microsoft Edge or Google Chrome.
Excel parsing, Chart.js and PDF libraries are loaded from CDN, so internet access is needed for those libraries.

Visibility controls:
- Hide/Show dashboard summary and charts
- Hide/Show result tables and analysis charts
- Hide/Show student table
- Hide/Show grade table and grade chart
- Hide/Show validation and data-description panels
- Global Hide Charts / Show Charts button

Developed by Dr. Vikas Panthi, PC CSE Core, SCOPE


EFFECTIVE VISUAL ANALYTICS
--------------------------
The application now includes a dedicated Visual Analytics page with:
- Percentile / quartile profile (P10, Q1, Median, Q3, P90)
- Performance-band distribution on normalized percentages
- Validated-data quality doughnut chart
- Statistical diagnostic chart
- Top Course/Faculty/Class mean vs median comparison
- Q1-Q3 floating range chart for group spread
- Existing marks histogram, status chart, grade chart, CAT-I/CAT-II change and scatter plots

Only valid Present marks are used in score-statistic graphs. Absent/Debarred and invalid/missing marks remain visible in validation/status graphs.

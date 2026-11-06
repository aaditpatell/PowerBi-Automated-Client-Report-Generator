CLIENT REPORT GENERATOR

Creates one Power BI report file per client from the quarterly billing workbooks.
Each client's report is built only from that client's own workbooks, so it contains
only their data.

HOW TO USE

1. Copy the client workbooks (Users_Report_*.xlsx) into the Input folder.
   Include every quarter each client needs:
     Q1 report -> Q1 workbook
     Q2 report -> Q1 and Q2 workbooks
     Q3 report -> Q1, Q2 and Q3 workbooks
     Q4 report -> Q1, Q2 and Q3 workbooks (Q3 is used as Q4; no Q4 workbook needed)

2. Double-click "Run Generator.bat".
   A window lists every client found and whether its template was created.
   The templates are saved in the Output folder, e.g. 2026-2027-LTC-Q3.pbit,
   with a generation_log.csv showing which workbooks each one uses.

3. For each .pbit in the Output folder:
     a. Double-click it. Power BI Desktop opens it and loads that client's data by
        itself - there are no file paths or parameters to enter.
     b. Check the report, then File > Save as > .pbix (same name).
   The saved .pbix files are the finished reports to send to clients.

   Note: the files in Output have the paths built in, so leave the workbooks in the
   Input folder until you have saved the .pbix files.

WHAT IT DOES AUTOMATICALLY

- Reads the reporting group, fiscal year and quarter from each file name.
  Names must look like Users_Report_<group>_<YYYY-YYYY>_<Qn>.xlsx, e.g.
  Users_Report_LTC_2026-2027_Q3.xlsx. A name without a group is the "Overall" report.
- Makes each client's report for the latest quarter it finds for that client.
- If a quarter has more than one version (e.g. "Q2" and "Q2 v2"), uses the most
  recently modified file and notes it in the log.
- Skips a client (and says why) if a needed quarter is missing, e.g. a Q3 report
  without the Q1 workbook.
- Ignores files that do not follow the naming pattern and Excel's temporary ~$ files.
- Names each output <FiscalYear>-<Client>-<Quarter>, e.g. 2026-2027-JTS-Q2.pbit. To use a
  different order, change the NameFormat line near the top of Generate-ClientReports.ps1,
  e.g. '{FiscalYear}-{Quarter}-{Client}'.

OPTIONAL: RUN FOR A SPECIFIC QUARTER OR FISCAL YEAR

Open PowerShell in this folder and run, for example:
  powershell -ExecutionPolicy Bypass -File .\Generate-ClientReports.ps1 -InputFolder .\Input -TemplatePath .\Automated_Client_User_Billing_Report_Template.pbit -OutputFolder .\Output -Quarter Q4 -FiscalYear 2026-2027

IF IT WILL NOT RUN

- If Windows says the file came from another computer, right-click
  Generate-ClientReports.ps1 > Properties > tick Unblock > OK.
- Keep the template file in this folder with the same name, or update the .bat.

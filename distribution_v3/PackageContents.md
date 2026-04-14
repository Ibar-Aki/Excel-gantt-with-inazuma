InazumaGantt v3 distribution package

Generated at: 2026-04-15 01:13:27 +09:00
Package root: .
Converter root: set at runtime

[Workbook]
excel\InazumaGantt_v3_20260415_0113.xlsm

[One-click launchers]
scripts\Run_OneClick_CreateLatestWorkbook.bat
scripts\OneClick_CreateLatestWorkbook.ps1
scripts\Run_RestoreWorkbookFromPayload.bat
scripts\RestoreWorkbookFromPayload.ps1
scripts\Run_CreateDistributionPackage.bat
scripts\CreateDistributionPackage.ps1

[Build helpers]
scripts\BuildInazumaGantt_UTF8.ps1
scripts\FixEncoding.ps1

[VBA modules]
vba\InazumaGantt_v3_UTF8.bas
vba\InazumaGantt_v3_SJIS.bas
vba\WBSParentRollup_UTF8.bas
vba\WBSParentRollup_SJIS.bas
vba\WBSRoadmapReport_UTF8.bas
vba\WBSRoadmapReport_SJIS.bas
vba\WBSSampleShowcase_UTF8.bas
vba\WBSSampleShowcase_SJIS.bas
vba\HierarchyColor_UTF8.bas
vba\HierarchyColor_SJIS.bas
vba\SetupWizard_UTF8.bas
vba\SetupWizard_SJIS.bas
vba\SheetModule_UTF8.bas
vba\SheetModule_SJIS.bas

[Workbook payload]
excel\WorkbookPayload.json

[Documents]
docs\RestoreGuide.md

[Usage]
1. Double-click scripts\Run_OneClick_CreateLatestWorkbook.bat to generate the latest workbook.
2. Open excel\InazumaGantt_v3_20260415_0113.xlsm to review the generated sample workbook.
3. If you restore from the converter bundle, run scripts\Run_RestoreWorkbookFromPayload.bat to recreate the xlsm from WorkbookPayload.json.
4. Use scripts\Run_CreateDistributionPackage.bat to rebuild this distribution folder.

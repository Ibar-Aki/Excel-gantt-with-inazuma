InazumaGantt Lite distribution package

Generated at: 2026-04-15 01:13:37 +09:00
Package root: .
Converter root: set at runtime

[Workbook]
excel\InazumaGantt_Lite_20260415_0113.xlsm

[One-click launchers]
scripts\Run_OneClick_CreateLiteWorkbook.bat
scripts\OneClick_CreateLiteWorkbook.ps1
scripts\Run_RestoreWorkbookFromPayload.bat
scripts\RestoreWorkbookFromPayload.ps1
scripts\Run_CreateDistributionPackage_Lite.bat
scripts\CreateDistributionPackage_Lite.ps1

[Build helpers]
scripts\BuildInazumaGantt_Lite_UTF8.ps1
scripts\FixEncoding.ps1

[VBA modules]
vba\InazumaGantt_Lite_UTF8.bas
vba\InazumaGantt_Lite_SJIS.bas
vba\WBSParentRollup_Lite_UTF8.bas
vba\WBSParentRollup_Lite_SJIS.bas
vba\WBSRoadmapReport_Lite_UTF8.bas
vba\WBSRoadmapReport_Lite_SJIS.bas
vba\WBSSampleShowcase_Lite_UTF8.bas
vba\WBSSampleShowcase_Lite_SJIS.bas
vba\HierarchyColor_Lite_UTF8.bas
vba\HierarchyColor_Lite_SJIS.bas
vba\SetupWizard_Lite_UTF8.bas
vba\SetupWizard_Lite_SJIS.bas
vba\SheetModule_Lite_UTF8.bas
vba\SheetModule_Lite_SJIS.bas

[Workbook payload]
excel\WorkbookPayload.json

[Documents]
docs\RestoreGuide_Lite.md

[Usage]
1. Double-click scripts\Run_OneClick_CreateLiteWorkbook.bat to generate the latest workbook.
2. Open excel\InazumaGantt_Lite_20260415_0113.xlsm to review the generated sample workbook.
3. If you restore from the converter bundle, run scripts\Run_RestoreWorkbookFromPayload.bat to recreate the xlsm from WorkbookPayload.json.
4. Use scripts\Run_CreateDistributionPackage_Lite.bat to rebuild this distribution folder.

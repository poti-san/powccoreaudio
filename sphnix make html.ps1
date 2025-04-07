.\_packagenames
foreach ($dirname in $packagenames) {
    Invoke-Expression "$dirname\docs\make.bat html"
}
pause

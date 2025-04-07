.\_packagenames
foreach ($dirname in $packagenames) {
	sphinx-apidoc -f -o $dirname\docs\source\apiref $dirname\src --no-toc --remove-old | Write-Output
}

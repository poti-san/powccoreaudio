.\_packagenames
foreach ($dirname in $packagenames) {
    pip uninstall $dirname
}

.\_packagenames
foreach ($dirname in $packagenames) {
    pip install -e ./$dirname
}

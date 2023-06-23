#!/bin/sh

set -e
set -x

stat "$1" || (echo "File $1 not found" && exit 1)
rm tmp-unprotecting -rf
mkdir tmp-unprotecting
cp "$1" tmp-unprotecting/in.xlsx

cd tmp-unprotecting
7z x in.xlsx xl/worksheets
find xl/worksheets -type f -name "*.xml" -exec sed -i 's/sheetProtection/sheepProtection/g' {} +
# I do this entire mkdir / cd / cp dance just because idk how to do `7z x -otmp-unprotecting xl/worksheets` and then `7z u xl/worksheets` (its in different dir)
7z u in.xlsx xl/worksheets
cd ..

cp "$1" "$1.bak"
cp tmp-unprotecting/in.xlsx "$1"
rm tmp-unprotecting -rf

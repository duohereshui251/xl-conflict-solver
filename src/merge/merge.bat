
echo "# Merge driver called #"


cp "%2" temp_a.xlsx
cp "%3" temp_b.xlsx
"C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe" temp_a.xlsx
"C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe" temp_b.xlsx

exit 0
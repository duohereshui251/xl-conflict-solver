
echo "# Merge driver called #"


cp "%2" temp_a.xlsx
cp "%3" temp_b.xlsx
rm "%1"
python.exe src/merge/merge.py temp_a.xlsx temp_b.xlsx
rm test/a.xlsx
cp temp_a.xlsx test/a.xlsx
rm temp_a.xlsx temp_b.xlsx
exit 0
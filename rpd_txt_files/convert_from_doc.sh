FILES=*.doc
for f in $FILES
do
  # extension="${f##*.}"
  filename="${f%.*}"
  echo "Converting $f"
  #`pandoc $f -t docx -o $filename.doc`
  soffice --convert-to txt $filename.doc
  # uncomment this line to delete the source file.
  mv $f ../bkp/
  mv $filename.txt ../rpd_txt_files/
done

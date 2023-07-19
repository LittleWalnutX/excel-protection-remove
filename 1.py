#!/bin/python3
import sys
import zipfile
import os
import shutil

extract_dir = "exrm-tmp"

in_file_name = sys.argv[1]

zf = zipfile.ZipFile(in_file_name)
zf.extractall(extract_dir)

with open(extract_dir + "/xl/workbook.xml", "r") as f:
    xml = f.read()

wp_start = xml.find(r"<workbookProtection")
wp_stop = xml.find(r"/>", wp_start + 1) + 2

out_xml = xml[:wp_start] + xml[wp_stop:]

#print(out_xml)

os.remove(extract_dir + "/xl/workbook.xml")



with open(extract_dir + "/xl/workbook.xml", "w") as f:
    f.write(out_xml)

shutil.make_archive(in_file_name + "r", "zip", extract_dir)
shutil.move(in_file_name + "r.zip", "rmed_pct.xlsx")

shutil.rmtree(path=extract_dir)




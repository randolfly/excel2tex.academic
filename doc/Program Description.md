---
tags: []
synthesis: 
date created: 2024-01-21
date modified: 2024-01-21
---

## 1 TLDR😄

The main idea of this program is to create "three-lines table" in latex format from excel file. 

> Sample Excel file

![Sample Excel File](assets/Program%20Description/image-20240121163717862.png)

> Result after execution

## 2 How to use 🧾

### 2.1 Console Application

1. `excel2tex -s "src.xlsx" -o "target.tex"`
	1. convert source `"src.xlsx"` to `"target.tex"`
2. `excel2tex -d "target_directory"`
	1. convert all excel files in directory `"target_directory"` to tex format with the same name and location
	2. `excel2tex -d "target_directory" --recursive`
		1. convert all excel files in directory and subdirectories `"target_directory"` to tex format with the same name and location

## 3 How to Develop 🤝

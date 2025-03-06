# auto_ref_word
A Python program that automatically marks references in a Word document.

## worning
Due to the complexity of Word documents, **please save a backup copy before running the script!!!!!!!!!**

## Installation
 pip install pywin32 python-docx  #(only in windows)

## Usage
```shell
python add.py
# change the fmt if you want
python use.py
```

## Description

格式要求，word中的参考文献用【】抱起来，里面是参考文献的内容，保证统一的参考文献文本一致。
add.py会把他们标记成编号，运行后，如果你需要更改格式，可以直接用word更改
use.py会把【】中的内容替换成交叉引用，并且会创造一个超链接

### Formatting Requirements  
- References in the Word document should be enclosed in 【】, with the reference content inside.  
- Ensure that identical reference texts remain consistent throughout the document.  

### Script Functions  
- `add.py` will replace the references with numbered citations. After running the script, you can manually adjust the format in Word if needed.  
- `use.py` will replace the content inside 【】 with cross-references and create a hyperlink.


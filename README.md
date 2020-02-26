这是一个用来从 mdx 字典中抓取所需的单词，并生成 html 或 pdf 文件的小工具。

## 用法
    usage: MdxConverter.py [-h] [--type [{pdf,html}]]
                       mdx_name input_name [output_name]

    positional arguments:
      mdx_name
      input_name
      output_name

    optional arguments:
      -h, --help           show this help message and exit
      --type [{pdf,html}]

例如：
    
    MdxConverter 某某词典.mdx input.xlsx output.pdf

## 依赖库
[mdict-query](https://github.com/mmjang/mdict-query)

[BeautifulSoup4](https://pypi.org/project/beautifulsoup4/)

[openpyxl](https://pypi.org/project/openpyxl/)

[pdfkit](https://github.com/JazzCore/python-pdfkit)


## 输入
### json 示例
```javascript
 [
 {
     "name": "Lesson 1",
     "words": [
         "hello",
         "world"
     ]
 },
 {
     "name": "Lesson 2",
     "words": [
         "only",
         "a",
         "test"
     ]
 }
 ]
```

### excel 示例
![](images/excel.jpg)

## 输出
### HTML
![](images/html.jpg)

### PDF
![](images/pdf.jpg)
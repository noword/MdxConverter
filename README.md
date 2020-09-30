这是一个用来从 mdx 字典中抓取所需的单词，并生成 html 或 pdf 文件的小工具。

## 用法
This Software changes the Text written in Python into Text in Excel as shown in the Example
    usage: MdxConverter [-h] [--type [{pdf,html}]] [--invalid {0,1,2}]
                            mdx_name input_name [output_name]

    positional arguments:
      mdx_name
      input_name
      output_name

    optional arguments:
      -h, --help           show this help message and exit
      --type [{pdf,html}]
      --invalid {0,1,2}    action for meeting invalid words
                           0: exit immediately
                           1: output warnning message to pdf/html
                           2: collect them to invalid_words.txt (default)

例如：
    
    MdxConverter 某某词典.mdx input.xlsx output.pdf

## 依赖库
[mdict-query](https://github.com/mmjang/mdict-query)

[BeautifulSoup4](https://pypi.org/project/beautifulsoup4/)

[openpyxl](https://pypi.org/project/openpyxl/)

[pdfkit](https://github.com/JazzCore/python-pdfkit)

[lxml](https://lxml.de/)

## 输入
### txt 示例
    #Lesson 1
    hello
    world

    #Lesson 2
    python
    is
    awesome


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
         "python",
         "is",
         "awesome"
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

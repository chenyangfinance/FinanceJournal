#coding:utf-8
import urllib.request
import re
import deepl
# from docx import Document
# from docx.shared import Pt
# from docx.enum.text import WD_ALIGN_PARAGRAPH

#初始化deepL
translator = deepl.Translator("e3c66233-3860-6d4e-22dc-eabe4408a3ca:fx")

url = 'https://www.sciencedirect.com/journal/journal-of-financial-economics/vol/143/issue/1'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.119 Safari/537.36',
}
request = urllib.request.Request(url=url, headers=headers)
content = urllib.request.urlopen(request).read().decode('utf8')
#print(content)

# 0. 文档信息
#<title data-react-helmet="true">Journal of Corporate Finance | Vol 70, October 2021 | ScienceDirect.com by Elsevier</title>
pattern0 = re.compile(r'<title data-react-helmet="true">(.*?)ScienceDirect.com')
date = pattern0.findall(content)[0].split("| ")[1]
issue = date.split(", ")[0] + " " + date.split(", ")[1]


# 1. title list
#<span class="js-article-title">Integrating corporate social responsibility criteria into executive compensation and firm innovation: International evidence</span></span></a></h3></dt><dd
pattern1 = re.compile(r'<span class="js-article-title">(.*?)</span></span>')
titleList = pattern1.findall(content)[1:] #删除一个Editorial Board
print(len(titleList))
#print(titleList)

# 2. author list
#<div class="text-s u-clr-grey8 js-article__item__authors">Albert Tsang, Kun Tracy Wang, Simeng Liu, Li Yu</div></dd>
pattern2 = re.compile(r'<div class="text-s u-clr-grey8 js-article__item__authors">(.*?)</div>')
authorList = pattern2.findall(content)[0:]
print(len(authorList))
#print(authorList)

# 3. id list 来生成每篇文章的URL,之后访问URL读取Abstract
# for="checkbox-S0165410121000434"><input type="checkbox
pattern3 = re.compile(r'for="checkbox-(.*?)"><input type="checkbox')
idList = pattern3.findall(content)[1:] #删除第一个是Editorial Board,
print(idList)

# 4. 循环idList 获取摘要
# Abstract</h2><div id="as0005"><p id="sp0115">I use the staggered adoption of state-level antitakeover laws to provide causal evidence that managerial agency problems reduce the allocative efficiency of conglomerate firms. I find that increases in control slack following the passage of antitakeover laws reduces <em>q</em>-sensitivity of investment by 64%. The adverse impact of the laws appears mostly at conglomerate firms that benefited from disciplinary takeover threats prior to the passage of the laws, lacked alternative sources of pressure on management, or had the structural makings to fuel wasteful influence activities and power struggles among managers. These findings suggest that takeover threats impact the efficiency of resource allocation.</p></div></div></div>
absList = []
abs_url = r"https://www.sciencedirect.com/science/article/pii/"
pattern4 = re.compile(r'A\D\D\D\D\D\D\D</h2><div id="\D+\d+"><p id="\D+\d+">(.*?)</p></div></div></div>')

for i in range(0, len(idList)):  # len(idList)
    id = abs_url + idList[i]
    # https://www.sciencedirect.com/science/article/pii/S0929119921001929
    abs_content = urllib.request.Request(url=id, headers=headers)
    abs = urllib.request.urlopen(abs_content).read().decode('utf8')
    # print(abs)
    a = pattern4.findall(abs)[0]  # 和JAE的不同，这里不切分字符串，而是用正则提取
    a = a.replace("&#x27;", "'")  # 单引号乱码
    a = re.sub(r'<.*?>', '', a)  # html乱码
    absList.append(a)
    print("正在获取摘要(%d)..." % (i+1))
#print(absList)

# 5. 写入txt
filename = "JFE_%s.txt" % issue
text = open(filename,"a", encoding="utf-8" )
text.write("刊名: Journal of Financial Economics")
text.write("\n")
text.write("刊号: " + date)
text.write("\n")
text.write("仅翻译用于学术交流，版权归期刊和作者所有")
text.write("\n")
text.write("原网页: " + url)
text.write("\n")
text.write("\n")
text.write("\n")

for i in range(0, len(idList)):
    print(i+1)
    text.write(str(i+1)+". "+titleList[i])  # 使用add_run添加文字
    text.write("\n")


    cn_title = translator.translate_text(titleList[i], target_lang="ZH")
    text.write(str(cn_title))
    text.write("\n")

    text.write("作者: "+authorList[i])
    text.write("\n")
    text.write("摘要: "+absList[i])
    text.write("\n")

    cn_abs = translator.translate_text(absList[i], target_lang="ZH")
    text.write(str(cn_abs))
    text.write("\n")

    text.write("原文链接: ")
    text.write( "\n"+ abs_url + idList[i])
    text.write("\n\n")

text.write("(关于本次推送：我编写代码从期刊官网摘取论文的题目，作者，和摘要，并调用 DeepL 的接口翻译题目和摘要。我对翻译内容进行审阅,补充和修正。代码公布在我的Github页面：https://github.com/chenyangfinance/FinanceJournal)")

text.close()


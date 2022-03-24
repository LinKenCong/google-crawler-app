# google-crawler-app

requests google results crawler app

![avatar](img\appimg.png)

The program to crawl Google search results written when learning Python, and packaged as exe software, realizes multi-threading and export table functions.

It may be unusable due to changes in Google's algorithm, and the code needs to be modified.

学习Python时写的爬取谷歌搜索结果程序，并打包为了exe软件，实现了多线程和导出表格功能。

可能因为Google的算法变更导致无法使用，需修改代码。

# 编译EXE步骤 Compile EXE step

```shell
pipenv install
pipenv shell
pipenv uninstall --all
pipenv install requests bs4 xlwings
pipenv install pyinstaller --dev
pyinstaller -F app.py -w --upx-exclude=vcruntime140.dll
```
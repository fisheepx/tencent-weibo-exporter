tencent-weibo-exporter
======================
腾讯微博导出工具，无需登录，执行时输入自己的微博ID即可将微博导出到Word文件，格式为docx.

    之前一直写Java,闲暇时间学习了两周Python,感觉Python确实是“为了让写程序变得更简单”的一门语言，
    好多之前用Java实现起来复杂的操作在Python中用几行代码就能够搞定的感觉还是很不错的。
    再加上原来就想把自己之前腾讯微博的内容备份出来，毕竟之前有一段写得还挺认真的，就算是当做回忆吧。
    但是在网上找了几个工具，发现都不能用了，而且还都需要登录，于是就自己试着写了一个。
    想备份腾讯微博的原因是感觉这货以目前的状况应该是马上就会关闭的节奏，
    只是不知道到时候会不会良心的出一个微博导出工具之类的了。
    根据网上的备份教程文章，原来应该是官方有一个备份工具，不过目前已经处于找不到的状态。
    代码逻辑不复杂，本人水平有限，而且只用了两个微博账号做了测试。
    示例中使用的人民网的账号是上传代码时候临时抓的。
    本来想找个明星账号的，一来是腾讯微博明星本来就少，二来实在是没有太喜欢的明星。

Requirements
------------
Python 2.7

Installation
------------
由于需要导出Word文件，所以需要Python的docx支持，安装时请使用 [pip](http://www.pip-installer.org/)

    $ pip install docx
    
或者 [easy_install](http://peak.telecommunity.com/DevCenter/EasyInstall)

    $ easy_install docx
    
Running
-------
本程序每次提交时，按版本号提交整个新文件夹，执行时下载最新版本号文件夹即可。

直接运行主文件tencent-weibo.py即可

    $ python tencent-weibo.py


运行后会在当前目录下生成Word文件，在当前目录下的 pic 文件夹下下载微博内的图片。目前每20页生成一个Word文件，可以代码内自行修改。

*※如果需要在Windows下运行，需要安装 [Python for Windows](https://www.python.org/downloads/) 并且正确配置环境变量。*

Code
----
1,如何修改为自己的微博ID？

version14 之前 在文件 tencent_weibo.py 最下方位置将 renminwangcom 替换为自己的微博ID即可。

2，如何指定备份的开始和结束页？

version8 之前在 *start()* 方法的 *while* 循环内通过注释 *test code* 部分实现。

version8 开始通过修改在类最开始的 *START_PAGE_INDEX* 和 *END_PAGE_INDEX* 两个常量实现。

3，如果修改多少页保存成一个文件？

从 version14 开始，通修改在类最开始的 *SAVE_FILE_PAGE* 常量实现。

*※此处所说的“页”为腾讯微博的页数，并非Word文件的页数。*

4,为什么最多只能备份100页的微博？

目前腾讯微博在不登录状态下只能查看前100页，不清楚是否100页之前的内容无法查看的原因，目前正在与客服沟通。

单独留下每个版本的代码文件夹，是为了记录一下每个版本的改进，也为了如果有人想通过本程序学习Python更方便。除各别版本生成文件会出错外，其它版均可正常生成文件。（目前已知version13出现异常，version14修复。）
代码非常简单（其实是过于简单:joy:），只是一点点的找到正规表达式的过程。发现不好的地方欢迎指正 :two_men_holding_hands: 代码会不定期更新。

About
-----
如果取得你的腾讯微博ID？

以人民网账号为例：http://t.qq.com/renminwangcom

则人民网的账号ID为：renminwangcom

在运行提示输入时输入即可

PS
-------
任何问题或者反馈发送至邮箱：poemfar@gmail.com

Enjoy it!

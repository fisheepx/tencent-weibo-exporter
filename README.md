tencent-weibo-exporter
======================
腾讯微博导出工具，无需登录，执行时输入自己的微博ID即可将微博导出到Word文件，格式为docx.

    之前一直写Java,闲暇时间学习了两周Python,感觉Python确实是“为了让写程序变得更简单”的一门语言，
    好多之前用Java实现起来复杂的操作在Python中用几行代码就能够搞定的感觉还是很不错的。
    再加上原来就想把自己之前腾讯微博的内容备份出来，毕竟之前有一段写得还挺认真的，就算是当做回忆吧。
    但是在网上找了几个工具，发现都不能用了，而且还都需要登录，于是就自己试着写了一个。
    想备份腾讯微博的原因是感觉这货以目前的状况应该是马上就会关闭的节奏，
    只是不知道到时候会不会良心的出一个微博导出工具之类的了。
    不过根据网卡的文章原来应该是官方有一个备份工具了，目前已经处于找不到的状态。
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
直接运行主文件tencent-weibo.py即可

    $ python tencent-weibo.py


运行后会在当前目录下生成Word文件，在当前目录下的 pic 文件夹下下载微博内的图片。目前每20页生成一个Word文件，可以代码内自行修改。

*※如果需要在Windows下运行，需要安装 [Python for Windows](https://www.python.org/downloads/) 并且正确配置环境变量。*

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

tencent-weibo-exporter
======================
腾讯微博导出工具，无需登录，执行时输入自己的微博ID即可将微博导出到Word文件，格式为docx.

:triangular_flag_on_post:※2017年9月15日：更新[登录备份版本](#triangular_flag_on_postlogin-versiontriangular_flag_on_post)代码，以对应不登录时100页以后无法备份的情况。

![image](https://github.com/coolcooldool/tencent-weibo-exporter/blob/master/images/logo.jpg)

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

备份文件效果

![image](https://github.com/coolcooldool/tencent-weibo-exporter/blob/master/images/content1.png)

![image](https://github.com/coolcooldool/tencent-weibo-exporter/blob/master/images/content2.png)

位置信息，视频，URL链接均可以点击打开相应网页，使用谷歌地图。

Requirements
------------
Python 2.x(version14以前)

Python 3.x(version15开始)

Installation
------------
由于需要导出Word文件，所以需要Python的docx支持，安装时请使用 [pip](http://www.pip-installer.org/)

Version15开始请使用以下命令：

    $ pip install --pre python-docx

Version14以前请使用以下命令：

    $ pip install docx

或者 [easy_install](http://peak.telecommunity.com/DevCenter/EasyInstall)

    $ easy_install docx

Running
-------
本程序每次提交时，按版本号提交整个新文件夹，执行时下载最新版本号文件夹即可。

Version15及以后的版本直接运行主文件tencent-weibo.py即可

    $ python tencent_weibo.py

:exclamation::exclamation::exclamation:Version14之前请使用 [PyCharm](https://www.jetbrains.com/pycharm/), [Eclipse(安装Python插件)](https://www.eclipse.org/downloads/)等IDE执行*tencent-weibo.py*，如果直接使用Windows控制台执行则会出现中文乱码问题，参见[issues1](https://github.com/coolcooldool/tencent-weibo-exporter/issues/1)，想直接用控制台执行可使用version15及以后的版本。

*※目前已不推荐使用version15之前的版本*


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
----

目前腾讯微博在不登录状态下只能查看前100页，由于找不到腾讯客服于是又花了两天时间写了一个登录备份版，参见下方登录备份版使用说明。

单独留下每个版本的代码文件夹，是为了记录一下每个版本的改进，也为了如果有人想通过本程序学习Python更方便。除各别版本生成文件会出错外，其它版均可正常生成文件。（目前已知version13出现异常，version14修复。）
代码非常简单（其实是过于简单:joy:），只是一点点的找到正规表达式的过程。发现不好的地方欢迎指正 :two_men_holding_hands: 代码会不定期更新。

:triangular_flag_on_post:Login Version:triangular_flag_on_post:
-------------
#### 1，为什么还要编写一个登录版的代码？

腾讯微博默认不登录时只能查看前100页内容，导致不登录状态下100页以后内容无法备份。而且不登录时的每页显示长度比较少，本人微博登录后总页数66页，而不登录时第100页的内容才相当于登录后的52页左右。

※.此版本备份过程中分自动打开FireFox浏览器，并执行自动点击操作，并不是闹鬼:ghost:请淡定:joy:

#### 2，登录版本的代码运行时有什么要求？

##### (1)请安装FireFox浏览器

因为代码使用Chrome浏览器测试时发生了无法响应事件的问题，所以请使用FireFox浏览器。[下载地址](https://www.mozilla.org/)

##### (2)下载FireFox浏览器的调试驱动程序并将其配置到环境变量

由于代码基于*python selenium*实现，需要浏览器的调试驱动程序。请下载驱动程序文件并将其配置到环境变量中。[下载地址](https://github.com/mozilla/geckodriver/releases)
※.将下载得到的*geckodriver.exe*文件直接放入Python安装目录即可快捷使用。(例如Windows：C:\Program Files\Python27)

##### (3)QQ必须处于登录状态并且网页快捷登录可用

由于在网页上通过账号和密码登录QQ的实现方式十分困难，并且于安全性考虑，决定程序采用快捷登录的方式进行登录，所以QQ必须要处于在线状态，并且允许使用QQ快捷登录的方式登录网页。如下图所示状态即可。
![image](https://github.com/coolcooldool/tencent-weibo-exporter/blob/master/images/login.jpg)

##### (4)将微博翻页状态切换为“页码翻页”语言切换为“简体中文”

登录版本的代码备份时需要通过模拟点击下一页进行翻页，所以需要事先登录自己的微博账号，并将显示方式切换为页码翻页；程序中翻页时查找的文字为*下一页*，所以需要简体中文显示。如遇到页面不是页码翻页，而是在页面最下方显示为“更多”时，通过拖拽浏览器的滚动条到页面最下方的方法，即可以页面最下方显示“试试页码翻页”链接，鼠标滚动到底部会自动加载下一页，无法显示页码翻页链接。如下图处示。
![image](https://github.com/coolcooldool/tencent-weibo-exporter/blob/master/images/footer.png)

#### 3，我的QQ号和密码有没有泄露的风险？

完全没有！如上所述，虽然是登录版，但代码中并没有利用账号和密码的方式进行登录，所以完全没有风险。

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

Change Log
----------
##### login version

version5:

    对应含有引用的内容,在非login版本基础上增加了含有视频的内容
    修改点击下一页的逻辑

version4:

    使用非login版代码对应含有主题，好友，Emoji，链接等全部内容

version3:

    添加位置信息
    对应带有视频有内容

version2:

    添加添加保存图片功能(单张与多张)

version1:

    保存纯文本内容

##### common version

version15:

    修改运行环境为Python3.x
    以对应控制台执行时中文LOG乱码导致异常结束

version14:

    修改文件为分割存储
    加入转帖与视频封面插入图片异常处理
    修改备份文件名生成规则
    
version13:

    对应其它各种表情
    对应",<符号的正确显示
    
version12:

    对应带有转发的内容（仅作者，内容，图片，时间）
    对应 &，>符号的正确显示
    
version11:

    对应含有QQ表情的内容
    
version10:

    对应含有URL的内容
    调整图片为统一宽度(对应某些图片显示过大的问题)

version9:

    对应含有@好友的内容

version8:

    对应分享视频链接的微博
    添加备份时开始与结束页的控制

version7:

    对应含有Emoji的内容
    对应Python的全名规则
    
version6:

    对应含有话题的内容
    
version5:

    对应位置信息(谷歌地图)
    
version4:

    优化生成Word文件格式
    下载图片时文件夹不存在则创建

version3:

    添加位置信息
    对应带有视频有内容
    
version2:

    添加添加保存图片功能
    
version1:

    保存纯文本内容

### 主要用途

- 合同台账做到excel上，到期需要提醒？


# ExcelRemainder



### 前言
本程序借助于企业微信的API以及cron计划任务。也可以自建API或别人的API进行推送。

本着分享的精神，任何人可无偿使用，请勿对源码进行二次售卖。

### 主要功能
定时推送，可提前N天（N天内每次计划任务促发均会推送）请修改python中的参数。

可自定义文件名，标题列，日期列，设置提前推送日期的时间区间。

### 使用方法
1.用电脑打开企业微信官网，注册一个企业微信(免费)。

2.注册成功后，点「管理企业」进入管理界面，选择「应用管理」 → 「自建」 → 「创建应用」。


3.完成创建企业微信APP后，可以得到应用ID( agentid )，应用Secret( secret )。进入「我的企业」页面，拉到最下边，可以看到企业ID
。

4.配置参数，应用ID( agentid )，应用Secret( secret )，企业ID。用户ID填 @all ，推送给全员，填写某个人则推送给某个人。

5.安装cron添加计划任务，以及依赖pip3 install openpyxl requests

6.定时执行目录下的python Remainder.py即可。设置几点执行就是几点推送。

#### 提供了2个脚本，另一个使用自建的API或者别人搭建好的API(可参考xpnas大佬的inotify)，https://github.com/xpnas/inotify

### 打赏

> 如果您觉得对您有帮助，欢迎给我打赏。

<img src="wxpay.jpg" width="400" />
# Transcript-Combiner
由于中国券商喜爱使用有道云笔记分发纪要，同时各个来源的纪要格式各不相同，使得整理收纳纪要的过程变得非常繁琐。这一工作可以交给廉价的实习生，但如果你就是廉价的实习生呢？  
**That‘s the Right Project for You!**

### 已实现功能：
  1. 一键导出有道云笔记文档为Markdown文件，主要依托于``pull_notes.py``
  2. 自动合并多个纪要文档，生成聚合的Word格式纪要集，主要依托于``main.py``
    a. 每篇纪要的标题命名为：时间+主题+来源，并尽量按照时间排序
    b. 自动生成纪要标题大纲  
    c. 目前支持合并的格式有：.md, .docx, .pdf, .html

### 待添加的功能：
  1. 均一化所有时间格式

# 如何使用
### 配置环境
  1. 将文件下载到本地并解压缩
  2. 在文件夹中创建名为“TS”的文件夹
  3. 使用``pip install -r requirements.txt``安装依赖库
  4. 配置cookies.json文件：获取``Cookies``方式：  
  - 在浏览器如 Chrome 中使用账号密码或者其他方式登录有道云笔记  
  - 打开 DevTools (F12)，Network 下找「主」请求（一般是第一个），再找``Cookie``  
  - 复制对应数据填入cookies.json
  5. 配置config.json文件  
- ``local_dir``：选填，本地存放导出文件的文件夹，不填则默认为当前文件夹，建议设置为刚刚创建的TS文件夹  
- ``ydnote_dir``：选填，有道云笔记指定导出文件夹名，不填则导出所有文件  
- ``smms_secret_token``：选填， SM.MS 的 Secret Token（注册后 -> Dashboard -> API Token），用于上传笔记中有道云图床图片到 SM.MS 图床，不填则只下载到本地（youdaonote-images 文件夹），Markdown 中使用本地链接  
- ``is_relative_path``：选填，在 MD 文件中图片 / 附件是否采用相对路径展示，不填或 false 为绝对路径，true 为相对路径
  6. 配置company.yaml文件
  
### 运行脚本  
**运行有道云下载脚本**  
``python pull_notes.py``  
**运行合并纪要脚本**  
``python main.py``  
**或者，也可以在idle中运行程序，比如Pycharm**

# 感谢/参考  
[YoudaoNoteExport](https://github.com/DeppWang/youdaonote-pull)


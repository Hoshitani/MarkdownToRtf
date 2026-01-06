# Markdown2Rtf
在网上搜到了这个库，但是发现是.net 6.0的，无法用在我的项目上。非常感谢GustavoHennig。这是原项目链接：https://github.com/GustavoHennig/MarkdownToRtf 

于是进行了一些修改，保留MIT许可，将.net版本下降为standard2.0，降无可降，这样就能同时用在.net .net framework，以及.net core上了，保证通用性。

加入了一些功能，这些是原项目没有，但是我却需要的。

使用效果如下：（在Winform的RichTextBox中展示，字体使用的是更纱等距黑体，否则加粗会影响对齐效果，导致表格无法对齐）

<img width="800" height="599" alt="微信图片_2026-01-06_113959_677" src="https://github.com/user-attachments/assets/1fd36730-8a3b-4f39-9c6a-e2c2210cef52" />

增加了颜色控制，格式为：```#123456 内容```
 > 注意，颜色控制与Markdown中兼容的<font color=red>内容</font>不一致，因为感觉标签识别太麻烦，就简单用#FF0000处理了

增加了代码段展示，格式为\`\`\`内容\`\`\`

增加了表格显示（核心），格式为：

|内容|内容|...

|---|----|...

|内容|内容|...

......

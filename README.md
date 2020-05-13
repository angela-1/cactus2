# cactus2

属于办公工具集officeutils。WPS文档审阅工具，基于JS加载项编写。


## 部署

加载项开发完成后，要使用离线打包，再放在临时的一个服务器上下载安装

1. 使用 `serve` 启动一个静态网站，提供 `cactus2.7z` 下载
2. 修改 `C:\Users\angela\AppData\Roaming\kingsoft\wps\jsaddons\jsplugins.xml` 配置文件为离线插件
```xml
<jsplugins>
  <jsplugin name="cactus2" type="wps" url="http://localhost:5000/cactus2.7z"
    version="0.0.1" />
</jsplugins>
```
3. 启动 wps 会自己添加加载项，若服务器提供了新的版本，会自动下载安装


## 使用

介绍各功能使用方法




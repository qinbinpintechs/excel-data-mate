# excel-data-mate
excel数据匹配
springboot3.1.2  jdk17

启动项目后浏览器打开 http://localhost:8848/swagger-ui/index.html#/ 访问swagger

1、handel接口:匹配数据<br/>
&emsp;参数解释:filterStr 匹配部分构件，不填则匹配全部<br/>
&emsp;&emsp;&emsp;&emsp;&nbsp;file匹配附件 sheet1为构件数据，sheet2为建模数据 ~~参考《示例.xlsx》文件~~文件涉及数据，暂不提供

2、handTableMapB:筛选建模表重复数据（遇到建模数据有重复可以使用筛选重复glId数据，注意筛选的是第二sheet页）

3、handleTable:合并多sheet建模到同一个sheet下（早期使用）
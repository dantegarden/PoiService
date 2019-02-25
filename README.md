# Excel子服务
**项目负责人**：李倞 @lijing

**镜像**：hub.docker.dvt.io:5000/devops/poi:latest

## 以镜像方式启动
```bash
docker run -d --name poi -p 7007:8080 -v /home/devops/poi/logs:/home/tomcat/apache-tomcat-8.0.44/logs hub.docker.dvt.io:5000/devops/poi:latest
```
### 需要暴露的端口
程序访问端口 8080
### 挂载卷
日志目录 /home/tomcat/apache-tomcat-8.0.44/logs"# PoiService" 

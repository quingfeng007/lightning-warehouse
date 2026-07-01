# 闪电仓计算工具 - Web 部署指南

把现有桌面 GUI 工具(PyQt5)改成 Web 版,部署到 Render 免费云平台,任何浏览器都能访问。

## 1. 准备 GitHub(2 分钟)

### 1.1 注册 GitHub
- 打开 https://github.com 注册账号(已有跳过)
- 用来托管代码,Render 平台从 GitHub 拉代码部署

### 1.2 安装 Git
- 下载 https://git-scm.com/download/win
- 安装时全部下一步

## 2. 上传代码到 GitHub(2 分钟)

### 2.1 在 GitHub 创建仓库
1. 登录 GitHub
2. 右上角 `+` -> `New repository`
3. 名称填 `lightning-warehouse`(随你起)
4. 选 `Public`(免费账户必须 public 才能给 Render 拉)
5. **不要**勾选 `Add a README file`
6. 点 `Create repository`

### 2.2 推送代码
在 PowerShell 里执行:

```powershell
cd C:\Users\33324\.qclaw\workspace-agent-3657e7c8\web
git init
git add .
git commit -m "init: lightning warehouse web app"
git branch -M main
git remote add origin https://github.com/你的用户名/lightning-warehouse.git
git push -u origin main
```

> 第一次 push 会要你登录 GitHub,用 Personal Access Token 登录(不要用密码):
> 1. GitHub 右上角头像 -> Settings -> Developer settings -> Personal access tokens -> Tokens (classic)
> 2. Generate new token,勾选 `repo`,生成后**复制保存**(只显示一次)
> 3. push 时用户名填 GitHub 用户名,密码填这个 token

## 3. 部署到 Render(3 分钟)

### 3.1 注册 Render
- 打开 https://render.com
- 点 `Get Started for Free`,用 GitHub 账号登录(**不需要绑卡**)

### 3.2 创建 Web Service
1. 登录后点 `New +` -> `Blueprint`
2. 在仓库列表找到 `quingfeng007/lightning-warehouse` 并选中
3. 点 `Connect`

### 3.3 确认配置
Render 会自动检测到 `render.yaml`,展示配置:
```
Name: lightning-warehouse
Runtime: Python
Build: pip install -r requirements.txt
Start: gunicorn --workers 2 --timeout 300 app:app
Plan: Free
```
点 `Apply`。

### 3.4 等待部署
等 3-5 分钟。Render 会:
1. 拉你的代码
2. 装依赖(Flask, pandas, openpyxl...)
3. 启动服务

进度在 Dashboard 上看,日志可点 `Logs` 标签看。

### 3.5 拿到链接
部署成功后,Render 会显示链接:
```
https://lightning-warehouse-xxxx.onrender.com
```
(**x** 是随机字符)

**这个链接就是入口**,浏览器、手机、发给同事,都能用。

---

## 4. 保活方案(关键!)

### 4.1 问题
Render **免费实例 15 分钟没人访问会休眠**,下一次访问要等 30-60 秒冷启动。

### 4.2 方案 2:cron-job.org 每 10 分钟 ping(本项目已支持)

**思路**:用第三方 cron 服务每 10 分钟访问一次 `/ping` 端点,实例就不会休眠。

#### 步骤 1:注册 cron-job.org
1. 打开 https://cron-job.org
2. 点 `Create account`,邮箱注册(免费)
3. 登录后进 Dashboard

#### 步骤 2:创建定时任务
1. 点 `CREATE CRONJOB`
2. 填表:
   - **Title**:`Keep Render Alive`(随便起)
   - **URL**:`https://lightning-warehouse-xxxx.onrender.com/ping`(**替换成你的真实链接**)
   - **Schedule**:`Every 10 minutes`
3. 点 `CREATE`

#### 步骤 3:验证
等 10 分钟,看 cron-job.org 的 `Executions history`:
- ✅ HTTP 200 表示服务在跑
- ❌ HTTP 000 或 timeout 表示服务没起来

---

### 4.3 方案 3:GitHub Actions 自动保活(推荐,更稳)

**思路**:用 GitHub 自己的定时任务(GitHub Actions 每月 2000 分钟免费额度,远超保活需求)。

#### 配置
本项目已包含 `.github/workflows/keepalive.yml`,**Push 到 GitHub 后自动生效**,无需额外操作。

#### 验证
1. 打开你的 GitHub 仓库
2. 点 `Actions` 标签
3. 能看到 `Keep Render Alive` workflow 列表
4. 点进去看执行记录

#### 修改频率
打开 `.github/workflows/keepalive.yml`:
```yaml
schedule:
  - cron: '*/10 * * * *'  # 每 10 分钟
```

可以改成 `*/14 * * * *`(每 14 分钟)等。

---

## 5. 限制

### 5.1 性能
- 单次处理 Excel 最大 100MB
- 处理超时 5 分钟(可调 gunicorn timeout)

### 5.2 文件
- 上传的 Excel 文件存在 `/tmp` 临时目录
- 处理完 1 小时后自动清

### 5.3 安全性
- 任何人访问链接都能用
- 适合小团队内部使用,不要公开传敏感数据

---

## 6. 故障排查

| 现象 | 解决方法 |
|---|---|
| 看不到 log | Render Dashboard -> 选服务 -> `Logs` 标签 |
| 第一次很慢 | 冷启动,等 30-60 秒 |
| 报错"文件读取失败" | 检查文件格式 .xlsx/.xls/.xlsm/.csv |
| 中差评数/推广费=0 | 没传文件 3/4,这是可选的 |
| 保活不生效 | 看 cron-job.org 的执行历史或 GitHub Actions 日志 |

---

## 7. 本地开发

```powershell
cd C:\Users\33324\.qclaw\workspace-agent-3657e7c8\web
# 改完后
git add .
git commit -m "fix: xxx"
git push
# Render 会自动部署,等 1-2 分钟

# 本地测试
py -3.12 app.py
# 访问 http://localhost:5000
```

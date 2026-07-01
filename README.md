# 闪电仓计算工具 - Web 部署指南

把现有的桌面 GUI 工具（PyQt5）改成 Web 版，部署到免费云平台。

## 1. 准备工作（5 分钟）

### 1.1 注册 GitHub
- 打开 https://github.com 注册账号（已有跳过）
- GitHub 用来托管代码，平台从 GitHub 拉代码部署

### 1.2 安装 Git
- 下载 https://git-scm.com/download/win
- 安装时全部下一步

## 2. 上传代码到 GitHub（5 分钟）

### 2.1 在 GitHub 上创建仓库
1. 登录 GitHub
2. 右上角 `+` → `New repository`
3. 名称填 `lightning-warehouse`（随便起）
4. 选 `Public`（免费账户必须 public 才能让 Render 拉）
5. 不要勾选 `Add a README file`
6. 点 `Create repository`

### 2.2 推送代码
在 PowerShell 里执行：

```powershell
cd C:\Users\33324\.qclaw\workspace-agent-3657e7c8\web
git init
git add .
git commit -m "init: lightning warehouse web app"
git branch -M main
git remote add origin https://github.com/你的用户名/lightning-warehouse.git
git push -u origin main
```

> 第一次 push 会要你登录 GitHub，用 Personal Access Token 登录（不要用密码）：
> 1. GitHub 右上角头像 → Settings → Developer settings → Personal access tokens → Tokens (classic)
> 2. Generate new token，勾选 `repo`，生成后**复制保存**（只显示一次）
> 3. push 时用户名填 GitHub 用户名，密码填这个 token

## 3. 部署到 Render（5 分钟）

### 3.1 注册 Render
- 打开 https://render.com
- 点 `Get Started for Free`，用 GitHub 账号登录（**不需要绑卡**）

### 3.2 创建 Web Service
1. 登录后点 `New +` → `Blueprint`
2. 选 `lightning-warehouse` 仓库
3. Render 自动检测到 `render.yaml`，点 `Apply`
4. 等 3-5 分钟部署完成

### 3.3 获取访问地址
- 部署成功后，Render 会显示一个 `https://lightning-warehouse-xxx.onrender.com` 的链接
- 浏览器打开这个链接就能用了

## 4. 注意事项

### 4.1 免费实例休眠
Render 免费实例**15 分钟无访问会休眠**，下次访问需要 30-60 秒冷启动。
- 解决方案 1：付费 $7/月保持 24/7 在线
- 解决方案 2：用 [cron-job.org](https://cron-job.org) 每 10 分钟 ping 一次保持唤醒

### 4.2 文件大小限制
- 单个 Excel 不超过 100MB
- 总处理时间超过 5 分钟（gunicorn timeout）会被中断

### 4.3 隐私
- 上传的 Excel 在服务器 `/tmp` 临时目录处理
- 任务完成后 1 小时自动清理
- **不要上传含敏感数据的真实文件**到免费服务上

## 5. 日常使用

部署完成后，你可以：
- ✅ 在手机浏览器打开链接上传
- ✅ 分享给同事
- ✅ 换电脑/换设备都通用
- ✅ 不用装 Python、不用装 PyQt5

## 6. 故障排查

| 现象 | 解决方案 |
|---|---|
| 部署失败，看 log | Render Dashboard → 你的服务 → Logs |
| 冷启动很久 | 免费实例正常，等 30-60 秒 |
| 上传后无反应 | 检查文件格式（必须 .xlsx/.xls/.xlsm/.csv）|
| 提示"缺少商家名称列" | 文件2/3/4 第一行要找的列名不对 |
| 匹配数 = 0 | 透视表里的简称和文件1 门店名对不上 |

## 7. 进阶：自己升级或改代码

```powershell
cd C:\Users\33324\.qclaw\workspace-agent-3657e7c8\web
# 改完代码后:
git add .
git commit -m "fix: xxx"
git push
# Render 会自动重新部署（约 2 分钟）
```

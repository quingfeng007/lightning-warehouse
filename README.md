# 闪电仓计算工具 - Web 部署指南

把现有命令行 GUI 工具(PyQt5)改成 Web 版,部署到 Render 免费云平台,任何浏览器都能访问。

## 🚀 一键部署到 Render(5 分钟)

### 1. 准备 GitHub(2 分钟)

如果还没有这个仓库,在 https://github.com/new 新建一个空仓库,名字随便(比如 `lightning-warehouse`),**不要勾选** Add README。

然后在本地这个 `web/` 目录下执行:

```bash
cd web
git init
git add .
git commit -m "init: web version"
git remote add origin https://github.com/你的用户名/你的仓库名.git
git push -u origin main
```

### 2. 在 Render 部署(3 分钟)

1. 打开 https://render.com → 用 GitHub 账号登录
2. 右上角 `New +` → `Blueprint`
3. 选你的 `lightning-warehouse` 仓库 → `Connect`
4. Render 自动读 `render.yaml`,点 `Apply`
5. 等待 3-5 分钟(看 Logs 标签),部署完成

### 3. 设置访问密码(强烈推荐)

1. Render Dashboard → 你的服务 → `Environment`
2. 点 `Add Environment Variable`:
   - Key: `APP_PASSWORD`
   - Value: 你的密码(比如 `mySecret123`)
3. 保存,Render 自动重启
4. 访问你的网址,会先看到登录页,输入密码才能用

**不设密码 = 任何人都能上传文件用你的工具**(虽然没风险,但可能被滥用)

## 🛡️ 安全特性

### 自动锁屏(5 分钟)
- 5 分钟无任何操作(鼠标移动/点击/键盘)自动锁定
- 锁定前 30 秒顶部出现橙色警告条
- 切回 tab/继续操作 自动重置计时
- 后端 session 同步计时(刷新页面也要重新计时)

### 鉴权
- 默认密码: `lightning`(改 `APP_PASSWORD` 环境变量)
- 所有 API 都要鉴权,前端绕过不了
- 401 自动跳登录页

### 文件安全
- 上传文件存到 `/tmp/lightning_<id>/`,处理完用户点下载,服务器保留 1 小时后自动清理
- 不会被别人访问(每个任务 ID 是随机的 12 位字符串)

## 📁 目录结构

```
web/
├── app.py                       # Flask 主程序(含锁屏、鉴权)
├── lightning_warehouse_tool.py  # 核心处理逻辑(从 PyQt5 版复用)
├── requirements.txt             # Python 依赖
├── render.yaml                  # Render 部署配置
└── README.md                    # 本文件
```

## 🔧 本地开发

```bash
cd web
pip install -r requirements.txt
python app.py
# 访问 http://localhost:5000
# 默认密码: lightning
```

## 🌐 Render 注意事项

- **免费实例会休眠**:15 分钟没人访问就休眠,下次访问要 30-60 秒唤醒
- **冷启动**:第一个访问最慢,之后 15 分钟内的访问秒开
- **超时**:Render 免费版请求超时 30 秒,大文件可能不够;要 5 分钟超时需升级($7/月)

## 📋 4 个文件要求

| # | 文件 | 必填 | 格式 |
|---|---|---|---|
| 1 | 闪电仓计算模板 | ✅ | .xlsx/.xls/.xlsm |
| 2 | 全店数据-门店成交明细 | ✅ | .xlsx/.xls/.xlsm/.csv |
| 3 | 评价分析明细 | ❌ 可选 | .xlsx/.xls/.xlsm/.csv |
| 4 | 门店推广费 | ❌ 可选 | .xlsx/.xls/.xlsm/.csv |

只填文件 1+2 也能用(只是没有中差评数和推广费)。

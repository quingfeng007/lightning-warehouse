# AGENTS.md — 项目协作约定（AI 助手/开发者进入本项目必读）

## 一、文档同步规则（强制，不可跳过）

**每次系统更新（任何影响功能、接口、页面、部署、运维的代码变更）发布时，必须同步更新《项目交接文档.md》**，缺一不可：

1. 更新文档头部的"最后更新"日期与"当前线上版本"；
2. 在第 11 节"版本历史"追加一行：版本号 / 时间 / 更新内容 / commit 号；
3. 逐项核对并更新受影响的章节（第 4 节业务规则、第 5 节代码结构、第 7 节发布流程、第 9 节已知问题等），一律以代码实际行为为准，不写"计划做"的事；
4. 同步更新 `app.py` 中三个版本常量：`APP_VERSION`（递增）、`APP_UPDATED`（发布时间）、`APP_RELEASE_NOTE`（页面底部展示的更新说明）；
5. 发布后用 `curl https://lightning-warehouse.onrender.com/health` 验证 `version` 与 `started` 已变化，才算发布完成。

**版本递增规则**：
- 功能性 commit → `APP_VERSION` 必须递增（哪怕是很小的改动）；
- 纯文档/纯注释 commit → 不递增版本，但仍更新交接文档的"最后更新"日期（如有必要）。

## 二、推送约定

本机凭据管理器中的 GitHub 账号（xiaofeixiang007）对仓库**只有读权限**，提交作者为 quingfeng007。推送命令固定使用：

```bash
git -c credential.username=quingfeng007 push origin main
```

首次可能弹出 GitHub 登录授权窗口，需仓库所有者完成授权；授权成功后凭据会被记住。

## 三、发布检查清单（每次发布过一遍）

- [ ] 本地测试通过（核心逻辑改动跑 `test_data/run_e2e_test.py`）
- [ ] `app.py` 三个版本常量已更新（功能性变更必须递增 APP_VERSION）
- [ ] 《项目交接文档.md》已同步（头部日期/版本、版本历史、受影响章节）
- [ ] commit 信息清晰（feat:/fix:/docs: 前缀）
- [ ] push 成功且 `/health` 的 version/started 已更新
- [ ] 页面底部版本栏与 `/health` 一致

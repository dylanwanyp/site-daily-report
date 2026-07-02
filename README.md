# 外勤工程地盘日报系统 — 部署说明（GitHub Pages）

## 系统架构

```
browser (employee.html / admin.html)
  -> HTTPS fetch -> Google Apps Script Web App
    -> Google Sheet (5 tabs)
GitHub Pages (免费): static hosting for HTML files
```

---

## Step 1: 建立 Google 试算表

1. 前往 [sheets.new](https://sheets.new) 建立一个新试算表
2. 将试算表改名（例如「地盘日报系统」）
3. 建立以下 5 个 tabs（改名 sheet 下方 tab）：

| Tab 名称 | 用途 | 栏位 (第一行) |
|---|---|---|
| config | 系统设定 | key | value |
| sites | 地盘清单 | 地盘名称 | 所属子公司 |
| employees | 员工清单 | 员工姓名 | 月结开始日 |
| subsidiaries | 子公司清单 | 子公司名称 |
| records | 每日记录 | 日期 | 员工姓名 | 地盘名称 | 子公司 | 工时(hr) | 备注 | 提交时间 |

4. 在 config tab 输入管理密码：
   - A1: admin_password
   - B1: 您设定的密码（例如 abc123）

---

## Step 2: 部署 Google Apps Script

1. 在试算表中按选单：延伸功能 -> Apps Script
2. 将预设的 Code.gs 内容全部删除
3. 贴上本项目 Code.gs 的完整内容
4. 按储存（或 Ctrl+S）
5. 按部署 -> 新增部署作业
   - 类型：网页应用程序
   - 执行身分：我
   - 存取权：任何人
6. 按部署，复制产生的网页应用程序网址

---

## Step 3: 设定 Frontend

本项目的 HTML 已内置正确的 GAS 网址，无需修改。
本项目使用的 GAS 网址：

```
https://script.google.com/macros/s/AKfycbw9wWpDVdY0GvvFbmPW3m8Zh2L7wUT0u47WMxYgBDg16Z2_5W2zPjft9tH7spNx7aERjQ/exec
```

若需更换 GAS 网址，请修改 employee.html 和 admin.html 最上方的 `GAS_URL` 变量。

---

## Step 4: 部署到 GitHub Pages

1. 前往 https://github.com/dylanwanyp/site-daily-report
2. 在 Repo Settings → Pages → 将 Source 设为「Deploy from branch」→ master → / (root)
3. GitHub 会自动部署，约 1-2 分钟生效
4. 员工页面：https://dylanwanyp.github.io/site-daily-report/employee.html
5. 管理后台：https://dylanwanyp.github.io/site-daily-report/admin.html
6. 每次 git push 到 master 分支，GitHub 会自动重新部署（无使用次数限制）

---

## Step 5: 开始使用

1. 开启管理后台 -> 输入密码解锁
2. 先在子公司管理 tab 加入子公司
3. 在地盘管理 tab 加入地盘
4. 在员工管理 tab 加入员工（设定月结开始日）
5. 员工即可用手机开启 employee.html 填写记录

---

## 日常管理

| 动作 | 方式 |
|---|---|
| 修改管理密码 | 直接在 config tab 改 B1 |
| 新增/删除地盘 | 管理后台 -> 地盘管理 |
| 新增/删除员工 | 管理后台 -> 员工管理 |
| 新增/删除子公司 | 管理后台 -> 子公司管理 |
| 查看历史记录 | 直接打开 Google 试算表 |
| 更改员工月结日 | 直接在 employees tab 修改 |

---

## 常见问题

Q: 需要 HTTPS 吗？
A: GitHub Pages 自动提供 HTTPS，GAS 也是 HTTPS，全程加密。

Q: 可以多人同时填写吗？
A: 可以，极少情况同时提交可能有 race condition，一般用量可接受。

Q: 手机字体够大吗？
A: 员工页面已针对手机优化：字体 18px，input 22px，按钮 48px 高。

Q: 可以不用 GitHub Pages 吗？
A: 可以。任何 static hosting 都得（Netlify, Cloudflare Pages, Vercel...）。但建议用 GitHub Pages，因为免费无限制。

---

## 档案清单

| 档案 | 说明 |
|---|---|
| Code.gs | Google Apps Script 后端 |
| employee.html | 员工填写页（手机优先） |
| admin.html | 管理后台（桌面设计） |
| README.md | 部署说明 |

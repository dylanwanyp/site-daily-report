# 外勤工程地盘日报系统 — 部署说明

## 系统架构

```
browser (employee.html / admin.html)
  -> HTTPS fetch -> Google Apps Script Web App
    -> Google Sheet (5 tabs)
Netlify 免费 plan: static hosting for HTML files
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

在两个 HTML 档案中，将 `REPLACE_WITH_YOUR_GAS_WEB_APP_URL` 替换成你的 GAS 网址：

- employee.html 的 GAS_URL
- admin.html 的 GAS_URL

```javascript
var GAS_URL = 'https://script.google.com/macros/s/你的ID/exec';
```

---

## Step 4: 部署到 Netlify

1. 前往 netlify.com 并登入
2. 拖放 employee.html 和 admin.html 到部署区域
3. Netlify 会自动产生网址
4. 员工页面：https://XXXX.netlify.app/employee.html
5. 管理后台：https://XXXX.netlify.app/admin.html

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
A: Netlify 自动提供 HTTPS，GAS 也是 HTTPS，全程加密。

Q: 可以多人同时填写吗？
A: 可以，极少情况同时提交可能有 race condition，一般用量可接受。

Q: 手机字体够大吗？
A: 员工页面已针对手机优化：字体 18px，input 22px，按钮 48px 高。

Q: 可以不用 Netlify 吗？
A: 可以。任何 static hosting 都得（GitHub Pages, Cloudflare Pages, Vercel...）。

---

## 档案清单

| 档案 | 说明 |
|---|---|
| Code.gs | Google Apps Script 后端 |
| employee.html | 员工填写页（手机优先） |
| admin.html | 管理后台（桌面设计） |
| README.md | 部署说明 |

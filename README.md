# ACS Outlook 统一签名加载项 — 部署与使用说明

| 项目 | 值 |
|---|---|
| 托管地址 | `https://sig.breezefreight.com`（GitHub Pages + 自定义子域名） |
| 发件邮箱 | `@aioncargo.com`（不受影响，不需要改 GoDaddy 上的任何设置） |
| 需要的权限 | breezefreight.com 的 DNS 管理权限、GitHub 账号、M365 全局管理员（或应用管理员 + Exchange 管理员） |

加载项在员工**新建 / 回复 / 转发**邮件时自动插入公司统一签名。经典 Outlook、新版 Outlook、网页版、Mac、iOS、Android 用的都是同一套模板。

- 个人信息（姓名、手机、电话、分机、职位）从 **Entra ID** 自动读取
- 公司信息、Logo、免责声明在 `src/config.js` 里统一修改，推送到 GitHub 后全员自动生效
- **完整签名**用于新邮件、会话中的第一次回复和转发；**精简签名**（无图片，含澳洲地址）用于同一会话中的后续回复
- 图片以内嵌方式插入，收件人**看不到** breezefreight.com，各个客户端显示一致

---

## 部署总览

| # | 步骤 | 在哪里操作 | 用时 |
|---|---|---|---|
| 0 | 补全员工 Entra 资料 | M365 管理中心 | 视人数而定 |
| 1 | 建 GitHub 仓库并上传代码 | GitHub | 10 分钟 |
| 2 | 添加 DNS 记录 `sig` | breezefreight.com 的 DNS 服务商 | 5 分钟（生效 5 分钟到数小时） |
| 3 | 开启 GitHub Pages 并绑定域名 | GitHub | 10 分钟（HTTPS 证书最多等 24 小时） |
| 4 | 在 Entra 注册应用 | Entra 管理中心 | 10 分钟 |
| 5 | 填 clientId / tenantId 并推送 | GitHub | 5 分钟 |
| 6 | 检查线上文件 | 浏览器 | 2 分钟 |
| 7 | 部署给试点用户 | M365 管理中心 → 集成应用 | 5 分钟（生效最长 24 小时） |
| 8 | 测试，然后推给全员 | Outlook 各客户端 | — |

---

## 步骤 0：补全员工的 Entra 资料

签名内容取决于 Entra 里填了什么，**缺哪个字段，签名里就少哪一行**。

| 签名项 | Entra 字段 | 填写示例 |
|---|---|---|
| 姓名 | Display name | Ivy Hu |
| 职位（精简签名显示） | Job title | Managing Director – Australia |
| Mob | Mobile phone | +61 425 666 802 |
| Tel + EXT | Office phone | `+61 2 9160 2300 EXT 601` |

在哪里填：[M365 管理中心](https://admin.microsoft.com) → 用户 → 活跃用户 → 选中用户 → **管理联系信息**。

批量修改用 PowerShell：

```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser   # 首次运行需要安装
Connect-MgGraph -Scopes "User.ReadWrite.All"
Update-MgUser -UserId ivy.hu@aioncargo.com `
  -MobilePhone "+61 425 666 802" `
  -BusinessPhones @("+61 2 9160 2300 EXT 601") `
  -JobTitle "Managing Director – Australia"
```

> 分机也可以不写进电话，改填在 Exchange 自定义属性 1（extensionAttribute1），例如 `601`。办公电话留空的人，签名里会显示总机号码。

---

## 步骤 1：建 GitHub 仓库并上传代码

1. 登录 GitHub → 右上角 **+ → New repository**
   - Repository name：`acs-outlook-signature`
   - 选 **Public**。免费版 GitHub 只有公开仓库能用 Pages；如果你们有 GitHub Team 或 Enterprise，可以选 Private。
   - 其他选项保持默认 → **Create repository**
2. 解压 `acs-outlook-signature.zip`，把**文件夹里的全部内容**上传到仓库根目录。

   **方法 A：网页上传**
   - 在仓库页面点 **uploading an existing file**，把解压后文件夹里的内容全部拖进去 → **Commit changes**
   - ⚠️ 网页拖拽**经常会漏掉以 `.` 开头的隐藏文件夹**。上传后请检查仓库里有没有 `.github/workflows/deploy.yml`。如果没有，点 **Add file → Create new file**，文件名输入 `.github/workflows/deploy.yml`，把压缩包里这个文件的内容粘贴进去再提交。

   **方法 B：命令行（推荐）**
   ```bash
   cd acs-outlook-signature
   git init -b main
   git add .
   git commit -m "ACS signature add-in"
   git remote add origin https://github.com/<你的账号>/acs-outlook-signature.git
   git push -u origin main
   ```

3. 第一次推送后，Actions 页面会出现一次失败的部署，这是正常的（Pages 还没开启）。步骤 3 完成后会重新运行。

> 仓库公开是安全的：里面只有 HTML、JS、Logo 和 clientId。clientId 不是密钥，本来就会出现在前端代码里。**员工资料不在仓库中**，是运行时从 Graph 读取的。

---

## 步骤 2：在 breezefreight.com 添加 DNS 记录

到 breezefreight.com 的 DNS 管理后台（域名在哪里注册或托管就去哪里，比如 GoDaddy、Cloudflare、阿里云），新增一条记录：

| 类型 | 主机 / 名称 | 值 / 指向 | TTL |
|---|---|---|---|
| **CNAME** | `sig` | `<你的 GitHub 用户名或组织名>.github.io` | 默认（1 小时） |

- 例如 GitHub 用户名是 `leonliu-acs`，值就填 `leonliu-acs.github.io`。**不要带仓库名，也不要加 `https://`**。
- 如果用的是 **Cloudflare**：必须把代理状态设为 **仅 DNS（灰色云朵）**，否则 GitHub 无法签发 HTTPS 证书。
- 检查是否生效（Windows 命令提示符）：`nslookup sig.breezefreight.com`，结果里能看到 `github.io` 就说明生效了。

**建议（可选）：在 GitHub 验证域名，防止子域名被别人抢占**

1. GitHub 右上角头像 → **Settings → Pages → Add a domain**，输入 `breezefreight.com`
2. 按页面提示在 DNS 里加一条 **TXT** 记录（名称类似 `_github-pages-challenge-<用户名>`）
3. 回到 GitHub 点 **Verify**

---

## 步骤 3：开启 GitHub Pages 并绑定域名

1. 进入仓库 → **Settings → Pages**
2. **Build and deployment → Source** 选 **GitHub Actions**
3. **Custom domain** 填 `sig.breezefreight.com` → **Save**
   - 页面会显示 “DNS check in progress”，通过后显示绿色的 “DNS check successful”
4. 等证书签发（通常几分钟，最多 24 小时），然后勾选 **Enforce HTTPS**
   - 如果这个选项是灰色的，说明证书还没签发好，稍后刷新再勾
5. 进入 **Actions** 标签 → 左侧选 **Build & deploy to GitHub Pages** → 右侧 **Run workflow**，重新运行一次

> 构建时会根据 `baseUrl` 自动生成 `CNAME` 文件，所以以后每次发布，自定义域名都不会丢失。

---

## 步骤 4：在 Entra 注册应用

打开 [Entra 管理中心](https://entra.microsoft.com) → **标识 → 应用程序 → 应用注册 → 新注册**：

| 字段 | 填写 |
|---|---|
| 名称 | `ACS Outlook Signature` |
| 受支持的账户类型 | **仅此组织目录中的账户（单一租户）** |
| 重定向 URI | 平台选 **单页应用程序 (SPA)**，值填 `brk-multihub://sig.breezefreight.com` |

注册后：

1. 在 **概述** 页复制两个值，步骤 5 要用：
   - **应用程序(客户端) ID**
   - **目录(租户) ID**
2. **API 权限 → 添加权限 → Microsoft Graph → 委托的权限**，搜索并勾选：
   - `User.Read`：读取本人资料
   - `Mail.ReadBasic`：只读邮件的发件人和时间等元数据，**读不到正文**，用来判断本人是否已在这个会话里回复过
3. 点击 **代表 ACS 授予管理员同意 → 是**。两个权限的状态都应显示绿色的“已授予”。

> 重定向 URI 必须**只写域名**，前缀是 `brk-multihub://`，不能写 `https://`，末尾也不能有 `/`。这里填错的话，所有人都读不到 Entra 资料。

---

## 步骤 5：填 clientId / tenantId 并推送

两种方法任选一种。

**方法 A：改文件**。在 GitHub 网页上打开 `addin.config.json` → 点铅笔图标编辑：

```json
{
  "baseUrl": "https://sig.breezefreight.com",
  "clientId": "<步骤 4 的应用程序(客户端) ID>",
  "tenantId": "<步骤 4 的目录(租户) ID>",
  "addinId": "7d1f3c2a-5b8e-4a61-9f0d-2c6e8b4a1d73",
  "version": "1.0.0.0",
  ...
}
```

然后 **Commit changes**。

**方法 B：用仓库变量**，不改文件。仓库 **Settings → Secrets and variables → Actions → Variables** 标签 → **New repository variable**：

- `ADDIN_CLIENT_ID` = 应用程序(客户端) ID
- `ADDIN_TENANT_ID` = 目录(租户) ID

设置完后，去 Actions 手动 **Run workflow** 一次。

**顺便修改 `src/config.js` 里的这两处：**

- `termsUrl`：改成你们保密声明 / T&C 的真实链接（现在是占位地址 `https://www.aioncargo.com/terms`）
- 如需调整总机号码、地址或结束语，也在这里改

**替换图片：** `src/assets/acs-logo.png` 和 `src/assets/wca-badge.png` 目前是从截图里裁出来的，建议换成设计原稿。
- 使用 PNG 格式，不能用 SVG
- Logo 尺寸 262×162，徽章 260×168（显示尺寸的 2 倍）
- 单张 50KB 以内
- 文件名保持不变

每次提交后，Actions 会自动运行测试、构建和发布，大约 1–2 分钟。在 Actions 页面看到绿色勾就说明成功了。

---

## 步骤 6：检查线上文件

在浏览器中逐个打开以下地址，都应该能正常显示，并且地址栏有锁形图标：

| 地址 | 应看到的内容 |
|---|---|
| https://sig.breezefreight.com/manifest.xml | XML 内容，里面所有地址都是 `sig.breezefreight.com` |
| https://sig.breezefreight.com/commands.html | 空白页（正常） |
| https://sig.breezefreight.com/launchevent.js | 一大段压缩过的 JS |
| https://sig.breezefreight.com/.well-known/microsoft-officeaddins-allowed.json | `{ "allowed": ["https://sig.breezefreight.com/launchevent.js"] }` |
| https://sig.breezefreight.com/assets/acs-logo.png | ACS Logo |

如果 `.well-known` 打不开，说明隐藏文件没有发布成功：检查 `.github/workflows/deploy.yml` 里有没有 `include-hidden-files: true`，以及 Actions 是否运行成功。

---

## 步骤 7：部署给试点用户

1. 在浏览器打开 https://sig.breezefreight.com/manifest.xml → 右键 **另存为** `manifest.xml`
2. [M365 管理中心](https://admin.microsoft.com) → **设置 → 集成应用 → 上传自定义应用**
3. 应用类型选 **Office 加载项** → **上传清单文件 (.xml)** → 选择刚才的 `manifest.xml`
4. **分配用户**：**先选 2–3 个试点用户**（比如你自己和 Ivy）
5. 接受权限 → **完成部署**
6. 最长 24 小时生效，通常 1–6 小时。生效后：
   - 网页版：刷新页面
   - 经典 / 新版 Outlook：完全退出后重新打开
   - 手机：强制关闭 Outlook App 后重新打开

---

## 步骤 8：测试，然后推给全员

在试点用户的每个客户端（网页版、新版 Outlook、经典 Outlook、手机）上逐项测试：

- [ ] 新建邮件 → 自动出现**完整签名**，Logo 和徽章显示正常
- [ ] 回复一封从未回复过的外部来信 → **完整签名**
- [ ] 同一会话再回复一次 → **精简签名**（姓名 | 职位 / ACS Aion Cargo Solutions | 地址 / 电话）
- [ ] 转发 → **完整签名**
- [ ] 发一封到外部邮箱（Gmail、QQ 邮箱），确认对方收到后图片正常显示
- [ ] 撰写邮件时点功能区的 **ACS Signature** 按钮 → 面板显示“✔ 已连接 Entra ID”，个人资料正确

测试通过后：

1. 集成应用 → 点击 **ACS Email Signature** → **编辑用户** → 改为 **整个组织**，或者只分配给澳洲员工所在的组
2. 通知员工：删除 Outlook 里原来的个人签名。路径是 Outlook → 设置 → 账户 → 签名，删除签名，并把“新邮件”和“回复 / 转发”的默认签名都设为“无”。手机端在 设置 → 签名 里清空。

> 加载项会自动屏蔽本地签名，所以不删也不会出现两份签名；删掉是为了避免员工混淆。

---

## 签名规则

| 场景 | 签名 |
|---|---|
| 新邮件 | 完整签名 |
| 回复：本人还没有在这个会话里发过邮件 | 完整签名 |
| 回复：本人已经在这个会话里发过邮件 | 精简签名 |
| 转发 | 完整签名（改成按回复规则处理：在 `config.js` 里设 `forwardAlwaysFull: false`） |

判断方式：通过 Graph 查本人“已发送邮件”里有没有同一会话的记录。Graph 不可用时，改为检查引用的历史邮件里有没有本人的手机号，或者 “Email 本人邮箱” 这一行。

---

## 日常维护

| 需求 | 操作 | 需要重新上传 manifest 吗 |
|---|---|---|
| 员工换手机、升职 | 在 Entra 修改；最迟 12 小时后签名自动更新，或让员工在面板点“刷新资料” | 否 |
| 新员工入职 | 补全 Entra 资料即可（加载项已分配给整个组织或组） | 否 |
| 改公司地址、网址、免责声明、结束语 | 修改 `src/config.js` → 提交 | 否 |
| 换 Logo 或徽章 | 替换 `src/assets/` 里的文件 → 提交 | 否 |
| 改签名样式 | 修改 `src/shared/signature.js` → 提交 | 否 |
| 改按钮名称、权限（manifest） | 修改 `manifest.template.xml`，把 `addin.config.json` 的 `version` 改为 `1.0.1.0` → 提交 → 集成应用里**更新**加载项 | **是** |

> WCA 徽章上印有到期日 **2027 年 4 月 7 日**，续期后记得更换图片。

---

## 故障排查

| 现象 | 可能原因 / 处理 |
|---|---|
| 完全没有签名 | ① 部署还没生效（最长 24 小时）② 该用户没有被分配 ③ Outlook 版本过低（见下表）④ 打开 `commands.html` 检查站点是否正常 |
| 只有姓名和邮箱，并提示“请打开签名面板” | 读不到 Entra 资料：① 是否已授予管理员同意 ② 重定向 URI 是否为 `brk-multihub://sig.breezefreight.com` ③ clientId / tenantId 是否正确 ④ 经典 Outlook 还要确认 `.well-known` 文件能打开 |
| GitHub Pages 显示 DNS 检查失败 | CNAME 记录填错（不要带仓库名或 `https://`），Cloudflare 没有关闭代理，或者 DNS 还没生效 |
| “Enforce HTTPS” 一直是灰色 | 证书还在签发中，最多等 24 小时；可以把 Custom domain 删掉后重新填一次，触发重新签发 |
| Actions 构建失败 | 点开失败的任务看日志，常见原因是 JSON 格式写错（多了或少了逗号） |
| 出现两份签名 | 客户端版本过低，不支持屏蔽本地签名 → 删除本地签名 |
| 签名里某一行缺失 | Entra 里对应字段为空，面板里会标红显示“未填写” |
| 第二次回复还是完整签名 | 第一次回复还在草稿或发件箱里，没有进入“已发送邮件” |
| 手机上点底部“回复”后看不到签名 | Outlook 手机端的已知行为：签名已经插入，展开为全屏编辑就能看到 |
| Apple Mail 收件人看到 Logo 显示为附件 | 内嵌图片的正常表现；介意的话可以把 `imageMode` 改为 `"link"`（代价是 Outlook 收件人默认可能不显示图片） |

**客户端最低版本**（低于这些版本只能用缓存资料或基本信息）：

| 客户端 | 版本 |
|---|---|
| 网页版 / 新版 Outlook | 无要求 |
| 经典 Outlook (Windows, Microsoft 365) | 2409 (Build 18025) 及以上 |
| Outlook for Mac | 16.89 及以上 |
| Outlook iOS / Android | 4.2433 及以上 |

查看日志：网页版按 F12 看控制台，搜索 `[ACS-SIG]`。经典 Outlook 可以开启 [runtime logging](https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging)。

---

## 已知限制

- 打开已有草稿时不会自动插入签名（Outlook 的设计），需要的话可以在面板里手动插入
- iOS 上通过“共享”新建的邮件不会触发
- 用共享邮箱或“代表他人发送”时，签名仍然是当前登录用户的
- 只处理邮件，不处理会议邀请
- **breezefreight.com 必须长期续费、DNS 必须保持不变**，否则全公司的签名都会失效。换域名时要按步骤 2–7 重新做一遍（包括在 Entra 里改重定向 URI）

---

## 本地开发（可选）

```bash
npm install
npm test          # 12 项单元测试：模板、分机解析、会话判断、事件处理流程
npm run preview   # 生成 preview.html，用示例资料预览两种签名
npm run build     # 生成 dist/
```

## 文件结构

```
addin.config.json            ★ 部署参数（域名、clientId、tenantId、版本）
src/config.js                ★ 公司信息与签名规则
src/assets/                  ★ Logo、徽章、按钮图标
src/shared/signature.js      完整签名 / 精简签名 HTML 模板
src/shared/profile.js        Entra 资料标准化、分机解析、会话判断
src/shared/graph.js          NAA 单点登录 + Microsoft Graph
src/shared/office.js         判断签名类型、插入签名
src/launchevent/             事件入口（自动插入）
src/taskpane/                ACS Signature 侧边面板
src/well-known/              经典 Outlook 需要的 SSO 白名单
manifest.template.xml        加载项清单模板
scripts/                     构建和预览脚本
test/                        单元测试
.github/workflows/deploy.yml 推送后自动测试、构建并发布到 GitHub Pages
```

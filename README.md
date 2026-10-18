# hongke-skills

鸿科内部 AI 技能仓库。这里维护我们在 Kimi 里使用的自定义技能（SOP 编写、流程图绘制、文档风格等），通过 `npx skills` 安装和同步。

## 当前包含的技能

| 技能 | 说明 |
|---|---|
| hongke-doc-style | 鸿科品牌风格（橙色系）+ SOP 标准框架 + 白底流程图规范（PPT 内流程图为原生可编辑形状，非贴图）。触发词："梳理一下流程"、"制定SOP"、"画个流程图"、"按我们的风格出文档" |

## 同事首次安装（只做一次）

前置条件：电脑上装有 Node.js 和 Git。如果仓库是私有的，还需要你的 GitHub 账号已被邀请为协作者；公共仓库无需账号，直接装。

```bash
npx skills add https://github.com/xuanzi-dev/hongke-skills -g
```

> 安装时会拉取仓库中全部技能；如果只要其中某一个，按提示选择即可。

## 以后同步更新

管理员更新技能后会 push 到仓库并在群里通知，收到通知后运行：

```bash
# 看看有没有新版本（可选）
npx skills check

# 同步到最新
npx skills update
```

## 常见问题

**Q：`npx skills add` 提示权限不足 / 仓库不存在？**
先确认仓库是公共还是私有。私有仓库走的是你本机的 Git 凭据：
1. 你的 GitHub 账号已被加为仓库协作者（在 GitHub 通知邮件里点接受邀请）；
2. 本机已配置 GitHub 凭据（执行 `git ls-remote https://github.com/xuanzi-dev/hongke-skills` 能通即可；不通就先按 GitHub 提示登录，推荐用 GitHub CLI `gh auth login` 或配置 Personal Access Token）。

**Q：更新后 Kimi 里没看到变化？**
新开一个会话再试；仍不行就 `npx skills remove hongke-doc-style -g` 后重新 `npx skills add https://github.com/xuanzi-dev/hongke-skills -g`。

**Q：想提修改建议？**
直接在群里反馈给管理员，由管理员统一修改发布，避免多人同时改动冲突。

## 管理员维护说明（写给仓库维护者）

1. 修改或新增技能：技能目录放在仓库**第一层**（不要套子目录，`npx skills` 对 SSH 地址 + 子目录的更新有已知 bug）。
2. 每次改动后提交并推送：
   ```bash
   git add -A
   git commit -m "更新说明"
   git push
   ```
3. 推送后在飞书群通知同事运行 `npx skills update`。
4. 新增同事：GitHub 仓库 → Settings → Collaborators → 邀请；同事接受邀请后即可安装。

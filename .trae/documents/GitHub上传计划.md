# GitHub 上传计划

## 当前状态分析

### Git 状态
- 分支：`master`
- 提交数：**0**（尚未有任何 commit）
- Remote：**未配置**
- Staging：仅 `.gitignore` 被 staging
- Untracked：`.python-version`、`.trae/`、`input/`、`main.py`、`pyproject.toml`、`src/`、`tests/`、`uv.lock`

### 目标仓库
- **URL**：`https://github.com/P4X666/ppt-transfor.git`
- **分支**：`master`

### .gitignore 规则（已配置，合理）
- `__pycache__/`、`*.py[oc]`、`build/`、`dist/`、`wheels/`、`*.egg-info` — Python 编译产物和构建目录
- `.venv` — 虚拟环境
- `out/` — 运行时产物（JSON、PPTX、视觉对比结果）

### 需上传的文件
- `main.py` — CLI 入口
- `pyproject.toml` — 项目配置
- `uv.lock` — 依赖锁定
- `.python-version` — Python 版本锁定
- `.gitignore` — Git 忽略规则
- `.trae/` — 项目计划文档（`PPT视觉差异修复与视觉对比集成计划.md` 等）
- `input/` — 17 个测试 PPT 文件
- `src/` — 源代码（`ppt_transfor/`）
- `tests/` — 测试代码

## 实施步骤

### 步骤 1：unstage 当前 staging 文件
当前 `.gitignore` 已在 staging，需先 unstage，然后重新 add 所有文件以确保 commit 完整。

```bash
git reset HEAD .gitignore
```

### 步骤 2：add 所有文件
```bash
git add .
```

验证：`git status` 应显示所有文件在 staged 状态，无 untracked 或 ignored 文件。

### 步骤 3：创建初始 commit
```bash
git commit -m "Initial commit: PPT ↔ JSON 双向转换服务"
```

commit 信息说明：清晰描述项目功能（PPT 解析为 JSON，JSON 渲染回 PPT，支持往返对比定位解析/渲染缺陷，含主题色固化和视觉对比工具）。

### 步骤 4：添加 remote
```bash
git remote add origin https://github.com/P4X666/ppt-transfor.git
```

验证：`git remote -v` 应显示 fetch/push 配置。

### 步骤 5：push 到 GitHub
```bash
git push -u origin master
```

**注意**：push 时需处理 GitHub 认证。可能的方式：
1. Personal Access Token (PAT) — 在 GitHub Settings → Developer Settings → Personal access tokens → Fine-grained tokens，生成有 repo 权限的 token
2. SSH key — 如本地配置了 SSH，改用 SSH URL：`git@github.com:P4X666/ppt-transfor.git`
3. GitHub CLI (`gh auth login`)

若 push 失败（权限不足），需先配置认证后重试。

### 步骤 6：验证
- 浏览器打开 `https://github.com/P4X666/ppt-transfor`
- 确认文件列表完整，commit 成功推送

## 风险与注意事项

### 风险 1：GitHub 认证
- **问题**：首次 push 可能因无认证失败
- **处理**：使用 Personal Access Token 或 SSH key。若是 PAT，Git push 时会提示输入用户名（GitHub 用户名）和密码（填入 PAT，而非账号密码）

### 风险 2：仓库不存在
- **问题**：GitHub 上尚未创建该仓库，push 会失败
- **处理**：浏览器登录 GitHub 手动创建空仓库（New Repository → Repository name: ppt-transfor → Public → Create），再重试 push。或者使用 `gh repo create ppt-transfor --public`

### 风险 3：GitHub 已有其他内容
- **问题**：目标仓库已有其他 commit，push 会冲突
- **处理**：`git push -u origin master` 如失败，可能需要 `git pull origin master --allow-unrelated-histories` 或改用 `git push -f`（会覆盖，谨慎使用）

### 注意事项
- `.gitignore` 已正确忽略 `out/`（运行时产物）和 `.venv`（虚拟环境），避免上传大量二进制文件
- `input/` 目录的 PPT 文件会被上传（约 17 个文件），文件大小应在 GitHub 合理范围内（单个文件 < 100MB 限制）
- `.trae/` 目录包含计划文档，会被上传，便于历史记录追溯
- `uv.lock` 锁定依赖版本，推荐上传，便于其他开发者复现环境

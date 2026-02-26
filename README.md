# RouteOps — 部署到 Railway（外网可访问）

## 文件结构（缺一不可）

```
eataway-router/
├── app.py
├── requirements.txt
├── Procfile
├── runtime.txt
├── coords.xlsx        ← 你的坐标数据
├── routes.xlsx        ← 你的路线数据
└── templates/
    └── index.html
```

---

## 部署步骤

### 第一步：安装 Git（如果没有）
```bash
brew install git
```

### 第二步：在项目文件夹初始化 Git 仓库
```bash
cd ~/eataway-router       # 进入你的项目文件夹
git init
git add .
git commit -m "first deploy"
```

### 第三步：创建 Railway 账号
访问 https://railway.app → 用 GitHub 登录（免费）

### 第四步：部署
```bash
# 安装 Railway CLI
brew install railway

# 登录
railway login

# 在项目文件夹里部署
railway init          # 选 "Empty Project"
railway up            # 上传并部署
```

### 第五步：获取外网地址
部署完成后运行：
```bash
railway domain
```
会生成类似 `https://eataway-router-production.up.railway.app` 的地址

---

## 部署后的使用

- 管理员面板：`https://你的地址.railway.app`
- 司机收到的链接：`https://你的地址.railway.app/links/Abbe`

WhatsApp 消息里的链接会自动变成 Railway 地址（`window.location.origin` 自动读取）

---

## 更新数据文件（coords.xlsx / routes.xlsx）

每次更新 Excel 文件后：
```bash
git add coords.xlsx routes.xlsx
git commit -m "update routes"
railway up
```

---

## 本地运行（开发调试用）
```bash
pip install -r requirements.txt
python3 app.py
# 访问 http://localhost:5050
```

# 返餐费审批提交工具

基于飞书开放平台 API 的退返餐费审批提交工具。财务人员上传 Excel 明细表，自动解析并提交飞书审批流程。

## 项目结构

```
MealRefundApproval/
├── app/                        # 应用代码
│   ├── app.py                  # Streamlit UI 入口
│   ├── feishu_api.py           # 飞书 API 封装（认证、上传、提交、查询）
│   ├── excel_parser.py         # Excel 解析（退返餐费明细表）
│   ├── form_builder.py         # 审批表单构建（字段映射）
│   └── logger_config.py        # 日志配置
├── test/                       # 测试代码
├── deploy/                     # 部署配置
│   └── meal-refund-approval.service  # systemd 服务单元示例
├── start.bat                   # Windows 启动脚本
├── start.sh                    # Linux 启动脚本（服务器部署）
├── requirements.txt            # Python 依赖
├── pyproject.toml              # 项目配置（uv + 清华镜像）
├── pytest.ini                  # pytest 配置
├── .env.example                # 环境变量模板
└── README.md                   # 项目说明
```

## 快速开始

### 1. 环境准备

```bash
# 创建虚拟环境
uv venv
source .venv/bin/activate   # Linux/macOS
.venv\Scripts\activate      # Windows

# 安装依赖（使用清华镜像加速）
uv pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

### 2. 配置凭证

```bash
cp .env.example .env
# 编辑 .env 填入飞书应用凭证
```

`.env` 配置项：

| 变量名 | 说明 | 示例值 |
|--------|------|--------|
| `FEISHU_APP_ID` | 飞书应用 App ID | `cli_aabb96a4a9b99bee` |
| `FEISHU_APP_SECRET` | 飞书应用 App Secret | `xxxxxxxxxxxxxxxx` |
| `FEISHU_APPROVAL_CODE` | 审批定义 Code（可选，默认使用内置值） | `081B5624-B8B7-4334-AF42-D1075F8E650A` |

### 3. 启动应用

**Windows 用户：**

双击 `start.bat` 即可启动。

**手动启动：**

```bash
streamlit run app/app.py
```

浏览器将自动打开 `http://localhost:8501`。

## Ubuntu 服务器部署

本节介绍如何在 Ubuntu 服务器上部署本工具，供团队通过浏览器访问。

### 环境要求

| 项目 | 要求 |
|------|------|
| 操作系统 | Ubuntu 20.04 / 22.04 / 24.04 |
| Python | 3.11 或更高版本 |
| 磁盘空间 | ~200 MB（含虚拟环境） |
| 网络 | 出站 HTTPS 访问 `open.feishu.cn` 和 `www.feishu.cn` |

### 1. 上传项目到服务器

```bash
# 从本地上传到服务器（假设服务器 IP 为 192.168.1.100）
scp -r MealRefundApproval user@192.168.1.100:/opt/

# 登录服务器后，重命名为规范路径
ssh user@192.168.1.100
sudo mv /opt/MealRefundApproval /opt/meal-refund-approval
```

> **提示**：建议创建专用部署用户（如 `deploy`），避免使用 root 运行服务。

### 2. 安装依赖

**方式一：系统 Python + venv（推荐）**

```bash
# 安装 Python 和 venv
sudo apt update
sudo apt install -y python3 python3-venv

# 进入项目目录，创建虚拟环境并安装依赖
cd /opt/meal-refund-approval
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

**方式二：uv（更快，可选）**

```bash
# 安装 uv
curl -LsSf https://astral.sh/uv/install.sh | sh

# 创建虚拟环境并安装依赖
cd /opt/meal-refund-approval
uv venv
uv pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

后续命令中，虚拟环境的 Python 路径为 `.venv/bin/python`。

### 3. 配置 .env

```bash
cd /opt/meal-refund-approval
cp .env.example .env
nano .env
```

填入飞书应用凭证：

```env
FEISHU_APP_ID=cli_aabb96a4a9b99bee
FEISHU_APP_SECRET=your_actual_secret_here
FEISHU_APPROVAL_CODE=081B5624-B8B7-4334-AF42-D1075F8E650A
```

> **首次运行前**：确保已在飞书开放平台开通 `contact:user.id:readonly` 权限（用于通过手机号查询用户 ID）。  
> 授权页面：[https://open.feishu.cn/app/cli_aabb96a4a9b99bee/auth?q=contact:user.id:readonly&op_from=openapi&token_type=tenant](https://open.feishu.cn/app/cli_aabb96a4a9b99bee/auth?q=contact:user.id:readonly&op_from=openapi&token_type=tenant)  
> 开通后需创建新版本并发布，等待几分钟生效。详见 [常见问题 #1](#1-查询发起人失败99991672-错误)。

### 4. 部署验证

```bash
cd /opt/meal-refund-approval

# 运行测试（67 个测试，无需网络）
.venv/bin/python -m pytest test/ -q

# 预期输出：67 passed
```

> **提示**：测试模板 `test/fixtures/返餐费2026.xls` 含真实姓名/身份证号，不随仓库分发。完整 67 个测试需自行将模板放入该路径；缺失时相关用例自动跳过，不会报错。

### 5. 启动方式一：start.sh 快速启动（适合试用/内网）

```bash
cd /opt/meal-refund-approval
chmod +x start.sh
./start.sh
```

`start.sh` 会自动检测虚拟环境、安装缺失依赖、启动 Streamlit 服务。

**保持后台运行：**

```bash
# 方式 A：使用 tmux（推荐，可随时 attach 查看）
tmux new -s mealrefund './start.sh'
# 按 Ctrl+B 然后 D  detach；重新 attach：tmux attach -t mealrefund

# 方式 B：使用 nohup
nohup ./start.sh > logs/start.log 2>&1 &
# 查看日志：tail -f logs/start.log
```

访问地址：`http://服务器IP:8501`

### 6. 启动方式二：systemd 服务（推荐，开机自启 + 崩溃自动重启）

**6.1 安装服务单元**

```bash
sudo cp /opt/meal-refund-approval/deploy/meal-refund-approval.service /etc/systemd/system/
```

**6.2 调整服务配置**

编辑 `/etc/systemd/system/meal-refund-approval.service`，根据实际情况修改：

```ini
# 修改为实际部署用户（如 ubuntu、deploy 等）
User=www-data
Group=www-data

# 确认路径正确
WorkingDirectory=/opt/meal-refund-approval
ExecStart=/opt/meal-refund-approval/.venv/bin/python -m streamlit run app/app.py ...
```

**6.3 启动服务**

```bash
# 重新加载 systemd 配置
sudo systemctl daemon-reload

# 启用并启动服务（开机自启）
sudo systemctl enable --now meal-refund-approval

# 查看服务状态
sudo systemctl status meal-refund-approval
```

**6.4 查看日志**

```bash
# systemd 日志（实时跟踪）
journalctl -u meal-refund-approval -f

# 应用日志（项目内）
tail -f /opt/meal-refund-approval/logs/app.log
```

**6.5 重启服务（更新代码后）**

```bash
sudo systemctl restart meal-refund-approval
```

### 7. 防火墙配置

```bash
# 允许 SSH 和 Streamlit 端口
sudo ufw allow 22
sudo ufw allow 8501/tcp
sudo ufw enable

# 仅允许内网访问（如 10.0.0.0/8 网段）
sudo ufw allow from 10.0.0.0/8 to any port 8501
```

### 8. 浏览器访问

在浏览器中打开：`http://服务器IP:8501`

> **注意**：服务器模式下 Streamlit 不会自动打开浏览器，需手动输入地址。

### 9. 可选：Nginx 反向代理

若需通过域名或 80 端口访问，可配置 Nginx 反向代理。

**9.1 安装 Nginx**

```bash
sudo apt install -y nginx
```

**9.2 创建站点配置**

```bash
sudo nano /etc/nginx/sites-available/meal-refund-approval
```

写入以下内容：

```nginx
server {
    listen 80;
    server_name refund.example.internal;  # 修改为实际域名或 IP

    # 飞书附件上传限制
    client_max_body_size 50m;

    location / {
        proxy_pass http://127.0.0.1:8501;
        proxy_http_version 1.1;
        
        # Streamlit WebSocket 支持（必需）
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        
        # 长连接超时（Streamlit 需要）
        proxy_read_timeout 86400;
    }
}
```

**9.3 启用站点并重载**

```bash
# 创建软链接启用站点
sudo ln -s /etc/nginx/sites-available/meal-refund-approval /etc/nginx/sites-enabled/

# 测试配置
sudo nginx -t

# 重载 Nginx
sudo systemctl reload nginx
```

访问地址：`http://refund.example.internal`

### 10. 更新部署

更新代码时，按以下步骤操作：

```bash
cd /opt/meal-refund-approval

# 1. 停止服务
sudo systemctl stop meal-refund-approval

# 2. 更新代码（选择其一）
# 方式 A：scp 上传
# 在本地执行：scp -r app user@server:/opt/meal-refund-approval/

# 方式 B：rsync（增量同步）
# rsync -avz --exclude='.venv' --exclude='logs' --exclude='.env' ./ user@server:/opt/meal-refund-approval/

# 方式 C：git pull（如使用 git 管理）
# git pull origin main

# 3. 更新依赖（如有变化）
.venv/bin/pip install -r requirements.txt --index-url https://pypi.tuna.tsinghua.edu.cn/simple

# 4. 运行测试
.venv/bin/python -m pytest test/ -q

# 5. 重启服务
sudo systemctl start meal-refund-approval
```

### 11. 安全建议

**⚠️ 重要提示**：本工具处理的 Excel 表格包含**身份证号码**等个人敏感信息，请注意数据安全。

**推荐部署方式：**

1. **内网/VPN 部署**（首选）：仅允许公司内网访问，不暴露到公网
2. **防火墙源 IP 限制**：使用 ufw 限制访问来源
   ```bash
   # 仅允许特定 IP 段访问
   sudo ufw allow from 10.0.0.0/8 to any port 8501
   ```
3. **Nginx 基础认证**：添加 HTTP Basic Auth
   ```bash
   # 安装 apache2-utils
   sudo apt install -y apache2-utils
   
   # 创建密码文件（按提示输入密码）
   sudo htpasswd -c /etc/nginx/.htpasswd admin
   
   # 在 Nginx 配置的 location 块中添加：
   # auth_basic "Restricted Access";
   # auth_basic_user_file /etc/nginx/.htpasswd;
   
   # 重载 Nginx
   sudo systemctl reload nginx
   ```

**不要**将本工具直接暴露到公网且无任何访问控制。

### 12. 故障排查

**问题：端口 8501 被占用**

```bash
# 查看占用端口的进程
ss -ltnp | grep 8501

# 停止占用进程或修改 Streamlit 端口
```

**问题：启动失败，提示 .env 缺失**

```bash
# 确认 .env 文件存在且配置正确
ls -la /opt/meal-refund-approval/.env
cat /opt/meal-refund-approval/.env
```

**问题：Token 缓存异常**

```bash
# 删除 token 缓存文件（应用会自动重新获取）
rm /opt/meal-refund-approval/.token_cache.json
sudo systemctl restart meal-refund-approval
```

**问题：99991672 权限错误**

参见 [常见问题 #1](#1-查询发起人失败99991672-错误)，确保已在飞书开放平台开通 `contact:user.id:readonly` 权限并发布新版本。

**问题：systemd 服务无法启动**

```bash
# 查看详细错误日志
journalctl -u meal-refund-approval -n 50

# 常见原因：
# - User= 指定的用户不存在
# - WorkingDirectory 或 ExecStart 路径错误
# - .venv/bin/python 不存在
```

## 使用说明

### 步骤 1：上传表格

点击"选择退返餐费明细表"按钮，上传 `.xls` 或 `.xlsx` 格式的 Excel 文件。

系统会自动解析表格内容，包括：
- 标题行
- 明细数据（序号、姓名、身份证号码、金额）
- 合计金额（数字 + 中文大写）
- 制表人信息

### 步骤 2：预览数据

- **审批标题**：可编辑，默认为表格标题
- **明细数据**：以表格形式展示所有行
- **合计金额**：显示数字和中文大写
- **警告信息**：如有数据异常会显示警告

**查询发起人：**

输入审批发起人的手机号，点击"查询发起人"按钮。系统会通过飞书 API 查找对应的员工信息。

- 成功：显示匹配的用户 ID
- 失败：提示未找到或权限不足

### 步骤 3：备注

填写可选的备注信息，默认为"制表：{制表人姓名}"。

### 步骤 4：操作

- **预览审批流程**：查看审批节点和审批人
- **提交审批**：上传 Excel 附件并提交审批实例

提交成功后会显示：
- 实例编号
- 审批详情链接（可点击跳转）

## 常见问题

### 1. 查询发起人失败：99991672 错误

**错误信息：**
```
权限不足，需要 contact:user.id:readonly 权限
```

**解决方法：**

1. 打开飞书开放平台授权页面：
   [https://open.feishu.cn/app/cli_aabb96a4a9b99bee/auth?q=contact:user.id:readonly&op_from=openapi&token_type=tenant](https://open.feishu.cn/app/cli_aabb96a4a9b99bee/auth?q=contact:user.id:readonly&op_from=openapi&token_type=tenant)

2. 勾选 `contact:user.id:readonly` 权限

3. 点击"批量开通"

4. **重要**：创建新版本并发布（权限变更后需要发布新版本才生效）

5. 等待几分钟后重试

### 2. Excel 模板格式说明

退返餐费明细表应包含以下结构：

**第 1 行（标题行）：**
```
退返餐费明细表
```

**第 2 行（表头）：**
| 序号 | 姓名 | 身份证号码 | 金额 |
|------|------|------------|------|
| 1    | 张三 | 110101199001011234 | 500 |
| 2    | 李四 | 110101199002022345 | 600 |

**最后一行（合计行）：**
```
总计 | | | 1100
```

**注意事项：**
- 表头必须为"序号"、"姓名"、"身份证号码"、"金额"
- 金额列应为数字格式
- 合计行第一列为"总计"
- 系统会自动识别中文大写金额（如"壹仟壹佰元整"）

### 3. 字段映射说明

Excel 数据会映射到飞书审批表单的以下字段：

| Excel 列 | 飞书表单字段 | Widget ID |
|----------|--------------|-----------|
| 序号 | 序号 | `widget17811629865160001` |
| 姓名 | 姓名 | `widget17920477234030001` |
| 身份证号码 | 身份证号码 | `widget17920477572400001` |
| 金额 | 金额 | `widget17822958110070001` |
| Excel 文件 | 附件 | `widget16496686388470001` |
| 备注 | 备注 | `widget17811635300840001` |

**明细列表字段：**
- 列表组件：`widget16496685286700001`（退返餐费明细）
- 标题组件：`widget17811628753290001`

### 4. 提交后如何查看审批进度？

提交成功后，页面会显示：
- **实例编号**：如 `instance_code_xxxxx`
- **查看链接**：点击可跳转到飞书审批详情页

也可以在飞书审批应用中搜索实例编号查看进度。

### 5. 重复提交保护

系统会防止重复提交：
- 提交成功后，"提交审批"按钮会变为禁用状态
- 如需重新提交，请刷新页面或重新上传文件

## 涉及的飞书 API

| 接口 | 用途 |
|------|------|
| `POST /auth/v3/tenant_access_token/internal/` | 获取访问令牌 |
| `GET /contact/v3/users/batch_get_id` | 通过手机号查询用户 ID |
| `POST /drive/v1/medias/upload_all` | 上传附件文件 |
| `GET /approval/v4/approvals/{code}` | 获取审批定义信息 |
| `POST /approval/v4/instances/preview` | 预览审批流程 |
| `POST /approval/v4/instances` | 创建审批实例 |
| `GET /approval/v4/instances/{code}` | 获取实例详情 |

## 技术栈

- **前端**：Streamlit 1.33+
- **后端**：Python 3.12
- **Excel 解析**：openpyxl（.xlsx）、xlrd（.xls）
- **HTTP 客户端**：requests
- **环境管理**：python-dotenv

## 许可证

内部工具，仅供公司财务部门使用。

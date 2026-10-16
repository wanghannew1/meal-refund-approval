# 问题排查：飞书欢迎语显示 open_id 而非姓名

## 元信息

- 日期：2026-10-16
- 项目：MealRefundApproval（飞书审批提交工具，Streamlit）
- 关联 issue：[#IKM93H](https://gitee.com/wanghannew1/meal-refund-approval/issues/IKM93H)
- 影响功能：手机号登录后的欢迎语姓名显示

## 现象

用户使用手机号 `18843112066` 登录本工具，欢迎语显示为「欢迎，`ou_83d9f792f9b78c95aca502d37cb92c59`」（open_id）而非姓名。

对比另一工具 `/home/ubuntu/coding/submit_approval_feishu/submit_approval_feishu`，同一手机号登录后显示「欢迎，王涵」。

两工具读取的飞书用户身份信息一致，表现却不同。

## 排查过程与根因

### 1. 对比应用凭证

对比两工具 `.env` 发现并非同一飞书应用：

- 本工具：`FEISHU_APP_ID=cli_aabb96a4a9b99bee`
- 对方工具：`FEISHU_APP_ID=cli_aab0fde2943a5bda`

两者在飞书开放平台开通的权限不同。

### 2. 对比取姓名代码

- 本工具 `get_user_name` 原调用：`GET /open-apis/contact/v3/users/{open_id}?user_id_type=open_id`（获取单个用户信息）
- 对方工具使用 lark_oapi SDK 的 `contact.v3.user.basic_batch`，即 `POST /open-apis/contact/v3/users/basic_batch`（获取用户基础信息）

### 3. 两应用 × 两接口实测矩阵

使用同一 open_id 查询，结果如下：

| 接口 | `cli_aabb96a4a9b99bee`（本工具） | `cli_aab0fde2943a5bda`（对方） |
| --- | --- | --- |
| `GET /contact/v3/users/:user_id` | 99991672 无权限 | code=0 但 name=None |
| `POST /contact/v3/users/basic_batch` | 99991672 无权限（开通后正常） | code=0，name=王涵 |

`99991672` 报错原文：

- `GET` 接口：缺 `[contact:contact.base:readonly, contact:contact:access_as_app, contact:contact:readonly, contact:contact:readonly_as_app]` 任一权限
- `basic_batch` 接口：缺 `[contact:user.basic_profile:readonly]`

关键发现：`GET` 接口即使调用成功（code=0）也拿不到姓名字段（name=None），姓名需要额外的字段级权限；`basic_batch` 只要 `contact:user.basic_profile:readonly` 就稳定返回姓名。这正是对方工具选择 `basic_batch` 的原因。

### 4. 回退逻辑触发

姓名查不到时 `get_user_name` 按设计返回 `None`，UI 回退显示 open_id（界面提示「未能获取飞书姓名（可能缺少权限），已显示 open_id」）。因此用户看到的是 open_id。

### 5. 额外发现：open_id 跨应用隔离

飞书 open_id 按应用隔离，同一用户在不同应用下 open_id 不同。`ou_83d9f792f9b78c95aca502d37cb92c59` 属于对方应用的命名空间，拿本工具应用的 token 查询它会报 `99992361 open_id cross app`；本工具自身链路解析出的 open_id 是 `ou_27a9f8182c5f068d68d9201d9ecf6b02`。

## 解决方案

### 代码改动（`app/feishu_api.py`）

`get_user_name` 改调 `POST /open-apis/contact/v3/users/basic_batch`：

- 请求参数：`user_id_type=open_id`
- 请求体：`{"user_ids": [open_id]}`
- 解析响应：`data.users[0].name`
- 任何失败（权限错误、空结果、网络错误）仍返回 `None` 优雅回退，不抛异常

### 界面改动（`app/app.py`）

缺权限提示文案中的权限名同步改为 `contact:user.basic_profile:readonly`。

### 测试改动（`test/test_feishu_api.py`）

`TestGetUserName` 全部用例改为 mock 新接口，新增「users 列表为空返回 None」用例，共 6 个用例。

### 配置改动

在飞书开放平台为 `cli_aabb96a4a9b99bee` 开通 `contact:user.basic_profile:readonly`。

授权链接：https://open.feishu.cn/app/cli_aabb96a4a9b99bee/auth?q=contact:user.basic_profile:readonly&op_from=openapi&token_type=tenant

已开通。

## 验证记录

- 全量测试：146 个全部通过（`python -m pytest test/ -q`）
- 实测完整登录链（本应用凭证）：`lookup_user_by_mobile("18843112066")` → `ou_27a9f8182c5f068d68d9201d9ecf6b02` → `get_user_name(...)` → 返回「王涵」

## 经验教训

- 飞书 open_id 是应用维度的用户标识，跨应用查询报 `99992361 open_id cross app`；排查「同一用户两工具表现不一致」时应先对比两边 `.env` 的应用凭证。
- 取用户姓名优先用 `POST /contact/v3/users/basic_batch`：所需权限最小（`contact:user.basic_profile:readonly`），且不受字段级权限影响；`GET /contact/v3/users/:user_id` 适合取完整资料但姓名字段常因字段级权限返回空。
- 业务代码对姓名查询失败应保持优雅回退（显示 open_id 并提示缺失权限），避免登录链路整体失败。

## 关联链接

- Issue：[#IKM93H](https://gitee.com/wanghannew1/meal-refund-approval/issues/IKM93H)
- 本仓库改动文件：
  - `app/feishu_api.py`
  - `app/app.py`
  - `test/test_feishu_api.py`

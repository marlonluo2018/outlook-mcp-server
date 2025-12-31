# Win32com vs Microsoft Graph API 跨平台对比

## 概述

本文档对比了使用Win32com和Microsoft Graph API两种方式访问Outlook邮件的差异，以及如何实现跨平台支持。

## 核心差异

| 特性 | Win32com | Microsoft Graph API |
|------|----------|-------------------|
| **平台支持** | ❌ 仅Windows | ✅ 跨平台 (Windows/macOS/Linux) |
| **依赖** | 需要本地Outlook安装 | ✅ 仅需网络连接 |
| **授权方式** | 本地COM接口 | OAuth 2.0 |
| **性能** | 本地访问，快 | 云端API，稍慢 |
| **功能覆盖** | 完整Outlook功能 | 大部分功能 |
| **部署难度** | 简单 | 需要授权配置 |
| **管理员权限** | 不需要 | 不需要 (设备代码流程) |

## 授权方式对比

### Win32com
```python
import win32com.client

outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
```

**优点**:
- 无需授权
- 直接访问本地Outlook
- 性能极佳

**缺点**:
- 仅限Windows
- 需要Outlook运行
- 无法在服务器环境使用

### Microsoft Graph API (设备代码流程)
```python
from outlook_graph_api import OutlookGraphAPI

# 1. 获取访问令牌 (运行 graph_api_auth_local.py)
access_token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6..."

# 2. 创建API客户端
outlook = OutlookGraphAPI(access_token)

# 3. 使用API
emails = outlook.list_recent_emails(days=7, top=10)
```

**优点**:
- 跨平台支持
- 无需本地Outlook
- 适合云部署
- 无需管理员权限

**缺点**:
- 需要网络连接
- 令牌有有效期 (1小时)
- 性能略慢于本地访问

## 功能映射表

| Win32com功能 | Graph API等效功能 | 说明 |
|-------------|------------------|------|
| `inbox.Items` | `list_recent_emails()` | 获取邮件列表 |
| `Items.Find()` | `search_emails_by_subject()` | 搜索邮件 |
| `MailItem.Reply()` | `reply_to_email()` | 回复邮件 |
| `MailItem.Forward()` | `forward_email()` | 转发邮件 |
| `CreateItem(0)` | `compose_email()` | 创建新邮件 |
| `Folders.Add()` | `create_folder()` | 创建文件夹 |
| `MailItem.Move()` | `move_email()` | 移动邮件 |
| `MailItem.Delete()` | `delete_email()` | 删除邮件 |

## 跨平台实现示例

### 1. 获取最近邮件

**Win32com (仅Windows)**:
```python
import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(6)

date_filter = (datetime.now() - timedelta(days=7)).strftime('%m/%d/%Y %H:%M')
emails = inbox.Items.Restrict(f"[ReceivedTime] >= '{date_filter}'")

for email in emails:
    print(email.Subject)
```

**Graph API (跨平台)**:
```python
from outlook_graph_api import OutlookGraphAPI

outlook = OutlookGraphAPI(access_token)
emails = outlook.list_recent_emails(days=7, top=10)

for email in emails:
    print(email['subject'])
```

### 2. 搜索邮件

**Win32com (仅Windows)**:
```python
emails = inbox.Items.Find("[Subject] LIKE '%Red Hat%'")
```

**Graph API (跨平台)**:
```python
emails = outlook.search_emails_by_subject('Red Hat', days=7)
```

### 3. 批量转发邮件

**Win32com (仅Windows)**:
```python
for recipient in recipient_list:
    email.Forward()
    email.To = recipient
    email.Send()
```

**Graph API (跨平台)**:
```python
results = outlook.batch_forward_emails(
    email_id,
    recipient_list,
    batch_size=500,
    custom_text="请查收"
)
```

## 部署场景对比

### 场景1: Windows桌面应用

**Win32com** - 推荐
- ✅ 性能最佳
- ✅ 无需网络
- ✅ 简单直接

### 场景2: macOS开发环境

**Graph API** - 唯一选择
- ✅ 跨平台支持
- ✅ 无需Windows
- ❌ 需要网络

### 场景3: 云服务器部署

**Graph API** - 唯一选择
- ✅ 无需GUI
- ✅ 可容器化
- ✅ 可扩展

### 场景4: 混合环境

**Graph API** - 推荐
- ✅ 统一代码库
- ✅ 简化维护
- ✅ 一致体验

## 性能对比

### 测试场景: 获取100封邮件

| 方式 | 耗时 | 说明 |
|------|------|------|
| Win32com | ~0.5秒 | 本地访问，极快 |
| Graph API | ~2-3秒 | 网络请求，稍慢 |

### 测试场景: 搜索1000封邮件

| 方式 | 耗时 | 说明 |
|------|------|------|
| Win32com | ~1-2秒 | 本地搜索 |
| Graph API | ~5-10秒 | 服务器端搜索 |

## 限制对比

### Win32com限制
- ❌ 仅Windows平台
- ❌ 需要Outlook运行
- ❌ 无法在服务器环境使用
- ❌ 不支持容器化部署

### Graph API限制
- ⚠️ 需要网络连接
- ⚠️ 访问令牌有有效期
- ⚠️ API调用频率限制
- ⚠️ 某些高级功能可能不支持

## 迁移建议

### 从Win32com迁移到Graph API

1. **获取授权**
   ```bash
   python graph_api_auth_local.py
   ```

2. **替换导入**
   ```python
   # 旧代码
   import win32com.client
   
   # 新代码
   from outlook_graph_api import OutlookGraphAPI
   ```

3. **替换API调用**
   ```python
   # 旧代码
   outlook = win32com.client.Dispatch('Outlook.Application')
   inbox = outlook.GetNamespace('MAPI').GetDefaultFolder(6)
   
   # 新代码
   outlook = OutlookGraphAPI(access_token)
   ```

4. **处理数据结构差异**
   ```python
   # Win32com返回对象
   email.Subject
   email.From
   
   # Graph API返回字典
   email['subject']
   email['from']['emailAddress']['name']
   ```

## 最佳实践

### 1. 令牌管理
```python
import json

# 保存令牌
with open('token.json', 'w') as f:
    json.dump({
        'access_token': access_token,
        'expires_at': expires_at
    }, f)

# 加载令牌
with open('token.json', 'r') as f:
    token_data = json.load(f)
    
# 检查令牌是否过期
if datetime.now() > datetime.fromtimestamp(token_data['expires_at']):
    # 重新获取令牌
    pass
```

### 2. 错误处理
```python
try:
    emails = outlook.list_recent_emails(days=7)
except Exception as e:
    if '401' in str(e):
        # 令牌过期，重新授权
        print('令牌过期，请重新授权')
    elif '429' in str(e):
        # API限流，等待重试
        print('API限流，请稍后重试')
    else:
        print(f'错误: {e}')
```

### 3. 跨平台兼容性检查
```python
import platform

def get_outlook_client(access_token=None):
    if platform.system() == 'Windows':
        try:
            import win32com.client
            return Win32OutlookClient()
        except:
            pass
    
    if access_token:
        return OutlookGraphAPI(access_token)
    
    raise Exception('无法创建Outlook客户端')
```

## 总结

### 何时使用Win32com
- ✅ Windows桌面应用
- ✅ 需要最佳性能
- ✅ 无需网络连接
- ✅ 本地数据处理

### 何时使用Graph API
- ✅ 跨平台需求
- ✅ 云部署
- ✅ 容器化应用
- ✅ 无需本地Outlook

### 推荐方案
对于新项目，**推荐使用Microsoft Graph API**，原因：
1. 跨平台支持，未来扩展性更好
2. 云原生，适合现代部署方式
3. 统一API，简化开发维护
4. 社区支持活跃，文档完善

对于现有Win32com项目，可以：
1. 保持Win32com用于Windows环境
2. 使用Graph API扩展到其他平台
3. 逐步迁移到Graph API

## 相关文件

- [graph_api_auth_local.py](file:///c:/Project/outlook-mcp-server/graph_api_auth_local.py) - 设备代码授权
- [outlook_graph_api.py](file:///c:/Project/outlook-mcp-server/outlook_graph_api.py) - Graph API封装
- [graph_api_auth.py](file:///c:/Project/outlook-mcp-server/graph_api_auth.py) - 传统授权方式

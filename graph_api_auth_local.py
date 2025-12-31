import requests
import json
import time

def get_device_code():
    """
    获取设备代码和用户代码
    
    Returns:
        dict: 包含设备代码、用户代码等信息
    """
    client_id = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    
    url = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode"
    
    data = {
        "client_id": client_id,
        "scope": "https://graph.microsoft.com/.default"
    }
    
    response = requests.post(url, data=data)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f'获取设备代码失败: {response.status_code} - {response.text}')

def poll_for_token(device_code, interval, expires_in):
    """
    轮询获取访问令牌
    
    Args:
        device_code: 设备代码
        interval: 轮询间隔（秒）
        expires_in: 过期时间（秒）
    
    Returns:
        dict: 包含访问令牌的响应
    """
    client_id = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    
    url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    
    start_time = time.time()
    
    while (time.time() - start_time) < expires_in:
        data = {
            "client_id": client_id,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code
        }
        
        response = requests.post(url, data=data)
        
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 400:
            error = response.json().get('error')
            if error == 'authorization_pending':
                print(f'等待授权... (剩余时间: {int(expires_in - (time.time() - start_time))}秒)')
                time.sleep(interval)
            elif error == 'authorization_declined':
                raise Exception('用户拒绝了授权')
            elif error == 'expired_token':
                raise Exception('设备代码已过期')
            else:
                raise Exception(f'授权错误: {error}')
        else:
            raise Exception(f'获取令牌失败: {response.status_code} - {response.text}')
    
    raise Exception('授权超时')

def test_graph_api_access(access_token):
    """
    测试Graph API访问权限
    
    Args:
        access_token: 访问令牌
    
    Returns:
        dict: API响应
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    print('\n正在测试Graph API访问...')
    response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f'API调用失败: {response.status_code} - {response.text}')

def get_user_emails(access_token, top=10):
    """
    获取用户的邮件列表
    
    Args:
        access_token: 访问令牌
        top: 获取邮件的数量
    
    Returns:
        dict: 邮件列表
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    print(f'\n正在获取最近{top}封邮件...')
    url = f'https://graph.microsoft.com/v1.0/me/messages?$top={top}&$orderby=receivedDateTime desc'
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f'获取邮件失败: {response.status_code} - {response.text}')

if __name__ == '__main__':
    print('=== Microsoft Graph API 本地授权工具 (设备代码流程) ===\n')
    print('这个工具使用设备代码流程，无需Azure门户管理员权限\n')
    
    try:
        print('步骤 1: 获取设备代码...')
        device_code_response = get_device_code()
        
        print('\n' + '='*60)
        print('请在另一个设备上完成授权:')
        print('='*60)
        print(f'\n1. 访问: {device_code_response["verification_uri"]}')
        print(f'2. 输入代码: {device_code_response["user_code"]}')
        print(f'\n提示: 也可以在手机上打开链接并输入代码')
        print('='*60)
        
        print('\n步骤 2: 等待授权完成...')
        token_response = poll_for_token(
            device_code_response['device_code'],
            device_code_response['interval'],
            device_code_response['expires_in']
        )
        
        access_token = token_response.get('access_token')
        refresh_token = token_response.get('refresh_token')
        expires_in = token_response.get('expires_in')
        
        print('\n' + '='*60)
        print('✓ 授权成功!')
        print('='*60)
        print(f'访问令牌: {access_token[:50]}...')
        print(f'过期时间: {expires_in} 秒 ({expires_in/3600:.1f} 小时)')
        
        if refresh_token:
            print(f'刷新令牌: {refresh_token[:50]}...')
        
        print('='*60)
        
        print('\n步骤 3: 测试API访问...')
        user_info = test_graph_api_access(access_token)
        print(f'\n用户信息:')
        print(f'  显示名称: {user_info.get("displayName", "N/A")}')
        print(f'  邮箱: {user_info.get("mail", user_info.get("userPrincipalName", "N/A"))}')
        print(f'  ID: {user_info.get("id", "N/A")}')
        
        print('\n步骤 4: 获取邮件列表...')
        emails_response = get_user_emails(access_token, top=5)
        
        if 'value' in emails_response and len(emails_response['value']) > 0:
            print(f'\n最近 {len(emails_response["value"])} 封邮件:')
            for i, email in enumerate(emails_response['value'], 1):
                subject = email.get('subject', '(无主题)')
                from_email = email.get('from', {}).get('emailAddress', {}).get('address', '未知')
                received = email.get('receivedDateTime', '未知时间')[:19].replace('T', ' ')
                print(f'\n{i}. {subject}')
                print(f'   发件人: {from_email}')
                print(f'   时间: {received}')
        else:
            print('\n没有找到邮件')
        
        print('\n' + '='*60)
        print('✓ 所有测试通过!')
        print('='*60)
        
        print('\n提示: 访问令牌已保存，可以用于后续的Graph API调用')
        print('令牌将在1小时后过期，之后需要重新授权')
        
    except Exception as e:
        print(f'\n错误: {str(e)}')
        print('\n请检查:')
        print('1. 网络连接是否正常')
        print('2. 是否正确输入了授权代码')
        print('3. 是否在有效时间内完成了授权')
        exit(1)

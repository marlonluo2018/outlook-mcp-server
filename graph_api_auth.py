import requests
import webbrowser
import json
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import threading
import time

class AuthCallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        query = urlparse(self.path).query
        params = parse_qs(query)
        
        if 'code' in params:
            self.server.auth_code = params['code'][0]
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wresponse = self.wfile
            self.wfile.write(b'<html><body><h1>Authentication successful!</h1><p>You can close this window.</p></body></html>')
            print("Authorization code received successfully!")
        else:
            self.send_response(400)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b'<html><body><h1>Authentication failed!</h1></body></html>')
            print("Authorization code not found in callback")

def get_graph_api_auth_code(client_id, tenant_id, redirect_uri='http://localhost:8080'):
    """
    获取Microsoft Graph API的授权码
    
    Args:
        client_id: Azure AD应用程序的客户端ID
        tenant_id: Azure AD租户ID
        redirect_uri: 回调URI (默认: http://localhost:8080)
    
    Returns:
        str: 授权码
    """
    
    scopes = [
        'https://graph.microsoft.com/.default'
    ]
    
    auth_url = (
        f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize?'
        f'client_id={client_id}&'
        f'response_type=code&'
        f'redirect_uri={redirect_uri}&'
        f'scope={"%20".join(scopes)}&'
        f'response_mode=query'
    )
    
    print(f'正在打开浏览器进行授权...')
    print(f'授权URL: {auth_url}')
    webbrowser.open(auth_url)
    
    server = HTTPServer(('localhost', 8080), AuthCallbackHandler)
    server.auth_code = None
    server.timeout = 300
    
    print('等待授权回调... (端口8080)')
    
    start_time = time.time()
    while server.auth_code is None and (time.time() - start_time) < 300:
        server.handle_request()
    
    if server.auth_code:
        return server.auth_code
    else:
        raise Exception('授权超时或失败')

def get_access_token(client_id, client_secret, tenant_id, auth_code, redirect_uri='http://localhost:8080'):
    """
    使用授权码获取访问令牌
    
    Args:
        client_id: Azure AD应用程序的客户端ID
        client_secret: Azure AD应用程序的客户端密钥
        tenant_id: Azure AD租户ID
        auth_code: 授权码
        redirect_uri: 回调URI
    
    Returns:
        dict: 包含访问令牌的响应
    """
    token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    
    data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'code': auth_code,
        'redirect_uri': redirect_uri,
        'grant_type': 'authorization_code'
    }
    
    response = requests.post(token_url, data=data)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f'获取访问令牌失败: {response.status_code} - {response.text}')

def get_access_token_client_credentials(client_id, client_secret, tenant_id):
    """
    使用客户端凭据流程获取访问令牌 (适用于后台服务)
    
    Args:
        client_id: Azure AD应用程序的客户端ID
        client_secret: Azure AD应用程序的客户端密钥
        tenant_id: Azure AD租户ID
    
    Returns:
        dict: 包含访问令牌的响应
    """
    token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    
    data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default',
        'grant_type': 'client_credentials'
    }
    
    response = requests.post(token_url, data=data)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f'获取访问令牌失败: {response.status_code} - {response.text}')

def refresh_access_token(client_id, client_secret, tenant_id, refresh_token):
    """
    使用刷新令牌获取新的访问令牌
    
    Args:
        client_id: Azure AD应用程序的客户端ID
        client_secret: Azure AD应用程序的客户端密钥
        tenant_id: Azure AD租户ID
        refresh_token: 刷新令牌
    
    Returns:
        dict: 包含新访问令牌的响应
    """
    token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    
    data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token'
    }
    
    response = requests.post(token_url, data=data)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f'刷新访问令牌失败: {response.status_code} - {response.text}')

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
    
    response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f'API调用失败: {response.status_code} - {response.text}')

if __name__ == '__main__':
    print('=== Microsoft Graph API 授权工具 ===\n')
    
    print('请提供以下信息 (从Azure门户获取):')
    client_id = input('客户端ID (Client ID): ').strip()
    tenant_id = input('租户ID (Tenant ID): ').strip()
    client_secret = input('客户端密钥 (Client Secret): ').strip()
    
    print('\n选择授权方式:')
    print('1. 授权码流程 (需要用户交互)')
    print('2. 客户端凭据流程 (适用于后台服务)')
    
    choice = input('\n请选择 (1 或 2): ').strip()
    
    try:
        if choice == '1':
            print('\n使用授权码流程获取令牌...\n')
            auth_code = get_graph_api_auth_code(client_id, tenant_id)
            print(f'\n授权码: {auth_code}')
            
            token_response = get_access_token(client_id, client_secret, tenant_id, auth_code)
            
        elif choice == '2':
            print('\n使用客户端凭据流程获取令牌...\n')
            token_response = get_access_token_client_credentials(client_id, client_secret, tenant_id)
            
        else:
            print('无效的选择')
            exit(1)
        
        access_token = token_response.get('access_token')
        refresh_token = token_response.get('refresh_token')
        expires_in = token_response.get('expires_in')
        
        print('\n=== 访问令牌获取成功 ===')
        print(f'访问令牌: {access_token[:50]}...')
        print(f'过期时间: {expires_in} 秒')
        
        if refresh_token:
            print(f'刷新令牌: {refresh_token[:50]}...')
        
        print('\n测试Graph API访问...')
        user_info = test_graph_api_access(access_token)
        print(f'用户信息: {json.dumps(user_info, indent=2, ensure_ascii=False)}')
        
        print('\n=== 授权完成 ===')
        
    except Exception as e:
        print(f'\n错误: {str(e)}')
        exit(1)

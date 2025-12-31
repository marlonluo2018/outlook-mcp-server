import requests
import json
from datetime import datetime, timedelta

class OutlookGraphAPI:
    def __init__(self, access_token):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
    
    def list_recent_emails(self, days=7, top=30):
        date_filter = (datetime.utcnow() - timedelta(days=days)).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        url = f'{self.base_url}/me/messages'
        params = {
            '$filter': f'receivedDateTime ge {date_filter}',
            '$top': top,
            '$orderby': 'receivedDateTime desc',
            '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments'
        }
        
        response = requests.get(url, headers=self.headers, params=params)
        
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            raise Exception(f'获取邮件失败: {response.status_code} - {response.text}')
    
    def search_emails_by_subject(self, search_term, days=7):
        date_filter = (datetime.utcnow() - timedelta(days=days)).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        url = f'{self.base_url}/me/messages'
        params = {
            '$filter': f"contains(subject, '{search_term}') and receivedDateTime ge {date_filter}",
            '$orderby': 'receivedDateTime desc',
            '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments'
        }
        
        response = requests.get(url, headers=self.headers, params=params)
        
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            raise Exception(f'搜索邮件失败: {response.status_code} - {response.text}')
    
    def search_emails_by_sender(self, sender_name, days=7):
        date_filter = (datetime.utcnow() - timedelta(days=days)).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        url = f'{self.base_url}/me/messages'
        params = {
            '$filter': f"contains(from/emailAddress/name, '{sender_name}') and receivedDateTime ge {date_filter}",
            '$orderby': 'receivedDateTime desc',
            '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments'
        }
        
        response = requests.get(url, headers=self.headers, params=params)
        
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            raise Exception(f'搜索邮件失败: {response.status_code} - {response.text}')
    
    def search_emails_by_body(self, search_term, days=7):
        date_filter = (datetime.utcnow() - timedelta(days=days)).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        url = f'{self.base_url}/me/messages'
        params = {
            '$search': f'"{search_term}"',
            '$filter': f'receivedDateTime ge {date_filter}',
            '$orderby': 'receivedDateTime desc',
            '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments'
        }
        
        response = requests.get(url, headers=self.headers, params=params)
        
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            raise Exception(f'搜索邮件失败: {response.status_code} - {response.text}')
    
    def get_email_details(self, email_id):
        url = f'{self.base_url}/me/messages/{email_id}'
        params = {
            '$select': 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,body,hasAttachments,attachments'
        }
        
        response = requests.get(url, headers=self.headers, params=params)
        
        if response.status_code == 200:
            return response.json()
        else:
            raise Exception(f'获取邮件详情失败: {response.status_code} - {response.text}')
    
    def compose_email(self, to_recipients, subject, body, cc_recipients=None, bcc_recipients=None):
        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": body
                },
                "toRecipients": [
                    {"emailAddress": {"address": email}}
                    for email in to_recipients
                ]
            }
        }
        
        if cc_recipients:
            email_data["message"]["ccRecipients"] = [
                {"emailAddress": {"address": email}}
                for email in cc_recipients
            ]
        
        if bcc_recipients:
            email_data["message"]["bccRecipients"] = [
                {"emailAddress": {"address": email}}
                for email in bcc_recipients
            ]
        
        url = f'{self.base_url}/me/sendMail'
        response = requests.post(url, headers=self.headers, json=email_data)
        
        if response.status_code == 202:
            return {"status": "success", "message": "邮件发送成功"}
        else:
            raise Exception(f'发送邮件失败: {response.status_code} - {response.text}')
    
    def reply_to_email(self, email_id, reply_text, reply_all=False):
        endpoint = 'replyAll' if reply_all else 'reply'
        url = f'{self.base_url}/me/messages/{email_id}/{endpoint}'
        
        reply_data = {
            "message": {
                "body": {
                    "contentType": "HTML",
                    "content": reply_text
                }
            }
        }
        
        response = requests.post(url, headers=self.headers, json=reply_data)
        
        if response.status_code == 202:
            return {"status": "success", "message": "回复成功"}
        else:
            raise Exception(f'回复邮件失败: {response.status_code} - {response.text}')
    
    def forward_email(self, email_id, to_recipients, comment=""):
        url = f'{self.base_url}/me/messages/{email_id}/forward'
        
        forward_data = {
            "message": {
                "toRecipients": [
                    {"emailAddress": {"address": email}}
                    for email in to_recipients
                ]
            }
        }
        
        if comment:
            forward_data["message"]["body"] = {
                "contentType": "HTML",
                "content": comment
            }
        
        response = requests.post(url, headers=self.headers, json=forward_data)
        
        if response.status_code == 202:
            return {"status": "success", "message": "转发成功"}
        else:
            raise Exception(f'转发邮件失败: {response.status_code} - {response.text}')
    
    def list_folders(self):
        url = f'{self.base_url}/me/mailFolders'
        
        response = requests.get(url, headers=self.headers)
        
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            raise Exception(f'获取文件夹失败: {response.status_code} - {response.text}')
    
    def create_folder(self, folder_name, parent_folder_id=None):
        if parent_folder_id:
            url = f'{self.base_url}/me/mailFolders/{parent_folder_id}/childFolders'
        else:
            url = f'{self.base_url}/me/mailFolders'
        
        folder_data = {
            "displayName": folder_name
        }
        
        response = requests.post(url, headers=self.headers, json=folder_data)
        
        if response.status_code == 201:
            return response.json()
        else:
            raise Exception(f'创建文件夹失败: {response.status_code} - {response.text}')
    
    def move_email(self, email_id, destination_folder_id):
        url = f'{self.base_url}/me/messages/{email_id}/move'
        
        move_data = {
            "destinationId": destination_folder_id
        }
        
        response = requests.post(url, headers=self.headers, json=move_data)
        
        if response.status_code == 201:
            return response.json()
        else:
            raise Exception(f'移动邮件失败: {response.status_code} - {response.text}')
    
    def delete_email(self, email_id):
        url = f'{self.base_url}/me/messages/{email_id}'
        
        response = requests.delete(url, headers=self.headers)
        
        if response.status_code == 204:
            return {"status": "success", "message": "邮件删除成功"}
        else:
            raise Exception(f'删除邮件失败: {response.status_code} - {response.text}')
    
    def batch_forward_emails(self, email_id, recipient_list, batch_size=500, custom_text=""):
        results = {
            "total_recipients": len(recipient_list),
            "batches": [],
            "successful": 0,
            "failed": 0
        }
        
        for i in range(0, len(recipient_list), batch_size):
            batch = recipient_list[i:i + batch_size]
            batch_num = (i // batch_size) + 1
            total_batches = (len(recipient_list) + batch_size - 1) // batch_size
            
            try:
                self.forward_email(email_id, batch, custom_text)
                results["batches"].append({
                    "batch": batch_num,
                    "total_batches": total_batches,
                    "recipients": len(batch),
                    "status": "success"
                })
                results["successful"] += len(batch)
                print(f'批次 {batch_num}/{total_batches} 发送成功 ({len(batch)} 封邮件)')
            except Exception as e:
                results["batches"].append({
                    "batch": batch_num,
                    "total_batches": total_batches,
                    "recipients": len(batch),
                    "status": "failed",
                    "error": str(e)
                })
                results["failed"] += len(batch)
                print(f'批次 {batch_num}/{total_batches} 发送失败: {str(e)}')
        
        return results

if __name__ == '__main__':
    print('=== Outlook Graph API 跨平台演示 ===\n')
    print('请先运行 python graph_api_auth_local.py 获取访问令牌\n')
    
    access_token = input('请输入访问令牌 (或按Enter使用演示模式): ').strip()
    
    if not access_token:
        print('\n演示模式: 展示API功能说明\n')
        print('可用功能:')
        print('1. list_recent_emails(days=7, top=30) - 获取最近邮件')
        print('2. search_emails_by_subject(search_term, days=7) - 按主题搜索')
        print('3. search_emails_by_sender(sender_name, days=7) - 按发件人搜索')
        print('4. search_emails_by_body(search_term, days=7) - 按正文搜索')
        print('5. get_email_details(email_id) - 获取邮件详情')
        print('6. compose_email(to_recipients, subject, body) - 发送邮件')
        print('7. reply_to_email(email_id, reply_text, reply_all=False) - 回复邮件')
        print('8. forward_email(email_id, to_recipients, comment="") - 转发邮件')
        print('9. list_folders() - 获取文件夹列表')
        print('10. create_folder(folder_name, parent_folder_id=None) - 创建文件夹')
        print('11. move_email(email_id, destination_folder_id) - 移动邮件')
        print('12. delete_email(email_id) - 删除邮件')
        print('13. batch_forward_emails(email_id, recipient_list, batch_size=500) - 批量转发')
        print('\n✓ 这些功能可以在Windows、macOS、Linux上运行!')
        print('✓ 无需Win32com，仅需网络连接和有效的访问令牌')
        exit(0)
    
    try:
        outlook = OutlookGraphAPI(access_token)
        
        print('\n1. 获取最近5封邮件...')
        emails = outlook.list_recent_emails(days=7, top=5)
        print(f'找到 {len(emails)} 封邮件\n')
        
        for i, email in enumerate(emails, 1):
            print(f'{i}. {email.get("subject", "(无主题)")}')
            print(f'   发件人: {email.get("from", {}).get("emailAddress", {}).get("address", "未知")}')
            print(f'   时间: {email.get("receivedDateTime", "未知时间")[:19].replace("T", " ")}')
            print()
        
        print('\n2. 获取文件夹列表...')
        folders = outlook.list_folders()
        print(f'找到 {len(folders)} 个文件夹\n')
        
        for folder in folders:
            print(f'- {folder.get("displayName", "未知")} (ID: {folder.get("id", "未知")})')
        
        print('\n3. 测试搜索功能...')
        search_results = outlook.search_emails_by_subject('Red Hat', days=7)
        print(f'找到 {len(search_results)} 封包含"Red Hat"的邮件\n')
        
        print('✓ 所有测试通过!')
        print('\n提示: 这个API可以在Windows、macOS、Linux上运行!')
        
    except Exception as e:
        print(f'\n错误: {str(e)}')
        print('\n如果令牌过期，请重新运行: python graph_api_auth_local.py')
        exit(1)

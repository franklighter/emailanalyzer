import win32com.client
from datetime import datetime, timedelta
import re
from collections import Counter, defaultdict
import sys

class OutlookEmailAnalyzer:
    def __init__(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            print("成功连接到Outlook")
        except Exception as e:
            print(f"无法连接到Outlook: {e}")
            sys.exit(1)
    
    def get_date_input(self, prompt):
        """获取用户输入的日期"""
        while True:
            try:
                date_str = input(prompt)
                return datetime.strptime(date_str, "%d-%m-%Y")
            except ValueError:
                print("日期格式错误，请使用DD-MM-YYYY格式")
    
    def list_available_accounts(self):
        """列出所有可用的邮箱账户"""
        try:
            accounts = self.namespace.Accounts
            print("可用的邮箱账户:")
            for i, account in enumerate(accounts, 1):
                try:
                    print(f"  {i}. {account.DisplayName} ({account.SmtpAddress})")
                except:
                    print(f"  {i}. {account.DisplayName}")
            return accounts
        except Exception as e:
            print(f"获取账户列表失败: {e}")
            return None
    
    def get_email_account(self, email_address):
        """获取指定邮箱账户"""
        try:
            accounts = self.namespace.Accounts
            print(f"正在查找邮箱账户: {email_address}")
            
            for account in accounts:
                try:
                    if account.SmtpAddress and account.SmtpAddress.lower() == email_address.lower():
                        print(f"找到匹配的账户: {account.DisplayName}")
                        return account
                except:
                    continue
            
            print(f"未找到匹配的账户，将使用默认账户")
            return None
        except Exception as e:
            print(f"获取邮箱账户时出错: {e}")
            return None
    
    def get_inbox_folders(self, email_address):
        """获取收件箱及其子文件夹"""
        try:
            account = self.get_email_account(email_address)
            if account and hasattr(account, 'DeliveryStore'):
                try:
                    inbox = account.DeliveryStore.GetDefaultFolder(6)  # olFolderInbox = 6
                    print(f"使用账户的收件箱: {account.DisplayName}")
                except:
                    inbox = self.namespace.GetDefaultFolder(6)
                    print("使用默认收件箱")
            else:
                inbox = self.namespace.GetDefaultFolder(6)
                print("使用默认收件箱")
            
            folders = [inbox]
            self._get_subfolders(inbox, folders)
            return folders
        except Exception as e:
            print(f"获取收件箱文件夹时出错: {e}")
            # 尝试使用默认收件箱
            try:
                inbox = self.namespace.GetDefaultFolder(6)
                return [inbox]
            except:
                return []
    
    def _get_subfolders(self, parent_folder, folder_list):
        """递归获取所有子文件夹"""
        try:
            for folder in parent_folder.Folders:
                folder_list.append(folder)
                self._get_subfolders(folder, folder_list)
        except:
            pass
    
    def get_emails_in_date_range(self, folders, start_date, end_date):
        """获取指定日期范围内的邮件"""
        emails = []
        start_str = start_date.strftime("%m/%d/%Y")
        end_str = (end_date + timedelta(days=1)).strftime("%m/%d/%Y")
        
        for folder in folders:
            try:
                print(f"正在读取文件夹: {folder.Name}")
                messages = folder.Items
                messages.Sort("[ReceivedTime]", True)
                
                # 使用过滤器提高性能
                filter_str = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] < '{end_str}'"
                filtered_messages = messages.Restrict(filter_str)
                
                count = 0
                for message in filtered_messages:
                    try:
                        if hasattr(message, 'ReceivedTime') and hasattr(message, 'Subject'):
                            emails.append(message)
                            count += 1
                    except:
                        continue
                print(f"  找到 {count} 封邮件")
            except Exception as e:
                print(f"读取文件夹 {folder.Name} 时出错: {e}")
                continue
        
        return emails
    
    def get_sent_emails_in_date_range(self, email_address, start_date, end_date):
        """获取指定日期范围内的发送邮件"""
        try:
            print("正在获取发送邮件...")
            
            # 方法1: 尝试使用指定账户的发送文件夹
            account = self.get_email_account(email_address)
            sent_folder = None
            
            if account:
                try:
                    if hasattr(account, 'DeliveryStore'):
                        sent_folder = account.DeliveryStore.GetDefaultFolder(5)  # olFolderSentMail = 5
                        print(f"使用账户 {account.DisplayName} 的发送文件夹")
                    else:
                        print("账户没有DeliveryStore属性")
                except Exception as e:
                    print(f"无法访问账户的发送文件夹: {e}")
            
            # 方法2: 如果账户方法失败，使用默认发送文件夹
            if not sent_folder:
                try:
                    sent_folder = self.namespace.GetDefaultFolder(5)
                    print("使用默认发送文件夹")
                except Exception as e:
                    print(f"无法访问默认发送文件夹: {e}")
                    return []
            
            # 方法3: 如果以上都失败，尝试遍历所有文件夹查找发送文件夹
            if not sent_folder:
                try:
                    print("尝试查找发送文件夹...")
                    stores = self.namespace.Stores
                    for store in stores:
                        try:
                            root_folder = store.GetRootFolder()
                            for folder in root_folder.Folders:
                                if "sent" in folder.Name.lower() or "已发送" in folder.Name or "寄件备份" in folder.Name:
                                    sent_folder = folder
                                    print(f"找到发送文件夹: {folder.Name}")
                                    break
                            if sent_folder:
                                break
                        except:
                            continue
                except Exception as e:
                    print(f"查找发送文件夹失败: {e}")
            
            if not sent_folder:
                print("无法找到发送文件夹，跳过发送邮件分析")
                return []
            
            sent_emails = []
            start_str = start_date.strftime("%m/%d/%Y")
            end_str = (end_date + timedelta(days=1)).strftime("%m/%d/%Y")
            
            try:
                messages = sent_folder.Items
                print(f"发送文件夹中总共有 {messages.Count} 封邮件")
                
                # 尝试使用过滤器
                try:
                    filter_str = f"[SentOn] >= '{start_str}' AND [SentOn] < '{end_str}'"
                    filtered_messages = messages.Restrict(filter_str)
                    print(f"过滤后有 {filtered_messages.Count} 封邮件")
                    
                    for message in filtered_messages:
                        try:
                            if hasattr(message, 'SentOn') and hasattr(message, 'Subject'):
                                sent_emails.append(message)
                        except:
                            continue
                
                except Exception as e:
                    print(f"过滤失败，尝试手动遍历: {e}")
                    # 如果过滤器失败，手动遍历
                    count = 0
                    for message in messages:
                        try:
                            if hasattr(message, 'SentOn') and hasattr(message, 'Subject'):
                                msg_date = message.SentOn.date()
                                if start_date.date() <= msg_date <= end_date.date():
                                    sent_emails.append(message)
                            count += 1
                            if count % 100 == 0:
                                print(f"已处理 {count} 封邮件...")
                        except:
                            continue
            
            except Exception as e:
                print(f"读取发送邮件失败: {e}")
                return []
            
            print(f"找到 {len(sent_emails)} 封发送邮件")
            return sent_emails
            
        except Exception as e:
            print(f"获取发送邮件时出现严重错误: {e}")
            print("将跳过发送邮件的相关分析")
            return []
    
    def analyze_read_status(self, emails):
        """分析邮件读取状态"""
        read_count = 0
        unread_count = 0
        
        for email in emails:
            try:
                if email.UnRead:
                    unread_count += 1
                else:
                    read_count += 1
            except:
                continue
        
        total = read_count + unread_count
        if total > 0:
            read_percentage = (read_count / total) * 100
            unread_percentage = (unread_count / total) * 100
        else:
            read_percentage = unread_percentage = 0
        
        return read_count, unread_count, read_percentage, unread_percentage
    
    def find_replied_emails(self, received_emails, sent_emails):
        """查找已回复的邮件"""
        if not sent_emails:
            print("没有发送邮件数据，跳过回复分析")
            return [], 0
            
        replied_emails = []
        same_day_replies = 0
        
        # 创建发送邮件的主题和时间映射
        sent_subjects = {}
        for sent_email in sent_emails:
            try:
                subject = sent_email.Subject.lower() if sent_email.Subject else ""
                # 移除 "re:" 前缀
                clean_subject = re.sub(r'^(re:|回复:|回覆:)\s*', '', subject, flags=re.IGNORECASE)
                sent_date = sent_email.SentOn.date()
                
                if clean_subject not in sent_subjects:
                    sent_subjects[clean_subject] = []
                sent_subjects[clean_subject].append(sent_date)
            except:
                continue
        
        for email in received_emails:
            try:
                if email.UnRead:  # 跳过未读邮件
                    continue
                
                subject = email.Subject.lower() if email.Subject else ""
                clean_subject = re.sub(r'^(re:|回复:|回覆:)\s*', '', subject, flags=re.IGNORECASE)
                received_date = email.ReceivedTime.date()
                
                # 检查是否有回复
                if clean_subject in sent_subjects:
                    replied_emails.append(email)
                    # 检查是否在同一天回复
                    for sent_date in sent_subjects[clean_subject]:
                        if sent_date == received_date:
                            same_day_replies += 1
                            break
            except:
                continue
        
        return replied_emails, same_day_replies
    
    def get_top_senders_and_recipients(self, received_emails, sent_emails):
        """获取前5名发件人和收件人"""
        senders = Counter()
        recipients = Counter()
        
        # 统计发件人
        for email in received_emails:
            try:
                # 尝试多种方式获取发件人信息
                sender = None
                if hasattr(email, 'SenderEmailAddress') and email.SenderEmailAddress:
                    sender = email.SenderEmailAddress
                elif hasattr(email, 'SenderName') and email.SenderName:
                    sender = email.SenderName
                
                if sender:
                    senders[sender] += 1
            except:
                continue
        
        # 统计收件人
        if sent_emails:
            for email in sent_emails:
                try:
                    # 获取所有收件人
                    to_recipients = []
                    cc_recipients = []
                    
                    if hasattr(email, 'To') and email.To:
                        to_recipients = [r.strip() for r in email.To.split(';')]
                    if hasattr(email, 'CC') and email.CC:
                        cc_recipients = [r.strip() for r in email.CC.split(';')]
                    
                    all_recipients = to_recipients + cc_recipients
                    for recipient in all_recipients:
                        if recipient:
                            recipients[recipient] += 1
                except:
                    continue
        
        return senders.most_common(5), recipients.most_common(5)
    
    def classify_emails(self, emails):
        """将邮件分类为三种类型"""
        info_keywords = ['通知', '信息', '更新', '公告', '新闻', 'newsletter', 'notification', 'update', 'info', '通告']
        approval_keywords = ['批准', '审批', '确认', '同意', '授权', 'approve', 'approval', 'authorize', 'confirm', '核准', '签核']
        response_keywords = ['回复', '回应', '反馈', '意见', '建议', 'reply', 'response', 'feedback', 'urgent', '紧急', '请回复', '请回覆']
        
        info_emails = []
        approval_emails = []
        response_emails = []
        
        for email in emails:
            try:
                subject = email.Subject.lower() if email.Subject else ''
                body = ""
                
                # 安全地获取邮件正文
                try:
                    if hasattr(email, 'Body') and email.Body:
                        body = email.Body.lower()
                except:
                    # 如果无法获取正文，只使用主题
                    pass
                
                # 检查是否包含批准关键词
                if any(keyword in subject or keyword in body for keyword in approval_keywords):
                    approval_emails.append(email)
                # 检查是否需要回复
                elif any(keyword in subject or keyword in body for keyword in response_keywords):
                    response_emails.append(email)
                # 默认为信息类
                else:
                    info_emails.append(email)
            except:
                # 如果出错，默认归类为信息类
                info_emails.append(email)
        
        return len(info_emails), len(approval_emails), len(response_emails)
    
    def run_analysis(self):
        """运行完整分析"""
        print("=== Outlook 邮件分析工具 ===\n")
        
        # 显示可用账户
        self.list_available_accounts()
        print()
        
        # 获取用户输入
        start_date = self.get_date_input("请输入开始日期 (DD-MM-YYYY): ")
        end_date = self.get_date_input("请输入结束日期 (DD-MM-YYYY): ")
        email_address = input("请输入邮箱地址: ").strip()
        
        print(f"\n正在分析 {start_date.strftime('%d-%m-%Y')} 到 {end_date.strftime('%d-%m-%Y')} 的邮件...")
        
        # 获取收件箱文件夹
        inbox_folders = self.get_inbox_folders(email_address)
        print(f"找到 {len(inbox_folders)} 个收件箱文件夹")
        
        if not inbox_folders:
            print("无法获取收件箱文件夹，程序退出")
            return
        
        # 获取收到的邮件
        received_emails = self.get_emails_in_date_range(inbox_folders, start_date, end_date)
        print(f"收到的邮件总数: {len(received_emails)}")
        
        # 获取发送的邮件
        sent_emails = self.get_sent_emails_in_date_range(email_address, start_date, end_date)
        print(f"发送的邮件总数: {len(sent_emails)}")
        
        # 分析读取状态
        read_count, unread_count, read_percentage, unread_percentage = self.analyze_read_status(received_emails)
        
        # 查找已回复的邮件
        replied_emails, same_day_replies = self.find_replied_emails(received_emails, sent_emails)
        
        # 获取前5名发件人和收件人
        top_senders, top_recipients = self.get_top_senders_and_recipients(received_emails, sent_emails)
        
        # 分类邮件
        info_count, approval_count, response_count = self.classify_emails(received_emails)
        
        # 打印结果
        self.print_results(
            len(received_emails), read_count, unread_count, read_percentage, unread_percentage,
            len(replied_emails), same_day_replies, top_senders, top_recipients,
            info_count, approval_count, response_count
        )
    
    def print_results(self, total_received, read_count, unread_count, read_percentage, 
                     unread_percentage, replied_count, same_day_replies, top_senders, 
                     top_recipients, info_count, approval_count, response_count):
        """打印分析结果"""
        print("\n" + "="*50)
        print("邮件分析结果")
        print("="*50)
        
        print(f"\n1. 收件箱邮件统计:")
        print(f"   总收到邮件数: {total_received}")
        
        print(f"\n2. 邮件读取状态:")
        print(f"   已读邮件: {read_count} ({read_percentage:.1f}%)")
        print(f"   未读邮件: {unread_count} ({unread_percentage:.1f}%)")
        
        print(f"\n3. 邮件回复统计:")
        print(f"   已回复邮件数: {replied_count}")
        print(f"   当天回复数: {same_day_replies}")
        
        print(f"\n4. 前5名发件人:")
        if top_senders:
            for i, (sender, count) in enumerate(top_senders, 1):
                print(f"   {i}. {sender}: {count} 封邮件")
        else:
            print("   无数据")
        
        print(f"\n5. 前5名回复对象:")
        if top_recipients:
            for i, (recipient, count) in enumerate(top_recipients, 1):
                print(f"   {i}. {recipient}: {count} 封邮件")
        else:
            print("   无数据")
        
        print(f"\n6. 邮件分类统计:")
        print(f"   a. 信息类邮件: {info_count} 封")
        print(f"   b. 需要批准的邮件: {approval_count} 封")
        print(f"   c. 需要回复的邮件: {response_count} 封")
        
        print("\n" + "="*50)

def main():
    try:
        analyzer = OutlookEmailAnalyzer()
        analyzer.run_analysis()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
    except Exception as e:
        print(f"程序运行出错: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main() 
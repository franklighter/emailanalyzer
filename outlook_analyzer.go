package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type OutlookEmailAnalyzer struct {
	outlook   *ole.IDispatch
	namespace *ole.IDispatch
}

type EmailInfo struct {
	Subject      string
	SenderEmail  string
	SenderName   string
	ReceivedTime time.Time
	SentTime     time.Time
	IsRead       bool
	Body         string
	To           string
	CC           string
}

type SenderCount struct {
	Email string
	Count int
}

func NewOutlookEmailAnalyzer() (*OutlookEmailAnalyzer, error) {
	// 初始化COM，使用单线程模式来减少权限需求
	err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	if err != nil {
		fmt.Printf("警告: COM初始化失败: %v\n", err)
		fmt.Println("尝试使用标准初始化...")
		ole.CoInitialize(0)
	}
	
	fmt.Println("正在尝试连接到Outlook...")
	fmt.Println("注意: 如果Outlook未运行，请先启动Outlook应用程序")
	
	// 尝试连接到已运行的Outlook实例
	outlook, err := oleutil.GetActiveObject("Outlook.Application")
	if err != nil {
		fmt.Printf("无法连接到运行中的Outlook实例: %v\n", err)
		fmt.Println("尝试创建新的Outlook实例...")
		
		// 如果获取失败，尝试创建新实例
		outlook, err = oleutil.CreateObject("Outlook.Application")
		if err != nil {
			return nil, fmt.Errorf(`无法连接到Outlook: %v

可能的解决方案:
1. 确保Microsoft Outlook已安装并已配置邮箱账户
2. 尝试先手动启动Outlook应用程序
3. 检查Outlook是否被防病毒软件阻止
4. 尝试以管理员身份运行此程序
5. 检查Windows安全策略设置

如果仍然无法连接，请联系IT管理员获取帮助。`, err)
		}
	}
	
	outlookApp := outlook.MustQueryInterface(ole.IID_IDispatch)
	namespace, err := oleutil.CallMethod(outlookApp, "GetNamespace", "MAPI")
	if err != nil {
		return nil, fmt.Errorf(`无法获取Outlook命名空间: %v

这通常意味着:
1. Outlook未正确初始化
2. 邮箱配置文件未加载
3. 需要用户交互来完成Outlook登录

请尝试:
1. 手动启动Outlook并确保完全加载
2. 确认所有邮箱账户都已连接
3. 关闭任何Outlook登录对话框
4. 重新运行此程序`, err)
	}
	
	fmt.Println("✓ 成功连接到Outlook")
	
	return &OutlookEmailAnalyzer{
		outlook:   outlookApp,
		namespace: namespace.ToIDispatch(),
	}, nil
}

func (oa *OutlookEmailAnalyzer) Close() {
	if oa.namespace != nil {
		oa.namespace.Release()
	}
	if oa.outlook != nil {
		oa.outlook.Release()
	}
	ole.CoUninitialize()
}

func (oa *OutlookEmailAnalyzer) checkOutlookSecurity() {
	fmt.Println("\n=== Outlook安全检查 ===")
	
	// 检查Outlook版本
	version, err := oleutil.GetProperty(oa.outlook, "Version")
	if err == nil {
		fmt.Printf("Outlook版本: %s\n", version.ToString())
		version.Clear()
	}
	
	// 检查安全设置
	fmt.Println("正在检查Outlook安全设置...")
	
	// 尝试访问基本功能来检查权限
	accounts, err := oleutil.GetProperty(oa.namespace, "Accounts")
	if err != nil {
		fmt.Printf("⚠️  警告: 无法访问邮箱账户信息: %v\n", err)
		fmt.Println("   这可能是由于Outlook安全策略限制")
		fmt.Println("   某些功能可能无法正常工作")
	} else {
		fmt.Println("✓ 可以访问邮箱账户信息")
		accounts.Clear()
	}
	
	// 检查默认文件夹访问权限
	inbox, err := oleutil.CallMethod(oa.namespace, "GetDefaultFolder", 6)
	if err != nil {
		fmt.Printf("⚠️  警告: 无法访问默认收件箱: %v\n", err)
	} else {
		fmt.Println("✓ 可以访问收件箱")
		inbox.Clear()
	}
	
	fmt.Println("=== 安全检查完成 ===\n")
}

func (oa *OutlookEmailAnalyzer) getDateInput(prompt string) (time.Time, error) {
	reader := bufio.NewReader(os.Stdin)
	
	for {
		fmt.Print(prompt)
		dateStr, err := reader.ReadString('\n')
		if err != nil {
			return time.Time{}, err
		}
		
		dateStr = strings.TrimSpace(dateStr)
		date, err := time.Parse("02-01-2006", dateStr)
		if err != nil {
			fmt.Println("日期格式错误，请使用DD-MM-YYYY格式 (例如: 01-03-2025)")
			continue
		}
		
		return date, nil
	}
}

func (oa *OutlookEmailAnalyzer) listAvailableAccounts() error {
	fmt.Println("正在获取邮箱账户列表...")
	
	accounts, err := oleutil.GetProperty(oa.namespace, "Accounts")
	if err != nil {
		fmt.Printf("⚠️  无法获取账户列表: %v\n", err)
		fmt.Println("   可能原因:")
		fmt.Println("   - Outlook安全策略限制")
		fmt.Println("   - 需要用户确认访问权限")
		fmt.Println("   - 邮箱配置文件未完全加载")
		fmt.Println("\n建议:")
		fmt.Println("   - 确保Outlook完全启动并登录所有账户")
		fmt.Println("   - 检查是否有Outlook安全提示需要确认")
		fmt.Println("   - 尝试在Outlook中手动发送一封测试邮件")
		return err
	}
	defer accounts.Clear()
	
	accountsDisp := accounts.ToIDispatch()
	defer accountsDisp.Release()
	
	count, err := oleutil.GetProperty(accountsDisp, "Count")
	if err != nil {
		return fmt.Errorf("无法获取账户数量: %v", err)
	}
	
	fmt.Printf("找到 %d 个邮箱账户:\n", int(count.Val))
	
	if int(count.Val) == 0 {
		fmt.Println("⚠️  未找到任何邮箱账户")
		fmt.Println("   请确保:")
		fmt.Println("   - Outlook中已配置邮箱账户")
		fmt.Println("   - 所有账户都已成功连接")
		fmt.Println("   - 没有未完成的登录流程")
		return nil
	}
	
	for i := 1; i <= int(count.Val); i++ {
		account, err := oleutil.GetProperty(accountsDisp, "Item", i)
		if err != nil {
			fmt.Printf("   %d. (无法访问账户信息)\n", i)
			continue
		}
		
		accountDisp := account.ToIDispatch()
		
		displayName, err := oleutil.GetProperty(accountDisp, "DisplayName")
		if err == nil {
			smtpAddress, err := oleutil.GetProperty(accountDisp, "SmtpAddress")
			if err == nil {
				fmt.Printf("   %d. %s (%s)\n", i, displayName.ToString(), smtpAddress.ToString())
				smtpAddress.Clear()
			} else {
				fmt.Printf("   %d. %s (邮箱地址不可用)\n", i, displayName.ToString())
			}
			displayName.Clear()
		} else {
			fmt.Printf("   %d. (账户信息不可用)\n", i)
		}
		
		accountDisp.Release()
		account.Clear()
	}
	
	return nil
}

func (oa *OutlookEmailAnalyzer) getEmailAccount(emailAddress string) (*ole.IDispatch, error) {
	fmt.Printf("正在查找邮箱账户: %s\n", emailAddress)
	
	accounts, err := oleutil.GetProperty(oa.namespace, "Accounts")
	if err != nil {
		return nil, fmt.Errorf("无法访问账户列表: %v", err)
	}
	defer accounts.Clear()
	
	accountsDisp := accounts.ToIDispatch()
	defer accountsDisp.Release()
	
	count, err := oleutil.GetProperty(accountsDisp, "Count")
	if err != nil {
		return nil, err
	}
	
	for i := 1; i <= int(count.Val); i++ {
		account, err := oleutil.GetProperty(accountsDisp, "Item", i)
		if err != nil {
			continue
		}
		
		accountDisp := account.ToIDispatch()
		
		smtpAddress, err := oleutil.GetProperty(accountDisp, "SmtpAddress")
		if err == nil && strings.EqualFold(smtpAddress.ToString(), emailAddress) {
			displayName, _ := oleutil.GetProperty(accountDisp, "DisplayName")
			fmt.Printf("✓ 找到匹配的账户: %s\n", displayName.ToString())
			displayName.Clear()
			smtpAddress.Clear()
			account.Clear()
			return accountDisp, nil
		}
		
		if err == nil {
			smtpAddress.Clear()
		}
		accountDisp.Release()
		account.Clear()
	}
	
	fmt.Println("⚠️  未找到匹配的账户，将使用默认账户")
	return nil, nil
}

func (oa *OutlookEmailAnalyzer) getInboxFolders(emailAddress string) ([]*ole.IDispatch, error) {
	var folders []*ole.IDispatch
	
	// 尝试获取指定账户的收件箱
	account, _ := oa.getEmailAccount(emailAddress)
	var inbox *ole.IDispatch
	
	if account != nil {
		deliveryStore, err := oleutil.GetProperty(account, "DeliveryStore")
		if err == nil {
			storeDisp := deliveryStore.ToIDispatch()
			inboxResult, err := oleutil.CallMethod(storeDisp, "GetDefaultFolder", 6) // olFolderInbox = 6
			if err == nil {
				inbox = inboxResult.ToIDispatch()
				displayName, _ := oleutil.GetProperty(account, "DisplayName")
				fmt.Printf("✓ 使用账户的收件箱: %s\n", displayName.ToString())
				displayName.Clear()
			} else {
				fmt.Printf("⚠️  无法访问账户收件箱: %v\n", err)
			}
			storeDisp.Release()
			deliveryStore.Clear()
		}
		account.Release()
	}
	
	// 如果账户收件箱失败，使用默认收件箱
	if inbox == nil {
		fmt.Println("尝试访问默认收件箱...")
		inboxResult, err := oleutil.CallMethod(oa.namespace, "GetDefaultFolder", 6)
		if err != nil {
			return nil, fmt.Errorf(`无法获取收件箱: %v

可能的原因:
1. Outlook未完全加载邮箱数据
2. 邮箱账户未正确配置
3. 网络连接问题导致邮箱离线
4. Outlook安全策略阻止程序访问

建议解决方案:
1. 确保Outlook完全启动并显示所有邮件
2. 检查邮箱连接状态（文件 > 账户设置）
3. 尝试在Outlook中手动刷新邮箱
4. 重新启动Outlook后再运行此程序`, err)
		}
		inbox = inboxResult.ToIDispatch()
		fmt.Println("✓ 使用默认收件箱")
	}
	
	folders = append(folders, inbox)
	
	// 尝试获取子文件夹，如果失败也不影响主要功能
	fmt.Println("正在查找子文件夹...")
	oa.getSubfolders(inbox, &folders)
	
	fmt.Printf("✓ 总共找到 %d 个文件夹\n", len(folders))
	return folders, nil
}

func (oa *OutlookEmailAnalyzer) getSubfolders(parentFolder *ole.IDispatch, folderList *[]*ole.IDispatch) {
	foldersProperty, err := oleutil.GetProperty(parentFolder, "Folders")
	if err != nil {
		return
	}
	defer foldersProperty.Clear()
	
	foldersDisp := foldersProperty.ToIDispatch()
	defer foldersDisp.Release()
	
	count, err := oleutil.GetProperty(foldersDisp, "Count")
	if err != nil {
		return
	}
	
	for i := 1; i <= int(count.Val); i++ {
		folder, err := oleutil.GetProperty(foldersDisp, "Item", i)
		if err != nil {
			continue
		}
		
		folderDisp := folder.ToIDispatch()
		*folderList = append(*folderList, folderDisp)
		oa.getSubfolders(folderDisp, folderList)
		folder.Clear()
	}
}

func (oa *OutlookEmailAnalyzer) getEmailsInDateRange(folders []*ole.IDispatch, startDate, endDate time.Time) ([]EmailInfo, error) {
	var emails []EmailInfo
	
	startStr := startDate.Format("01/02/2006")
	endStr := endDate.AddDate(0, 0, 1).Format("01/02/2006")
	
	fmt.Printf("正在分析 %d 个文件夹的邮件...\n", len(folders))
	
	for folderIndex, folder := range folders {
		folderName, err := oleutil.GetProperty(folder, "Name")
		if err != nil {
			fmt.Printf("文件夹 %d: (无法获取名称)\n", folderIndex+1)
			continue
		}
		
		fmt.Printf("正在读取文件夹 %d/%d: %s\n", folderIndex+1, len(folders), folderName.ToString())
		
		items, err := oleutil.GetProperty(folder, "Items")
		if err != nil {
			fmt.Printf("  ⚠️  无法访问文件夹内容: %v\n", err)
			folderName.Clear()
			continue
		}
		
		itemsDisp := items.ToIDispatch()
		
		// 尝试使用过滤器来提高性能
		filterStr := fmt.Sprintf("[ReceivedTime] >= '%s' AND [ReceivedTime] < '%s'", startStr, endStr)
		filteredItems, err := oleutil.CallMethod(itemsDisp, "Restrict", filterStr)
		var targetItems *ole.IDispatch
		
		if err == nil {
			targetItems = filteredItems.ToIDispatch()
			fmt.Printf("  ✓ 使用日期过滤器\n")
		} else {
			fmt.Printf("  ⚠️  过滤器失败，将手动检查所有邮件: %v\n", err)
			targetItems = itemsDisp
		}
		
		count, err := oleutil.GetProperty(targetItems, "Count")
		if err == nil {
			folderCount := 0
			totalCount := int(count.Val)
			fmt.Printf("  处理 %d 封邮件...\n", totalCount)
			
			for i := 1; i <= totalCount; i++ {
				if i%50 == 0 {
					fmt.Printf("  进度: %d/%d (%.1f%%)\n", i, totalCount, float64(i)/float64(totalCount)*100)
				}
				
				item, err := oleutil.GetProperty(targetItems, "Item", i)
				if err != nil {
					continue
				}
				
				itemDisp := item.ToIDispatch()
				emailInfo := oa.extractEmailInfo(itemDisp, false, startDate, endDate)
				if emailInfo.Subject != "" {
					emails = append(emails, emailInfo)
					folderCount++
				}
				
				itemDisp.Release()
				item.Clear()
			}
			fmt.Printf("  ✓ 找到 %d 封符合条件的邮件\n", folderCount)
		} else {
			fmt.Printf("  ⚠️  无法获取邮件数量: %v\n", err)
		}
		
		if filteredItems != nil {
			targetItems.Release()
			filteredItems.Clear()
		}
		itemsDisp.Release()
		items.Clear()
		folderName.Clear()
	}
	
	fmt.Printf("✓ 总共找到 %d 封邮件\n", len(emails))
	return emails, nil
}

func (oa *OutlookEmailAnalyzer) getSentEmailsInDateRange(emailAddress string, startDate, endDate time.Time) ([]EmailInfo, error) {
	fmt.Println("正在获取发送邮件...")
	
	var sentFolder *ole.IDispatch
	
	// 尝试获取指定账户的发送文件夹
	account, _ := oa.getEmailAccount(emailAddress)
	if account != nil {
		deliveryStore, err := oleutil.GetProperty(account, "DeliveryStore")
		if err == nil {
			storeDisp := deliveryStore.ToIDispatch()
			sentResult, err := oleutil.CallMethod(storeDisp, "GetDefaultFolder", 5) // olFolderSentMail = 5
			if err == nil {
				sentFolder = sentResult.ToIDispatch()
				displayName, _ := oleutil.GetProperty(account, "DisplayName")
				fmt.Printf("✓ 使用账户 %s 的发送文件夹\n", displayName.ToString())
				displayName.Clear()
			} else {
				fmt.Printf("⚠️  无法访问账户发送文件夹: %v\n", err)
			}
			storeDisp.Release()
			deliveryStore.Clear()
		}
		account.Release()
	}
	
	// 如果失败，使用默认发送文件夹
	if sentFolder == nil {
		sentResult, err := oleutil.CallMethod(oa.namespace, "GetDefaultFolder", 5)
		if err != nil {
			fmt.Printf("⚠️  无法访问发送文件夹: %v\n", err)
			fmt.Println("   跳过发送邮件分析，继续其他功能...")
			return []EmailInfo{}, nil
		}
		sentFolder = sentResult.ToIDispatch()
		fmt.Println("✓ 使用默认发送文件夹")
	}
	defer sentFolder.Release()
	
	var sentEmails []EmailInfo
	startStr := startDate.Format("01/02/2006")
	endStr := endDate.AddDate(0, 0, 1).Format("01/02/2006")
	
	items, err := oleutil.GetProperty(sentFolder, "Items")
	if err != nil {
		return sentEmails, fmt.Errorf("读取发送邮件失败: %v", err)
	}
	defer items.Clear()
	
	itemsDisp := items.ToIDispatch()
	defer itemsDisp.Release()
	
	count, err := oleutil.GetProperty(itemsDisp, "Count")
	if err == nil {
		totalCount := int(count.Val)
		fmt.Printf("发送文件夹中总共有 %d 封邮件\n", totalCount)
		
		// 尝试使用过滤器
		filterStr := fmt.Sprintf("[SentOn] >= '%s' AND [SentOn] < '%s'", startStr, endStr)
		filteredItems, err := oleutil.CallMethod(itemsDisp, "Restrict", filterStr)
		
		var targetItems *ole.IDispatch
		if err == nil {
			targetItems = filteredItems.ToIDispatch()
			defer targetItems.Release()
			defer filteredItems.Clear()
			
			filteredCount, _ := oleutil.GetProperty(targetItems, "Count")
			fmt.Printf("✓ 过滤后有 %d 封邮件\n", int(filteredCount.Val))
		} else {
			fmt.Printf("⚠️  过滤失败，手动遍历: %v\n", err)
			targetItems = itemsDisp
		}
		
		itemCount, _ := oleutil.GetProperty(targetItems, "Count")
		processCount := int(itemCount.Val)
		
		for i := 1; i <= processCount; i++ {
			if i%20 == 0 {
				fmt.Printf("处理发送邮件进度: %d/%d\n", i, processCount)
			}
			
			item, err := oleutil.GetProperty(targetItems, "Item", i)
			if err != nil {
				continue
			}
			
			itemDisp := item.ToIDispatch()
			emailInfo := oa.extractEmailInfo(itemDisp, true, startDate, endDate)
			if emailInfo.Subject != "" {
				sentEmails = append(sentEmails, emailInfo)
			}
			
			itemDisp.Release()
			item.Clear()
		}
	}
	
	fmt.Printf("✓ 找到 %d 封发送邮件\n", len(sentEmails))
	return sentEmails, nil
}

func (oa *OutlookEmailAnalyzer) extractEmailInfo(item *ole.IDispatch, isSent bool, startDate, endDate time.Time) EmailInfo {
	var emailInfo EmailInfo
	
	// 获取主题
	subject, err := oleutil.GetProperty(item, "Subject")
	if err == nil {
		emailInfo.Subject = subject.ToString()
	}
	subject.Clear()
	
	// 获取时间并验证日期范围
	if isSent {
		sentOn, err := oleutil.GetProperty(item, "SentOn")
		if err == nil {
			emailInfo.SentTime = sentOn.ToDateTime()
			// 验证日期范围
			if emailInfo.SentTime.Before(startDate) || emailInfo.SentTime.After(endDate.AddDate(0, 0, 1)) {
				sentOn.Clear()
				return EmailInfo{} // 返回空结构体表示不符合条件
			}
		}
		sentOn.Clear()
	} else {
		receivedTime, err := oleutil.GetProperty(item, "ReceivedTime")
		if err == nil {
			emailInfo.ReceivedTime = receivedTime.ToDateTime()
			// 验证日期范围
			if emailInfo.ReceivedTime.Before(startDate) || emailInfo.ReceivedTime.After(endDate.AddDate(0, 0, 1)) {
				receivedTime.Clear()
				return EmailInfo{} // 返回空结构体表示不符合条件
			}
		}
		receivedTime.Clear()
		
		// 获取已读状态
		unread, err := oleutil.GetProperty(item, "UnRead")
		if err == nil {
			emailInfo.IsRead = !unread.ToBool()
		}
		unread.Clear()
		
		// 获取发件人信息
		senderEmail, err := oleutil.GetProperty(item, "SenderEmailAddress")
		if err == nil {
			emailInfo.SenderEmail = senderEmail.ToString()
		}
		senderEmail.Clear()
		
		senderName, err := oleutil.GetProperty(item, "SenderName")
		if err == nil {
			emailInfo.SenderName = senderName.ToString()
		}
		senderName.Clear()
	}
	
	// 获取收件人信息（用于发送邮件）
	if isSent {
		to, err := oleutil.GetProperty(item, "To")
		if err == nil {
			emailInfo.To = to.ToString()
		}
		to.Clear()
		
		cc, err := oleutil.GetProperty(item, "CC")
		if err == nil {
			emailInfo.CC = cc.ToString()
		}
		cc.Clear()
	}
	
	// 尝试获取邮件正文（可能比较慢，所以可以选择跳过）
	// 为了提高性能，只获取前500个字符用于分类
	body, err := oleutil.GetProperty(item, "Body")
	if err == nil {
		bodyText := body.ToString()
		if len(bodyText) > 500 {
			bodyText = bodyText[:500]
		}
		emailInfo.Body = bodyText
	}
	body.Clear()
	
	return emailInfo
}

func (oa *OutlookEmailAnalyzer) analyzeReadStatus(emails []EmailInfo) (int, int, float64, float64) {
	readCount := 0
	unreadCount := 0
	
	for _, email := range emails {
		if email.IsRead {
			readCount++
		} else {
			unreadCount++
		}
	}
	
	total := readCount + unreadCount
	var readPercentage, unreadPercentage float64
	if total > 0 {
		readPercentage = float64(readCount) / float64(total) * 100
		unreadPercentage = float64(unreadCount) / float64(total) * 100
	}
	
	return readCount, unreadCount, readPercentage, unreadPercentage
}

func (oa *OutlookEmailAnalyzer) findRepliedEmails(receivedEmails, sentEmails []EmailInfo) (int, int) {
	if len(sentEmails) == 0 {
		fmt.Println("⚠️  没有发送邮件数据，跳过回复分析")
		return 0, 0
	}
	
	repliedCount := 0
	sameDayReplies := 0
	
	// 创建发送邮件的主题和时间映射
	sentSubjects := make(map[string][]time.Time)
	rePrefix := regexp.MustCompile(`^(re:|回复:|回覆:)\s*`)
	
	for _, sentEmail := range sentEmails {
		subject := strings.ToLower(sentEmail.Subject)
		cleanSubject := rePrefix.ReplaceAllString(subject, "")
		sentDate := sentEmail.SentTime.Truncate(24 * time.Hour)
		
		if _, exists := sentSubjects[cleanSubject]; !exists {
			sentSubjects[cleanSubject] = []time.Time{}
		}
		sentSubjects[cleanSubject] = append(sentSubjects[cleanSubject], sentDate)
	}
	
	for _, email := range receivedEmails {
		if !email.IsRead {
			continue
		}
		
		subject := strings.ToLower(email.Subject)
		cleanSubject := rePrefix.ReplaceAllString(subject, "")
		receivedDate := email.ReceivedTime.Truncate(24 * time.Hour)
		
		if sentDates, exists := sentSubjects[cleanSubject]; exists {
			repliedCount++
			for _, sentDate := range sentDates {
				if sentDate.Equal(receivedDate) {
					sameDayReplies++
					break
				}
			}
		}
	}
	
	return repliedCount, sameDayReplies
}

func (oa *OutlookEmailAnalyzer) getTopSendersAndRecipients(receivedEmails, sentEmails []EmailInfo) ([]SenderCount, []SenderCount) {
	senderCounts := make(map[string]int)
	recipientCounts := make(map[string]int)
	
	// 统计发件人
	for _, email := range receivedEmails {
		sender := email.SenderEmail
		if sender == "" {
			sender = email.SenderName
		}
		if sender != "" {
			senderCounts[sender]++
		}
	}
	
	// 统计收件人
	for _, email := range sentEmails {
		recipients := strings.Split(email.To, ";")
		ccRecipients := strings.Split(email.CC, ";")
		allRecipients := append(recipients, ccRecipients...)
		
		for _, recipient := range allRecipients {
			recipient = strings.TrimSpace(recipient)
			if recipient != "" {
				recipientCounts[recipient]++
			}
		}
	}
	
	// 转换为切片并排序
	var topSenders []SenderCount
	for sender, count := range senderCounts {
		topSenders = append(topSenders, SenderCount{Email: sender, Count: count})
	}
	sort.Slice(topSenders, func(i, j int) bool {
		return topSenders[i].Count > topSenders[j].Count
	})
	if len(topSenders) > 5 {
		topSenders = topSenders[:5]
	}
	
	var topRecipients []SenderCount
	for recipient, count := range recipientCounts {
		topRecipients = append(topRecipients, SenderCount{Email: recipient, Count: count})
	}
	sort.Slice(topRecipients, func(i, j int) bool {
		return topRecipients[i].Count > topRecipients[j].Count
	})
	if len(topRecipients) > 5 {
		topRecipients = topRecipients[:5]
	}
	
	return topSenders, topRecipients
}

func (oa *OutlookEmailAnalyzer) classifyEmails(emails []EmailInfo) (int, int, int) {
	infoKeywords := []string{"通知", "信息", "更新", "公告", "新闻", "newsletter", "notification", "update", "info", "通告"}
	approvalKeywords := []string{"批准", "审批", "确认", "同意", "授权", "approve", "approval", "authorize", "confirm", "核准", "签核"}
	responseKeywords := []string{"回复", "回应", "反馈", "意见", "建议", "reply", "response", "feedback", "urgent", "紧急", "请回复", "请回覆"}
	
	infoCount := 0
	approvalCount := 0
	responseCount := 0
	
	for _, email := range emails {
		subject := strings.ToLower(email.Subject)
		body := strings.ToLower(email.Body)
		
		foundApproval := false
		foundResponse := false
		
		// 检查是否包含批准关键词
		for _, keyword := range approvalKeywords {
			if strings.Contains(subject, keyword) || strings.Contains(body, keyword) {
				approvalCount++
				foundApproval = true
				break
			}
		}
		
		if !foundApproval {
			// 检查是否需要回复
			for _, keyword := range responseKeywords {
				if strings.Contains(subject, keyword) || strings.Contains(body, keyword) {
					responseCount++
					foundResponse = true
					break
				}
			}
		}
		
		if !foundApproval && !foundResponse {
			infoCount++
		}
	}
	
	return infoCount, approvalCount, responseCount
}

func (oa *OutlookEmailAnalyzer) printResults(totalReceived, readCount, unreadCount int, readPercentage, unreadPercentage float64,
	repliedCount, sameDayReplies int, topSenders, topRecipients []SenderCount, infoCount, approvalCount, responseCount int) {
	
	fmt.Println("\n" + strings.Repeat("=", 60))
	fmt.Println("📊 邮件分析结果")
	fmt.Println(strings.Repeat("=", 60))
	
	fmt.Printf("\n📧 1. 收件箱邮件统计:\n")
	fmt.Printf("   总收到邮件数: %d 封\n", totalReceived)
	
	fmt.Printf("\n👁️ 2. 邮件读取状态:\n")
	fmt.Printf("   已读邮件: %d 封 (%.1f%%)\n", readCount, readPercentage)
	fmt.Printf("   未读邮件: %d 封 (%.1f%%)\n", unreadCount, unreadPercentage)
	
	fmt.Printf("\n↩️ 3. 邮件回复统计:\n")
	fmt.Printf("   已回复邮件数: %d 封\n", repliedCount)
	fmt.Printf("   当天回复数: %d 封\n", sameDayReplies)
	if repliedCount > 0 {
		sameDayPercentage := float64(sameDayReplies) / float64(repliedCount) * 100
		fmt.Printf("   当天回复率: %.1f%%\n", sameDayPercentage)
	}
	
	fmt.Printf("\n📬 4. 前5名发件人:\n")
	if len(topSenders) > 0 {
		for i, sender := range topSenders {
			fmt.Printf("   %d. %s: %d 封邮件\n", i+1, sender.Email, sender.Count)
		}
	} else {
		fmt.Printf("   无数据\n")
	}
	
	fmt.Printf("\n📤 5. 前5名回复对象:\n")
	if len(topRecipients) > 0 {
		for i, recipient := range topRecipients {
			fmt.Printf("   %d. %s: %d 封邮件\n", i+1, recipient.Email, recipient.Count)
		}
	} else {
		fmt.Printf("   无发送邮件数据\n")
	}
	
	fmt.Printf("\n📋 6. 邮件分类统计:\n")
	fmt.Printf("   a. 信息类邮件: %d 封 (%.1f%%)\n", infoCount, float64(infoCount)/float64(totalReceived)*100)
	fmt.Printf("   b. 需要批准的邮件: %d 封 (%.1f%%)\n", approvalCount, float64(approvalCount)/float64(totalReceived)*100)
	fmt.Printf("   c. 需要回复的邮件: %d 封 (%.1f%%)\n", responseCount, float64(responseCount)/float64(totalReceived)*100)
	
	fmt.Println("\n" + strings.Repeat("=", 60))
	
	// 添加一些有用的建议
	fmt.Println("💡 分析建议:")
	if unreadPercentage > 20 {
		fmt.Println("   - 未读邮件较多，建议及时处理重要邮件")
	}
	if repliedCount > 0 && sameDayReplies < repliedCount/2 {
		fmt.Println("   - 考虑提高邮件回复及时性")
	}
	if responseCount > 0 {
		fmt.Printf("   - 有 %d 封邮件可能需要您的回复\n", responseCount)
	}
	if approvalCount > 0 {
		fmt.Printf("   - 有 %d 封邮件可能需要您的批准\n", approvalCount)
	}
}

func (oa *OutlookEmailAnalyzer) runAnalysis() error {
	fmt.Println("=== 📧 Outlook 邮件分析工具 ===")
	fmt.Println("此工具将分析您的Outlook邮件数据")
	fmt.Println("请确保Outlook已完全启动并登录所有账户\n")
	
	// 执行安全检查
	oa.checkOutlookSecurity()
	
	// 显示可用账户
	if err := oa.listAvailableAccounts(); err != nil {
		fmt.Printf("⚠️  获取账户列表时遇到问题: %v\n", err)
		fmt.Println("程序将尝试继续运行...")
	}
	fmt.Println()
	
	// 获取用户输入
	startDate, err := oa.getDateInput("请输入开始日期 (DD-MM-YYYY): ")
	if err != nil {
		return err
	}
	
	endDate, err := oa.getDateInput("请输入结束日期 (DD-MM-YYYY): ")
	if err != nil {
		return err
	}
	
	// 验证日期范围
	if endDate.Before(startDate) {
		return fmt.Errorf("结束日期不能早于开始日期")
	}
	
	// 检查日期范围是否过大
	daysDiff := endDate.Sub(startDate).Hours() / 24
	if daysDiff > 365 {
		fmt.Printf("⚠️  警告: 日期范围超过一年 (%.0f 天)，分析可能需要较长时间\n", daysDiff)
		fmt.Print("是否继续? (y/n): ")
		reader := bufio.NewReader(os.Stdin)
		response, _ := reader.ReadString('\n')
		if !strings.HasPrefix(strings.ToLower(strings.TrimSpace(response)), "y") {
			return fmt.Errorf("用户取消操作")
		}
	}
	
	reader := bufio.NewReader(os.Stdin)
	fmt.Print("请输入邮箱地址 (或按回车使用默认账户): ")
	emailAddress, err := reader.ReadString('\n')
	if err != nil {
		return err
	}
	emailAddress = strings.TrimSpace(emailAddress)
	
	if emailAddress == "" {
		emailAddress = "default"
		fmt.Println("使用默认邮箱账户")
	}
	
	fmt.Printf("\n🔍 正在分析 %s 到 %s 的邮件...\n", 
		startDate.Format("2006-01-02"), endDate.Format("2006-01-02"))
	
	// 获取收件箱文件夹
	inboxFolders, err := oa.getInboxFolders(emailAddress)
	if err != nil {
		return err
	}
	defer func() {
		for _, folder := range inboxFolders {
			folder.Release()
		}
	}()
	
	if len(inboxFolders) == 0 {
		return fmt.Errorf("无法获取收件箱文件夹")
	}
	
	// 获取收到的邮件
	receivedEmails, err := oa.getEmailsInDateRange(inboxFolders, startDate, endDate)
	if err != nil {
		return err
	}
	
	if len(receivedEmails) == 0 {
		fmt.Println("⚠️  在指定日期范围内未找到任何邮件")
		fmt.Println("   请检查:")
		fmt.Println("   - 日期范围是否正确")
		fmt.Println("   - Outlook是否已同步邮件")
		fmt.Println("   - 邮箱账户是否正常工作")
		return nil
	}
	
	// 获取发送的邮件
	sentEmails, err := oa.getSentEmailsInDateRange(emailAddress, startDate, endDate)
	if err != nil {
		fmt.Printf("⚠️  获取发送邮件时出错: %v\n", err)
		fmt.Println("   继续进行其他分析...")
		sentEmails = []EmailInfo{} // 继续执行，但没有发送邮件数据
	}
	
	// 执行各项分析
	fmt.Println("\n📊 正在进行数据分析...")
	
	readCount, unreadCount, readPercentage, unreadPercentage := oa.analyzeReadStatus(receivedEmails)
	repliedCount, sameDayReplies := oa.findRepliedEmails(receivedEmails, sentEmails)
	topSenders, topRecipients := oa.getTopSendersAndRecipients(receivedEmails, sentEmails)
	infoCount, approvalCount, responseCount := oa.classifyEmails(receivedEmails)
	
	// 打印结果
	oa.printResults(len(receivedEmails), readCount, unreadCount, readPercentage, unreadPercentage,
		repliedCount, sameDayReplies, topSenders, topRecipients, infoCount, approvalCount, responseCount)
	
	return nil
}

func main() {
	fmt.Println("正在启动Outlook邮件分析工具...")
	fmt.Println("版本: 2.0 (增强权限处理)")
	
	analyzer, err := NewOutlookEmailAnalyzer()
	if err != nil {
		fmt.Printf("\n❌ 初始化失败: %v\n\n", err)
		fmt.Println("🔧 故障排除建议:")
		fmt.Println("1. 确保Microsoft Outlook已安装并配置")
		fmt.Println("2. 尝试手动启动Outlook并确保完全加载")
		fmt.Println("3. 如果使用企业版Outlook，联系IT部门检查安全策略")
		fmt.Println("4. 尝试以管理员身份运行此程序")
		fmt.Println("5. 检查防病毒软件是否阻止了程序访问")
		
		fmt.Print("\n按回车键退出...")
		bufio.NewReader(os.Stdin).ReadString('\n')
		return
	}
	defer analyzer.Close()
	
	if err := analyzer.runAnalysis(); err != nil {
		fmt.Printf("\n❌ 分析过程中出错: %v\n", err)
	} else {
		fmt.Println("\n✅ 分析完成!")
	}
	
	fmt.Print("\n按回车键退出...")
	bufio.NewReader(os.Stdin).ReadString('\n')
} 
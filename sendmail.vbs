function Send_mail(To_Account,Send_Topic,Send_Body,Send_Attachment) 

You_Account	="oliuzhiyuo@163.com"                             '发件人163邮箱
You_Password="password"                                       '发件人邮箱密码

Set Email 		= CreateObject("CDO.Message")						        '发件人
Email.From 		= You_Account
Email.To 		= To_Account			      				 		              '收件人
Email.Subject 	= Send_Topic        								          '邮件主题
Email.Textbody 	= Send_Body        									          '邮件内容

If Send_Attachment <> "" Then										              '邮件附件
Email.AddAttachment Send_Attachment     							
End If

You_ID   = Split(You_Account, "@", -1, vbTextCompare) 			  '帐号和服务器分离
MS_Space = "http://schemas.microsoft.com/cdo/configuration/"	'必要
With Email.Configuration.Fields
.Item(MS_Space&"sendusing") 		    = 2       						    '发信端口
.Item(MS_Space&"smtpserver") 		    = "smtp."&You_ID(1) 			'SMTP服务器地址
.Item(MS_Space&"smtpserverport") 	  = 465     						    'SMTP服务器端口
.Item(MS_Space&"smtpusessl") 		    = true							      'SMTP服务器是否使用了SSL
.Item(MS_Space&"smtpauthenticate") 	= 1     						      '认证方式
.Item(MS_Space&"sendusername") 		  = You_ID(0)    					  '发件帐号
.Item(MS_Space&"sendpassword") 		  = You_Password   				  '发件密码
.Update
End With

Email.Send															                      '发送邮件
Set Email=Nothing													                    '关闭组件

Send_Mail=True 														                    '如果没有任何错误信息，则表示发送成功,否则发送失败 
If Err Then 
Err.Clear 
Send_Mail=False 
End If 
End Function


'调用函数发送带附件的邮件   Send_Mail(收件人，标题，正文，附件)
If Send_Mail("oliuzhiyuo@163.com","toplic","test","C:\Users\ZhiYuLiu\Desktop\test.txt")=True Then
Wscript.Echo "successful"
Else
Wscript.Echo "faild"
End If

# data_excel_mail
导出数据为Excel格式，并作为附件发送邮件。数据库链接参数和SQL语法，以及邮箱可配置。


直接打成jar包使用，相关的运行参数复制 data_excel_mail.properties 文件到jar包同目录

```
"驱动，目前只打包了oracle的驱动
driver=oracle.jdbc.driver.OracleDriver
"数据库连接参数
conn=jdbc:oracle:thin:@1.18.21.18:1521:orcl
“数据库用户名
u=root
“数据库密码
p=root

”需要执行的SQL语法
sql=select * from student where rownum < 20

“生成的Excel文件名称，也用作email的标题
file=学生统计

"接收邮件的邮件地址，不配置的话，表示仅生成本地文件，文件在jar包相同目录
mail=aaa@163.com
”需要发送邮件时，需要以下三个参数，邮件服务器,发件人的账号和密码
"PS：如果使用163发邮件，邮件密码不使用真正的密码，需要去163的网站上拿一个验证码下来做密码，“设置->POP3/SMTP/IMAP”
mailserver=smtp.163.com
sender=
senderpsw=
```

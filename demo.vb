Private Sub 全自动发送邮件_Click()
    '要能正确发送并需要对Microseft Outlook进行有效配置
    On Error Resume Next
    Dim rowCount, endRowNo, endColumnNo, sFile$, sFile1$, A&, B&
    Dim objOutlook As Object
    Dim objMail As MailItem
    '取得当前工作表数据区行数列数
    endRowNo = ActiveSheet.UsedRange.Rows.Count
    endColumnNo = ActiveSheet.UsedRange.Columns.Count
    
    '取得当前工作表的名称，用来作为邮件主题进行发送
    sFile1 = ActiveSheet.Name
    '创建objOutlook为Outlook应用程序对象
    Set objOutlook = CreateObject("Outlook.Application")
    
   '开始循环发送电子邮件
    For rowCount = 2 To endRowNo
   '创建objMail为一个邮件对象
    Set objMail = objOutlook.CreateItem(olMailItem)
 
    With objMail
    
    '设置收件人地址，数据源所在列数
    .To = Cells(rowCount, 5)
    
    '设置抄送人地址（从通讯录表的'E-mail地址'字段中获得）
    '.CC = Cells(rowCount, 0)
    '设置邮件主题,取值工作表名，
    .Subject = sFile1
  '设置邮件内容(从通讯录表的“内容”字段中获得)
  'align  单元格文本显示方式 left(向左)、center(居中)、right(向右)，默认是center, width-宽 height-高  border 单元格线粗细,bordercolor返回或设置对象的边框颜色
  'colSpan是一种编程语言，其属性可设置或返回表元横跨的列数
  
  
 sFile = "<tr>您好！<br> 以下是您" + sFile1 + "，请查收！</tr>"
    sFile = sFile + "<table align='left' width='500' height='25' border= 1   bordercolor='#000000'> <tbody> "
    sFile = sFile + "<tr>  <td colspan ='4' align='center'> 工资表</td> </tr> "
    B = 1
    For A = 1 To endColumnNo
    '数据表头中添加“X”后将不发送此字段
       If Application.WorksheetFunction.CountIf(Cells(1, A), "*X*") = 0 Then
       If B = 1 Then
         sFile = sFile + "<tr>  <td width='20%' height='25'> " + Cells(1, A).Text + "   </td> <td  width='30%' height='25'> " + Cells(rowCount, A).Text + "</td>"
         B = 0
                 
       Else
        sFile = sFile + "<td width='20%' height='25'> " + Cells(1, A).Text + "   </td> <td  width='30%' height='25'> " + Cells(rowCount, A).Text + "</td> </tr>"
        B = 1
       End If
     End If
    Next
    
   .HTMLBody = sFile
 
  
    '设置附件(从通讯录表的“附件”字段中获得)
    .Attachments.Add Cells(rowCount, 24).Value
    '自动发送邮件
    .Send
     End With
    
    '销毁objMail对象
    Set objMail = Nothing
    Next
    '销毁objOutlook对象
    Set objOutlook = Nothing
    '所有电子邮件发送完成时提示
     MsgBox rowCount - 2 & "个员工的工资单发送成功！"
 
End Sub

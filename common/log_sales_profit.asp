<%
Dim userIP, userInfo, userUrl, menuName, excelYn, fileName

UserIP = Request.ServerVariables("REMOTE_ADDR")
UserInfo = Request.ServerVariables("HTTP_USER_AGENT")
userUrl = Request.ServerVariables("HTTP_URL")
fileName = f_LogFileName(Request.ServerVariables("SCRIPT_NAME"))

'response.write fileName

menuName = "영업관리 > 손익현황"

Select Case fileName
    Case "saupbu_profit_loss_total_std_excel.asp"
        excelYn = "Y"
    Case "saupbu_profit_loss_excel_std.asp", "cost_center_detail_excel_std.asp", "saupbu_sales_detail_excel2.asp"
        excelYn = "Y"
    Case "saupbu_profit_loss_total_excel.asp"
        excelYn = "Y"
    Case "saupbu_profit_loss_excel.asp", "cost_center_detail_excel.asp"
        excelYn = "Y"
    Case Else
        excelYn = "N"
End Select

'접근 정보 저장
objBuilder.Append "INSERT INTO emp_sys_log(emp_no, remote_ip, user_url, menu_name, menu_title, excel_yn)"
objBuilder.Append "VALUES('"&emp_no&"', '"&UserIP&"', '"&userUrl&"', '"&menuName&"', '"&title_line&"', '"&excelYn&"')"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
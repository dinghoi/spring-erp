<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'on Error resume next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim company_tab(50)
Dim title_name, savefilename, condi_sql, order_sql, where_sql, base_sql
Dim rs, i, alldata, numcols, numrows, j, in_process, thisfield, k
Dim rs_trade, com_sql, kk

'Dim sql

title_name = array("접수번호","접수일자","접수자","직급","사용자","전화번호","핸드폰","회사","조직명","주소","CE명","장애내역","요청일","요청시간","처리방법","진행","고객요청","제조사","장애장비","모델명","입고사유","입고상태")

savefilename = user_id & ".xls"

'Response.Buffer = True
'Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
'Response.CacheControl = "public"
'Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename
Call ViewExcelType(savefilename)

If reside = "9" Then
	k = 0

	'Sql="select * from trade where use_sw = 'Y' and group_name = '"+user_name+"' order by trade_name asc"
	objBuilder.Append "SELECT trade_name FROM trade "
	objBuilder.Append "WHERE use_sw = 'Y' AND group_name = '"&user_name&"' "
	objBuilder.Append "ORDER BY trade_name ASC "

	Set rs_trade = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Do Until rs_trade.EOF
		k = k + 1

		company_tab(k) = rs_trade("trade_name")

		rs_trade.MoveNext()
	Loop
	rs_trade.Close() : Set rs_trade = Nothing
End If

If reside = "9" Then
	com_sql = "company = '" & company_tab(1) & "'"

	For kk = 2 To k
		com_sql = com_sql & " OR company = '" & company_tab(kk) & "'"
	Next

	condi_sql = " OR " & com_sql & ") "
Else
	condi_sql = " OR company = '" & reside_company & "' OR company = '" & user_name & "') "
End If

'//2017-06-07 아이티퓨처(사번:900002) 로그인시 웅진관련 기업 검색하게 수정
If user_id = "900002" Then
	condi_sql = " OR company in ('웅진식품','웅진씽크빅','코웨이') " & condi_sql
End If

order_Sql = " ORDER BY acpt_date desc"

'	where_sql = " WHERE (acpt_man = '" + user_name + "' or company = '" + reside_company + "' or company = '" + user_name + "') and "
where_sql = " WHERE (acpt_man = '" & user_name & "'" & condi_sql
base_sql = " AND (as_process = '접수' OR as_process = '입고' OR as_process = '연기' OR as_process = '대체입고') "

'sql = "select acpt_no,acpt_date,acpt_man,acpt_grade,acpt_user,concat(tel_ddd,'-',tel_no1,'-',tel_no2),concat(hp_ddd,'-',hp_no1,'-',hp_no2),company,dept,concat(sido,' ',gugun,' ',dong,' ',addr),mg_ce,as_memo,request_date,request_time,as_type,as_process,visit_request_yn,maker,as_device,model_no,into_reason from as_acpt "
objBuilder.Append "SELECT acpt_no,acpt_date,acpt_man,acpt_grade,acpt_user, "
objBuilder.Append "	CONCAT(tel_ddd, '-', tel_no1, '-', tel_no2), CONCAT(hp_ddd, '-', hp_no1, '-', hp_no2), company, dept,  "
objBuilder.Append "	CONCAT(sido, ' ', gugun, ' ', dong, ' ', addr), mg_ce, as_memo, request_date, request_time, as_type, "
objBuilder.Append "	as_process, visit_request_yn, maker, as_device, model_no, into_reason "
objBuilder.Append "FROM as_acpt "

'sql = sql + where_sql + base_sql + order_sql
objBuilder.Append where_sql & base_sql & order_sql

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rs.EOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('다운 할 자료가 없습니다.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
</head>
<body>
<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
	<tr>
	<%=Chr(13)&Chr(10)%>
<%
	i = 0

	For i = 0 To 21
'	for each whatever in rs.fields
'		if i < 21 then
%>
		<td><b><%=title_name(i)%></b></TD><%=Chr(13)&Chr(10)%>
<%
	Next
'		end if
'		i = i + 1
'	next
%>
	</tr>
	<%=Chr(13)&Chr(10)%>
<%
alldata = rs.getRows()

numcols = UBound(alldata, 1) + 1
numrows = UBound(alldata, 2)

For j= 0 To numrows
	in_process = ""

	If alldata(15,j) = "입고" Then
		'sql = "select into_date,in_process,in_place from as_into where acpt_no="&alldata(0,j)&" and in_seq="&"(select max(in_seq) from as_into where acpt_no="&alldata(0,j)&")"
		objBuilder.Append "SELECT into_date, in_process, in_place FROM as_into "
		objBuilder.Append "WHERE acpt_no="&alldata(0, j)&" "
		objBuilder.Append "	AND in_seq = (SELECT MAX(in_seq) FROM as_into WHERE acpt_no = "&alldata(0, j)&") "

		Set rs_in = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_in.EOF Then
			in_process = "없음"
		Else
			in_process = rs_in("in_process")
		End If
	End If

	If alldata(16, j) = "Y" Then
		alldata(16, j) = "방문요청"
	Else
		alldata(16, j) = ""
	End If
%>
	<tr>
	<%=Chr(13)&Chr(10)%>
<%
	For i = 0 To numcols
		If i = 21 Then
    		thisfield = in_process
		Else
			thisfield = alldata(i, j)
		End If

		If IsNull(thisfield) Then
			thisfield=""
		End If

		If Trim(thisfield) = "" Then
			thisfield = ""
		End If

		If i = 1 Or i = 11 Then
%>
		<td valign="top"><%=thisfield%></td>
		<%=Chr(13)&Chr(10)%>
<%		Else	%>
		<td style="mso-number-format:'\@'" valign=top><%=thisfield%> </td><%=chr(13)&chr(10)%>
<%
		End If
	NEXT
%>
	</tr>
	<%=Chr(13)&Chr(10)%>
<%Next%>
</table>
</body>
</html>

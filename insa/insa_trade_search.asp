<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
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
Dim gubun, title_line, rs, trade_name, i, trade_code
Dim trade_no, trade_person, trade_email, group_name
Dim group_trade_code

gubun = f_Request("gubun")
trade_name = f_Request("trade_name")

objBuilder.Append "SELECT trdt.trade_code, trade_name, trade_no, trade_person, trade_email, group_name, "
objBuilder.Append "	CASE WHEN group_name IS NULL OR group_name = '' THEN trdt.trade_code "
objBuilder.Append "	ELSE (SELECT trade_code FROM trade WHERE trade_name = trdt.group_name) "
objBuilder.Append "	END AS group_trade_code "
objBuilder.Append "FROM trade AS trdt "

If gubun = "1" Or gubun = "5" Then
   If trade_name = "" Then
	   objBuilder.Append "WHERE (trade_id = '일반' or trade_id = '매출') AND trade_name = '%" & trade_name & "%' ORDER BY trade_name ASC"
   Else
	   objBuilder.Append "WHERE (trade_id = '일반' or trade_id = '매출') AND trade_name LIKE '%" & trade_name & "%' ORDER BY trade_name ASC"
   End If
End If

If gubun = "2" Or gubun = "3" Or gubun = "4" Then
   If trade_name = "" Then
	   objBuilder.Append "WHERE trade_name = '" & trade_name & "' ORDER BY trade_name ASC"
   Else
	   objBuilder.Append "WHERE trade_name LIKE '%" & trade_name & "%' ORDER BY trade_name ASC"
   End If
End If

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "거래처 검색"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사 관리 시스템</title>
	<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<script type="text/javascript">
		function trade_list(trade_code,trade_name,trade_no,trade_person,trade_email,group_name){
			console.log(group_name);
			//인사 조직
			if(document.frm.gubun.value =="1") {
				opener.document.frm.org_reside_company.value = trade_name;
				opener.document.frm.org_cost_group.value = group_name;
//					opener.document.frm.trade_no.value = trade_no;
//					opener.document.frm.trade_person.value = trade_person;
//					opener.document.frm.trade_email.value = trade_email;

				opener.document.frm.trade_code.value = trade_code;

				window.close();
			}
			if(document.frm.gubun.value =="2") {
				opener.document.frm.cost_company.value = trade_name;
				window.close();
			}
			if(document.frm.gubun.value =="3") {
				opener.document.frm.customer.value = trade_name;
				opener.document.frm.customer_no.value = trade_no;
				window.close();
			}
			if(document.frm.gubun.value =="4") {
				opener.document.frm.company.value = trade_name;
				window.close();
			}
			if(document.frm.gubun.value =="5") {
				opener.document.frm.emp_reside_company.value = trade_name;
				opener.document.frm.cost_group.value = group_name;
				/*opener.document.frm.trade_no.value = trade_no;
				opener.document.frm.trade_person.value = trade_person;
				opener.document.frm.trade_email.value = trade_email;*/
				window.close();
			}
		}

		function frmcheck(){
			if(chkfrm()){
				document.frm.submit();
			}
		}

		function chkfrm(){
			if(document.frm.trade_name.value =="") {
				alert('거래처명을 입력하세요');
				frm.trade_name.focus();
				return false;}
			{
				return true;
			}
		}
	</script>

</head>
<body>
<div id="container">
	<h3 class="insa"><%=title_line%></h3><br/>
	<form action="/insa/insa_trade_search.asp" method="post" name="frm">
		<fieldset class="srch">
			<legend>조회영역</legend>
			<dl>
				<dd>
					<p>
					<strong>거래처명을 입력하세요 </strong>
						<label>
							<input type="text" name="trade_name" id="trade_name" value="<%=trade_name%>" style="width:150px;text-align:left;ime-mode:active" />
						</label>
						<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
					</p>
				</dd>
			</dl>
		</fieldset>
		<div class="gView">
			<table cellpadding="0" cellspacing="0" class="tableList">
				<colgroup>
					<col width="25%" >
					<col width="20%" >
					<col width="20%" >
					<col width="20%" >
					<col width="*" >
				</colgroup>
				<thead>
					<tr>
						<th class="first" scope="col">거래처명</th>
						<th scope="col">그룹명</th>
						<th scope="col">사업자번호</th>
						<th scope="col">담당자</th>
						<th scope="col">이메일</th>
					</tr>
				</thead>
				<tbody>
				<%
				i = 0

				Do Until rs.EOF Or rs.BOF
					trade_code = rs("trade_code")
					trade_name = rs("trade_name")
					trade_no = Mid(rs("trade_no"), 1, 3)&"-"&Mid(rs("trade_no"), 4, 2)&"-"&Mid(rs("trade_no"), 6)
					trade_person = rs("trade_person")
					trade_email = rs("trade_email")
					group_name = rs("group_name")

					If Trim(group_name) = "" Or IsNull(group_name) Then
						group_name = trade_name
					End If

					group_trade_code = rs("group_trade_code")
				%>
					<tr>
						<td class="first">
							<a href="#" onClick="trade_list('<%=group_trade_code%>','<%=trade_name%>','<%=rs("trade_no")%>','<%=trade_person %>','<%=trade_email%>','<%=group_name%>');"><%=rs("trade_name")%></a>
						</td>
						<td><%=group_name%>&nbsp;</td>
						<td><%=trade_no%>&nbsp;</td>
						<td><%=trade_person%>&nbsp;</td>
						<td><%=group_trade_code%><%'=trade_email%>&nbsp;</td>
					</tr>
				<%
					i = i + 1
					rs.MoveNext()
				Loop

				rs.Close() : Set rs = Nothing
				DBConn.Close() : Set DBConn = Nothing

				If i = 0 Then
				%>
					<tr>
						<td class="first" colspan="4" style="height:30px;">조회된 내역이 없습니다.</td>
					</tr>
				<%
				End If
				%>
				</tbody>
			</table>
		</div>
		<input type="hidden" name="gubun" value="<%=gubun%>" />
	</form>
</div>
</body>
</html>
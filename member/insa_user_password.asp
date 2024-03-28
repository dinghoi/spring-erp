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
Dim owner_view, view_condi, title_line, rsMem

owner_view = f_Request("owner_view")
view_condi = f_Request("view_condi")

title_line = "사용자 비밀번호 확인(초기화) "

If view_condi = "" Then
	view_condi = user_id
	owner_view = "T"
End If

objBuilder.Append "SELECT memt.user_name, memt.user_id, memt.pass, memt.emp_company, "
objBuilder.Append "	memt.org_name, team, memt.hp,"
objBuilder.Append "	eomt.org_name AS orgName, eomt.org_company, eomt.org_team "
objBuilder.Append "FROM memb AS memt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON memt.user_id = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

'하위 조건 제거(로직 상 사용자명으로 검색 어려움)[허정호_20220323]
'If owner_view = "C" Then
'	objBuilder.Append "WHERE memt.user_name LIKE '%"&view_condi&"%' "
'Else
'	objBuilder.Append "WHERE memt.user_id = '"&view_condi&"'"
'End If
objBuilder.Append "WHERE memt.user_id = '"&view_condi&"'"

Set rsMem = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사관리 시스템</title>
	<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<script type="text/javascript">
		function goAction(){
		   window.close();
		}

		function frmcheck () {
			if(formcheck(document.frm) && chkfrm()){
				document.frm.submit();
			}
		}

		function chkfrm(){
			if(document.frm.view_condi.value == ""){
				alert ("조건을 입력하시기 바랍니다");
				return false;
			}
			return true;
		}

		function user_password_modify(val){
			if (!confirm("사용자 비밀번호를 초기화 하시겠습니까 ?")) return;

			var frm = document.frm;

			document.frm.view_condi1.value = document.getElementById(val).value;
			document.frm.action = "/member/insa_user_password_ok.asp";
			document.frm.submit();
		}
	</script>
</head>
<body>
	<div id="container">
		<h3 class="insa"><%=title_line%></h3>
		<form action="/member/insa_user_password.asp" method="post" name="frm">
		<fieldset class="srch">
			<legend>조회영역</legend>
			<dl>
				<dd>
					<p>
						<!--<label>
							<input name="owner_view" type="radio" value="T" <%'If owner_view = "T" Then %>checked<%'End If %> style="width:25px;">사번
						</label>
						<strong>조건 : </strong>-->
						<strong>사번 : </strong>
						<label>
							<input type="text" name="view_condi" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left;">
						</label>
						<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
					</p>
				</dd>
			</dl>
		</fieldset>
		<div class="gView">
			<table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="30%" >
					<col width="*" >
				</colgroup>
				<tbody>
				<%If rsMem.EOF Or rsMem.BOF Then%>
					<tr>
						<td colspan="2" style="align:center;height:30px;">조회된 내역이 없습니다.</td>
					</tr>
				<%Else%>
					<tr>
						<th class="first">사용자</th>
						<td class="left"><%=rsMem("user_name")%>(<%=rsMem("user_id")%>)</td>
					</tr>
					<tr>
						<th class="first">기존비밀번호</th>
						<td class="left"><%=rsMem("pass")%>&nbsp;</td>
					</tr>
					<tr>
						<th class="first">소속회사</th>
						<td class="left"><%=rsMem("org_company")%>&nbsp;</td>
					</tr>
					<tr>
						<th class="first">소속</th>
						<td class="left"><%=rsMem("orgName")%>(<%=rsMem("org_team")%>)&nbsp;</td>
					</tr>
					<tr>
						<th class="first">핸드폰번호</th>
						<td class="left"><%=rsMem("hp")%>&nbsp;</td>
					</tr>
				<%End If%>
				</tbody>
			</table>
			<%
				rsMem.Close() : Set rsMem = Nothing
				DBConn.Close() : Set DBConn = Nothing
			%>
		</div>
		<br>
		<div align="center">
			<span class="btnType01"><input type="button" value="변경" onclick="user_password_modify('view_condi');return false;"/></span>
			<span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
		</div>
		<input type="hidden" name="view_condi1" value="<%=view_condi%>"/>
		</form>
	</div>
</body>
</html>


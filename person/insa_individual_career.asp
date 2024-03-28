<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim rsCrr, title_line

in_name = user_name
in_empno = user_id

If f_toString(Request.Form("in_empno"), "") <> "" Then
   objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no = '"&in_empno&"';"

   in_name = rs_emp("emp_name")
   rs_emp.Close() : Set rs_emp = Nothing
End If

objBuilder.Append "SELECT career_join_date, career_end_date, career_office, career_dept, career_position, "
objBuilder.Append "	career_task, career_seq, career_empno "
objBuilder.Append "FROM emp_career "
objBuilder.Append "WHERE career_empno = '"&in_empno&"' "
objBuilder.Append "ORDER BY career_empno,career_seq ASC "

Set rsCrr = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "경력 사항"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			function goAction(){
			   window.close();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.in_empno.value == ""){
					alert ("사번을 입력하시기 바랍니다");
					return false;
				}
				return true;
			}
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_individual_career.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>사번 : </strong>
							<label>
        						<input type="text" name="in_empno" id="in_empno" value="<%=in_empno%>" style="width:80px;" class="no-input"/>
							</label>
                            <strong>성명 : </strong>
                            <label>
								<input type="text" name="in_name" id="in_name" value="<%=in_name%>" readonly="true" style="width:80px;" class="no-input"/>
							</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="5%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                            <th colspan="3">재직기간</th>
                            <th colspan="2">회사명</th>
                            <th colspan="2">부서</th>
                            <th colspan="1">직위</th>
                            <th colspan="3">담당업무</th>
                            <th>순번</th>
                            <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsCrr.EOF Or rsCrr.BOF Then
							Response.Write "<tr><td colspan='13' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsCrr.EOF
						%>
							<tr>
                              <td colspan="3"><%=rsCrr("career_join_date")%>∼<%=rsCrr("career_end_date")%>&nbsp;</td>
                              <td colspan="2"><%=rsCrr("career_office")%>&nbsp;</td>
                              <td colspan="2"><%=rsCrr("career_dept")%>&nbsp;</td>
                              <td colspan="1"><%=rsCrr("career_position")%>&nbsp;</td>
                              <td colspan="3"><%=rsCrr("career_task")%>&nbsp;</td>
                              <td class="right"><%=rsCrr("career_seq")%>&nbsp;</td>
							  <td>
								<a href="#" onClick="pop_Window('/person/insa_career_add.asp?career_empno=<%=rsCrr("career_empno")%>&career_seq=<%=rsCrr("career_seq")%>&emp_name=<%=in_name%>&u_type=U','경력사항 변경','scrollbars=yes,width=750,height=300')">수정</a>
							  </td>
							</tr>
						<%
								rsCrr.MoveNext()
							Loop
						End If
						rsCrr.Close() : Set rsCrr = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="pop_Window('/person/insa_career_add.asp?career_empno=<%=in_empno%>&emp_name=<%=in_name%>','경력사항 추가','scrollbars=yes,width=750,height=300')" class="btnType04">경력등록</a>
					</div>
                    </td>
			      </tr>
				</table>
                <input type="hidden" name="career_empno" value="<%=in_empno%>"/>
			</form>
		</div>
	</div>
	</body>
</html>
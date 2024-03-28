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
Dim title_line, rsFamily, arrFamily

emp_no = f_Request("emp_no")

title_line = "◈ 가족 사항 ◈"

objBuilder.Append "CALL USP_INSA_FAMILY_INFO('"&emp_no&"')"

Call Rs_Open(rsFamily, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsFamily.EOF Then
	arrFamily = rsFamily.getRows()
End If

Call Rs_Close(rsFamily)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<style type="text/css">
		.no-input{
			color:gray;
			background-color:#E0E0E0;
			border:1px solid #999999;
		}
	</style>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:60px;" class="no-input" readonly/>
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%Call EmpInfo_Name(emp_no)%>" style="width:100px;" class="no-input" readonly/>
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="6%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">관계</th>
                                <th scope="col">성&nbsp;&nbsp;명</th>
                                <th scope="col">생년월일</th>
                                <th scope="col">직&nbsp;&nbsp;&nbsp;업</th>
                                <th scope="col">전화번호</th>
                                <th scope="col">주민번호</th>
                                <th scope="col">동거여부</th>
 							</tr>
						</thead>
						<tbody>
						<%
						DBConn.Close() : Set DBConn = Nothing
						Dim i

						If IsArray(arrFamily) Then
							Dim family_rel, family_name, family_birthday, family_birthday_id, family_job
							Dim family_tel_ddd, family_tel_no1, family_tel_no2, family_person1, family_person2
							Dim family_live

							For i = LBound(arrFamily) To UBound(arrFamily, 2)
								family_rel = arrFamily(0, i)
								family_name = arrFamily(1, i)
								family_birthday = arrFamily(2, i)
								family_birthday_id = arrFamily(3, i)
								family_job = arrFamily(4, i)
								family_tel_ddd = arrFamily(5, i)
								family_tel_no1 = arrFamily(6, i)
								family_tel_no2 = arrFamily(7, i)
								family_person1 = arrFamily(8, i)
								family_person2 = arrFamily(9, i)
								family_live = arrFamily(10, i)
						%>
							<tr>
								<td><%=family_rel%>&nbsp;</td>
								<td><%=family_name%>&nbsp;</td>
                                <td><%=family_birthday%>&nbsp;(<%=family_birthday_id%>)&nbsp;</td>
                                <td><%=family_job%>&nbsp;</td>
                                <td><%=family_tel_ddd%>-<%= family_tel_no1%>-<%=family_tel_no2%>&nbsp;</td>
                                <td><%=family_person1%>-<%= family_person2%>&nbsp;</td>
                                <td style="border-bottom:1px solid #e3e3e3;"><%=family_live%>&nbsp;</td>
							</tr>
						<%
							Next
						Else
						%>
							<tr>
								<td class="first" colspan="6" style="height:30px;">조회된 내역이 없습니다.</td>
							</tr>
                        <%
						End If
						%>
						</tbody>
					</table>
				</div>
			</div>
		</div>
		<br/>
		<div align="right">
			<a href="#" class="btnType04" onclick="close_win();" >닫기</a>&nbsp;&nbsp;
		</div>
		<br/>
	</body>
</html>
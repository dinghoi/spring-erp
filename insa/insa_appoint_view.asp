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
Dim title_line, rsApp, arrApp

emp_no = f_Request("emp_no")

objBuilder.Append "CALL USP_INSA_APPOINT_INFO('"&emp_no&"')"

Call Rs_Open(rsApp, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsApp.EOF Then
	arrApp = rsApp.getRows()
End If

Call Rs_Close(rsApp)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
        						<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:80px;" class="no-input" readonly/>
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%Call EmpInfo_Name(emp_no)%>" style="width:80px;" class="no-input" readonly/>
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="9%" >
							<col width="10%" >
							<col width="9%" >
							<col width="9%" >
							<col width="10%" >
                            <col width="9%" >
                            <col width="*" >
						</colgroup>
						<thead>
                            <tr>
				                <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령일</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령구분</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령유형</th>
                                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령전</th>
				                <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령후</th>
			                </tr>
                            <tr>
                                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">발령내용</th>
                            </tr>
						</thead>
						<tbody>
						<%
							DBConn.Close() : Set DBConn = Nothing
							Dim i

						    If IsArray(arrApp) Then
								Dim app_date, app_id, app_id_type, app_to_company, app_to_orgcode
								Dim app_to_org, app_to_grade, app_to_position, app_be_company, app_be_orgcode
								Dim app_be_org, app_be_grade, app_be_position, app_start_date, app_finish_date
								Dim app_be_enddate, app_reward, app_comment

								For i = LBound(arrApp) To UBound(arrApp, 2)
									app_date = arrApp(0, i)
									app_id = arrApp(1, i)
									app_id_type = arrApp(2, i)
									app_to_company = arrApp(3, i)
									app_to_orgcode = arrApp(4, i)
									app_to_org = arrApp(5, i)
									app_to_grade = arrApp(6, i)
									app_to_position = arrApp(7, i)
									app_be_company = arrApp(8, i)
									app_be_orgcode = arrApp(9, i)
									app_be_org = arrApp(10, i)
									app_be_grade = arrApp(11, i)
									app_be_position = arrApp(12, i)
									app_start_date = arrApp(13, i)
									app_finish_date = arrApp(14, i)
									app_be_enddate = arrApp(15, i)
									app_reward = arrApp(16, i)
									app_comment = arrApp(17, i)
						%>
							<tr>
								<td><%=app_date%>&nbsp;</td>
								<td><%=app_id%>&nbsp;</td>
                                <td><%=app_id_type%>&nbsp;</td>
                                <td><%=app_to_company%>&nbsp;</td>
                                <td><%=app_to_orgcode%>)<%=app_to_org%>&nbsp;</td>
                                <td><%=app_to_grade%>-<%=app_to_position%>&nbsp;</td>
                                <td><%=app_be_company%>&nbsp;</td>
                                <td><%=app_be_orgcode%>)<%=app_be_org%>&nbsp;</td>
                                <td><%=app_be_grade%>-<%=app_be_position%>&nbsp;</td>
                                <td class="left"><%=app_start_date%>&nbsp;-&nbsp;<%=app_finish_date%>&nbsp;<%=app_be_enddate%>&nbsp;<%=app_reward%>&nbsp;:&nbsp;<%=app_comment%>&nbsp;</td>
							</tr>
						<%
							Next
						  Else
						%>
							<tr>
								<td class="first" colspan="10" style="font-weight:bold;">해당 내역이 없습니다</td>
							</tr>
                        <%
						  End If
						%>

						</tbody>
					</table>
				</div>
			</div>
	   </div>
        <br>
        <div align="right">
            <a href="#" class="btnType04" onclick="close_win();">닫기</a>&nbsp;&nbsp;
        </div>
        <br>
	</body>
</html>
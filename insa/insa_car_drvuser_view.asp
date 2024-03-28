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
Dim car_no, car_name, car_year, car_reg_date, title_line
Dim page, start_page, stpage, pgsize, be_pg, total_page
Dim rsCount, total_record, rsCarDrv, str_param

car_no = f_Request("car_no")
car_name = f_Request("car_name")
car_year = f_Request("car_year")
car_reg_date = f_Request("car_reg_date")
page = f_Request("page")

title_line = " 차량 운행자 현황 "
be_pg = "/insa/insa_car_drvuser_view.asp"
pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

'Sql = "SELECT count(*) FROM car_drive_user where use_car_no = '"&car_no&"'"
objBuilder.Append "SELECT COUNT(*) FROM car_drive_user "
objBuilder.Append "WHERE use_car_no = '"&car_no&"' "

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

str_param = "&car_no="&car_no&"&car_name="&car_name&"&car_year="&car_year&"&car_reg_date="&car_reg_date

'sql = "select * from car_drive_user where use_car_no = '" + car_no + "' ORDER BY use_car_no,use_date,use_owner_emp_no DESC limit "& stpage & "," &pgsize
objBuilder.Append "SELECT use_date, use_emp_name, use_owner_emp_no, use_company, "
objBuilder.Append "	use_org_name, use_org_code, use_emp_grade, use_end_date "
objBuilder.Append "FROM car_drive_user "
objBuilder.Append "WHERE use_car_no = '"&car_no&"' "
objBuilder.Append "ORDER BY use_car_no, use_date, use_owner_emp_no DESC "
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rsCarDrv = Server.CreateObject("ADODB.RecordSet")
rsCarDrv.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<!--<form action="insa_car_drvuser_view.asp?car_no=<%'=car_no%>" method="post" name="frm">-->
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>차량번호 : </strong>
								<label>
        						<input name="in_carno" type="text" id="in_carno" value="<%=car_no%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            <strong>차종/연식/취득일 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=car_name%>" style="width:100px; text-align:left" readonly="true">
                                -
                                <input name="in_year" type="text" id="in_year" value="<%=car_year%>" style="width:70px; text-align:left" readonly="true">
                                 -
                                <input name="car_reg_date" type="text" id="car_reg_date" value="<%=car_reg_date%>" style="width:70px; text-align:left" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="12%" >
                            <col width="10%" >
                            <col width="12%" >
                            <col width="8%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">시작일</th>
                                <th scope="col">운행자</th>
                                <th scope="col">소속회사</th>
                                <th scope="col">부서</th>
                                <th scope="col">직위</th>
                                <th scope="col">종료일</th>
							</tr>
						</thead>
						<tbody>
						<%
							Do Until rsCarDrv.EOF Or rsCarDrv.BOF
						%>
							<tr>
								<td><%=rsCarDrv("use_date")%>&nbsp;</td>
								<td><%=rsCarDrv("use_emp_name")%>(<%=rsCarDrv("use_owner_emp_no")%>)&nbsp;</td>
                                <td><%=rsCarDrv("use_company")%>&nbsp;</td>
                                <td><%=rsCarDrv("use_org_name")%>(<%=rsCarDrv("use_org_code")%>_&nbsp;</td>
                                <td><%=rsCarDrv("use_emp_grade")%>&nbsp;</td>
                                <td><%=rsCarDrv("use_end_date")%>&nbsp;</td>
							</tr>
							<%
								rsCarDrv.MoveNext()
							Loop
							rsCarDrv.Close() : Set rsCarDrv = Nothing
							%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<%
					'page navigator[허정호_20210721]
					Call Page_Navi(page, be_pg, str_param, total_page)
					%>
                    </td>
				    <td width="20%">
					<div align=right>
						<a href="#" class="btnType04" onclick="javascript:toclose();" >닫기</a>&nbsp;&nbsp;
					</div>
                    </td>
			      </tr>
			  </table>
         </div>
		<!--</form>-->
	  </div>
	</body>
</html>
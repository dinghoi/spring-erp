<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### 작업 내역
'===================================================
' 허정호_20210721 :
'	- 신규 페이지 작성 및 코드 정리

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
Dim car_no, car_name, car_year, car_reg_date, str_param
Dim pgsize, page, start_page, stpage, be_pg, total_page
Dim rsCount, total_record, title_line, rsIns

car_no = f_Request("car_no")
car_name = f_Request("car_name")
car_year = f_Request("car_year")
car_reg_date = f_Request("car_reg_date")
page = f_Request("page")

title_line = " 차량 보험가입 현황 "
be_pg = "/insa/insa_car_ins_view.asp"
pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
str_param = "&car_no="&car_no&"&car_name="&car_name&"&car_year="&car_year&"&car_reg_date="&car_reg_date

'Sql = "SELECT count(*) FROM car_insurance where ins_car_no = '"&car_no&"'"
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM car_insurance "
objBuilder.Append "where ins_car_no = '"&car_no&"' "

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

'sql = "select * from car_insurance where ins_car_no = '" + car_no + "' ORDER BY ins_car_no,ins_date DESC limit "& stpage & "," &pgsize
objBuilder.Append "SELECT ins_date, ins_company, ins_last_date, ins_amount, ins_man1, "
objBuilder.Append "	ins_man2, ins_object, ins_self, ins_injury, ins_self_car,"
objBuilder.Append "	ins_age, ins_scramble, ins_contract_yn, ins_comment "
objBuilder.Append "FROM car_insurance "
objBuilder.Append "WHERE ins_car_no = '"&car_no&"' "
objBuilder.Append "ORDER BY ins_car_no,ins_date DESC "
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rsIns = Server.CreateObject("ADODB.RecordSet")
rsIns.Open objBuilder.ToString(), DBConn, 1
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
				<form action="insa_car_ins_view.asp?car_no=<%=car_no%>" method="post" name="frm">
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
                                <input name="in_year" type="text" id="in_year" value="<%=car_year%>" style="width:100px; text-align:left" readonly="true">
                                 -
                                <input name="car_reg_date" type="text" id="car_reg_date" value="<%=car_reg_date%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="10%" >
                            <col width="6%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">가입일</th>
                                <th scope="col">보험사</th>
                                <th scope="col">보험기간</th>
                                <th scope="col">보험료</th>
                                <th scope="col">대인1</th>
                                <th scope="col">대인2</th>
                                <th scope="col">대물</th>
                                <th scope="col">자기보험</th>
                                <th scope="col">무상해</th>
                                <th scope="col">자차</th>
                                <th scope="col">연령</th>
                                <th scope="col">긴급<br>출동</th>
                                <th scope="col">계약내용</th>
 							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsIns.EOF or rsIns.BOF
						%>
							<tr>
								<td><%=rsIns("ins_date")%>&nbsp;</td>
								<td><%=rsIns("ins_company")%>&nbsp;</td>
                                <td><%=rsIns("ins_last_date")%>&nbsp;</td>
                                <td><%=FormatNumber(rsIns("ins_amount"), 0)%>&nbsp;</td>
                                <td><%=rsIns("ins_man1")%>&nbsp;</td>
                                <td><%=rsIns("ins_man2")%>&nbsp;</td>
                                <td><%=rsIns("ins_object")%>&nbsp;</td>
                                <td><%=rsIns("ins_self")%>&nbsp;</td>
                                <td><%=rsIns("ins_injury")%>&nbsp;</td>
                                <td><%=rsIns("ins_self_car")%>&nbsp;</td>
                                <td><%=rsIns("ins_age")%>&nbsp;</td>
                                <td><%=rsIns("ins_scramble")%>&nbsp;</td>
							<%If rsIns("ins_contract_yn") = "Y" Then %>
                                <td>계약내용포함&nbsp;</td>
							<%Else %>
                                <td>계약내용미포함(<%=rsIns("ins_comment")%>)&nbsp;</td>
							<%End If %>
							</tr>
						<%
							rsIns.MoveNext()
						Loop
						rsIns.Close() : Set rsIns = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<%
					'page navigator[허정호_20210720]
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
	</form>
	  </div>
	</body>
</html>


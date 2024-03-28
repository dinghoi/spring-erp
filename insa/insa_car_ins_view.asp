<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### �۾� ����
'===================================================
' ����ȣ_20210721 :
'	- �ű� ������ �ۼ� �� �ڵ� ����

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

title_line = " ���� ���谡�� ��Ȳ "
be_pg = "/insa/insa_car_ins_view.asp"
pgsize = 10 ' ȭ�� �� ������

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
		<title>�λ�޿� �ý���</title>
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
					<legend>��ȸ����</legend>
					<dl>
                        <dd>
                            <p>
							<strong>������ȣ : </strong>
								<label>
        						<input name="in_carno" type="text" id="in_carno" value="<%=car_no%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            <strong>����/����/����� : </strong>
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
                                <th class="first" scope="col">������</th>
                                <th scope="col">�����</th>
                                <th scope="col">����Ⱓ</th>
                                <th scope="col">�����</th>
                                <th scope="col">����1</th>
                                <th scope="col">����2</th>
                                <th scope="col">�빰</th>
                                <th scope="col">�ڱ⺸��</th>
                                <th scope="col">������</th>
                                <th scope="col">����</th>
                                <th scope="col">����</th>
                                <th scope="col">���<br>�⵿</th>
                                <th scope="col">��೻��</th>
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
                                <td>��೻������&nbsp;</td>
							<%Else %>
                                <td>��೻�������(<%=rsIns("ins_comment")%>)&nbsp;</td>
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
					'page navigator[����ȣ_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)
					%>
                    </td>
				    <td width="20%">
					<div align=right>
						<a href="#" class="btnType04" onclick="javascript:toclose();" >�ݱ�</a>&nbsp;&nbsp;
					</div>
                    </td>
			      </tr>
			  </table>
         </div>
	</form>
	  </div>
	</body>
</html>


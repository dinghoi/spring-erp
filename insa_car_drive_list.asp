<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim as_process
Dim field_check
Dim field_view
Dim win_sw

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	field_check=Request("field_check")
	field_view=Request("field_view")
  else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If to_date = "" or from_date = "" Then
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-7),1,10)
	field_check = "total"
End If

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * from car_drv"
date_sql = " where (drv_date >= '" + from_date  + "' and drv_date <= '" + to_date  + "')"

if field_check <> "total" then
	field_sql = " and ( " + field_check + " = '" + field_view + "' ) "
  else
  	field_sql = " "
end if
order_sql = " ORDER BY drv_date ASC"

sql = "select count(*) from car_drv" + date_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "������������"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
		<script type="text/javascript">
			function getPageCode(){
				return "8 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.field_check.value == "") {
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_drive_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>������ : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>������ : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <label>
								<strong>�ʵ�����</strong>
                                <select name="field_check" id="field_check" style="width:70px">
                                  <option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                  <option value="user_name" <% if field_check = "user_name" then %>selected<% end if %>>�۾���</option>
                                  <option value="belong" <% if field_check = "belong" then %>selected<% end if %>>�Ҽ�</option>
                                  <option value="acpt_no" <% if field_check = "acpt_no" then %>selected<% end if %>>AS��ȣ</option>
                                  <option value="work_item" <% if field_check = "work_item" then %>selected<% end if %>>�׸�</option>
                                  <option value="cancel" <% if field_check = "cancel" then %>selected<% end if %>>��Ұ�</option>
                                  <option value="company" <% if field_check = "company" then %>selected<% end if %>>ȸ�纰</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:80px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="10%" >
							<col width="10%" >
							<col width="5%" >
							<col width="10%" >
							<col width="10%" >
							<col width="5%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">��������</th>
								<th rowspan="2" scope="col">������</th>
								<th rowspan="2" scope="col">����</th>
								<th rowspan="2" scope="col">����<br>/<br>����<br>����</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� ��</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� ��</th>
								<th rowspan="2" scope="col">�������</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� �� </th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">��ü��</th>
								<th scope="col">�����</th>
								<th scope="col">���KM</th>
								<th scope="col">��ü��</th>
								<th scope="col">������</th>
								<th scope="col">����KM</th>
								<th scope="col">���߱���</th>
								<th scope="col">�����ݾ�</th>
								<th scope="col">������</th>
								<th scope="col">�����</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							sql="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
							set rs_memb=dbconn.execute(sql)
						
							if	rs_memb.eof or rs_memb.bof then
								user_name = "�̵��"
							  else
								user_name = rs_memb("user_name")
							end if
							rs_memb.close()
						%>
							<tr>
								<td class="first"><%=rs("drv_date")%></td>
								<td><%=user_name%></td>
								<td><%=rs("car_owner")%></td>
								<td>
								<% if rs("car_owner") = "���߱���" then %>
								<%=rs("transit")%>
								<%   else	%>                                
								<%=rs("oil_kind")%>
								<% end if %>
                                </td>
								<td><%=rs("start_company")%></td>
								<td><%=rs("start_point")%></td>
								<td class="right"><%=formatnumber(rs("start_km"),0)%></td>
								<td><%=rs("end_company")%></td>
								<td><%=rs("end_point")%></td>
								<td class="right"><%=formatnumber(rs("end_km"),0)%></td>
								<td><%=rs("drv_memo")%></td>
								<td class="right"><%=formatnumber(rs("fare"),0)%></td>
								<td class="right"><%=formatnumber(rs("oil_price"),0)%></td>
								<td class="right"><%=formatnumber(rs("parking"),0)%></td>
								<td class="right"><%=formatnumber(rs("toll"),0)%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="excel_down_condi.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_car_drive_list.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_car_drive_list.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_car_drive_list.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="insa_car_drive_list.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_car_drive_list.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="10%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('car_drive_add.asp','car_drive_add_popup','scrollbars=yes,width=750,height=420')" class="btnType04">������������</a>
					</div>                  
                    </td>
				    <td width="10%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('mass_transit_add.asp','mass_transit_add_popup','scrollbars=yes,width=750,height=300')" class="btnType04">�������</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>


<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_car_drv_list.asp"

from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date")
end if

if view_condi = "" then
	view_condi = "��ü"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

pgsize = 10 ' ȭ�� �� ������ 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "��ü" then
   Sql = "select count(*) from transit_cost where run_date >= '"+from_date+"' and run_date <= '"+to_date+"'"
   else  
   Sql = "select count(*) from transit_cost where car_no='"+view_condi+"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"'"
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_condi = "��ü" then
   Sql = "select * from transit_cost where run_date >= '"+from_date+"' and run_date <= '"+to_date+"' "
   else  
   Sql = "select * from transit_cost where car_no = '"+view_condi+"' and run_date >= '"+from_date+"' and run_date <= '"+to_date+"' "
end If

	'//2017-09-07 ���ļ��� ����
   'Sql = Sql & " ORDER BY car_no,run_date,run_seq DESC "
   Sql = Sql & " ORDER BY car_no,run_date,run_seq ASC"
   Sql = Sql & " limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - ���� ������Ȳ "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "7 1";
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
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_drv_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                               <strong>������ȣ : </strong>
                              <%
								Sql="select * from car_info where (end_date = '1900-01-01' or isNull(end_date)) ORDER BY car_no ASC"
	                            rs_car.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
                                  <option value="��ü" <%If view_condi = "��ü" then %>selected<% end if %>>��ü</option>
                			  <% 
								do until rs_car.eof 
			  				  %>
                					<option value='<%=rs_car("car_no")%>' <%If view_condi = rs_car("car_no") then %>selected<% end if %>><%=rs_car("car_no")%></option>
                			  <%
									rs_car.movenext()  
								loop 
								rs_car.Close()
							  %>
            					</select>
                                </label>
								<label>
								<strong>������ : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>������ : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
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
                            <col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="8%" >
							<col width="9%" >
							<col width="5%" >
							<col width="8%" >
							<col width="*" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="4%" >
                		</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">������ȣ</th>
                                <th rowspan="2" scope="col">��������</th>
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
                            emp_no = rs("mg_ce_id")
							Sql = "select * from emp_master where emp_no = '"+emp_no+"'"
	                        Set Rs_emp = DbConn.Execute(SQL)
	                        if not Rs_emp.EOF or not Rs_emp.BOF then
			                       drv_owner_emp_name = rs_emp("emp_name")
                               else
                                   drv_owner_emp_name = emp_no
							end if
							
							if rs("start_km") = "" or isnull(rs("start_km")) then
								start_view = 0
							  else
							  	start_view = rs("start_km")
							end if
							if rs("end_km") = "" or isnull(rs("end_km")) then
								end_view = 0
							  else
							  	end_view = rs("end_km")
							end if
							run_km = rs("far")
	           			%>
							<tr>
								<td class="first"><%=rs("car_no")%></td>
                                <td><%=rs("run_date")%></td>
								<td><%=drv_owner_emp_name%></td>
								<td><%=rs("car_owner")%></td>
								<td>
								<% if rs("car_owner") = "���߱���" then %>
								       <%=rs("transit")%>
								<%   else	%>                                
								       <%=rs("oil_kind")%>
								<% end if %>
                                </td>
								<td><%=rs("start_company")%>&nbsp;</td>
								<td class="left"><%=rs("start_point")%></td>
								<td class="right"><%=formatnumber(start_view,0)%></td>
								<td><%=rs("end_company")%>&nbsp;</td>
								<td class="left"><%=rs("end_point")%></td>
								<td class="right"><%=formatnumber(end_view,0)%></td>
								<td><%=rs("run_memo")%></td>
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
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_car_drv.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_car_drv_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_car_drv_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_car_drv_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_car_drv_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_car_drv_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>

			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>


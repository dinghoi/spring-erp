<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
car_no = request("car_no")

car_name = ""
car_year = ""
oil_kind = ""
car_owner = ""
insurance_company = ""
insurance_date = ""
insurance_amt = 0
buy_gubun = "����"
rental_company = ""
car_reg_date = ""
car_use_dept = ""
car_company = ""
car_use = ""
owner_emp_no = ""
owner_emp_name = ""
emp_name = ""
emp_grade = ""
start_date = ""
end_date = ""
last_km = 0
last_check_date = ""
car_status = ""
car_comment = ""

view_condi = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "���� ���"

if u_type = "U" then

	sql = "select * from car_info where car_no = '" + car_no + "'"
	set rs = dbconn.execute(sql)

    car_no = rs("car_no")
	car_old_no = rs("car_no")
    car_name = rs("car_name")
	
	car_year = rs("car_year")
    oil_kind = rs("oil_kind")
    car_owner = rs("car_owner")
    insurance_company = rs("insurance_company")
    insurance_date = rs("insurance_date")
    insurance_amt = rs("insurance_amt")
    buy_gubun = rs("buy_gubun")
    rental_company = rs("rental_company")
    car_reg_date = rs("car_reg_date")
    car_use_dept = rs("car_use_dept")
    car_company = rs("car_company")
    car_use = rs("car_use")
    owner_emp_no = rs("owner_emp_no")
	owner_emp_name = rs("owner_emp_name")
    start_date = rs("start_date")
    end_date = rs("end_date")
    last_km = rs("last_km")
    last_check_date = rs("last_check_date")
    car_status = rs("car_status")
    car_comment = rs("car_comment")
	if rs("last_check_date") = "1900-01-01"  then
           last_check_date = ""
	   else 
	       last_check_date = rs("last_check_date")
    end if
    if rs("end_date") = "1900-01-01" then
           end_date = ""
	   else 
           end_date = rs("end_date")
	end if
	
	owner_emp_name = ""
    owner_emp_no = rs("owner_emp_no")
	if owner_emp_name = "" or isnull(owner_emp_name) then
	     Sql="select * from emp_master where emp_no = '"&owner_emp_no&"'"
	     Set rs_emp=DbConn.Execute(Sql)
	     if not rs_emp.eof then
		        owner_emp_name = rs_emp("emp_name")
		        emp_grade = rs_emp("emp_job")
		        emp_org_name = rs_emp("emp_org_name")
	        else 
	            owner_emp_name = rs("owner_emp_name")
		        emp_grade = ""
		 end if
    end if
	if car_use_dept = "" then
	   car_use_dept = emp_org_name
	end if
	
	rs.close()

	title_line = "���� ����"
end if
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=car_reg_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=last_check_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=car_year%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.car_no.value =="" ) {
					alert('������ȣ�� �Է��ϼ���');
					frm.car_no.focus();
					return false;}
				if(document.frm.car_name.value =="") {
					alert('������ �Է��ϼ���');
					frm.car_name.focus();
					return false;}
				if(document.frm.oil_kind.value =="") {
					alert('������ �����ϼ���');
					frm.oil_kind.focus();
					return false;}			
				if(document.frm.car_owner.value =="") {
					alert('�����ڸ� �����ϼ���');
					frm.car_owner.focus();
					return false;}			
				if(document.frm.car_reg_date.value =="") {
					alert('����������� �Է��ϼ���');
					frm.car_reg_date.focus();
					return false;}			
				if(document.frm.owner_emp_no.value =="" ) {
					alert('�����˻��� �ϼ���');
					frm.emp_name.focus();
					return false;}
			
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
			
			function num_chk(txtObj){
				lst_km = parseInt(document.frm.last_km.value.replace(/,/g,""));		
				lst_km = String(lst_km);
				num_len = lst_km.length;
				sil_len = num_len;
				lst_km = String(lst_km);
				if (lst_km.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) lst_km = lst_km.substr(0,num_len -3) + "," + lst_km.substr(num_len -3,3);
				if (sil_len > 6) lst_km = lst_km.substr(0,num_len -6) + "," + lst_km.substr(num_len -6,3) + "," + lst_km.substr(num_len -2,3);

				document.frm.last_km.value = lst_km; 

				if (txtObj.value.length >= 2) {
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}
				if (txtObj.value.length<5) {
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
				}
				var num = txtObj.value;
				if (num == "--" ||  num == "." ) num = "";
				if (num != "" ) {
					temp=new String(num);
					if(temp.length<1) return "";
					
					// ����ó��
					if(temp.substr(0,1)=="-") minus="-";
					else minus="";
					
					// �Ҽ�������ó��
					dpoint=temp.search(/\./);
					
					if(dpoint>0)
					{
					// ù��° ������ .�� �������� �ڸ��� ���������� ���� ����
					dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
					temp=temp.substr(0,dpoint);
					}else dpointVa="";
					
					// �����ܹ̿��� ����
					temp=temp.replace(/\D/g,"");
					zero=temp.search(/[1-9]/);
					
					if(zero==-1) return "";
					else if(zero!=0) temp=temp.substr(zero);
					
					if(temp.length<4) return minus+temp+dpointVa;
					buf="";
					while (true)
					{
					if(temp.length<3) { buf=temp+buf; break; }
				
					buf=","+temp.substr(temp.length-3)+buf;
					temp=temp.substr(0, temp.length-3);
					}
					if(buf.substr(0,1)==",") buf=buf.substr(1);
				
					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;
				}else txtObj.value = "0";					
			}	
			
		function delcheck() 
				{
				a=confirm('���� �����Ͻðڽ��ϱ�?')
				if (a==true) {
					document.frm.method = "post";
					document.frm.action = "insa_car_info_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_info_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
                                <th class="first">������ȣ</th>
								<td class="left">
                                <input name="car_no" type="text" value="<%=car_no%>" style="width:150px" onKeyUp="checklength(this,20)"></td>
								<th>����</th>
								<td class="left">
                                <input name="car_name" type="text" value="<%=car_name%>" style="width:150px" onKeyUp="checklength(this,30)"></td>
							</tr>
                           	<tr>
								<th class="first">��������</th>
								<td colspan="3" class="left"><input name="car_year" type="text" value="<%=car_year%>" style="width:70px" id="datepicker3"></td>
							</tr>             
							<tr>
								<th class="first">����</th>
								<td class="left">
                                <select name="oil_kind" id="oil_kind" style="width:150px">
								  <option value="">����</option>
								  <option value="�ֹ���" <%If oil_kind = "�ֹ���" then %>selected<% end if %>>�ֹ���</option>
								  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
								  <option value="����" <%If oil_kind = "����" then %>selected<% end if %>>����</option>
							    </select>
                                </td>
								<th>����</th>
								<td class="left"><select name="car_owner" id="car_owner" style="width:150px">
								  <option value="">����</option>
								  <option value="ȸ��" <%If car_owner = "ȸ��" then %>selected<% end if %>>ȸ��</option>
								  <option value="����" <%If car_owner = "����" then %>selected<% end if %>>����</option>
							    </select></td>
							</tr>
							<tr>
								<th class="first">���ű���</th>
								<td class="left">
                                <input type="radio" name="buy_gubun" value="����" <% if buy_gubun = "����" then %>checked<% end if %> style="width:40px" id="Radio1">����
                                <input type="radio" name="buy_gubun" value="����" <% if buy_gubun = "����" then %>checked<% end if %> style="width:40px" id="Radio2">����
                                <input type="radio" name="buy_gubun" value="��Ʈ" <% if buy_gubun = "��Ʈ" then %>checked<% end if %> style="width:40px" id="Radio2">��Ʈ
                                </td>
								<th>��Ʈȸ��</th>
                                <td class="left">
								<input name="rental_company" type="text" value="<%=rental_company%>" style="width:150px" onKeyUp="checklength(this,30)"></td>
							</tr>
							<tr>
								<th class="first">�Ҽ�ȸ��</th>
								<td class="left">
                            <%
					            Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = 'ȸ��') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1
				         	%>
					            <select name="car_company" id="car_company" type="text" style="width:150px">
                                  <option value="" <% if car_company = "" then %>selected<% end if %>>����</option>
                		    <% 
						 		do until rs_org.eof 
			  			    %>
                				  <option value='<%=rs_org("org_name")%>' <%If car_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			<%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							%>
            		            </select>    
                                </td>
								<th>���������</th>
								<td class="left"><input name="car_reg_date" type="text" value="<%=car_reg_date%>" style="width:70px" id="datepicker"></td>
							</tr>    
							<tr>
								<th class="first">�뵵</th>
								<td class="left">
                                <input name="car_use" type="text" value="<%=car_use%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
								<th>���μ�</th>
								<td class="left">
                                <input name="car_use_dept" type="text" id="car_use_dept" style="width:80px" value="<%=car_use_dept%>" readonly="true">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="car"%>&mg_level=<%=org_level%>&view_condi=<%=view_condi%>','orgselect','scrollbars=yes,width=850,height=400')">�μ�ã��</a>
                                </td>
							</tr>                                                    
							<tr>
								<th class="first">������</th>
								<td colspan="3" class="left">
                                <input name="emp_name" type="text" id="emp_name" style="width:80px" value="<%=owner_emp_name%>" readonly="true">
                                <input name="emp_grade" type="text" id="emp_grade" style="width:80px" value="<%=emp_grade%>" readonly="true">
                                <input name="owner_emp_no" type="text" id="owner_emp_no" style="width:80px" value="<%=owner_emp_no%>" readonly="true">
                            <% if u_type = "" then %>
                                <a href="#" class="btnType03" onClick="pop_Window('insa_emp_select.asp?gubun=<%="car"%>&view_condi=<%=view_condi%>','orgempselect','scrollbars=yes,width=600,height=400')">�����˻�</a>
                            <% end if %>
                                <input name="emp_company" type="hidden" id="emp_company" value="<%=emp_company%>">
                                <input name="emp_org_code" type="hidden" id="emp_org_code" value="<%=emp_org_code%>">
                                <input name="emp_org_name" type="hidden" id="emp_org_name" value="<%=emp_org_name%>">
							</tr>
							<tr>
								<th class="first">��������</th>
								<td class="left">
                                <input name="car_status" type="text" value="<%=car_status%>" style="width:150px" onKeyUp="checklength(this,20)"></td>
								<th>��������</th>
								<td class="left">
                                <input name="car_comment" type="text" value="<%=car_comment%>" style="width:170px" onKeyUp="checklength(this,50)"></td>
							</tr>
                        	<tr>
								<th class="first">������km</th>
								<td class="left">
                                <input name="last_km" type="text" id="last_km" style="width:70px;text-align:right" value="<%=formatnumber(last_km,0)%>" onKeyUp="num_chk(this);"></td>
								<th>�����˻���</th>
                                <td class="left"><input name="last_check_date" type="text" value="<%=last_check_date%>" style="width:70px" id="datepicker1"></td>
							</tr>
                        	<tr>
								<th class="first">ó������</th>
								<td colspan="3" class="left"><input name="end_date" type="text" value="<%=end_date%>" style="width:70px" id="datepicker2"></td>
							</tr>                            
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
            <% if u_type = "U" and user_id = "100787" or user_id = "900002" Or user_id = "101168" Or user_id = "101100" then	%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();"></span>
			<% end if	%>                           
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="car_old_no" value="<%=car_old_no%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>


<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

to_date=Request.form("to_date")
team = request.form("team")

If to_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	team = "��ü"
End If

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_tot = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select memb.user_id,memb.team,memb.user_name,memb.reside from as_acpt inner join memb on as_acpt.mg_ce_id = memb.user_id "
sql = sql + " Where (as_acpt.as_process='����' or as_acpt.as_process='����' or as_acpt.as_process='�԰�')"
sql = sql + " GROUP BY memb.user_id,memb.team,memb.user_name,memb.reside Order By memb.team, memb.user_name Asc"

Rs.Open Sql, Dbconn, 1

title_line = "CE�� ������ ��ó����Ȳ"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
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
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.to_date.value == "") {
					alert ("�������� �Է��ϼ���");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/ce_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=ce_mi.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								<strong>������ : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                    			<a href="ce_mi_excel.asp?to_date=<%=to_date%>" class="btnType04">�����ٿ�ε�</a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="4%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">�Ҽ�</th>
								<th scope="col" rowspan="2">CE��</th>
								<th scope="col" rowspan="2">����</th>
								<th scope="col" colspan="13" style=" border-bottom:1px solid #e3e3e3;">�����ϱ��� ��ó��</th>
								<th scope="col" colspan="13" style=" border-bottom:1px solid #e3e3e3;">��ü ��ó��</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">�Ұ�</th>
								<th scope="col">����</th>
								<th scope="col">�湮</th>
								<th scope="col">�԰�</th>
								<th scope="col">�ű�<br>��ġ</th>
								<th scope="col">�ż�<br>����</th>
								<th scope="col">����<br>��ġ</th>
								<th scope="col">�̼�<br>����</th>
								<th scope="col">��<br>����</th>
								<th scope="col">����<br>��</th>
								<th scope="col">ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">��Ÿ</th>
								<th scope="col">�Ұ�</th>
								<th scope="col">����</th>
								<th scope="col">�湮</th>
								<th scope="col">�԰�</th>
								<th scope="col">�ű�<br>��ġ</th>
								<th scope="col">�ż�<br>����</th>
								<th scope="col">�̼�<br>��ġ</th>
								<th scope="col">�̼�<br>����</th>
								<th scope="col">��<br>����</th>
								<th scope="col">����<br>��</th>
								<th scope="col">ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">��Ÿ</th>
							</tr>
						</thead>
						<tbody>
						<% 
                        dim day_sum(12)
                        dim month_sum(12)
                        dim day_tot(12)
                        dim month_tot(12)
                        for i = 0 to 12
                            day_sum(i) = 0
                            month_sum(i) = 0
                            day_tot(i) = 0
                            month_tot(i) = 0
                        next
						                
                        do until rs.eof                 
' ���� ��ó�� �԰�
                            sql = "select count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (as_process='�԰�') and (mg_ce_id='"+rs("user_id")+"') "
                            set rs_in=dbconn.execute(sql)
                            if rs_in.eof then
                                month_sum(3) = 0
                              else
                                month_sum(3) = cint(rs_in("end_cnt"))
                            end if
                            rs_in.close()
                
                ' ���� ��ó�� �԰�
                            sql = "select count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (as_process='�԰�') and (mg_ce_id='"+rs("user_id")+"') and (request_date <= '"+to_date+"')"
                            set rs_in=dbconn.execute(sql)
                            if rs_in.eof then
                                day_sum(3) = 0
                              else
                                day_sum(3) = cint(rs_in("end_cnt"))
                            end if
                            rs_in.close()
                
                ' ���� ������ ��ó��
                            sql = "select as_type, count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (as_process='����' or as_process='����') and (mg_ce_id='"+rs("user_id")+"') GROUP BY as_type"		
                            rs_as.Open Sql, Dbconn, 1
                            do until rs_as.eof
                                select case rs_as("as_type")
                                    case "����ó��"
                                        month_sum(1) = cint(rs_as("end_cnt"))	
                                    case "�湮ó��"
                                        month_sum(2) = cint(rs_as("end_cnt"))	
                                    case "�űԼ�ġ"
                                        month_sum(4) = cint(rs_as("end_cnt"))	
                                    case "�űԼ�ġ����"
                                        month_sum(5) = cint(rs_as("end_cnt"))	
                                    case "������ġ"
                                        month_sum(6) = cint(rs_as("end_cnt"))	
                                    case "������ġ����"
                                        month_sum(7) = cint(rs_as("end_cnt"))	
                                    case "������"
                                        month_sum(8) = cint(rs_as("end_cnt"))	
                                    case "����������"
                                        month_sum(9) = cint(rs_as("end_cnt"))	
                                    case "���ȸ��"
                                        month_sum(10) = cint(rs_as("end_cnt"))	
                                    case "��������"
                                        month_sum(11) = cint(rs_as("end_cnt"))	
                                    case "��Ÿ"
                                        month_sum(12) = cint(rs_as("end_cnt"))	
                                end select												
                                rs_as.movenext()
                            loop
                            rs_as.close()
                            
                ' ���� ������ ��ó��
                            sql = "select as_type, count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (as_process='����' or as_process='����') and (mg_ce_id='"+rs("user_id")+"') and (request_date <= '"+to_date+"') GROUP BY as_type"		
                            rs_as.Open Sql, Dbconn, 1
                            do until rs_as.eof
                                select case rs_as("as_type")
                                    case "����ó��"
                                        day_sum(1) = cint(rs_as("end_cnt"))	
                                    case "�湮ó��"
                                        day_sum(2) = cint(rs_as("end_cnt"))	
                                    case "�űԼ�ġ"
                                        day_sum(4) = cint(rs_as("end_cnt"))	
                                    case "�űԼ�ġ����"
                                        day_sum(5) = cint(rs_as("end_cnt"))	
                                    case "������ġ"
                                        day_sum(6) = cint(rs_as("end_cnt"))	
                                    case "������ġ����"
                                        day_sum(7) = cint(rs_as("end_cnt"))	
                                    case "������"
                                        day_sum(8) = cint(rs_as("end_cnt"))	
                                    case "����������"
                                        day_sum(9) = cint(rs_as("end_cnt"))	
                                    case "���ȸ��"
                                        day_sum(10) = cint(rs_as("end_cnt"))	
                                    case "��������"
                                        day_sum(11) = cint(rs_as("end_cnt"))	
                                    case "��Ÿ"
                                        day_sum(12) = cint(rs_as("end_cnt"))	
                                end select												
                                rs_as.movenext()
                            loop
                            rs_as.close() 
                
                            for i = 1 to 12
                                day_sum(0) = day_sum(0) + day_sum(i)
                                month_sum(0) = month_sum(0) + month_sum(i)
                                day_tot(0) = day_tot(0) + day_tot(i)
                                month_tot(0) = month_tot(0) + month_tot(i)			
                            next
                            for i = 1 to 12
                                day_tot(i) = day_tot(i) + day_sum(i)
                                month_tot(i) = month_tot(i) + month_sum(i)			
                            next
                
                            if day_sum(0) <> 0 or month_sum(0) <> 0 then
                                if rs("reside") = "0" then
                                    reside = "."
                                  else
                                    reside = "����"
                                end if
                    %>
							<tr>
                              <td><%=rs("team")%></td>
                              <td><%=rs("user_name")%></td>
                              <td><%=reside%></td>
                              <td bgcolor="#FFFFCA" class="right"><%=formatnumber(day_sum(0),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(1),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(2),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(3),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(4),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(5),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(6),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(7),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(8),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(9),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(10),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(11),0)%></td>
                              <td class="right"><%=formatnumber(day_sum(12),0)%></td>
                              <td bgcolor="#FFE8E8" class="right"><%=formatnumber(month_sum(0),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(1),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(2),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(3),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(4),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(5),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(6),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(7),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(8),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(9),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(10),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(11),0)%></td>
                              <td class="right"><%=formatnumber(month_sum(12),0)%></td>
							</tr>
							<%
                                end if
                                
                                for i = 0 to 12
                                    day_sum(i) = 0
                                    month_sum(i) = 0
                                next
                                rs.movenext()
                            loop
                            rs.close()
                            day_tot(0) = day_tot(1) + day_tot(2) + day_tot(3) + day_tot(4) + day_tot(5) + day_tot(6) + day_tot(7) + day_tot(8) + day_tot(9) + day_tot(10) + day_tot(11) + day_tot(12)
                            month_tot(0) = month_tot(1) + month_tot(2) + month_tot(3) + month_tot(4) + month_tot(5) + month_tot(6) + month_tot(7) + month_tot(8) + month_tot(9) + month_tot(10) + month_tot(11) + month_tot(12)
                            %>
							<tr>
                              <th colspan="3">�Ѱ�</th>
                              <th><%=formatnumber(day_tot(0),0)%></th>
                              <th><%=formatnumber(day_tot(1),0)%></th>
                              <th><%=formatnumber(day_tot(2),0)%></th>
                              <th><%=formatnumber(day_tot(3),0)%></th>
                              <th><%=formatnumber(day_tot(4),0)%></th>
                              <th><%=formatnumber(day_tot(5),0)%></th>
                              <th><%=formatnumber(day_tot(6),0)%></th>
                              <th><%=formatnumber(day_tot(7),0)%></th>
                              <th><%=formatnumber(day_tot(8),0)%></th>
                              <th><%=formatnumber(day_tot(9),0)%></th>
                              <th><%=formatnumber(day_tot(10),0)%></th>
                              <th><%=formatnumber(day_tot(11),0)%></th>
                              <th><%=formatnumber(day_tot(12),0)%></th>
                              <th><%=formatnumber(month_tot(0),0)%></th>
                              <th><%=formatnumber(month_tot(1),0)%></th>
                              <th><%=formatnumber(month_tot(2),0)%></th>
                              <th><%=formatnumber(month_tot(3),0)%></th>
                              <th><%=formatnumber(month_tot(4),0)%></th>
                              <th><%=formatnumber(month_tot(5),0)%></th>
                              <th><%=formatnumber(month_tot(6),0)%></th>
                              <th><%=formatnumber(month_tot(7),0)%></th>
                              <th><%=formatnumber(month_tot(8),0)%></th>
                              <th><%=formatnumber(month_tot(9),0)%></th>
                              <th><%=formatnumber(month_tot(10),0)%></th>
                              <th><%=formatnumber(month_tot(11),0)%></th>
                              <th><%=formatnumber(month_tot(12),0)%></th>
							</tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>


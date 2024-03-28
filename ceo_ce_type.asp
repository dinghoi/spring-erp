<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

from_date=Request.form("from_date")
to_date=Request.form("to_date")
team = "��ü"
company_sum = 0

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
End If

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name from ce_work inner join memb on ce_work.mg_ce_id=memb.user_id where (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') GROUP BY ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name Order By memb.team, memb.user_name Asc"
Rs.Open Sql, Dbconn, 1

title_line = "CE�� ���� ��Ȳ"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�ӿ� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
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
				if (document.frm.from_date.value > document.frm.to_date.value) {
					alert ("�������� �����Ϻ��� Ŭ���� �����ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
			<!--#include virtual = "/include/ceo_as_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=ceo_ce_type.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
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
							<col width="7%" >
							<col width="5%" >
							<col width="*" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">�Ҽ�</th>
								<th scope="col" rowspan="2">CE��</th>
								<th scope="col" rowspan="2">����ó</th>
								<th scope="col" colspan="14" style=" border-bottom:1px solid #e3e3e3;">
                                ������ ó�� ��Ȳ ( ��ü����/���ϱٹ����� )
                                </th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">�Ұ�</th>
								<th scope="col">����</th>
								<th scope="col">�湮</th>
								<th scope="col">�űԼ�ġ</th>
								<th scope="col">�űԼ�ġ<br>����</th>
								<th scope="col">������ġ</th>
								<th scope="col">������ġ<br>����</th>
								<th scope="col">������</th>
								<th scope="col">���� ��<br>����</th>
								<th scope="col">ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">��Ÿ</th>
								<th scope="col">�԰�<br>�Ϸ�</th>
								<th scope="col">�԰�</th>
							</tr>
						</thead>
						<tbody>
						<% 
                        dim month_sum(13)
                        dim month_tot(13)
                        dim overtime_sum(13)
                        dim overtime_tot(13)
                        for i = 0 to 13
                            month_sum(i) = 0
                            month_tot(i) = 0
                            overtime_sum(i) = 0
                            overtime_tot(i) = 0
                        next
                
						ce_cnt = 0
                        do until rs.eof 
							ce_cnt = ce_cnt + 1
				' ���� ������ ó��
                            sql = "select as_type, holiday_yn, count(*) as end_cnt from ce_work WHERE (ce_work.work_id='2') and (ce_work.mg_ce_id='"+rs("mg_ce_id")+"') and (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') GROUP BY ce_work.as_type,holiday_yn"		
                            rs_as.Open Sql, Dbconn, 1
                            do until rs_as.eof
                                select case rs_as("as_type")
                                    case "����ó��"
                                        month_sum(1) = month_sum(1) + cint(rs_as("end_cnt"))	
                                    case "�湮ó��"
                                        month_sum(2) = month_sum(2) + cint(rs_as("end_cnt"))	
                                    case "�űԼ�ġ"
                                        month_sum(3) = month_sum(3) + cint(rs_as("end_cnt"))	
                                    case "�űԼ�ġ����"
                                        month_sum(4) = month_sum(4) + cint(rs_as("end_cnt"))	
                                    case "������ġ"
                                        month_sum(5) = month_sum(5) + cint(rs_as("end_cnt"))	
                                    case "������ġ����"
                                        month_sum(6) = month_sum(6) + cint(rs_as("end_cnt"))	
                                    case "������"
                                        month_sum(7) = month_sum(7) + cint(rs_as("end_cnt"))	
                                    case "����������"
                                        month_sum(8) = month_sum(8) + cint(rs_as("end_cnt"))	
                                    case "���ȸ��"
                                        month_sum(9) = month_sum(9) + cint(rs_as("end_cnt"))	
                                    case "��������"
                                        month_sum(10) = month_sum(10) + cint(rs_as("end_cnt"))	
                                    case "��Ÿ"
                                        month_sum(11) = month_sum(11) + cint(rs_as("end_cnt"))	
                                end select												
								if rs_as("holiday_yn") = "Y" then
									select case rs_as("as_type")
										case "����ó��"
											overtime_sum(1) = cint(rs_as("end_cnt"))	
										case "�湮ó��"
											overtime_sum(2) = cint(rs_as("end_cnt"))	
										case "�űԼ�ġ"
											overtime_sum(3) = cint(rs_as("end_cnt"))	
										case "�űԼ�ġ����"
											overtime_sum(4) = cint(rs_as("end_cnt"))	
										case "������ġ"
											overtime_sum(5) = cint(rs_as("end_cnt"))	
										case "������ġ����"
											overtime_sum(6) = cint(rs_as("end_cnt"))	
										case "������"
											overtime_sum(7) = cint(rs_as("end_cnt"))	
										case "����������"
											overtime_sum(8) = cint(rs_as("end_cnt"))	
										case "���ȸ��"
											overtime_sum(9) = cint(rs_as("end_cnt"))	
										case "��������"
											overtime_sum(10) = cint(rs_as("end_cnt"))	
										case "��Ÿ"
											overtime_sum(11) = cint(rs_as("end_cnt"))	
									end select												
								end if
                                rs_as.movenext()
                            loop
                            rs_as.close()
                ' �԰��� ó�� �Ϸ�
                            sql = "select count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (in_date <> '') and (as_process='�Ϸ�') and (mg_ce_id='"+rs("mg_ce_id")+"') and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"') GROUP BY mg_ce_id"		
							Set rs_as = Dbconn.Execute (sql)
							if rs_as.eof or rs_as.bof then
								month_sum(12) = 0
							  else
                                month_sum(12) = cint(rs_as("end_cnt"))	
							end if
							rs_as.close()
                ' �԰�
                            sql = "select count(*) as end_cnt from as_acpt "
                            sql = sql + "WHERE (as_process='�԰�') and (mg_ce_id='"+rs("mg_ce_id")+"') and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"') GROUP BY mg_ce_id"		
                            rs_as.Open Sql, Dbconn, 1
							Set rs_as = Dbconn.Execute (sql)
							if rs_as.eof or rs_as.bof then
								month_sum(13) = 0
							  else
                                month_sum(13) = cint(rs_as("end_cnt"))	
							end if
							rs_as.close()
                
                            for i = 1 to 13
                                month_sum(0) = month_sum(0) + month_sum(i)
                                month_tot(0) = month_tot(0) + month_tot(i)			
                                overtime_sum(0) = overtime_sum(0) + overtime_sum(i)
                                overtime_tot(0) = overtime_tot(0) + overtime_tot(i)			
                            next
                            for i = 1 to 13
                                month_tot(i) = month_tot(i) + month_sum(i)			
                                overtime_tot(i) = overtime_tot(i) + overtime_sum(i)			
                            next
                
                            if month_sum(0) <> 0 then
								if rs("team") = "" or isnull(rs("team")) then
									org_view = rs("org_name") 
								  else
								  	org_view = rs("team")
								end if
								
                    %>
							<tr>
                              <td><%=org_view%></td>
                              <td><%=rs("user_name")%></td>
                              <td><%=rs("reside_place")%>&nbsp;</td>
					<% if company_cnt = 0 then	%>
                              <%   else	%>
                              <td bgcolor="#FFD8B0"><strong><%=company_cnt%>/<%=company_over%></strong></td>
                    <% end if	%>
                              <td bgcolor="#FFFFCA" class="right"><%=formatnumber(clng(month_sum(0)),0)%>/<%=overtime_sum(0)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(1)),0)%>/<%=overtime_sum(1)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(2)),0)%>/<%=overtime_sum(2)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(3)),0)%>/<%=overtime_sum(3)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(4)),0)%>/<%=overtime_sum(4)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(5)),0)%>/<%=overtime_sum(5)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(6)),0)%>/<%=overtime_sum(6)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(7)),0)%>/<%=overtime_sum(7)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(8)),0)%>/<%=overtime_sum(8)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(9)),0)%>/<%=overtime_sum(9)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(10)),0)%>/<%=overtime_sum(10)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(11)),0)%>/<%=overtime_sum(11)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(12)),0)%></td>
                              <td class="right"><%=formatnumber(clng(month_sum(13)),0)%></td>
							</tr>
						    <%
                                end if
                                
                                for i = 0 to 13
                                    month_sum(i) = 0
                                    overtime_sum(i) = 0
                                next
                    
                                rs.movenext()
                            loop
                            rs.close()
                            month_tot(0) = month_tot(1) + month_tot(2) + month_tot(3) + month_tot(4) + month_tot(5) + month_tot(6) + month_tot(7) + month_tot(8) + month_tot(9) + month_tot(10) + month_tot(11) + month_tot(12) + month_tot(13)
                            overtime_tot(0) = overtime_tot(1) + overtime_tot(2) + overtime_tot(3) + overtime_tot(4) + overtime_tot(5) + overtime_tot(6) + overtime_tot(7) + overtime_tot(8) + overtime_tot(9) + overtime_tot(10) + overtime_tot(11) + overtime_tot(12) + overtime_tot(13)
                            %>
							<tr>
                              <th>�Ѱ�</th>
                              <th><%=ce_cnt%></th>
                              <th>&nbsp;</th>
                              <th><%=formatnumber(clng(month_tot(0)),0)%>/<%=overtime_tot(0)%></th>
                              <th><%=formatnumber(clng(month_tot(1)),0)%>/<%=overtime_tot(1)%></th>
                              <th><%=formatnumber(clng(month_tot(2)),0)%>/<%=overtime_tot(2)%></th>
                              <th><%=formatnumber(clng(month_tot(3)),0)%>/<%=overtime_tot(3)%></th>
                              <th><%=formatnumber(clng(month_tot(4)),0)%>/<%=overtime_tot(4)%></th>
                              <th><%=formatnumber(clng(month_tot(5)),0)%>/<%=overtime_tot(5)%></th>
                              <th><%=formatnumber(clng(month_tot(6)),0)%>/<%=overtime_tot(6)%></th>
                              <th><%=formatnumber(clng(month_tot(7)),0)%>/<%=overtime_tot(7)%></th>
                              <th><%=formatnumber(clng(month_tot(8)),0)%>/<%=overtime_tot(8)%></th>
                              <th><%=formatnumber(clng(month_tot(9)),0)%>/<%=overtime_tot(9)%></th>
                              <th><%=formatnumber(clng(month_tot(10)),0)%>/<%=overtime_tot(10)%></th>
                              <th><%=formatnumber(clng(month_tot(11)),0)%>/<%=overtime_tot(11)%></th>
                              <th><%=formatnumber(clng(month_tot(12)),0)%></th>
                              <th><%=formatnumber(clng(month_tot(13)),0)%></th>
							</tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>


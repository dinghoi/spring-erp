<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim u_type, mg_ce_id, mg_ce, start_point, start_hh, start_mm
Dim company, end_point, end_hh, end_mm, transit, payment, fare
Dim run_memo, cancel_yn, end_yn, curr_date, run_date, strNowWeek
Dim title_line, week, rs_end, end_date, new_date, end_saupbu
Dim run_seq

u_type = f_Request("u_type")

mg_ce_id = user_id
mg_ce = user_name
start_point = ""
start_hh = ""
start_mm = ""
company = ""
end_point = ""
end_hh = ""
end_mm = ""
transit = ""
payment = ""
fare = 0
run_memo = ""
cancel_yn = "N"
end_yn = "N"

curr_date = mid(cstr(now()),1,10)
run_date = mid(cstr(now()),1,10)

strNowWeek = WeekDay(run_date)

Select Case (strNowWeek)
   Case 1
       week = "�Ͽ���"
   Case 2
       week = "������"
   Case 3
       week = "ȭ����"
   Case 4
       week = "������"
   Case 5
       week = "�����"
   Case 6
       week = "�ݿ���"
   Case 7
       week = "�����"
End Select

company = "����"

title_line = "���� ����� ���"

'include => end_check.asp =============
If saupbu = "" Then
	end_saupbu = "����οܳ�����"
Else
  	end_saupbu = saupbu
End If

'sql = "SELECT MAX(end_month) as max_month " &_
'      "  FROM cost_end                    " &_
'     " WHERE saupbu = '"&end_saupbu&"'   " &_
'     "   AND end_yn ='Y'                 "
objBuilder.Append "SELECT MAX(end_month) AS max_month "
objBuilder.Append "FROM cost_end "
objBuilder.Append "WHERE saupbu = '"&end_saupbu&"' AND end_yn = 'Y';"

Set rs_end = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_end("max_month")) Then
	end_date = "2014-08-31"
Else
	new_date = DateAdd("m", 1, DateValue(Mid(rs_end("max_month"), 1, 4) & "-" & Mid(rs_end("max_month"), 5, 2) & "-01"))
	end_date = DateAdd("d", -1, new_date)
End If

rs_end.Close() : Set rs_end = Nothing
'========================================

Dim rsTran, reg_id, reg_date, reg_user, mod_id, mod_date, mod_user, end_view
Dim cancel_view

If u_type = "U" Then
	run_date = f_Request("run_date")
	mg_ce_id = f_Request("mg_ce_id")
	run_seq = f_Request("run_seq")

	'sql = "select * from transit_cost where run_date ='"&run_date&"' and mg_ce_id ='"&mg_ce_id&"' and run_seq ='"&run_seq&"'"
	objBuilder.Append "SELECT start_point, start_time, company, end_point, end_time, "
	objBuilder.Append "	transit, payment, fare, run_memo, cancel_yn, end_yn, trct.reg_id, "
	objBuilder.Append "	trct.reg_date, trct.reg_user, trct.mod_id, trct.mod_date, trct.mod_user, "
	objBuilder.Append "	memt.user_name "
	objBuilder.Append "FROM transit_cost AS trct "
	objBuilder.Append "LEFT OUTER JOIN memb AS memt ON trct.mg_ce_id = memt.user_id AND memt.grade < '5' "
	objBuilder.Append "WHERE run_date ='"&run_date&"' AND mg_ce_id ='"&mg_ce_id&"' AND run_seq ='"&run_seq&"';"

	Set rsTran = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsTran.EOF Or rsTran.BOF Then
		mg_ce = "ERROR"
	Else
		'sql = "select * from memb where user_id = '"&mg_ce_id&"'"
		'set rs_memb=dbconn.execute(sql)

		'if	rs_memb.eof or rs_memb.bof then
		'	mg_ce = "ERROR"
		'  else
		'	mg_ce = rs_memb("user_name")
		'end if
		'rs_memb.close()
		If f_toString(rsTran("user_name"), "") = "" Then
			mg_ce = "ERROR"
		Else
			mg_ce = rsTran("user_name")
		End If
	End If

	start_point = rsTran("start_point")
	start_hh = Mid(rsTran("start_time"),1,2)
	start_mm = Mid(rsTran("start_time"),3,2)
	company = rsTran("company")
	end_point = rsTran("end_point")
	end_hh = Mid(rsTran("end_time"),1,2)
	end_mm = Mid(rsTran("end_time"),3,2)
	transit = rsTran("transit")
	payment = rsTran("payment")
	fare = Int(rsTran("fare"))
	run_memo = rsTran("run_memo")
	cancel_yn = rsTran("cancel_yn")
	end_yn = rsTran("end_yn")
	reg_id = rsTran("reg_id")
	reg_date = rsTran("reg_date")
	reg_user = rsTran("reg_user")
	mod_id = rsTran("mod_id")
	mod_date = rsTran("mod_date")
	mod_user = rsTran("mod_user")

	rsTran.Close() : Set rsTran = Nothing

	title_line = "���� ����� ����"
End If

If end_yn = "Y" Then
	end_view = "����"
Else
  	end_view = "����"
End If

If cancel_yn = "Y" Then
	cancel_view = "���"
Else
  	cancel_view = "����"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=run_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.run_date.value <= document.frm.end_date.value){
					alert('�̿����ڰ� ������ �Ǿ� �ִ� �����Դϴ�');
					frm.run_date.focus();
					return false;
				}

				if(document.frm.run_date.value > document.frm.curr_date.value){
					alert('�̿����ڰ� �����Ϻ��� Ŭ���� �����ϴ�.');
					frm.run_date.focus();
					return false;
				}

				if(document.frm.company.value == "" ){
					alert('��ü�� �����ϼ���');
					frm.company.focus();
					return false;
				}

				if(document.frm.mg_ce.value == "" ){
					alert('�̿��� �����Դϴ�. �����ڿ��� ���� �ٶ��ϴ�');
					frm.mg_ce.focus();
					return false;
				}

				if(document.frm.start_point.value == "" ){
					alert('����ּ��� �Է��ϼ���');
					frm.start_point.focus();
					return false;
				}

				if(document.frm.start_hh.value > "23" || document.frm.start_hh.value < "00"){
					alert('��߽ð��� �߸��Ǿ����ϴ�');
					frm.start_hh.focus();
					return false;
				}

				if(document.frm.start_mm.value > "59" || document.frm.start_mm.value < "00"){
					alert('��ߺ��� �߸��Ǿ����ϴ�');
					frm.start_mm.focus();
					return false;
				}

				if(document.frm.end_point.value == "" ){
					alert('�����ּ��� �Է��ϼ���');
					frm.end_point.focus();
					return false;
				}

				if(document.frm.end_hh.value > "23" || document.frm.end_hh.value < "00"){
					alert('�����ð��� �߸��Ǿ����ϴ�');
					frm.end_hh.focus();
					return false;
				}

				if(document.frm.end_mm.value > "59" || document.frm.end_mm.value < "00"){
					alert('�������� �߸��Ǿ����ϴ�');
					frm.end_mm.focus();
					return false;
				}

				if(document.frm.start_hh.value > document.frm.end_hh.value){
					alert('�����ð��� ��߽ð� ���� �����ϴ�');
					frm.end_hh.focus();
					return false;
				}

				if(document.frm.start_hh.value == document.frm.end_hh.value){
					if(document.frm.start_mm.value > document.frm.end_mm.value){
						alert('�����ð��� ��߽ð� ���� �����ϴ�');
						frm.end_mm.focus();
						return false;
					}
				}

				if(document.frm.transit.value == "" ){
					alert('�������� �����ϼ���');
					frm.transit.focus();
					return false;
				}

				if(document.frm.fare.value <= 0 ){
					alert('����� �Է��ϼ���');
					frm.fare.focus();
					return false;
				}

				if(document.frm.run_memo.value == "" ){
					alert('�۾������� �����ϼ���');
					frm.run_memo.focus();
					return false;
				}

				a=confirm('���� �Ͻðڽ��ϱ�?');

				if(a == true){
					return true;
				}
				return false;
			}

 			function week_check(){
				a = document.frm.run_date.value.substring(0,4);
				b = document.frm.run_date.value.substring(5,7);
				c = document.frm.run_date.value.substring(8,10);

				var newDate = new Date(a,b-1,c);
				var s = newDate.getDay();

				switch(s){
					case 0: str = "�Ͽ���" ; break;
					case 1: str = "������" ; break;
					case 2: str = "ȭ����" ; break;
					case 3: str = "������" ; break;
					case 4: str = "�����" ; break;
					case 5: str = "�ݿ���" ; break;
					case 6: str = "�����" ; break;
				}
				document.frm.week.value = str;
			}

			function update_view(){
				var c = document.frm.u_type.value;

				if(c == 'U'){
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}

			function delcheck(){
				a=confirm('���� �����Ͻðڽ��ϱ�?');

				if(a == true){
					document.frm.action = "/cost/mass_transit_del_ok.asp";
					document.frm.submit();
					return true;
				}
				return false;
			}
       </script>
	</head>
	<body onLoad="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/mass_transit_add_save.asp" method="post" name="frm">
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
								<th class="first">�̿�����</th>
								<td class="left">
									<input name="run_date" type="text" id="datepicker" style="width:70px" value="<%=run_date%>" readonly="true"/>&nbsp;
									�������� : <%=end_date%>
								<%If u_type = "U" Then %>
									<input name="old_date" type="hidden" value="<%=run_date%>"/>
								<%End If %>
                                </td>
								<th>�̿���</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)
									<input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=mg_ce_id%>"/>
									<input name="mg_ce" type="hidden" id="mg_ce" value="<%=mg_ce%>"/>
                                </td>
							</tr>
							<tr>
								<th class="first">��ü</th>
								<td class="left">
								<%
								Dim rsTrade

								If f_toString(reside_company, "") = "" Then

									'Sql="select * from trade where (trade_id ='����' or trade_id ='����') and use_sw = 'Y' order by trade_name asc"
									'Rs_etc.Open Sql, Dbconn, 1
									objBuilder.Append "SELECT trade_name FROM trade "
									objBuilder.Append "WHERE (trade_id ='����' OR trade_id ='����') AND use_sw = 'Y' "
									objBuilder.Append "ORDER BY trade_name ASC;"

									Set rsTrade = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()
                                %>
									<select name="company" id="company" style="width:150px">
										<option value="">����</option>
										<option value='����' <%If company = "����" Then %>selected<%End If %>>����</option>
                                    <%
                                        Do Until rsTrade.EOF
                                    %>
										<option value='<%=rsTrade("trade_name")%>' <%If rsTrade("trade_name") = company Then %>selected<%End If %>><%=rsTrade("trade_name")%></option>
                                    <%
                                        	rsTrade.MoveNext()
                                        Loop
                                        rsTrade.Close() : Set rsTrade = Nothing
                                    %>
									</select>
								<%Else	%>
                                    <input name="company" type="text" id="company" style="width:100px" value="<%=reside_company%>" readonly="true"/>
								<%End If	%>
                                </td>
								<th>����ּ�</th>
								<td class="left">
									<input name="start_point" type="text" id="start_point" style="width:200px" onKeyUp="checklength(this,50)" value="<%=start_point%>"/>
								</td>
							</tr>
							<tr>
								<th class="first">��߽ð�</th>
								<td class="left">
									<input name="start_hh" type="text" id="start_hh" size="2" maxlength="2" value="<%=start_hh%>"/>��
									<input name="start_mm" type="text" id="start_mm" size="2" maxlength="2" value="<%=start_mm%>"/>��
                                </td>
								<th>�����ּ�</th>
								<td class="left">
									<input name="end_point" type="text" id="end_point" style="width:200px" onKeyUp="checklength(this,50)" value="<%=end_point%>"/>
								</td>
							</tr>
							<tr>
								<th class="first">�����ð�</th>
								<td class="left">
									<input name="end_hh" type="text" id="end_hh" size="2" maxlength="2" value="<%=end_hh%>"/>��
									<input name="end_mm" type="text" id="end_mm" size="2" maxlength="2" value="<%=end_mm%>"/>��
                                </td>
								<th>������</th>
								<td class="left">
                                <select name="transit" id="transit" style="width:80px">
                                    <option value="">����</option>
									<option value='����' <%If transit= "����" Then %>selected<% End If %>>����</option>
								  	<option value='����ö' <%If transit= "����ö" Then %>selected<% End If %>>����ö</option>
								  	<option value='�ý�' <%If transit= "�ý�" Then %>selected<% End If %>>�ý�</option>
								  	<option value='����' <%If transit= "����" Then %>selected<% End If %>>����</option>
								  	<option value='�����' <%If transit= "�����" Then %>selected<% End If %>>�����</option>
								  	<option value='��' <%If transit= "��" Then %>selected<% End If %>>��</option>
								  	<option value='��Ÿ' <%If transit= "��Ÿ" Then %>selected<% End If %>>��Ÿ</option>
							    </select></td>
							</tr>
							<tr>
								<th class="first">�����</th>
								<td class="left">���ҹ��
									<select name="payment" id="select" style="width:80px">
										<option value='����' <%If payment= "����" Then %>selected<%End If %>>����</option>
									</select>
								<%If u_type = "U" Then	%>
									<input name="fare" type="text" id="far2" style="width:80px;text-align:right" value="<%=formatnumber(fare,0)%>" onKeyUp="plusComma(this);"/>
								<%Else	%>
									<input name="fare" type="text" id="far2" style="width:80px;text-align:right" onKeyUp="plusComma(this);"/>
								<%End If	%>
                                </td>
								<th>�۾�����</th>
								<td class="left">
								<%
								Dim rs_etc

                                'Sql="select * from etc_code where etc_type = '42' and used_sw = 'Y' order by etc_code asc"
                                'Rs_etc.Open Sql, Dbconn, 1
								objBuilder.Append "SELECT etc_name FROM etc_code "
								objBuilder.Append "WHERE etc_type = '42' AND used_sw = 'Y' "
								objBuilder.Append "ORDER BY etc_code ASC; "

								Set rs_etc = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
                                %>
									<select name="run_memo" id="select" style="width:150px">
										<option value="">����</option>
                                    <%
                                        Do Until rs_etc.EOF
                                    %>
										<option value='<%=rs_etc("etc_name")%>' <%If rs_etc("etc_name") = run_memo Then %>selected<%End If %>><%=rs_etc("etc_name")%></option>
                                    <%
                                        	rs_etc.MoveNext()
                                        Loop
                                        rs_etc.Close() : Set rs_etc = Nothing
										DBConn.Close() : Set DBConn = Nothing
                                    %>
									</select>
                                </td>
							</tr>
    					<tr id="cancel_col" style="display:none">
							<th class="first">��ҿ���</th>
							<td class="left"><%=cancel_view%><input type="hidden" name="cancel_yn" value="<%=cancel_yn%>"/></td>
							<th>��������</th>
							<td class="left"><%=end_view%></td>
						</tr>
						<tr id="info_col" style="display:none">
							<th class="first">�������</th>
							<td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
							<th>��������</th>
							<td class="left"><%=mod_user%>&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
						</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();"/></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
				<%
				If u_type = "U" And user_id = mg_ce_id Then
					If end_yn = "N" Or end_yn = "C" Then
				%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%
					End If
				End If
				%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>"/>
                <input type="hidden" name="curr_date" value="<%=curr_date%>"/>
                <input type="hidden" name="end_date" value="<%=end_date%>"/>
				<input type="hidden" name="run_seq" value="<%=run_seq%>"/>
				<input type="hidden" name="end_yn" value="<%=end_yn%>"/>
                <input type="hidden" name="mod_id" value="<%=mod_id%>"/>
                <input type="hidden" name="mod_user" value="<%=mod_user%>"/>
                <input type="hidden" name="mod_date" value="<%=mod_date%>"/>
			</form>
		</div>
	</body>
</html>
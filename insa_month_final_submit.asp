<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim month_tab(24,2)


user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = now()
be_date = dateadd("m",-1,curr_date)
be_month = cstr(mid(be_date,1,4)) + cstr(mid(be_date,6,2))

inc_yyyy = cstr(mid(be_date,1,4)) + " �� " + cstr(mid(be_date,6,2))
inc_yyyy1 = cstr(mid(be_date,1,4)) + "" + cstr(mid(be_date,6,2))


'Response.Write inc_yyyy1

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' ��� ���̺����
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if
	view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

title_line = " ���� �� �λ� �� ����ó�� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
	    </script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}

			function insa_month_final(val) {

            if (!confirm("���� �� �λ� �� ����ó���� �Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
			//document.frm.be_month1.value = document.getElementById(val).value;
            document.frm.action = "insa_month_final_submit_ok.asp";
            document.frm.submit();
            }
		</script>

	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_month_final_submit.asp" method="post" name="frm">
                <fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
                        <dd>
                            <p>
							<strong>
							<% // =inc_yyyy%>


								<select name="inc_yyyy1" id="inc_yyyy1" type="text" value="<%=inc_yyyy1%>" style="width:110px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If inc_yyyy1 = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>



							�� ���� �� �λ� ���� </strong>
								<label>
        						<input name="emp_no" type="hidden" id="emp_no" value="<%=emp_no%>" style="width:40px" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <h3 class="stit"> ����ó���� Ŭ���Ͻø� <%=inc_yyyy%>  �� ���� �� �λ��ڷᰡ �����˴ϴ�.<br>&nbsp;<br></h3>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th colspan="2" class="left" style=" border-bottom:1px solid #ffffff;">����ó���� �ſ� 5�ϰ� ���� ���� �� �λ��ڷḦ �����մϴ�.<br><br><strong>�� ������ ���� ���� �� �λ�߷ɵ��� 5������ ��� ���� �Ͻð� ó���� �ϼž� �մϴ�.</strong><br><br><strong>�� ��� ���� �� �λ�߷��� �Ͻñ� ���� ���� ����ó���� �ϼž� �մϴ�.</strong></th>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����ó��" onclick="insa_month_final('be_month');return false;" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="be_month" value="<%=be_month%>" ID="Hidden1">
                <input type="hidden" name="be_month1" value="<%=be_month1%>" ID="Hidden1">
				</form>
		</div>
	</body>
</html>


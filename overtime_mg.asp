<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim from_date
Dim to_date

' ��Ư�� ���α��� ID ����Ʈ
allowerIDs = Array("100125","100029","100015","100031","100020","100018") ' "����","�����","������","�ֱ漺','ȫ����','������'

treeDayAgo = DateAdd("d",-60,now())

view_c     = Request.form("view_c")
mg_ce      = Request.form("mg_ce")
from_date  = Request.form("from_date")
to_date    = Request.form("to_date")

work_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)

If view_c = "" Then
	view_c = "total"
End If

if from_date = "" then
    from_date = mid(work_month,1,4) + "-" + mid(work_month,5,2) + "-01"
end if
if to_date = "" then
    to_date = cstr(dateadd("d",-1, dateadd("m",1,datevalue(from_date)) ))
end if


Set Dbconn  = Server.CreateObject("ADODB.Connection")
Set Rs      = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' �����Ǻ�
posi_sql = " and overtime.mg_ce_id = '" + user_id + "'"

if position = "����" then
	view_condi = "����"
end if

if position = "��Ʈ��" then
	if view_c = "total" then
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (overtime.org_name = '��ȭ����ȣ��' or overtime.org_name = '��ȭ��������') "
		  else
			posi_sql = " and overtime.org_name = '"&org_name&"'"
		end if
	  else
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (overtime.org_name = '��ȭ����ȣ��' or overtime.org_name = '��ȭ��������') and memb.user_name like '%"&mg_ce&"%'"
		  else
			posi_sql = " and overtime.org_name = '"&org_name&"' and memb.user_name like '%"&mg_ce&"%'"
		end if
	end if
end if

if position = "����" then
	if view_c = "total" then
		posi_sql = " and overtime.team = '"&team&"'"
	  else
		posi_sql = " and overtime.team = '"&team&"' and memb.user_name like '%"&mg_ce&"%'"
	end if
end if

if position = "�������" or cost_grade = "2" then
	if view_c = "total" then
        'posi_sql = " and overtime.saupbu = '"&saupbu&"'"
        posi_sql = " and overtime.saupbu = emp_master.emp_saupbu "&chr(13)
	  else
        'posi_sql = " and overtime.saupbu = '"&saupbu&"' and memb.user_name like '%"&mg_ce&"%'"
        posi_sql = " and overtime.saupbu = emp_master.emp_saupbu and memb.user_name like '%"&mg_ce&"%'"&chr(13)
	end if
end if

if position = "������" or cost_grade = "1" then
  	if view_c = "total" then
		posi_sql = " and overtime.bonbu = '"&bonbu&"'"
 	else
		posi_sql = " and overtime.bonbu = '"&bonbu&"' and memb.user_name like '%"&mg_ce&"%'"
	end if
end if

view_grade = position

if cost_grade = "0" then
	view_grade = "��ü"
  	if view_c = "total" then
		posi_sql = ""
 	else
		posi_sql = " and memb.user_name like '%"&mg_ce&"%'"
	end if
end if

base_sql = "    SELECT overtime.cancel_yn                            "&chr(13)&_
           "         , overtime.acpt_no                              "&chr(13)&_
           "         , overtime.you_yn                               "&chr(13)&_
           "         , overtime.org_name                             "&chr(13)&_
           "         , overtime.user_name                            "&chr(13)&_
           "         , overtime.user_grade                           "&chr(13)&_
           "         , overtime.mg_ce_id                             "&chr(13)&_
           "         , overtime.work_date                            "&chr(13)&_
           "         , overtime.company                              "&chr(13)&_
           "         , overtime.dept                                 "&chr(13)&_
           "         , overtime.work_gubun                           "&chr(13)&_
           "         , overtime.work_memo                            "&chr(13)&_
           "         , overtime.overtime_amt                         "&chr(13)&_
           "         , overtime.end_yn                               "&chr(13)&_
           "         , overtime.reg_id                               "&chr(13)&_
           "         , overtime.allow_yn                             "&chr(13)&_
           "         , ifnull(overtime.delta_minute,0) delta_minute  "&chr(13)&_
           "         , ifnull(overtime.rest_minute,0) rest_minute    "&chr(13)&_
           "         , memb.user_name                                "&chr(13)&_
           "         , memb.user_grade                               "&chr(13)&_
		   "         , emp_org_mst.org_name                          "&chr(13)&_
           "      FROM overtime                                      "&chr(13)&_
           "INNER JOIN memb                                          "&chr(13)&_
           "        ON overtime.mg_ce_id = memb.user_id              "&chr(13)&_
           "inner join emp_master                                    "&chr(13)&_
           "        ON emp_master.emp_no = overtime.mg_ce_id         "&chr(13)&_
		   "inner join emp_org_mst									 "&chr(13)&_
		   "        ON emp_org_mst.org_code = emp_master.emp_org_code        "&chr(13)
date_sql = "     WHERE work_date >= '" + from_date  + "' "&chr(13)&_
           "       AND work_date <= '" + to_date  + "'   "&chr(13)

sql = base_sql & date_sql & posi_sql & chr(13)&_
    " ORDER BY overtime.org_name, memb.user_name, work_date"


Rs.Open Sql, Dbconn, 1

title_line = "��Ư�� ���� (�� �Ҽ����� �˻� �ȵ� ��� �λ����� ��û�Ͽ� �ҼӺ���, ����� ��ġ�� ��û �Ͻñ� �ٶ��ϴ�.) �۾���3���� �����Ұ�"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
		    $(function() {

				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=from_date%>" );

				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {

			  var fDate = $("#datepicker1").val();
				var lDate = $("#datepicker2").val();

				if (fDate = "")
				{
				  alert("�˻� ���۳������ �����ϴ�.");
					return false;
				}

				if (lDate = "")
				{
				  alert("�˻� ���������� �����ϴ�.");
					return false;
				}

				if ((fDate != "") && (lDate != "") && (fDate > lDate)) {
					alert("�˻� ���۳������ ���� ����� ���� ���� �� �����ϴ�.");
					return false;
				}

				return true;
			}

			function condi_view()
            {
            <%
                if not (position = "����" and cost_grade <> "0") then
                        %>
                    if (eval("document.frm.view_c[0].checked")) {
                        document.getElementById('mg_ce_view').style.display = 'none';
                    }
                    if (eval("document.frm.view_c[1].checked")) {
                        document.getElementById('mg_ce_view').style.display = '';
                    }
                    <%
                end if
            %>
			}

		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="overtime_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ǰ˻�</dt>
						<dd>
							<p style="position:relative">

                                &nbsp;&nbsp;<strong>�۾����&nbsp;</strong>
                                <input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker1">
                                    ~
                                <input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker2">

                                <label><strong>��ȸ���� : </strong><%=view_grade%></label>
                                <label><strong>��ȸ���� : </strong>
                                <%
                                if position = "����" and cost_grade <> "0" then
                                        Response.write view_condi
                                else
                                    %>
                                    <label><input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">������ü</label>
                                    <label><input type="radio" name="view_c" value="reg_id" <% if view_c = "reg_id" then %>checked<% end if %> style="width:25px" onClick="condi_view()">���κ�</label>
                                    <%
                                end if
                                %>
                                </label>
                                <label>
                                    <input name="mg_ce" type="text" value="<%=mg_ce%>" style="width:70px; display:none" id="mg_ce_view">
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
                                <span style="position:absolute;right:5px; cursor: pointer;" class="btnType04" onclick="pop_Window('overtime_stats.asp?','asview_pop','scrollbars=yes,width=1200,height=700')">�� 52�ð� ��Ȳ����</span>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" />
							<col width="7%" />
							<col width="7%" />
							<col width="11%" />
							<col width="5%" />
							<col width="11%" />
							<col width="11%" />
							<col width="13%" />
							<col width="13%" />
							<%
                            find = False
                            For i = 0 To uBound(allowerIDs)
                            if  user_id = allowerIDs(i) then
                                find =True
                            end if
                            Next

                            if find = True then
                                %><col width="7%" /><%
                            end if
							%>
							<col width="5%" />
							<col width="5%" />
							<col width="4%" />
							<col width="4%" />
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">������</th>
								<th scope="col">�۾���</th>
								<th scope="col">�ٹ�����</th>
								<th scope="col">�� �ð�</th>
								<th scope="col">AS NO</th>
								<th scope="col">ȸ��</th>
								<th scope="col">������</th>
								<th scope="col">��Ư�ٱ���</th>
								<th scope="col">�۾�����</th>
								<%
  								find = False
                                For i = 0 To uBound(allowerIDs)
                                    if  user_id = allowerIDs(i) then
                                        find =True
                                    end if
                                Next

                                if find = True then
                                    %><th scope="col">��û�ݾ�</th><%
                                end if
  							    %>
								<th scope="col">������</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
						<%
                        cost_sum = 0
                        end_sum = 0
                        cancel_sum = 0

                        do until rs.eof

                            delta_minute = Cint( Rs("delta_minute") ) ' �Ѱ���ð��� �Ѻ����� .. (����,�̽��� ������� �Ѵ�)
                                rest_minute  = Cint( Rs("rest_minute") )  ' ���ްԽð��� �Ѻ����� .. (����,�̽��� ������� �Ѵ�)
                                if (delta_minute > rest_minute) then
                                delta_minute = delta_minute - rest_minute
                            else
                                delta_minute = 0
                            end if
                            work_time   = Fix(delta_minute / 60) ' ���۾��ð��� �÷� ..  (����,�̽��� ������� �Ѵ�)
                            work_minute = delta_minute mod 60    ' ���۾��ð��� �÷� �������� ������ ..  (����,�̽��� ������� �Ѵ�)

                            if  rs("cancel_yn") = "Y" then
                                cancel_yn = "���"
                                else
                                cancel_yn = "����"
                            end if
                            if rs("acpt_no") = 0 or rs("acpt_no") = null then
                                acpt_no = "����"
                                else
                                acpt_no = rs("acpt_no")
                            end if

                            cost_sum = cost_sum + rs("overtime_amt")
                            if rs("cancel_yn") = "Y" then
                                cancel_sum = cancel_sum + rs("overtime_amt")
                            else
                                end_sum = end_sum + rs("overtime_amt")
                            end if
                            if rs("you_yn") = "Y" then
                                you_view = "����"
                                else
                                you_view = "����"
                            end if
                            %>
                            <tr>
                                <td class="first"><%=rs("org_name")%></td>
                                <td><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%><input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=rs("mg_ce_id")%>"></td>
                                <td><%=rs("work_date")%><input name="work_date" type="hidden" id="work_date" value="<%=rs("work_date")%>"></td>
                                <td><%=work_time%>�ð� <%=work_minute%>��</td>
                                <td>
                                    <%
                                        if acpt_no = "����" then
                                            Response.write acpt_no
                                        else
                                    %>
                                    <a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=acpt_no%>','asview_pop','scrollbars=yes,width=800,height=700')"><%=acpt_no%></a>
                                    <% end if	%>
                                </td>
                                <td><%=rs("company")%></td>
                                <td><%=rs("dept")%></td>
                                <td><%=rs("work_gubun")%></td>
                                <td><%=rs("work_memo")%></td>
                                <%
                                find = False
                                For i = 0 To uBound(allowerIDs)
                                if  user_id = allowerIDs(i) then
                                    find =True
                                end if
                                Next

                                if find = True then
                                    %><td class="right"><%=formatnumber(rs("overtime_amt"),0)%></td><%
                                end if
                                %>
                                <td><%=you_view%></td>
                                <td><%=cancel_yn%></td>
                                <td>
                                <%
                                if rs("end_yn") = "Y" then
                                    Response.write "����"
                                else
                                    if treeDayAgo < rs("work_date") then
                                        if rs("mg_ce_id") = user_id or rs("reg_id") = user_id then ' ce�� ����ڰ� �������� ���..
                                            if rs("acpt_no") = 0 then ' AS���� ��ȣ�� ���ٸ� '����'
                                                %><a href="#" onClick="pop_Window('overtime_hanjin_add.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime__hanjinadd_popup','scrollbars=yes,width=1100,height=500')">����</a><%
                                            else
                                                if rs("work_date") > "2014-12-31" then
                                                    %><a href="#" onClick="pop_Window('overtime_as_mod_15.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>','overtime_as_mod_15_popup','scrollbars=yes,width=1000,height=660')">����</a><%
                                                else
                                                    %><a href="#" onClick="pop_Window('overtime_add.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime_add_popup','scrollbars=yes,width=750,height=300')">����</a><%
                                                end if
                                            end if
                                        else
                                            %><a href="#" onClick="pop_Window('overtime_cancel.asp?work_date=<%=rs("work_date")%>&mg_ce_id=<%=rs("mg_ce_id")%>&u_type=<%="U"%>','overtime_cancel_popup','scrollbars=yes,width=750,height=300')">����</a><%
                                        end if
                                    else
                                        Response.write "�Ұ�" ' 3�� �ʰ��� �����Ұ�
                                    end if
                                end if
                                %>
                                </td>
                                <td><%=rs("allow_yn")%></td>
                                </tr>
                                <%
                                rs.movenext()
                            loop
                            rs.close()
                            %>
							<!-- �հ� �׸��� �Ǽ��� ���Ǿ� �̳��� ó��(�λ翡�� �޿��� �����ؼ� ���� ó����)[����ȣ_20210705]
                            <tr>
								<th colspan="2" class="first">�� ��</th>
                                <th colspan="3">��û�ݾ� :&nbsp;<%''=formatnumber(cost_sum,0)%></th>
                                <th colspan="3">���ޱݾ� :&nbsp;<%''=formatnumber(end_sum,0)%></th>
                                <%
  								'find = False
                                'For i = 0 To uBound(allowerIDs)
                                '    if  user_id = allowerIDs(i) then
                                '        find =True
                                '    end if
                                'Next

                                'if find = True then
                                '    width = 6
                                'else
                                '    width = 5
                                'end if
                                %>
                                <th colspan="<%''=width%>">��ұݾ� :&nbsp;<%''=formatnumber(cancel_sum,0)%></th>
						    </tr>-->
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
                        <a href="/cost/excel/overtime_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&view_c=<%=view_c%>&mg_ce=<%=mg_ce%>" class="btnType04">�����ٿ�ε�</a>
					    </div>
                    </td>
				    <td width="85%">
    					<div class="btnRight">
                            <!-- 2019.02.07 ������ ��û ������,������ "���� ���ְ���","���� ��������","�����ǻ���� �� ��� '��������������'�� ���÷��� �ϵ��� ���� -->
                            <!-- 2019.02.08 ���ͼ� ��û ���׼��� -->
                            <!-- [<%=cost_grade%>] [<%=saupbu%>] [<%=org_name%>] -->
                            <% if cost_grade = "0" or (saupbu <> "KAL���������" and saupbu <> "�������������" and ((org_name<>"���� ���ְ���" and org_name<>"���� ��������") or saupbu<>"�����ǻ����")) then	%>
                                <a href="#" onClick="pop_Window('overtime_as_add_15.asp','overtime_as_add_15_popup','scrollbars=yes,width=1000,height=660')" class="btnType04">A/S���� ��Ư�ٵ��</a>
                            <% end if	%>
                            <% if cost_grade = "0" or saupbu = "KAL���������" or saupbu = "�������������" or org_name = "��ȸ�繫ó" or org_name="���� �λ�����" or org_name="���� ��ũ����" or org_name="���� ������" or org_name="���� ���װ���" or org_name="���� ��������" or org_name="���� ��������" or org_name="���� �λ����" or org_name="���� �뱸����" or (org_name="���� û�ְ���" and saupbu="��û�����") or ((org_name="���� ���ְ���" or org_name="���� ��������" ) and saupbu="�����ǻ����") Or org_name = "����" then	%>
                                <a href="#" onClick="pop_Window('overtime_hanjin_add.asp','overtime_hanjin_as_add_popup','scrollbars=yes,width=1100,height=500')" class="btnType04"> ���������׽����ٵ��</a>
                            <% end if	%>
    					</div>
                    </td>
			    </tr>
				</table>
			</form>
		</div>
	</div>
	</body>
</html>

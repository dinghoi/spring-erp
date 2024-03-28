<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
acpt_no = request("acpt_no")
win_sw = request("win_sw")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_work = Server.CreateObject("ADODB.Recordset")
Set rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_mod = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

Sql = "select a.* "
sql = sql & ", (select concat(emp_hp_ddd,'-',emp_hp_no1, '-', emp_hp_no2) from emp_master where emp_no=a.mg_ce_id) AS ce_tel "
sql = sql & ", (SELECT b.au_name FROM as_unitprice_month b WHERE b.use_yn = 'Y' AND b.au_code = a.as_unit) AS au_name "
sql = sql & ", CASE a.night WHEN 'Y' THEN 'YES' ELSE 'NO' END AS night_name "
sql = sql & ", CASE a.weekend_work WHEN 'Y' THEN 'YES' ELSE 'NO' END AS weekend_work_name "
sql = sql & "from as_acpt a where a.acpt_no = "&int(acpt_no)
'Response.write sql
Set rs=DbConn.Execute(SQL)


as_memo = replace(rs("as_memo"),chr(10),"<br>")





if rs("overtime") = "Y" then
	overtime_view = "��Ư�ٽ�û"
  else
  	overtime_view = ""
end if

request_date = cstr(rs("request_date")) + " " + mid(cstr(rs("request_time")),1,2) + ":" + mid(cstr(rs("request_time")),3)
if rs("visit_date") = "" or isnull(rs("visit_date")) then
	visit_date = "."
  else
	visit_date = cstr(rs("visit_date")) + " " + mid(cstr(rs("visit_time")),1,2) + ":" + mid(cstr(rs("visit_time")),3)
end if




sql_etc = "select * from memb where user_id = '" + rs("mg_ce_id") +"'"
'response.write(sql_etc)
set rs_etc=dbconn.execute(sql_etc)
if rs_etc.eof then
	hp_no = "�����"
  else
  	hp_no = rs_etc("hp")
end if
rs_etc.close()

if rs("visit_request_yn") = "Y" then
	visit_request_view = "���湮��û"
  else
  	visit_request_view = ""
end if


sql = "select a.* "
sql = sql & ", (select concat(emp_hp_ddd,'-',emp_hp_no1, '-', emp_hp_no2) from emp_master where emp_no=a.mg_ce_id) AS ce_tel "
sql = sql & " from as_acpt a"
sql = sql & where_sql & base_sql & order_sql & " limit "& stpage & "," &pgsize


title_line = "A/S ���γ���"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� ����</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>

		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}

			function printWindow(){
        //		viewOff("button");
                factory.printing.header = ""; //�Ӹ��� ����
                factory.printing.footer = ""; //������ ����
                factory.printing.portrait = false; //��¹��� ����: true - ����, false - ����
                factory.printing.leftMargin = 13; //���� ���� ����
                factory.printing.topMargin = 10; //���� ���� ����
                factory.printing.rightMargin = 13; //�����P ���� ����
                factory.printing.bottomMargin = 15; //�ٴ� ���� ����
        //		factory.printing.SetMarginMeasure(2); //�׵θ� ���� ������ ������ ��ġ�� ����
        //		factory.printing.printer = ""; //������ �� ������ �̸�
        //		factory.printing.paperSize = "A4"; //��������
        //		factory.printing.pageSource = "Manusal feed"; //���� �ǵ� ���
        //		factory.printing.collate = true; //������� ����ϱ�
        //		factory.printing.copies = "1"; //�μ��� �ż�
        //		factory.printing.SetPageRange(true,1,1); //true�� �����ϰ� 1,3�̸� 1���� 3������ ���
        //		factory.printing.Printer(true); //����ϱ�
                factory.printing.Preview(); //�����츦 ���ؼ� ���
                factory.printing.Print(false); //�����츦 ���ؼ� ���
            }
        </script>

	</head>

	<style media="print">
    .noprint     { display: none }
    </style>

	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="12%" >
							<col width="13%" >
							<col width="13%" >
							<col width="12%" >
							<col width="12%" >
							<col width="14%" >
							<col width="10%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>������ȣ</th>
							  <td class="left"><%=acpt_no%></td>
							  <th>��������</th>
							  <td class="left" colspan="3"><%=rs("acpt_date")%></td>
					      	</tr>
							<tr>
							  <th>������</th>
							  <td class="left"><%=rs("acpt_man")%>&nbsp;<%=rs("acpt_grade")%></td>
							  <th>�����</th>
							  <td class="left"><%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%></td>
							  <th>���CE</th>
							  <td class="left"><%=rs("mg_ce")%>
							  <th>CE TEL</th>
							  <td class="left"><%=rs("ce_tel")%></td>
   		      	</tr>
							<tr>
							  <th>��ȭ��ȣ</th>
							  <td class="left"><%=rs("tel_ddd")%>-<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
							  <th>ȸ��</th>
							  <td class="left"><%=rs("company")%></td>
							  <th>������</th>
							  <td class="left" colspan="3"><%=rs("dept")%></td>
					      	</tr>
							<tr>
							  <th>�ּ�</th>
							  <td class="left" colspan="7"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>&nbsp;<%=rs("addr")%></td>
					      	</tr>
							<tr>
							  <th>��ֳ���</th>
							  <td class="left" colspan="7"><%=as_memo%></td>
					      	</tr>
							<tr>
							  <th>ASǥ�شܰ�<br>����</th>
							  <td class="left" colspan="2"><%=rs("au_name")%></td>
							   <th>�߰��۾�</th>
  							  <td class="left"><%=rs("night_name")%></td>
							  <th>�ָ��۾�</th>
							  <td class="left" colspan="2"><%=rs("weekend_work_name")%></td>
					      	</tr>
							<tr>
							  <th>��û��</th>
							  <td class="left"><%=request_date%></td>
							  <th>ó����</th>
							  <td class="left"><%=visit_date%></td>
							  <th>ó������</th>
							  <td class="left"><%=rs("as_type")%>&nbsp;<%=visit_request_view%></td>
							  <th>��������</th>
							  <td class="left">
							  	<% if rs("cowork_yn") = "Y" then	%>
                      	<%="����"%>
                  <% else	%>
                      	<%="�Ϲ�"%>
						      <% end if	%>
								</td>
					    </tr>
							<tr>
							  <th>ó����Ȳ</th>
							  <td class="left"><%=rs("as_process")%></td>
							  <th>����/�԰����</th>
							  <td class="left" colspan="3">&nbsp;<%=rs("into_reason")%></td>
							   <th> </th>
							  <td class="left"> </td>
					      	</tr>
							<tr>
							  <th>������</th>
							  <td class="left"><%=rs("maker")%></td>
							  <th>������</th>
							  <td class="left"><%=rs("as_device")%></td>
							  <th>�𵨹�ȣ</th>
							  <td class="left">&nbsp;<%=rs("model_no")%></td>
							   <th> </th>
							  <td class="left"> </td>
					      	</tr>
							<tr>
							  <th>�ø���NO</th>
							  <td class="left">&nbsp;<%=rs("serial_no")%></td>
							  <th>�ڻ��ȣ</th>
							  <td class="left">&nbsp;<%=rs("asets_no")%></td>
							  <th>����ǰ</th>
							  <td class="left">&nbsp;<%=rs("as_parts")%></td>
							   <th> </th>
							  <td class="left"> </td>
					      	</tr>
					      	<tr>
							  <th>�۾�����</th>
							  <td class="left" colspan="5">&nbsp;<%=rs("as_history")%></td>
							  <th> </th>
							  <td class="left"> </td>
					      	</tr>
							<%
                                dim error_pro
                                dim err_name

                                error_pro = rs("err_pc_sw")+rs("err_pc_hw")+rs("err_monitor")+rs("err_printer")+rs("err_network")+rs("err_server")+rs("err_adapter")+rs("err_etc")

                                if error_pro <> "" then
                                    error_pro = replace(error_pro,",","")
                                    error_pro = replace(error_pro," ","")

                                    j = len(error_pro)

                                    for i = 4 to j step 4

                                        err_code = mid(error_pro,i-3,4)

                                        sql_etc = "select * from etc_code where etc_code = '"&err_code&"'"
                                        set rs_etc=dbconn.execute(sql_etc)

										if rs_etc.eof or rs_etc.bof then
											etc_name = ""
										  else
											etc_name = rs_etc("etc_name")
											if err_memo = "" then
												err_memo = etc_name
											  else
												err_memo = err_memo + "," +etc_name
											end if
										end if
                                        rs_etc.close()
                                    next
                                end if

							path = "/att_file/" + rs("company")

							sql_att = "select * from att_file where acpt_no = "&int(acpt_no)
							set rs_att=dbconn.execute(sql_att)
							if rs_att.eof or rs_att.bof then
								not_att = "Y"
							  else
								not_att = "N"
							end if
							if rs("dev_inst_cnt") = "" or isnull(rs("dev_inst_cnt")) then
								dev_inst_cnt = "0"
							  else
							  	dev_inst_cnt = rs("dev_inst_cnt")
							end if
                            if rs("as_process") = "�Ϸ�" and ( rs("as_type") = "�űԼ�ġ" or rs("as_type") = "�űԼ�ġ����" or rs("as_type") = "������ġ" or rs("as_type") = "������ġ����" or rs("as_type") = "������" or rs("as_type") = "������" ) then
								err_name = " ��ġ : " + cstr(dev_inst_cnt) + "��, ����: " + cstr(rs("ran_cnt")) + "��, �۾��ο�: " + cstr(rs("work_man_cnt")) + "��, �˹�: " + cstr(rs("alba_cnt"))
							end if
                            if rs("as_process") = "�Ϸ�" and ( rs("as_type") = "���ȸ��" or rs("as_type") = "��������" ) then
								err_name = "�۾�: " + cstr(dev_inst_cnt) + "��"
							end if
                            %>
					      	<tr>
							  <th>��ġ����</th>
							  <td class="left" colspan="5">&nbsp;<%=err_name%></td>
							  <th> </th>
							  <td class="left"> </td>
					      	</tr>
					      	<tr>
							  <th>÷������</th>
							  <td colspan="3" class="left">&nbsp;
								<%
                                if not_att = "N" then
                                    if rs_att("att_file1") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file1")%>">÷��1</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file2") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file2")%>">÷��2</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file3") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file3")%>">÷��3</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file4") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file4")%>">÷��4</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file5") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file5")%>">÷��5</a>&nbsp;
                                <%
                                    end if
                                end if
                                %>
                              </td>
							  <th>��Ư�ٵ��</th>
							  <td class="left"><%=overtime_view%>&nbsp;</td>
							  <th> </th>
							  <td class="left"> </td>
					      	</tr>
					      	<tr>
					      	  <th>�۾��η�</th>
					      	  <td colspan="7" class="left">
						<%
							j = 0
							sql_work = "select * from ce_work where acpt_no = "&int(acpt_no)
							'Response.write sql_work & "<br>"
							rs_work.Open sql_work, Dbconn, 1
							do until rs_work.eof
								j = j + 1
								sql_etc = "select * from memb where user_id = '" + rs_work("mg_ce_id") +"'"
								'Response.write sql_etc & "<br>"
								set rs_etc=dbconn.execute(sql_etc)
								if rs_etc.eof then
									work_man = "ERROR"
								  else
									work_man = rs_etc("user_name") + " " + rs_etc("user_grade")
								end if
								rs_etc.close()
						%>
 								<%=j%>.&nbsp;<%=work_man%>(<%=rs_work("mg_ce_id")%>)&nbsp;<%=rs_work("org_name")%>&nbsp;&nbsp;
                        <%
								rs_work.movenext()
							loop
						%>
                              &nbsp;
                              </td>
			      	      </tr>
						</tbody>
					</table>
        <h3 class="stit">* �԰� History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">��������</th>
								<th scope="col">�԰�ó</th>
								<th scope="col">�԰�����</th>
								<th scope="col" class="left">�԰��γ���</th>
							</tr>
						</thead>
						<tbody>
						<%
                            Sql_in="select * from as_into where acpt_no = " & int(acpt_no) & " order by in_seq asc"
                            Rs_in.Open Sql_in, Dbconn, 1
                            i = 0
                            do until rs_in.eof
                        %>
							<tr>
								<td class="first"><%=rs_in("into_date")%></td>
								<td><%=rs_in("in_place")%></td>
								<td><%=rs_in("in_process")%></td>
								<td style="text-align:left" class="left"><%=rs_in("in_remark")%></td>
							</tr>
						<%
							rs_in.movenext()
						loop
						rs_in.close()
						%>
						</tbody>
					</table>
					<br>
      <h3 class="stit">* ���� History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���泻��</th>
								<th scope="col">������</th>
								<th scope="col">��������</th>
							</tr>
						</thead>
						<tbody>
						<%
                            Sql_mod="select * from as_mod where acpt_no = " & int(acpt_no) & " order by mod_date asc"
                            ' response.write(Sql_mod)
                            Rs_mod.Open Sql_mod, Dbconn, 1
                            i = 0
                            do until Rs_mod.eof
                        %>
							<tr>
								<td class="first"><%=Rs_mod("mod_pg")%></td>
								<td><%=Rs_mod("mod_name")%>(<%=Rs_mod("mod_id")%>)</td>
								<td><%=Rs_mod("mod_date")%></td>
							</tr>
						<%
							Rs_mod.movenext()
						loop
						Rs_mod.close()
						%>
						</tbody>
					</table>
					<br>
				</form>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>
                    		<span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
                    	</div>
    				</div>
				<br>
		  </div>
			</div>
	</body>
</html>


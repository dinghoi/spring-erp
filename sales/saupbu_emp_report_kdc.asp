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
Dim cost_month, before_date
Dim cost_date, start_date, cost_year
Dim rs, title_line

cost_month = Request.Form("cost_month")

'==����
Dim sales_bonbu

sales_bonbu = Request.Form("sales_bonbu")

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) + Mid(CStr(before_date), 6, 2)
	sales_bonbu = "��ü"
End If

cost_date = Mid(CStr(cost_month), 1, 4) + "-" + Mid(CStr(cost_month), 5, 2) + "-01"
start_date = DateAdd("m", -1, cost_date)
cost_year = Mid(cost_month, 1, 4)

title_line = "����κ� �ο� ��Ȳ"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
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
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("�ٹ������ �Է��ϼ���.");
					return false;
				}
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}

            $(document).ready(function(){
                $("input[name=cost_except]").change(function(){
					var bonbu = "<%=sales_bonbu%>";

					if (bonbu !== "����SI����" && bonbu !== "SI���ົ��")
                    {
                        alert("����SI���� Ȥ�� SI���ົ�θ� �����մϴ�!");
                        return ;
                    }
                    var emp_month = $(this).attr("emp_month"); //
                    var emp_no    = $(this).attr("emp_no");    //
                    var chked     = $(this).is(":checked");    // üũ����

                    // alert("emp_month= "+emp_month+", emp_no= "+emp_no);

                    $.ajax({
                             url: "/ajax_set_empMasterMonth_costExcept.asp"
                            ,type: 'post'
                            ,data:  { "emp_month" : emp_month
                                    , "emp_no"    : emp_no
                                    , "chked"     : chked
                                    }
                            ,dataType: "json"
                            ,success: function(data){
        						var result = data.result;
        						if( result=="succ"){
        							if(chked)
                                    {
                                        alert("�������� ����!");
                                    }
                                    else
                                    {
                                        alert("�������� ����!");
                                    }
                                }
                            }
                            ,error: function(jqXHR, status, errorThrown){
                                alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
                            }
                    });
                });
            });
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>

                <form action="saupbu_emp_report_kdc.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�ٹ����&nbsp;</strong>(��201401) :
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>
                                <label>
								<strong>���� &nbsp;:</strong>
								<%
									Dim rs_org

                                    'sql_org="select saupbu from sales_org order by sort_seq"
                                    'sql_org="select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"

									objBuilder.Append "SELECT org_bonbu FROM emp_org_mst "
									objBuilder.Append "WHERE org_code > '6505' "
									objBuilder.Append "	AND org_bonbu NOT IN (' ', '�濵����', '����Ÿ������', '�����ι�', '���������') "
									objBuilder.Append "GROUP BY org_bonbu "
									objBuilder.Append "ORDER BY org_code "

									Set rs_org = Server.CreateObject("ADODB.RecordSet")
                                    rs_org.Open objBuilder.ToString(), DBConn, 1
									objBuilder.Clear()
								%>
                                <select name="sales_bonbu" id="sales_bonbu" style="width:150px">
                                    <option value="��ü" <% if sales_bonbu = "��ü" then %>selected<% end if %>>��ü</option>
                                    <%
                                    Do Until rs_org.EOF
                                    %>
                                    <option value='<%=rs_org("org_bonbu")%>' <%If rs_org("org_bonbu") = sales_bonbu  then %>selected<% end if %>>
										<%=rs_org("org_bonbu")%>
									</option>
                                    <%
                                        rs_org.MoveNext()
                                    Loop
                                    rs_org.Close() : Set rs_org = Nothing
                                    %>
                                </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

                <table cellpadding="0" cellspacing="0" width="100%">
                <tr>
                    <td>
                    <DIV id="topLine2" style="width:1200px;overflow:hidden;">
                    <div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="*" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
                            <col width="2%" >
							<col width="1%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">��</th>
								<th scope="col">����ó</th>
								<th scope="col">����ȸ��</th>
								<th scope="col">���</th>
								<th scope="col">�����</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��뱸��</th>
								<th scope="col">��������</th>
								<th scope="col">�޿��Ѿ�</th>
								<th scope="col">��Ư��</th>
                                <th scope="col">���� ����</th>
								<th scope="col"></th>
							</tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="3%">
							<col width="*" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
                            <col width="2%" >
							<col width="1%" >
						</colgroup>
						<tbody>
						<%
						Dim i, j, team_sum, team_overtime_sum, tot_sum, tot_overtime_sum, bi_team

						i = 0
						j = 0
						team_sum = 0
						team_overtime_sum = 0
						tot_sum = 0
						tot_overtime_sum = 0
						bi_team = "first"

						objBuilder.Append "SELECT eomt.org_company, eomt.org_bonbu, eomt.org_team, emmt.emp_reside_place, emmt.emp_reside_company, "
						objBuilder.Append "	emmt.emp_no, emmt.emp_name, emmt.emp_job, emmt.cost_center, emmt.mg_saupbu, "
						objBuilder.Append "	emmt.emp_month, emmt.cost_except, "
						objBuilder.Append " pmgt.pmg_give_total, pmgt.pmg_job_support "
						objBuilder.Append "FROM emp_org_mst AS eomt "
						objBuilder.Append "INNER JOIN emp_master_month as emmt ON eomt.org_code = emmt.emp_org_code "
						objBuilder.Append "	AND emmt.emp_month = '202201' AND emmt.emp_pay_id <> '2' "
						objBuilder.Append "INNER JOIN pay_month_give as pmgt ON emmt.emp_no = pmgt.pmg_emp_no "
						objBuilder.Append "	AND pmgt.pmg_yymm = '"&cost_month&"' AND pmgt.pmg_id = '1' "

						If sales_bonbu <> "��ü" then
							'objBuilder.Append "	AND emmt.mg_saupbu = '"&sales_bonbu&"' "
							objBuilder.Append "WHERE eomt.org_bonbu = '"&sales_bonbu&"' "
						End If

						objBuilder.Append "ORDER BY eomt.org_company, eomt.org_bonbu DESC, eomt.org_team, emmt.cost_except, "
						objBuilder.Append "FIELD(emmt.emp_job, "
						objBuilder.Append "'����', '�λ���', '�Ѱ���ǥ', '�̻�', '�����̻�', '���̻�', '��������', '��������', '����', '����', "
						objBuilder.Append "'����', '����������', 'å�ӿ�����', '�븮', '�븮1��', '�븮2��', '���ӿ�����', '������', '���', ' ') "

						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open objBuilder.ToString(), DBConn, 1
						objBuilder.Clear()

						Do Until rs.EOF
							If bi_team = "first" Then
								bi_team = rs("org_team")
							End If

							If bi_team <> rs("org_team") Then
                                %>
                                <tr bgcolor="#FFFFCC">
                                    <td class="first">�Ұ�</td>
                                    <td>�ο���&nbsp;&nbsp;<%=j%></td>
                                    <td><%=bI_team%>&nbsp;</td>
                                    <td colspan="8">&nbsp;</td>
                                    <td class="right">
                                    <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" Then%>
                                        <%=FormatNumber(team_sum, 0)%>
                                    <%Else	%>
                                        ********
                                    <%End If%>

                                    </td>
                                    <td class="right">
                                    <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" Then%>
                                        <%=FormatNumber(team_overtime_sum, 0)%>
                                    <%Else %>
                                        ********
                                    <%End If %>
                                    </td>
                                    <td></td>
                                    <td></td>
                                </tr>
                                <%
								j = 0
								bi_team = rs("org_team")
								team_sum = 0
								team_overtime_sum = 0
							End If

                            ' �������ܰ��� ���� '2019.08.27
                            If rs("cost_except") <> "2" Then
                                i = i + 1
                                j = j + 1
                            End If

							Dim pmg_give_total, pmg_job_support, emp_end_date

						  	pmg_give_total = rs("pmg_give_total")
						  	pmg_job_support = rs("pmg_job_support")

							team_sum = team_sum + pmg_give_total
							team_overtime_sum = team_overtime_sum + pmg_job_support
							tot_sum = tot_sum + pmg_give_total
							tot_overtime_sum = tot_overtime_sum + pmg_job_support
                            %>
                            <tr>
                                <td class="first"><%=i%></td>
								<td><%=rs("org_company")%>&nbsp;</td>
                                <td><%=rs("org_bonbu")%>&nbsp;</td>
                                <td><%=rs("org_team")%>&nbsp;</td>
                                <td><%=rs("emp_reside_place")%>&nbsp;</td>
                                <td><%=rs("emp_reside_company")%>&nbsp;</td>
                                <td><%=rs("emp_no")%></td>
                                <td><%=rs("emp_name")%></td>
                                <td><%=rs("emp_job")%></td>
                                <td><%=emp_end_date%>&nbsp;</td>
                                <td><%=rs("cost_center")%></td>
                                <td><%=rs("mg_saupbu")%>&nbsp;</td>
                                <td class="right">
                                <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" Then%>
                                    <%=FormatNumber(pmg_give_total, 0)%>
                                <%Else%>
                                    ********
                                <%End If%>
                                </td>
                                <td class="right">
                                <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" then%>
                                    <%=FormatNumber(pmg_job_support, 0)%>
                                <%Else %>
                                    ********
                                <%End If%>
                                </td>
                                <td>
                                    <!-- �������� ���θ� ǥ�� (2019.08.27) -->
                                    <input type="checkbox" name="cost_except" emp_month="<%=rs("emp_month")%>" emp_no="<%=rs("emp_no")%>" <% If  (rs("cost_except") = "2") Then %>checked<% End If %>>
                                </td>
                                <td></td>
                            </tr>
                            <%
							rs.MoveNext()
						Loop
						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
							<tr bgcolor="#FFFFCC">
								<td class="first">�Ұ�</td>
								<td>�ο���&nbsp;&nbsp;<%=j%></td>
								<td><%=bI_team%>&nbsp;</td>
								<td colspan="8">&nbsp;</td>
								<td class="right">
                                <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" Then %>
                                    <%=FormatNumber(team_sum, 0)%>
                                <%Else	%>
                                    ********
                                <%End If%>
                                </td>
                                <td class="right">
                                <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" Then%>
                                    <%=FormatNumber(team_overtime_sum, 0)%>
                                <%Else	%>
                                    ********
                                <%End If%>
                                </td>
                                <td></td>
                            </tr>
                            <tr bgcolor="#FFE8E8">
                                <td colspan="2" class="first">�Ѱ�</td>
                                <td>�ο���&nbsp;&nbsp;<%=i%></td>
                                <td>&nbsp;</td>
                                <td colspan="8">&nbsp;</td>
                                <td class="right">
                                <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" Then	%>
                                    <%=FormatNumber(tot_sum, 0)%>
                                <%Else%>
                                    ********
                                <%End If%>
                                </td>
                                <td class="right">
                                <%If (position = "�������" And sales_bonbu = bonbu) Or user_id = "102592" Then	%>
                                    <%=FormatNumber(tot_overtime_sum, 0)%>
                                <%Else%>
                                    ********
                                <%End If%>
								</td>
								<td></td>
                                <td></td>
							</tr>
						</tbody>
						</table>
                        </DIV>
						</td>
                    </tr>
					</table>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td width="25%">
                            <div class="btnCenter">
                            <a href="/saupbu_emp_excel.asp?cost_month=<%=cost_month%>&sales_bonbu=<%=sales_bonbu%>" class="btnType04">�����ٿ�ε�</a>
                            </div>
                        </td>
                        <td width="50%"></td>
                        <td width="25%"></td>
                    </tr>
				    </table>
			    </form>
			    <br>
		    </div>
	    </div>
	</body>
</html>


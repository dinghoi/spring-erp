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
Dim page, page_cnt, pg_cnt, be_pg, curr_date
Dim ck_sw, srchWord, srchCategory, view_sort
Dim view_condi, pgsize, start_page, stpage
Dim listCostCenter
Dim order_sql, where_sql, total_record, total_page, title_line
Dim rs_etc, rsCount, rs

be_pg = "/insa/insa_emp_master_mg.asp"

page = Request("page")
page_cnt = Request.Form("page_cnt")
pg_cnt = CInt(Request("pg_cnt"))
curr_date = DateValue(Mid(CStr(Now()), 1, 10))
ck_sw = Request("ck_sw")
srchWord = Request("srchWord")
srchCategory = Request("srchCategory")
view_sort = Request("view_sort")

'srchEmpName = Request("srchEmpName")
'srchEmpMonth = Request("srchEmpMonth")

If ck_sw = "y" Then
	view_condi = Request("view_condi")
Else
	view_condi = Request.Form("view_condi")
End if

If view_condi = "" Then
	view_condi = "���̿�"
End If

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

'//������ ��
'If Trim(srchEmpMonth&"") = "" Then
'	srchEmpMonth = Left(Replace(DateAdd("m", -1, Now()), "-", ""), 6)
'End If

'sql="select max(pmg_yymm) as max_pmg_yymm from pay_month_give "
'set rs_max=dbconn.execute(sql)
'If Not(rs_max.bof Or rs_max.eof) Then
'	empMonth = rs_max("max_pmg_yymm")
'End If
'rs_max.close : Set rs_max = Nothing

'//��뱸��
objBuilder.Append "SELECT emp_etc_name "
objBuilder.Append "FROM emp_etc_code "
objBuilder.Append "WHERE emp_etc_type = '70' "
objBuilder.Append "ORDER BY emp_etc_code ASC "

Set rs_etc = Server.CreateObject("ADODB.Recordset")
rs_etc.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Set listCostCenter = getRsToDic(rs_etc)
rs_etc.close : Set rs_etc = Nothing

If view_sort = "" Then
	view_sort = "ASC"
End If

'order_sql = "ORDER BY eomt.org_company " & view_sort & " "

'where_sql = where_sql & " AND A.pmg_id=1 " & chr(13) ' �� �� ���ο� �ּ��� ��Ҿ���? �����ź��� ���� 2180-08-14 (�����Ȳ����/����������Ȳ ���� 2���̻󳪿��� ����..)
'where_sql = where_sql & " AND B.cost_except in ('0','1') " & chr(13)	'���� �ּ� ó�� �ڵ�

'ȸ��� �˻�
'If view_condi <> "��ü" Then
'    where_sql = where_sql & " AND A.pmg_company='" & view_condi & "' " & chr(13)
'End If

'�̸� �˻�
'If Trim(srchWord & "") <>"" Then
'    where_sql = where_sql & " AND B." & srchCategory & " like '%" & srchWord & "%' " & chr(13)
'End If

'Sql = "SELECT count(*) FROM pay_month_give  A ,emp_master_month B  " & where_sql

'objBuilder.Append "SELECT COUNT(*) "
'objBuilder.Append "FROM pay_month_give AS pmgt "
'objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
'objBuilder.Append "	AND emmt.emp_month = '" & srchEmpMonth & "' "
'objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
'objBuilder.Append "WHERE pmgt.pmg_id = '1' "
'objbuilder.Append "	AND pmgt.pmg_yymm = '" & srchEmpMonth & "' "

objBuilder.Append "SELECT COUNT(*) FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emtt.emp_pay_id <> '2' "

If view_condi <> "��ü" Then
	objBuilder.Append "AND eomt.org_company = '" & view_condi & "' "
End If

If Trim(srchWord) <> "" Then
	objBuilder.Append "AND emtt."& srchCategory &" LIKE '%" & srchWord & "%' "
End If

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(RsCount(0)) 'Result.RecordCount

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

'sql = " SELECT  A.pmg_yymm           " & chr(13) & _
'      "       , A.pmg_company        " & chr(13) & _
'      "       , A.pmg_saupbu         " & chr(13) & _
'      "       , A.pmg_give_total     " & chr(13) & _
'      "       , A.cost_group         " & chr(13) & _
'      "       , A.cost_center        " & chr(13) & _
'      "       , B.emp_no             " & chr(13) & _
'      "       , B.emp_name           " & chr(13) & _
'      "       , B.emp_job            " & chr(13) & _
'      "       , B.emp_type           " & chr(13) & _
'      "       , B.emp_saupbu         " & chr(13) & _
'      "       , B.emp_org_name       " & chr(13) & _
'      "       , B.emp_company        " & chr(13) & _
'      "       , B.emp_bonbu          " & chr(13) & _
'      "       , B.emp_team           " & chr(13) & _
'      "       , B.emp_reside_company " & chr(13) & _
'      "       , B.emp_reside_place   " & chr(13) & _
'      "    FROM pay_month_give A     " & chr(13) & _
'      "       , emp_master_month B   " & chr(13) & _
'      where_sql                        & chr(13) & _
'      order_sql                        & chr(13) & _
'      " LIMIT "& stpage & ", " & pgsize & chr(13)

objBuilder.Append "SELECT * "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emtt.emp_pay_id <> '2' "

If view_condi <> "��ü" Then
	objBuilder.Append "AND eomt.org_company = '" & view_condi & "' "
End If

If Trim(srchWord) <> "" Then
	objBuilder.Append "AND emtt."& srchCategory &" LIKE '%" & srchWord & "%' "
End If

objBuilder.Append "ORDER BY eomt.org_name " & view_sort & ", "
objBuilder.Append "	FIELD(emtt.emp_job, '����', '�λ���', '�Ѱ���ǥ', '�����̻�', '���̻�', '�̻�', "
objBuilder.Append "		'��������', '��������', '����', '����', '����', '����������', 'å�ӿ�����', "
objBuilder.Append "		'�븮', '�븮1��', '�븮2��', '���ӿ�����', '������', '���') "

objBuilder.Append "LIMIT "& stpage & ", " & pgsize

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

title_line = " �λ� ������ ��Ȳ "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "1 1"
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if (document.frm.view_condi.value == ""){
					alert ("�ʵ� ������ �����Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}

			function callbackOrgSelect(empNo, costGroup){
				var tEmpNo = "";
				$(".tableList tbody tr").each(function(){
					tEmpNo = $(this).find("td").eq(0).text();
					if( tEmpNo == empNo ){
						$(this).find("input[name='cost_group']").val(costGroup);
					}
				});
			}

            function changeEmpMasterMonth(empNo){
				//empMonth = $("#srchEmpMonth").val();

				var tEmpNo = "";
				var tCostCenter = "";
				var tCostGroup = "";

				var tEmpOrgCode = "";
				var tEmpOrgName = "";
				var tEmpCompany = "";
				var tEmpBonbu = "";
				var tEmpSaupbu = "";
				var tEmpTeam = "";
				var tEmpResideCompany = "";
				var tEmpResidePlace = "";

				$(".tableList tbody tr").each(function(){
					tEmpNo = $(this).find("td").eq(0).text();

					if( tEmpNo == empNo ){
						tCostCenter = $(this).find("select option:selected").val();
						//tCostGroup = $(this).find("input[name='cost_group']").val();
						tEmpOrgCode = $(this).find("input[name='emp_org_code']").val();
						tEmpOrgName = $(this).find("input[name='emp_org_name']").val();
						tEmpCompany = $(this).find("input[name='emp_company']").val();
						tEmpBonbu = $(this).find("input[name='emp_bonbu']").val();
						tEmpSaupbu = $(this).find("input[name='emp_saupbu']").val();
						tEmpTeam = $(this).find("input[name='emp_team']").val();
						tEmpResideCompany = $(this).find("input[name='emp_reside_company']").val();
						return false;
					}
				});

				/*if(empMonth==null || empMonth==""){
					alert("��� ������ �����ϴ�.");
					return false;
				}*/

				if(tEmpNo==null || tEmpNo==""){
					alert("��������� �����ϴ�.");
					return false;
				}

				if(tCostCenter==null || tCostCenter==""){
					alert("��뱸���� �������ּ���.");
					return false;
				}
				/*
				if(tCostGroup==null || tCostGroup==""){
					alert("���׷��� �Է����ּ���.");
					return false;
				}*/

				var params = {
								"empNo"      : tEmpNo
                                ,"costCenter" : escape(tCostCenter)
                                //,"costGroup"  : escape(tCostGroup)
                                ,"empOrgCode" : escape(tEmpOrgCode)
                                ,"empOrgName" : escape(tEmpOrgName)
                                ,"empCompany" : escape(tEmpCompany)
                                ,"empBonbu"   : escape(tEmpBonbu)
                                ,"empSaupbu"  : escape(tEmpSaupbu)
                                ,"empTeam"    : escape(tEmpTeam)
								, "emp_reside_company" : escape(tEmpResideCompany)
                             };
				$.ajax({
 					 url: "/insa/insa_emp_master_month_mg_save.asp"
					,type: 'post'
					,data: params
					,dataType: "json"
					,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
					,beforeSend: function(jqXHR){
							jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
						}
					//,success:function(data, status, request){
					,success: function(data){
						var result = data.result;
						if( result=="succ"){
							alert("����ƽ��ϴ�.");
						}else if( result=="invalid" ){
							alert("�Է��Ͻ� ������ ��Ȯ���� �ʽ��ϴ�.");
						}else if(result=="fail"){
							alert("���� �����߽��ϴ�.");
						}
					}
					,error: function(jqXHR, status, errorThrown){
						alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
					}
				});
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--include virtual = "/include/insa_asses_promo_menu.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="<%=be_pg%>" method="post" name="frm">

				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>�˻�</dt>
						<dd>
							<p>
								<label for="view_condi"><strong>ȸ�� : </strong></label>
								<%
								Call SelectEmpOrgList("view_condi", "view_condi", "width:150px", view_condi)
								%>
								<!--<label for="srchEmpName"><strong>�̸� : </strong></label>-->
								<select name="srchCategory" id="srchCategory">
									<option value="emp_name"<%If srchCategory="emp_name" Then%>selected<%End If%>>�̸�</option>
									<option value="emp_no"<%If srchCategory="emp_no" Then%>selected<%End If%>>���</option>
								</select>
								<input type="text" name="srchWord" id="srchWord" style="width: 100px; text-align: left; -ms-ime-mode: active;" value="<%=srchWord%>" />
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="13%" >
							<col width="35%" >
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�� ��</th>
								<th scope="col">�� ��</th>
								<th scope="col">�� ��</th>
								<th scope="col">�� �� (������)</th>
								<th scope="col">�� �� (ȸ��/����/�����/��)</th>
								<th scope="col">���� ȸ��</th>
								<th scope="col">����ó</th>
								<th scope="col">�� �ٹ���</th>
								<th scope="col">��� ����</th>
								<th scope="col"></th>
							</tr>
						</thead>
					<tbody>
						<%
						Dim j
						Dim tEmpNo, tCostCenter, tCostGroup, tEmpOrgName, tEmpCompany
						Dim tEmpBonbu,tEmpTeam, tEmpResideCompany, tEmpResidePlace
						Dim tEmpOrgCode, row
						Dim tEmpSaupbu, tEmpStayName

						Int j = 0

						Do Until rs.EOF
							tEmpNo 				= rs("emp_no")
							tCostCenter 		= rs("cost_center")
							tCostGroup 			= rs("cost_group")

							'tEmpOrgName 		= rs("emp_org_name")
							'tEmpCompany 		= rs("emp_company")
							'tEmpBonbu 			= rs("emp_bonbu")
							'tEmpSaupbu 			= rs("emp_saupbu")
							'tEmpTeam 			= rs("emp_team")
							tEmpOrgName 		= rs("org_name")
							tEmpCompany 		= rs("org_company")
							tEmpBonbu 			= rs("org_bonbu")
							tEmpSaupbu 			= rs("org_saupbu")
							tEmpTeam 			= rs("org_team")

							tEmpResideCompany	= rs("emp_reside_company")
							tEmpResidePlace 	= rs("emp_reside_place")
							tEmpStayName = rs("emp_stay_name")

							j = j + 1
						    %>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
								<td><a href="#" onClick="pop_Window('/insa/insa_emp_master_modify.asp?view_condi=<%=tEmpCompany%>&emp_no=<%=tEmpNo%>&u_type=U','�λ�⺻���� ����','scrollbars=yes,width=1250,height=600')"><%=rs("emp_name")%></td>
								<td><%=rs("emp_job")%></td>
								<td>
									<input type="hidden" id="emp_org_code<%=j%>" name="emp_org_code" value="<%=tEmpOrgCode%>" />
									<input id="emp_org_name<%=j%>" name="emp_org_name" type="text" style="width:90px" readonly="true" value="<%=tEmpOrgName%>">
                                    <a href="#" class="btnType03" onClick="pop_Window('/insa/popup/insa_emp_master_org_select.asp?gubun=org&view_condi=<%=view_condi%>&org_id=<%=j%>','orgselect','scrollbars=yes,width=800,height=400')">����</a>
								</td>
								<td>
									<input type="text" id="emp_company<%=j%>" name="emp_company" readonly="true" value="<%=tEmpCompany%>" style="width:80px" />
									<input type="text" id="emp_bonbu<%=j%>" name="emp_bonbu" readonly="true" value="<%=tEmpBonbu%>" style="width:100px" />
									<input type="text" id="emp_saupbu<%=j%>" name="emp_saupbu" readonly="true" value="<%=tEmpSaupbu%>" style="width:100px" />
									<input type="text" id="emp_team<%=j%>" name="emp_team" readonly="true" value="<%=tEmpTeam%>" style="width:100px" />
									<input type="hidden" id="emp_reside_company<%=j%>" name="emp_reside_company" readonly="true" value="<%=tEmpResideCompany%>" />
									<input type="hidden" id="emp_reside_place<%=j%>" name="emp_reside_place" readonly="true" value="<%=tEmpResidePlace%>"   />
									<input type="hidden" id="emp_org_level<%=j%>" name="emp_org_level" readonly="true" value="" />
									<input type="hidden" id="emp_type<%=j%>" name="emp_type">
                                </td>
								<td>
									<!--<input type="text" id="cost_group<%=j%>" name="cost_group" value="<%=tCostGroup%>" readonly="readonly" />-->
									<%=rs("emp_reside_company")%>
								</td>
								<td><%=rs("emp_reside_place")%></td>
								<td>
									<input type="text" id="emp_stay_name<%=j%>" name="emp_stay_name" value="<%=tEmpStayName%>" style="width:110px" />
								</td>
                                <td>
									<select id="cost_center<%=j%>" name="cost_center" style="width:90px">
										<option value="">����</option>
										<%
											If IsObject(listCostCenter) Then
												If listCostCenter.count > 0 Then
													For i=0 to listCostCenter.count-1
														Set row = listCostCenter.item(i)
                                                        %>
                                                        <option value='<%=row("emp_etc_name")%>' <%If tCostCenter = row("emp_etc_name") Then %>selected<%End If %>><%=row("emp_etc_name")%></option>
                                                        <%
													Next
												End If
											End If
										%>
									</select>
								</td>
								<td><a href="#" class="btnType04" onClick="changeEmpMasterMonth('<%=tEmpNo%>')">����</a></td>
							</tr>
						    <%
							rs.MoveNext()
						Loop

						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (Int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<!--<div class="btnCenter">

                    <a href="insa_excel_emp.asp?view_condi=<%=view_condi%>" class="btnType04">�����ٿ�ε�</a>

					</div>-->
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="<%=be_pg%>?page=<%=first_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&ck_sw=y">[ó��]</a>
                  	<%If intstart > 1 Then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&ck_sw=y">[����]</a>
                    <%End If %>
                    <%For i = intstart To intend %>
                  	<%	If i = int(page) Then %>
                        <b>[<%=i%>]</b>
                    <%	Else %>
                        <a href="<%=be_pg%>?page=<%=i%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&ck_sw=y">[<%=i%>]</a>
                    <%	End If %>
                    <%Next %>

                  	<%If intend < total_page Then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&ck_sw=y">[����]</a>
						<a href="<%=be_pg%>?page=<%=total_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&ck_sw=y">[������]</a>
                    <%Else %>
                        [����]&nbsp;[������]
                    <%End If %>
                    </div>
                    </td>
				    <td width="20%">
					<!--<div class="btnCenter">

                    <a href="#" onClick="pop_Window('insa_emp_add01.asp?view_condi=<%=view_condi%>&u_type=<%=""%>','insa_emp_add01_popup','scrollbars=yes,width=1250,height=600')" class="btnType04">�ű�ä����</a>

					</div>-->
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>


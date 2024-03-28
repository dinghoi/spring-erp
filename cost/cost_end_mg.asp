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
Dim rs_insa, max_org_month
Dim org_bonbu, rsOrgCode
Dim rsEmpOrgList
Dim title_line

objBuilder.Append "SELECT MAX(org_month) AS max_org_month "
objBuilder.Append "FROM emp_org_mst_month "

Set rs_insa = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_insa("max_org_month")) Then
    max_org_month = "000000"
Else
    max_org_month = rs_insa("max_org_month")
End If
rs_insa.Close() : Set rs_cost = Nothing

'��� ������ 0�� �ƴ� ��� ���θ� �˻�[����ȣ_20210306]
If cost_grade <> "0" Then
	objBuilder.Append "SELECT eomt.org_bonbu "
	objBuilder.Append "FROM emp_master_month AS emmt "
	objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
	objBuilder.Append "WHERE emmt.emp_month = '"&max_org_month&"' "
	objBuilder.Append "	AND emmt.emp_no = '"&emp_no&"' "

	Set rsOrgCode = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	org_bonbu = rsOrgCode("org_bonbu")

	rsOrgCode.Close() : Set rsOrgCode = Nothing
End If

' 2019.02.22 ������ �䱸 'N/W 1�����','N/W 2�����',"SI3�����","�ַ�ǻ����"	�� �������ʵ��� �������� ó��..
'sql = "SELECT *                                " & chr(13) & _
'      "  FROM emp_org_mst                      " & chr(13) & _
'      " WHERE (org_level = '�����')           " & chr(13) & _
'      "   AND (org_name <> '�Ѱ���ǥ')         " & chr(13) & _
'      "   AND (    ISNULL(org_end_date)        " & chr(13) & _
'      "         OR org_end_date = '0000-00-00' " & chr(13) & _
'      "       )                                " & chr(13)
' org_end_date = '' or   ....   date���� '' ���� ���Ҽ�����.   Warning: Incorrect date value: '' for column 'org_end_date' at row 1

objBuilder.Append "SELECT org_name, org_date "
objBuilder.Append "FROM emp_org_mst "
objBuilder.Append "WHERE org_level = '����' "
objBuilder.Append "	AND (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
'objBuilder.Append "	AND org_name NOT IN ('�����ι�', 'ICT������', '����Ÿ������', '���������', '�����׷�������')"
objBuilder.Append "	AND org_name NOT IN ('�����ι�', 'ICT������', '���������', '�����׷�������', 'SI���ົ��', '����Ʈ����')"

If cost_grade = "0" Then
	objBuilder.Append "GROUP BY org_bonbu, org_name "
	objBuilder.Append "ORDER BY FIELD(org_company, '���̿�', '���̳�Ʈ����', '���̽ý���'), "
	objBuilder.Append "	FIELD(org_bonbu, '����Ÿ������', '����Ʈ����', 'DI����ι�', '����SI����', '����SI����', 'ICT����', '��������', 'NI����', 'SI2����', 'SI1����') DESC "
Else
	objBuilder.Append "	AND (org_name = '"&org_bonbu&"' Or org_empno = '"&emp_no&"') "
	objBuilder.Append "GROUP BY org_name "
End If

Set rsEmpOrgList = DBConn.Execute(objBuilder.ToSTring)
objBuilder.Clear()

title_line = "��� ���� ����"
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
				return "1 1";
			}

			function frmcheck(){
				document.frm.submit();
			}

			//��븶�� ó��
			function setCostEnd(end_url, end_month, dept, end_yn, type){
				if(type == 'A'){
					if(!confirm('�ش� ��(' +end_month + ')�� ��� ������ �ϰ� �����Ͻðڽ��ϱ�?')){
						return false;
					}
				}

				var param = {"end_month":end_month, "saupbu":encodeURIComponent(dept), "end_yn":end_yn};

				let start_time = new Date();

				$.ajax({
					type : "GET"
					, dataType : 'html'
					, contentType: "application/x-www-form-urlencoded; charset=EUC-KR"
					, url: end_url
					, data: param
					, async: true
					, error: function(request, status, error){
						console.log("code = "+ request.status + " message = " + request.responseText + " error = " + error);
					}
					, success: function(data){
						let end_time = new Date();
						var elapedMin = (end_time.getTime() - start_time.getTime()) / 1000 / 60;

						console.log('����ð�(��) : ' + elapedMin);
						console.log(data);
						console.log($(window).scrollTop());

						alert(data);
						location.href="/cost/cost_end_mg.asp";
						return;
					}
					, beforeSend: function(){
						var width = 0;
						var height = 0;
						var left = 0;
						var top = 0;

						width = 220;
						height = 118;
						top = ( $(window).height() - height ) / 2 + $(window).scrollTop();
						left = ( $(window).width() - width ) / 2 + $(window).scrollLeft();

						if($("#div_ajax_load_image").length != 0){
							$("#div_ajax_load_image").css({
								"top": top+"px",
								"left": left+"px"
							});
							$("#div_ajax_load_image").show();
						}else{
							$('body').append('<div id="div_ajax_load_image" style="position:absolute; top:' + top + 'px; left:' + left + 'px; width:' + width + 'px; height:' + height + 'px; z-index:9999; background:#f0f0f0; filter:alpha(opacity=50); opacity:alpha*0.5; margin:auto; padding:0; "><img src="/image/wait.gif" style="width:220px; height:118px;"></div>');
						}
					}
					, complete: function(){
						$("#div_ajax_load_image").hide();
					}
				});
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/cost_end_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ǰ˻�</dt>
						<dd>
						<p>
							<label>&nbsp;&nbsp;<strong>�ֽ������� �ٽ� ��ȸ�ϱ�&nbsp;</strong></label>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
							(������ ������,�λ�� �����ڵ帶��, �λ縶��[<%=max_org_month%>] Ȯ��)
						</p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="13%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">�����</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� �� �� �� �� Ȳ</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">���ο� ���� ó��</th>
								<th rowspan="2" scope="col">�����ڷ�</th>
								<th rowspan="2" scope="col">�����庸��</th>
								<th rowspan="2" scope="col">CEO����</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">�������</th>
							  <th scope="col">��������</th>
							  <th scope="col">������</th>
							  <th scope="col">ó������</th>
							  <th scope="col">�������</th>
							  <th scope="col">�������</th>
							  <th scope="col">����ó��</th>
              				</tr>
						</thead>
						<tbody>
						<%
							Dim rsCostEndMax, rsCostEndList

							Dim cancel_yn, rs_cost, new_date
							Dim new_month, now_month, end_view, end_yn, end_month
							Dim reg_name, reg_id, reg_date, batch_view, bonbu_view
							Dim ceo_view, batch_yn, bonbu_yn, ceo_yn

							Dim jik_yn

							'=====	����� �� ���� �׸� ����Ʈ	=====
							Do Until rsEmpOrgList.EOF
								cancel_yn = "N"

								'If rs("org_bonbu") = "���һ����" Then
								'	If rs("org_saupbu") = "�������������" Or rs("org_saupbu") = "KAL���������" Then
								'		jik_yn = "N"
								'	Else
								'		jik_yn = "Y"
							  	'	End If
								'Else
							  	'	jik_yn = "N"
								'End If

								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '"&rsEmpOrgList("org_name")&"' "

								Set rsCostEndMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '"&rsEmpOrgList("org_name")&"' "
								objBuilder.Append "	AND end_month = '"&rsCostEndMax("max_month")&"' "

								rsCostEndMax.Close()

								Set rsCostEndList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndList.EOF Or rsCostEndList.BOF Then
									new_date = DateAdd("m", -1, Now())
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "����"

									If end_month = "����" Then
										new_date = rsEmpOrgList("org_date")
										new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									End If

									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								 Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndList("end_month"), 1, 4) & "-" & Mid(rsCostEndList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndList("end_yn") = "Y" Then
										end_view = "����"
									ElseIf rsCostEndList("end_yn") = "C" Then
										new_month = rsCostEndList("end_month")
										end_view = "���"
									Else
										end_view = "����"
									End If

									end_yn = rsCostEndList("end_yn")
									end_month = rsCostEndList("end_month")
									reg_name = rsCostEndList("reg_name")
									reg_id = rsCostEndList("reg_id")
									reg_date = rsCostEndList("reg_date")

									If rsCostEndList("batch_yn") = "Y" Then
										batch_view = "�ڷ����"
									Else
								  		batch_view = "�̻���"
									End If

									If rsCostEndList("bonbu_yn") = "Y" Then
										bonbu_view = "���οϷ�"
									End If

									If rsCostEndList("ceo_yn") = "Y" Then
										ceo_view = "���οϷ�"
									End If

									If rsCostEndList("batch_yn") = "Y" And rsCostEndList("bonbu_yn") = "N" Then
										bonbu_view = "������"
									  	ceo_view = ""
									End If

									If rsCostEndList("bonbu_yn") = "Y" And rsCostEndList("ceo_yn") = "N" Then
										ceo_view = "������"
									End If

									If rsCostEndList("batch_yn") = "N" And rsCostEndList("bonbu_yn") = "N" And rsCostEndList("ceo_yn") = "N" Then
										bonbu_view = ""
										ceo_view = ""
									End If

									batch_yn = rsCostEndList("batch_yn")
									bonbu_yn = rsCostEndList("bonbu_yn")
									ceo_yn = rsCostEndList("ceo_yn")
								End If

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
							  		If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr>
								<td class="first"><%=rsEmpOrgList("org_name")%></td>
								<td><%=end_month%></td>
								<td>
									<%
									If end_view = "���" Then
										Response.Write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.Write end_view
									End If
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
									<%
									If cancel_yn = "Y" Then
										Response.write "<a href='/cost/cost_end_cancel.asp?saupbu="&rsEmpOrgList("org_name")&"&end_month="&end_month&"' class='btnType03'>�������</a>"
									Else
										Response.write "��ҺҰ�"
									End If
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>"
									<%If SysAdminYn <> "Y" Then%> readonly="true" <%End If%> />
								</td>
								<td>
									<%
									if now_month > new_month then
										'Response.write "<a href='/cost/cost_end_pro.asp?saupbu="&rsEmpOrgList("org_name")&"&end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>����</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/cost_end_pro.asp"", """&new_month&""", """&rsEmpOrgList("org_name")&""", """&end_yn&""", ""S"");' class='btnType03'>����</a>"
									else
										Response.write "�����Ұ�"
									end if
									%>
                				</td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
							<%
								rsEmpOrgList.MoveNext()
							Loop
							rsCostEndList.Close() : Set rsCostEndList = Nothing
							Set rsCostEndMax = Nothing
							rsEmpOrgList.Close() : Set rsEmpOrgList = Nothing

							'=====	����οܳ�����	=====
							Dim rsCostEndEtcList, rsCostEndEtcMax

							If cost_grade = "0" Then
								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu='����οܳ�����' "

								Set rsCostEndEtcMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '����οܳ�����' "
								objBuilder.Append "	AND end_month = '"&rsCostEndEtcMax("max_month")&"' "

								rsCostEndEtcMax.Close() : Set rsCostEndEtcMax = Nothing

								Set rsCostEndEtcList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndEtcList.EOF Or rsCostEndEtcList.BOF Then
									new_date = "2015-01-01"
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "����"
									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndEtcList("end_month"), 1, 4) & "-" & Mid(rsCostEndEtcList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndEtcList("end_yn") = "Y" Then
										end_view = "����"
									ElseIf rsCostEndEtcList("end_yn") = "C" Then
										new_month = rsCostEndEtcList("end_month")
										end_view = "���"
									Else
										end_view = "����"
									End If

									end_yn = rsCostEndEtcList("end_yn")
									end_month = rsCostEndEtcList("end_month")
									reg_name = rsCostEndEtcList("reg_name")
									reg_id = rsCostEndEtcList("reg_id")
									reg_date = rsCostEndEtcList("reg_date")

									If rsCostEndEtcList("batch_yn") = "Y" Then
										batch_view = "�ڷ����"
									Else
										batch_view = "�̻���"
									End If

									If rsCostEndEtcList("bonbu_yn") = "Y" Then
										bonbu_view = "���οϷ�"
									End If

									If rsCostEndEtcList("ceo_yn") = "Y" Then
										ceo_view = "���οϷ�"
									End If

									If rsCostEndEtcList("batch_yn") = "Y" And rsCostEndEtcList("bonbu_yn") = "N" Then
										bonbu_view = "������"
									  ceo_view = ""
									End If

									If rsCostEndEtcList("bonbu_yn") = "Y" And rsCostEndEtcList("ceo_yn") = "N" Then
										ceo_view = "������"
									End If

									If rsCostEndEtcList("batch_yn") = "N" And rsCostEndEtcList("bonbu_yn") = "N" And rsCostEndEtcList("ceo_yn") = "N" Then
										bonbu_view = ""
									  ceo_view = ""
									End If

									batch_yn = rsCostEndEtcList("batch_yn")
									bonbu_yn = rsCostEndEtcList("bonbu_yn")
									ceo_yn = rsCostEndEtcList("ceo_yn")
								End If
								rsCostEndEtcList.Close() : Set rsCostEndEtcList = Nothing

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
									If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr bgcolor="#FFE8E8">
								<td class="first">����οܳ�����</td>
								<td><%=end_month%></td>
								<td>
									<%
									If end_view = "���" Then
										Response.Write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.Write end_view
									End If
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
									<%
									If cancel_yn = "Y" Then
										Response.write "<a href='/cost/cost_bonbu_end_cancel.asp?saupbu=����οܳ�����&end_month="&end_month&"' class='btnType03'>�������</a>"
									Else
										Response.write "��ҺҰ�"
									End If
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									If now_month > new_month Then
										'Response.write "<a href='/cost/cost_bonbu_end_pro.asp?saupbu=����οܳ�����&end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>����</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/cost_bonbu_end_pro.asp"", """&new_month&""", ""����οܳ�����"", """&end_yn&""", ""S"");' class='btnType03'>����</a>"
									Else
										Response.write "�����Ұ�"
									End If
									%>
								</td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
							<%
							End If

							'=====	���� ���	=====
							If resideEndViewYn = "Y" Then
								Dim rsCostEndMonthMax, rsCostEndResideList

								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '���ֺ��' "

								Set rsCostEndMonthMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end  "
								objBuilder.Append "WHERE saupbu = '���ֺ��' "
								objBuilder.Append "	AND end_month = '"&rsCostEndMonthMax("max_month")&"'"

								Set rsCostEndResideList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndResideList.EOF Or rsCostEndResideList.BOF Then
									new_date = "2015-01-01"
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "����"
									end_yn = ""
									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndResideList("end_month"), 1, 4) & "-" & Mid(rsCostEndResideList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndResideList("end_yn") = "Y" Then
										end_view = "����"
									ElseIf rsCostEndResideList("end_yn") = "C" Then
										new_month = rsCostEndResideList("end_month")
										end_view = "���"
									Else
										end_view = "����"
									End If

									end_yn = rsCostEndResideList("end_yn")
									end_month = rsCostEndResideList("end_month")
									reg_name = rsCostEndResideList("reg_name")
									reg_id = rsCostEndResideList("reg_id")
									reg_date = rsCostEndResideList("reg_date")

									If rsCostEndResideList("batch_yn") = "Y" Then
										batch_view = "�ڷ����"
									Else
										batch_view = "�̻���"
									End If

									If rsCostEndResideList("bonbu_yn") = "Y" Then
										bonbu_view = "���οϷ�"
									End If

									If rsCostEndResideList("ceo_yn") = "Y" Then
										ceo_view = "���οϷ�"
									End If

									If rsCostEndResideList("batch_yn") = "Y" And rsCostEndResideList("bonbu_yn") = "N" Then
										bonbu_view = "������"
									  ceo_view = ""
									End If

									If rsCostEndResideList("bonbu_yn") = "Y" And rsCostEndResideList("ceo_yn") = "N" Then
										ceo_view = "������"
									End If

									If rsCostEndResideList("batch_yn") = "N" And rsCostEndResideList("bonbu_yn") = "N" And rsCostEndResideList("ceo_yn") = "N" Then
										bonbu_view = ""
										ceo_view = ""
									End If

									batch_yn = rsCostEndResideList("batch_yn")
									bonbu_yn = rsCostEndResideList("bonbu_yn")
									ceo_yn = rsCostEndResideList("ceo_yn")
								End If

								rsCostEndResideList.Close() : Set rsCostEndResideList = Nothing

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
									If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr bgcolor="#FFFFCC">
								<td class="first">���ֺ��</td>
								<td><%=end_month%></td>
								<td>
									<%
									If end_view = "���" Then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.write end_view
									End If
									%>&nbsp;
                				</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  <td>
									<%
									If cancel_yn = "Y" Then
										Response.Write "<a href='/cost/company_cost_end_cancel.asp?end_month="&end_month&"'  class='btnType03'>�������</a>"
									Else
										Response.Write "��ҺҰ�"
									End If
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									If now_month > new_month Then
										'Response.Write "<a href='/cost/company_cost_end_pro.asp?end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>����</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/company_cost_end_pro.asp"", """&new_month&""", """", """&end_yn&""", ""S"");' class='btnType03'>����</a>"
									Else
										Response.Write "�����Ұ�"
									End If
									%>
								</td>
							  	<td colspan="3">&nbsp;</td>
							</tr>
								<%'=====	�����/��������		=====
								Dim rsCostEndCommList, rsCostEndCommMax

								objBuilder.Append "SELECT MAX(end_month) AS max_month "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu='�����/��������' "

								Set rsCostEndCommMax = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								objBuilder.Append "SELECT end_month, end_yn, reg_name, reg_id, reg_date, "
								objBuilder.Append "	batch_yn, ceo_yn, bonbu_yn "
								objBuilder.Append "FROM cost_end "
								objBuilder.Append "WHERE saupbu = '�����/��������' "
								objBuilder.Append "	AND end_month ='"&rsCostEndCommMax("max_month")&"' "

								rsCostEndCommMax.Close() : Set rsCostEndCommMax = Nothing

								Set rsCostEndCommList = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCostEndCommList.EOF Or rsCostEndCommList.BOF Then
									new_date = "2015-01-01"
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
									end_month = "����"
									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
								Else
									new_date = DateAdd("m", 1, DateValue(Mid(rsCostEndCommList("end_month"), 1, 4) & "-" & Mid(rsCostEndCommList("end_month"), 5, 2) & "-01"))
									new_month = Mid(CStr(new_date), 1, 4) & Mid(CStr(new_date), 6, 2)
									now_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

									If rsCostEndCommList("end_yn") = "Y" Then
										end_view = "����"
									ElseIf rsCostEndCommList("end_yn") = "C" Then
										new_month = rsCostEndCommList("end_month")
										end_view = "���"
									Else
										end_view = "����"
									End If

									end_yn = rsCostEndCommList("end_yn")
									end_month = rsCostEndCommList("end_month")
									reg_name = rsCostEndCommList("reg_name")
									reg_id = rsCostEndCommList("reg_id")
									reg_date = rsCostEndCommList("reg_date")

									If rsCostEndCommList("batch_yn") = "Y" Then
										batch_view = "�ڷ����"
									Else
										batch_view = "�̻���"
									End If

									If rsCostEndCommList("bonbu_yn") = "Y" Then
										bonbu_view = "���οϷ�"
									End If

									If rsCostEndCommList("ceo_yn") = "Y" Then
										ceo_view = "���οϷ�"
									End If

									If rsCostEndCommList("batch_yn") = "Y" And rsCostEndCommList("bonbu_yn") = "N" Then
										bonbu_view = "������"
										ceo_view = ""
									End If

									If rsCostEndCommList("bonbu_yn") = "Y" And rsCostEndCommList("ceo_yn") = "N" Then
									  ceo_view = "������"
									End If

									If rsCostEndCommList("batch_yn") = "N" And rsCostEndCommList("bonbu_yn") = "N" And rsCostEndCommList("ceo_yn") = "N" Then
										bonbu_view = ""
										ceo_view = ""
									End If

									batch_yn = rsCostEndCommList("batch_yn")
									bonbu_yn = rsCostEndCommList("bonbu_yn")
									ceo_yn = rsCostEndCommList("ceo_yn")
								End If

								rsCostEndCommList.Close() : Set rsCostEndCommList = Nothing
								DBConn.Close() : Set DBConn = Nothing

								If jik_yn = "Y" Then
									If ceo_yn = "N" Then
										cancel_yn = "Y"
									End If
								Else
									If bonbu_yn = "N" Then
										cancel_yn = "Y"
									End If
								End If
							%>
							<tr bgcolor="#CCFFFF">
								<td class="first">�����/��������</td>
					  	  		<td><%=end_month%></td>
								<td>
									<%
									If end_view = "���" Then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									Else
										Response.write end_view
									End If
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
							  		<%
							  		If cancel_yn = "Y" Then
							  			Response.Write "<a href='/cost/company_as_sum_cancel.asp?end_month="&end_month&"' class='btnType03'>�������</a>"
									Else
										Response.Write "��ҺҰ�"
									End If
									%>
                				</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									If now_month > new_month Then
										'Response.Write "<a href='/cost/company_as_sum_pro.asp?end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>����</a>"
										Response.Write "<a href='#' onclick='setCostEnd(""/cost/company_as_sum_pro.asp"", """&new_month&""", """", """&end_yn&""", ""S"");' class='btnType03'>����</a>"
									Else
										Response.Write "�����Ұ�"
									End If
									%>
								</td>
								<td colspan="3">&nbsp;</td>
						  	</tr>
							<%
							End If
							%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>
<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

be_pg = "insa_pay_total_info.asp"

curr_date		= mid(cstr(now()),1,10)
curr_year		= mid(cstr(now()),1,4)
curr_month	= mid(cstr(now()),6,2)
curr_day		= mid(cstr(now()),9,2)


ck_sw	= replaceXSS(Request("ck_sw"))

if ck_sw = "n" then
	srchCompany		= replaceXSS(request.form("srchCompany"))
	emp_month_start	= replaceXSS(Request.form("emp_month_start"))
	emp_month_end	= replaceXSS(Request.form("emp_month_end"))
else
	srchCompany		= replaceXSS(request("srchCompany"))
	emp_month_start	= replaceXSS(request("emp_month_start"))
	emp_month_end	= replaceXSS(request("emp_month_end"))
end if

'if srchCompany = "" then
'	srchCompany	= "��ü"
'end if

curr_dd = cstr(datepart("d",now))
If emp_month_start = "" Then
	emp_month_start = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
End If
If emp_month_end = "" Then
	emp_month_end = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
End If

'//�� ��� ���ϱ�
emp_month_start_prev = cstr(CInt(Left(emp_month_start,4)-1)) + cstr(mid(emp_month_start,5,2))
emp_month_end_prev = cstr(CInt(Left(emp_month_end,4)-1)) + cstr(mid(emp_month_end,5,2))


' ��� ���̺����
cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
'cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
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

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


'//�˻��� ����� �ΰǺ�,�ο��� ��ȸ
sql = ""
sql = sql & "select "
sql = sql & "		emp_saupbu "
sql = sql & "		,emp_org_name "
'sql = sql & "		,sum(IF( emp_month = '" & emp_month_prev & "', sum_pmg_give_total, '0')) as prev_pmg_give_total "
'sql = sql & "		,sum(IF( emp_month = '" & emp_month_prev & "', cnt_emp, '0')) as prev_cnt_emp "
'sql = sql & "		,sum(IF( emp_month = '" & emp_month & "', sum_pmg_give_total, '0')) as cur_pmg_give_total "
'sql = sql & "		,sum(IF( emp_month = '" & emp_month & "', cnt_emp, '0')) as cur_cnt_emp "
sql = sql & "		,sum(IF( emp_month between '" & emp_month_start_prev & "' and '" & emp_month_end_prev & "', sum_pmg_give_total, '0')) as prev_pmg_give_total "
sql = sql & "		,sum(IF( emp_month between '" & emp_month_start_prev & "' and '" & emp_month_end_prev & "', cnt_emp, '0')) as prev_cnt_emp "
sql = sql & "		,sum(IF( emp_month between '" & emp_month_start & "' and '" & emp_month_end & "', sum_pmg_give_total, '0')) as cur_pmg_give_total "
sql = sql & "		,sum(IF( emp_month between '" & emp_month_start & "' and '" & emp_month_end & "', cnt_emp, '0')) as cur_cnt_emp "
sql = sql & "from ( "
sql = sql & "		select "
sql = sql & "			emm.emp_month "
sql = sql & "			,emm.emp_saupbu "
sql = sql & "			,emm.emp_org_name "
sql = sql & "			,sum(pmg.pmg_give_total) as sum_pmg_give_total "
sql = sql & "			,count(emm.emp_no) as cnt_emp "
sql = sql & "		from emp_master_month emm "
sql = sql & "		inner join pay_month_give pmg on emm.emp_no=pmg.pmg_emp_no "
sql = sql & "		where 1=1 "
sql = sql & "		and emm.emp_month = pmg.pmg_yymm "

'//ȸ�纰 �˻�
If srchCompany <> "" Then
sql = sql & "		and emm.emp_company='" & srchCompany & "' "
End If

sql = sql & "		and ( "
sql = sql & "		emm.emp_month between '" & emp_month_start & "' and '" & emp_month_end & "' "
sql = sql & "		or emm.emp_month between '" & emp_month_start_prev & "' and '" & emp_month_end_prev & "' "
sql = sql & "		) "
sql = sql & "		group by emm.emp_month, emm.emp_saupbu, emm.emp_org_name "
sql = sql & ") v "
sql = sql & "group by emp_saupbu, emp_org_name "
sql = sql & "order by emp_saupbu asc, emp_org_name asc, emp_month asc "

Response.write "<!-- sql::" & sql & " -->"

Rs.Open Sql, Dbconn, 1

Set List	= getRsToDic(Rs)

Rs.close() : Set rs = Nothing

'//title
curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " ����κ��ΰǺ���ȸ"

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
			function getPageCode(){
				return "6 2";
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
		</script>
		<script type="text/javascript">

			function frmcheck() {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit();
				}
			}
			
			//�˻��� üũ
			function chkfrm() {
				if ($("#srchCompany option:selected").val() == "") {
					alert ("ȸ�縦 �����Ͻñ� �ٶ��ϴ�");
					$("#srchCompany").focus();
					return false;
				}
				if ($("#emp_month_start option:selected").val() == "") {
					alert ("�˻��Ⱓ�� �����Ͻñ� �ٶ��ϴ�");
					$("#emp_month_start").focus();
					return false;
				}
				if ($("#emp_month_end option:selected").val() == "") {
					alert ("�˻��Ⱓ�� �����Ͻñ� �ٶ��ϴ�");
					$("#emp_month_end").focus();
					return false;
				}
				if( parseInt( $("#emp_month_start option:selected").val(), 10 ) > parseInt( $("#emp_month_end option:selected").val(), 10 ) ){
					alert("�˻��Ⱓ �������� �����Ϻ��� Ŭ �� �����ϴ�.");
					$("#emp_month_start").focus();
					return false;
				}
				return true;
			}
			
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}

			$(document).ready(function(){
				//ȸ��� �˻�����Ʈ ó��
				getOrg("1", "", "", "", "", "<%=srchCompany%>", "srchCompany", setCompanySelect);
			});
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_asses_promo_menu.asp" -->
			<div id="container">
				<h3 class="insa">����κ� �ΰǺ���ȸ</h3>
				<form action="<%=be_pg%>?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>�� �˻���</dt>
                        <dd>
                            <p>
                                <label>
								<strong>ȸ�� : </strong>
								<select name="srchCompany" id="srchCompany" type="text" style="width:130px">
            					</select>
                                </label>
								<label>
								<strong>�˻��Ⱓ : </strong>
                                  <select name="emp_month_start" id="emp_month_start" type="text" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If emp_month_start = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
									~
                                <label>
                                  <select name="emp_month_end" id="emp_month_end" type="text" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If emp_month_end = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();return false;"><image src="/image/but_ser1.jpg" alt="�˻�"></a>

								<a href="insa_pay_total_info_excel.asp?srchCompany=<%=srchCompany%>&emp_month_start=<%=emp_month_start%>&emp_month_end=<%=emp_month_end%>" class="btnType04">�����ٿ�ε�</a>

                            </p>
						</dd>
					</dl>
				</fieldset>
                <table cellpadding="0" cellspacing="0">
				  <tr>
                   	<td>
      				<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="150px" >
                            <col width="150px" >
							<col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2" style=" border-bottom:1px solid #e3e3e3;">�����</th>
								<th scope="col" rowspan="2" style=" border-bottom:1px solid #e3e3e3;">�Ҽ�</th>
								<th scope="col" colspan="2" style=" border-bottom:1px solid #e3e3e3;">
								<%=mid(emp_month_start_prev,1,4)%>��&nbsp;<%=mid(emp_month_start_prev,5,2)%>��
									~ <%=mid(emp_month_end_prev,1,4)%>��&nbsp;<%=mid(emp_month_end_prev,5,2)%>��

								</th>
								<th scope="col"  colspan="2" style=" border-bottom:1px solid #e3e3e3;">
								<%=mid(emp_month_start,1,4)%>��&nbsp;<%=mid(emp_month_start,5,2)%>��
									~ <%=mid(emp_month_end,1,4)%>��&nbsp;<%=mid(emp_month_end,5,2)%>��
								</th>
                                <th scope="col" colspan="2" style=" border-bottom:1px solid #e3e3e3;">���</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">�ΰǺ�</th>
								<th scope="col">�ο���</th>
								<th scope="col">�ΰǺ�</th>
								<th scope="col">�ο���</th>
								<th scope="col">�ΰǺ� �߰�</th>
								<th scope="col">�ο��� �߰�</th>
							</tr>
                        </thead>
                    </table>
                    </DIV>
					</td>
                  </tr>
                  <tr>
                    <td valign="top">
				    <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll;overflow-x:hidden;" onscroll="scrollAll()">
					<table cellpadding="0" cellspacing="0" class="scrollList">
                        <colgroup>
							<col width="150px" >
                            <col width="150px" >
							<col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
                            <col width="150px" >
						</colgroup>                        
                        <tbody>
                        <%
							Dim idxBgColor : idxBgColor = 0
							Dim bgColor : bgColro = ""

							If IsObject(List) Then
								If List.count > 0 Then

									saupbuRowColor = ""

									For i=0 to List.count-1
										Set row = List.item(i)
										emp_saupbu					= row("emp_saupbu")
										emp_org_name				= row("emp_org_name")
										cur_pmg_give_total			= CDbl(row("cur_pmg_give_total"))
										cur_cnt_emp					= CDbl(row("cur_cnt_emp"))
										prev_pmg_give_total		= CDbl(row("prev_pmg_give_total"))
										prev_cnt_emp				= CDbl(row("prev_cnt_emp"))
										
										'//���հ�
										sum_prev_pmg_give_total	= CDbl(sum_prev_pmg_give_total) + prev_pmg_give_total 
										sum_prev_cnt_emp			= CDbl(sum_prev_cnt_emp) + prev_cnt_emp
										sum_cur_pmg_give_total	= CDbl(sum_cur_pmg_give_total) + cur_pmg_give_total
										sum_cur_cnt_emp			= CDbl(sum_cur_cnt_emp) + cur_cnt_emp

										'//�ΰǺ� ó��
										give_total						= cur_pmg_give_total - prev_pmg_give_total
										sum_give_total				= CDbl(sum_give_total) + give_total

										If give_total<0 Then
											give_mark	= "��"
										ElseIf give_total>0 Then
											give_mark	= "��"
										Else
											give_mark	= ""
										End If

										'//�ο��� ó��
										cnt_emp						= cur_cnt_emp - prev_cnt_emp
										sum_cnt_emp				= CDbl(sum_cnt_emp) + cnt_emp

										If cnt_emp<0 Then
											emp_mark	= "��"
										ElseIf cnt_emp>0 Then
											emp_mark	= "��"
										Else
											emp_mark	= ""
										End If

										'//���հ� �ΰǺ� ó��
										If sum_give_total<0 Then
											sum_give_mark	= "��"
										ElseIf sum_give_total>0 Then
											sum_give_mark	= "��"
										Else
											sum_give_mark	= ""
										End If

										'//���հ� �ο��� ó��
										If sum_cnt_emp<0 Then
											sum_emp_mark	= "��"
										ElseIf sum_cnt_emp>0 Then
											sum_emp_mark	= "��"
										Else
											sum_emp_mark	= ""
										End If

										'//��� ����θ� ó��
										out_emp_saupbu = ""
										If bigo_saupbu <> emp_saupbu Then
											'bigo_saupbu = emp_saupbu
											out_emp_saupbu = emp_saupbu
											bigo_saupbu = emp_saupbu

											idxBgColor = idxBgColor + 1
											If idxBgColor Mod 2 = 0 Then
												rowBgColor = "#ffffff"
											Else
												rowBgColor = "#eeffff"
											End If
										End If
						%>	
							<tr style="background-color:<%=rowBgColor%>;border-bottom-color:<%=rowBgColor%>">
								<td class="first"><%= out_emp_saupbu%></td>
                                <td><%=emp_org_name%></td>
								<td class="right"><%=formatnumber(prev_pmg_give_total,0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(prev_cnt_emp,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cur_pmg_give_total,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(cur_cnt_emp,0)%>&nbsp;</td>
                                <td class="right"><%=give_mark & " " & formatnumber(Abs(give_total),0)%>&nbsp;</td>
                                <td class="right"><%=emp_mark & " " & formatnumber(Abs(cnt_emp),0)%>&nbsp;</td>
							</tr>
                        <%

									Next
								End If
							End If
						%>
						<!-- ���հ� start -->
							<tr>
								<th class="first"></th>
                                <th>���հ�</th>
								<th class="right"><%=formatnumber(sum_prev_pmg_give_total,0)%>&nbsp;</th>
								<th class="right"><%=formatnumber(sum_prev_cnt_emp,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_cur_pmg_give_total,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_cur_cnt_emp,0)%>&nbsp;</th>
                                <th class="right"><%=sum_give_mark & " " & formatnumber(Abs(sum_give_total),0)%>&nbsp;</th>
                                <th class="right"><%=sum_emp_mark & " " & formatnumber(Abs(sum_cnt_emp),0)%>&nbsp;</th>
							</tr>
						<!-- ���հ� end -->
						</tbody>
					</table>
                    </DIV>
					</td>
                  </tr>
				</table>
			</form>
		</div>				
	</div>        				
	</body>
</html>


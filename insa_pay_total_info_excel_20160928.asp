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
	srchCompany	= replaceXSS(request.form("srchCompany"))
	emp_month		= replaceXSS(Request.form("emp_month"))
else
	srchCompany	= replaceXSS(request("srchCompany"))
	emp_month		= replaceXSS(request("emp_month"))
end if

'if srchCompany = "" then
'	srchCompany	= "��ü"
'end if

curr_dd = cstr(datepart("d",now))
If emp_month = "" Then
	emp_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
End If

'//�� ��� ���ϱ�
emp_month_prev = cstr(CInt(Left(emp_month,4)-1)) + cstr(mid(emp_month,5,2))

'//title
curr_yyyy = mid(cstr(emp_month),1,4)
curr_mm = mid(cstr(emp_month),5,2)
title_line = srchCompany & " " & cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " ����κ��ΰǺ���ȸ"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


'//�˻��� ����� �ΰǺ�,�ο��� ��ȸ
sql = ""
sql = sql & "select "
sql = sql & "		emp_saupbu "
sql = sql & "		,emp_org_name "
sql = sql & "		,sum(IF( emp_month = '" & emp_month_prev & "', sum_pmg_give_total, '0')) as prev_pmg_give_total "
sql = sql & "		,sum(IF( emp_month = '" & emp_month_prev & "', cnt_emp, '0')) as prev_cnt_emp "
sql = sql & "		,sum(IF( emp_month = '" & emp_month & "', sum_pmg_give_total, '0')) as cur_pmg_give_total "
sql = sql & "		,sum(IF( emp_month = '" & emp_month & "', cnt_emp, '0')) as cur_cnt_emp "
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
sql = sql & "		emm.emp_month='" & emp_month & "' "
sql = sql & "		or emm.emp_month='" & emp_month_prev & "' "
sql = sql & "		) "
sql = sql & "		group by emm.emp_month, emm.emp_saupbu, emm.emp_org_name "
sql = sql & ") v "
sql = sql & "group by emp_saupbu, emp_org_name "
sql = sql & "order by emp_saupbu asc, emp_org_name asc, emp_month asc "

'Response.write sql

Rs.Open Sql, Dbconn, 1

Set List	= getRsToDic(Rs)

Rs.close() : Set rs = Nothing


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
	</head>
	<body>
					<table cellpadding="0" cellspacing="0" border="1" >
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
								<th scope="col" rowspan="2" >�����</th>
								<th scope="col" rowspan="2" >�Ҽ�</th>
								<th scope="col" colspan="2" ><%=mid(emp_month_prev,1,4)%>��&nbsp;<%=mid(emp_month_prev,5,2)%>��
								<!--
								&nbsp;<%=mid(pmg_yymm_start,5,2)%>��
									~ <%=mid(pmg_yymm_end,1,4)%>��&nbsp;<%=mid(pmg_yymm_end,5,2)%>��
								-->
								</th>
								<th scope="col"  colspan="2" ><%=mid(emp_month,1,4)%>��&nbsp;<%=mid(emp_month,5,2)%>��
								<!--
									~ <%=mid(pmg_yymm_end,1,4)%>��&nbsp;<%=mid(pmg_yymm_end,5,2)%>��
								-->
                                <th scope="col" colspan="2" >���</th>
							</tr>
							<tr>
								<th scope="col" >�ΰǺ�</th>
								<th scope="col" >�ο���</th>
								<th scope="col" >�ΰǺ�</th>
								<th scope="col" >�ο���</th>
								<th scope="col" >�ΰǺ� �߰�</th>
								<th scope="col" >�ο��� �߰�</th>
							</tr>
                        </thead>
                    </table>
					<table cellpadding="0" cellspacing="0" border="1"  >
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
								<td style=" border-bottom:1px solid <%=rowBgColor%>;"><%= out_emp_saupbu%></td>
                                <td style=" border-bottom:1px solid <%=rowBgColor%>;"><%=emp_org_name%></td>
								<td style=" border-bottom:1px solid <%=rowBgColor%>;"><%=formatnumber(prev_pmg_give_total,0)%>&nbsp;</td>
								<td style=" border-bottom:1px solid <%=rowBgColor%>;"><%=formatnumber(prev_cnt_emp,0)%>&nbsp;</td>
                                <td style=" border-bottom:1px solid <%=rowBgColor%>;"><%=formatnumber(cur_pmg_give_total,0)%>&nbsp;</td>
                                <td style=" border-bottom:1px solid <%=rowBgColor%>;"><%=formatnumber(cur_cnt_emp,0)%>&nbsp;</td>
                                <td style=" border-bottom:1px solid <%=rowBgColor%>;"><%=give_mark & " " & formatnumber(Abs(give_total),0)%>&nbsp;</td>
                                <td style=" border-bottom:1px solid <%=rowBgColor%>;"><%=emp_mark & " " & formatnumber(Abs(cnt_emp),0)%>&nbsp;</td>
							</tr>
                        <%

									Next
								End If
							End If
						%>
						<!-- ���հ� start -->
							<tr>
								<th ></th>
                                <th>���հ�</th>
								<th ><%=formatnumber(sum_prev_pmg_give_total,0)%>&nbsp;</th>
								<th ><%=formatnumber(sum_prev_cnt_emp,0)%>&nbsp;</th>
                                <th ><%=formatnumber(sum_cur_pmg_give_total,0)%>&nbsp;</th>
                                <th ><%=formatnumber(sum_cur_cnt_emp,0)%>&nbsp;</th>
                                <th ><%=sum_give_mark & " " & formatnumber(Abs(sum_give_total),0)%>&nbsp;</th>
                                <th ><%=sum_emp_mark & " " & formatnumber(Abs(sum_cnt_emp),0)%>&nbsp;</th>
							</tr>
						<!-- ���հ� end -->
						</tbody>
					</table>

	</body>
</html>


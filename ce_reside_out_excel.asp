<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim end_cnt(200,10,2)
dim ce_tab(200,3)

from_date=Request.form("from_date")
to_date=Request.form("to_date")

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
End If

savefilename = "������ �ܰ� ó�� ��Ȳ" + from_date + "_" + to_date + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

for i = 0 to 200
	for j = 0 to 10
		end_cnt(i,j,1) = 0
		end_cnt(i,j,2) = 0
	next
next

sql = "select ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name from ce_work inner join memb on ce_work.mg_ce_id=memb.user_id where (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') and (memb.reside = '1') GROUP BY ce_work.mg_ce_id,memb.team,memb.org_name,memb.reside,memb.reside_place,memb.user_name Order By memb.team, memb.user_name Asc"
Rs.Open Sql, Dbconn, 1

i = 0
do until rs.eof
	i = i + 1
	if rs("team") = "" or isnull(rs("team")) then
		org_view = rs("org_name") 
	  else
	  	org_view = rs("team")
	end if
	ce_tab(i,1) = org_view
	ce_tab(i,2) = rs("user_name")
	ce_tab(i,3) = rs("reside_place")
	
    sql = "select ce_work.company,ce_work.reside_company,ce_work.mg_ce_id,ce_work.as_type,ce_work.holiday_yn,count(*) as end_cnt from ce_work WHERE (ce_work.company <> ce_work.reside_company) and (ce_work.reside_company<>'') and (ce_work.as_type<>'����ó��') and (ce_work.work_id='2') and (ce_work.mg_ce_id='"+rs("mg_ce_id")+"') and (ce_work.work_date >= '" + from_date + "' AND ce_work.work_date <= '"+to_date+"') GROUP BY ce_work.company,ce_work.reside_company,ce_work.as_type,ce_work.holiday_yn"		
    rs_as.Open Sql, Dbconn, 1
	do until rs_as.eof
		sql_trade = "select support_company from trade where trade_id ='����' and trade_name = '"&rs_as("company")&"'"
		Set rs_trade = Dbconn.Execute (sql_trade)
		if rs_trade.eof or rs_trade.bof then
			company1 = rs_as("company")
		  else
			if rs_trade("support_company") = "����" then
				company1 = rs_as("company")
			  else												
				company1 = rs_trade("support_company")
			end if
		end if
		rs_trade.close()
		
		sql_trade = "select support_company from trade where trade_id ='����' and trade_name = '"&rs_as("reside_company")&"'"
		Set rs_trade = Dbconn.Execute (sql_trade)
		if rs_trade.eof or rs_trade.bof then
			company2 = rs_as("reside_company")
		  else
			if rs_trade("support_company") = "����" then
				company2 = rs_as("reside_company")
			  else												
				company2 = rs_trade("support_company")
			end if
		end if
		rs_trade.close()									
		
        select case rs_as("as_type")
        	case "�湮ó��"
            	j = 1
        	case "�űԼ�ġ"
            	j = 2
        	case "�űԼ�ġ����"
            	j = 3
        	case "������ġ"
            	j = 4
        	case "������ġ����"
            	j = 5
        	case "������"
            	j = 6
        	case "����������"
            	j = 7
        	case "���ȸ��"
            	j = 8
        	case "��������"
            	j = 9
        	case "��Ÿ"
            	j = 10
        end select												

		if company1 <> company2 then
			end_cnt(i,j,1) = end_cnt(i,j,1) + cint(rs_as("end_cnt"))
			end_cnt(i,0,1) = end_cnt(i,0,1) + cint(rs_as("end_cnt"))
			end_cnt(0,j,1) = end_cnt(0,j,1) + cint(rs_as("end_cnt"))
			end_cnt(0,0,1) = end_cnt(0,0,1) + cint(rs_as("end_cnt"))
		end if
		if rs_as("holiday_yn") = "Y" then
			if company1 <> company2 then
				end_cnt(i,j,2) = end_cnt(i,j,2) + cint(rs_as("end_cnt"))
				end_cnt(i,0,2) = end_cnt(i,0,2) + cint(rs_as("end_cnt"))
				end_cnt(0,j,2) = end_cnt(0,j,2) + cint(rs_as("end_cnt"))
				end_cnt(0,0,2) = end_cnt(0,0,2) + cint(rs_as("end_cnt"))
			end if
		end if
		rs_as.movenext()
	loop
	rs_as.close()

	rs.movenext()
loop
title_line = "������ �ܰ� ó�� ��Ȳ"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title></title>
		<style type="text/css">
        <!--
        .style12 {font-size: 12px; font-family: "����ü", "����ü", Seoul; }
        .style12B {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; }
        .style12BW {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; color: #FFFFFF; }
        .style14 {font-size: 14px; font-family: "����ü", "����ü", Seoul; }
        .style14B {font-size: 14px; font-weight: bold; font-family: "����ü", "����ü", Seoul; }
        .style14BW {font-size: 14px; font-weight: bold; font-family: "����ü", "����ü", Seoul; color: #FFFFFF; }
        .style11 {font-size: 11px; font-family: "����ü", "����ü", Seoul; }
        .style11B {font-size: 11px; font-weight: bold; font-family: "����ü", "����ü", Seoul; }
        .style_red {color: #FF0000; font-weight: bold}
        .style_green {color: #006600; font-weight: bold}
        .style_blue {color: #000099; font-weight: bold}
        -->
        </style>
	</head>
	<body>
<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" bgcolor="#6699CC">&nbsp;<span class="style14BW"><%=title_line%></span></td>
  </tr>
  <tr>
    <td><table border="0" cellspacing="3">
      <tr>
        <td>
          <table width="100%"  border="0">
            <tr>
              <td>&nbsp;<%=from_date%>&nbsp;~&nbsp;<%=to_date%></td>
            </tr>
            <tr>
              <td><table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
					<td scope="col" rowspan="3"><div align="center">�Ҽ�</div></td>
					<td scope="col" rowspan="3"><div align="center">CE��</div></td>
					<td scope="col" rowspan="3"><div align="center">����ó</div></td>
					<td colspan="22" scope="col"><div align="center">������ ó�� ��Ȳ ( ��ü����/���ϱٹ����� )</div></td>
				</tr>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
					<td colspan="2" scope="col"><div align="center">�Ұ�</div></td>
					<td colspan="2" scope="col"><div align="center">�湮</div></td>
					<td colspan="2" scope="col"><div align="center">�űԼ�ġ</div></td>
					<td colspan="2" scope="col"><div align="center">�űԼ�ġ����</div></td>
					<td colspan="2" scope="col"><div align="center">������ġ</div></td>
					<td colspan="2" scope="col"><div align="center">������ġ����</div></td>
					<td colspan="2" scope="col"><div align="center">������</div></td>
					<td colspan="2" scope="col"><div align="center">����������</div></td>
					<td colspan="2" scope="col"><div align="center">ȸ��</div></td>
					<td colspan="2" scope="col"><div align="center">����</div></td>
					<td colspan="2" scope="col"><div align="center">��Ÿ</div></td>
				</tr>
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                	<td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                    <td>�Ѽ���</td>
                    <td>��Ư��</td>
                </tr>
			<% 
			ce_cnt = 0
			for  i = 1 to 200
				if end_cnt(i,0,1) > 0 then
					ce_cnt = ce_cnt + 1
           	%>
                <tr class="style12">
                  <td><div align="center"><%=ce_tab(i,1)%></div></td>
                  <td bgcolor="#FFFFCC"><div align="center"><%=ce_tab(i,2)%></div></td>
                  <td><div align="center"><%=ce_tab(i,3)%></div></td>
                  <td bgcolor="#CCCCCC"><%=formatnumber(end_cnt(i,0,1),0)%></td>
                  <td bgcolor="#CCCCCC"><%=end_cnt(i,0,2)%></td>
			<%
            		for j = 1 to 10                        
            %>
                  <td><%=formatnumber(end_cnt(i,j,1),0)%></td>
                  <td><%=end_cnt(i,j,2)%></td>
			<%
            		next                     
			%>
            	</tr>
			<%
            	end if
			next
			%>
                <tr valign="middle" bgcolor="#CCCCCC" class="style12B">
                  <td>�Ѱ�</td>
                  <td><%=ce_cnt%></td>
                  <td>&nbsp;</td>
                  <td><%=formatnumber(end_cnt(0,0,1),0)%></td>
                  <td><%=end_cnt(0,0,2)%></td>
			<%
            for j = 1 to 10                        
            %>
                  <td><%=formatnumber(end_cnt(0,j,1),0)%></td>
                  <td><%=end_cnt(0,j,2)%></td>
			<%
            next                     
            %>
                </tr>
              </table></td>
            </tr>
          </table>
        </td>
      </tr>
    </table></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>

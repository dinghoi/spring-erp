<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'   on Error resume next

Server.ScriptTimeOut = 1200

dim saupbu_tab(10,3) ' 1.����θ�, 2.������� �ο���, 3.������� ����

end_month=Request("end_month")
end_yn=Request("end_yn")

from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

mm = mid(end_month,5,2)
cost_year = mid(end_month,1,4)
cost_month = mid(end_month,5)

' ���� 5%, �湮 95%
won_per = 5
bang_per = 95

for i = 1 to 10
    saupbu_tab(i,1) = "" ' 1.����θ�
    saupbu_tab(i,2) = 0  ' 2.������� �ο���
    saupbu_tab(i,3) = 0  ' 3.������� ����
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from cost_end where end_month = '"&end_month&"' and (end_yn = 'Y') and (saupbu = '���ֺ��')"
Set rs_check = Dbconn.Execute (sql)	
if rs_check.eof or rs_check.bof then
	check_sw = "N"
  else
  	check_sw = "Y"
end if

if check_sw = "N" then
	response.write"<script language=javascript>"
	response.write"alert('���ֺ�� ������ �����ϼž� �մϴ� !!');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End
  else		
	response.write"<script language=javascript>"
	response.write"alert('����ó����!!!');"
	response.write"</script>"

dbconn.BeginTrans

sql = "SELECT sum(cost_amt_"&mm&") as tot_cost  " & chr(13) &_    
      "  FROM company_cost                      " & chr(13) &_    
      " WHERE cost_year   = '"&cost_year&"'     " & chr(13) &_    
      "   AND cost_center = '�ι������'        "
'Response.write "<pre>"& sql &"</pre><br>"          
Set rs=DbConn.Execute(SQL)
if not (rs.eof or rs.bof) then
    tot_cost = clng(rs("tot_cost")) ' (12) �� �ι������
end if
rs.close()

' 4�� ����� �Ѹ����
sales_date = left(end_month,4) & "-" &right(end_month,2)
sql = "  SELECT sum(cost_amt) as sum_cost_amt                " & chr(13) &_
      "    FROM saupbu_sales                                 " & chr(13) &_
      "   WHERE (    sales_date >= '" + from_date + "'       " & chr(13) &_    
      "          AND sales_date <= '" + to_date + "'  )      " & chr(13) &_    
      "     AND saupbu IN ( 'SI1�����'                      " & chr(13) &_
      "                   , 'SI2�����'                      " & chr(13) &_
      "                   , 'N/W �����'                     " & chr(13) &_
      "                   , '���������' )                   "
'Response.write "<pre>"& sql &"</pre><br>"          
Rs.Open Sql, Dbconn, 1
if not (rs.eof or rs.bof) then
    tot_sum_cost_amt = CCur(rs("sum_cost_amt")) ' (10)��ü ����� (4���� �հ�)
end if  
rs.close()

' ��ü AS�Ǽ� (To Be)
sql = "    SELECT count(*) as tot_cnt                               " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_process <> '���'                         " & chr(13) &_    
      "            and as_type    <> '��Ư��' )                     " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='�ι������'                        "
'Response.write "<pre>"& sql &"</pre><br>"          
Set rs=DbConn.Execute(SQL)
tot_as_cnt = clng(rs("tot_cnt")) ' (11)��ü AS�Ǽ� (To Be)
if tot_as_cnt = "" or isnull(tot_as_cnt) then
    tot_as_cnt = 0
end if
rs.close()

over_less =  tot_cost - (tot_as_cnt * 40000) ' (2)�������� ((12)�� �ι������ - ((11)�� AS �Ǽ� * 40000) )


' ���纰 AS �Ǽ��� ���Ѵ�
sql = "    SELECT company, count(*) as as_cnt                       " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_process <> '���'                         " & chr(13) &_    
      "            AND as_type    <> '��Ư��'  )                    " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='�ι������'                        " & chr(13) &_    
      "  GROUP BY company                                           " & chr(13) &_    
      "  ORDER BY company Asc                                       "
'Response.write "<pre>"& sql &"</pre><br>"          
Rs.Open Sql, Dbconn, 1
do until rs.eof 

    sql = "select saupbu from trade where trade_name = '"&rs("company")&"'" 
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        saupbu = "Error"
    else
        saupbu = rs_trade("saupbu")
    end if
    rs_trade.close()

    sql = "  SELECT company, saupbu                              " & chr(13) &_
          "       , ifnull(sum(cost_amt),0) as sum_cost_amt      " & chr(13) &_
          "    FROM saupbu_sales                                 " & chr(13) &_
          "   WHERE (    sales_date >= '" + from_date + "'       " & chr(13) &_    
          "          AND sales_date <= '" + to_date + "'  )      " & chr(13) &_    
          "     AND saupbu  = '"&saupbu&"'                       " & chr(13) &_
          "     AND company = '"&rs("company")&"'                " 
'Response.write "<pre>"& sql &"</pre><br>"
    rs_etc.Open Sql, Dbconn, 1
    if not (rs_etc.eof or rs_etc.bof) then
        sum_cost_amt = Clng(rs_etc("sum_cost_amt"))  ' ȸ�纰,����κ� �����
        if (sum_cost_amt=0) then
            sum_cost_per = 0
        else
            sum_cost_per = sum_cost_amt / tot_sum_cost_amt * 100.0  ' ȸ�纰,����κ� ����� / 4�� ����� �Ѹ���� (%)
        end if
    end if
    rs_etc.close

    divide_amt_1 = CInt(rs("as_cnt")) * 40000                    ' (3)1�� ��αݾ� = (ȸ�纰����κ�)AS�Ǽ�*ǥ�شܰ�
    divide_amt_2 = (sum_cost_amt / tot_sum_cost_amt) * over_less ' (4)2�� ��αݾ� = (ȸ�纰����κ�)�����/4��������Ѹ����)*��������
'Response.write "<pre>(sum_cost_amt / tot_sum_cost_amt) * over_less "& "("&sum_cost_amt&" / "&tot_sum_cost_amt&") *"& over_less &"</pre><br>" 
    cost_amt = divide_amt_1 + divide_amt_2 ' (5)�ι������ = 1�� ��αݾ� + 2�� ��αݾ�

    sql = "SELECT ifnull(sum(cost_amt_07),0) as saupbu_cost " & chr(13) &_    
          "  FROM company_cost                              " & chr(13) &_    
          " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_    
          "   AND cost_center = '�ι������'                " & chr(13) &_    
          "   AND saupbu      = '"&saupbu&"'                " & chr(13) &_    
          "   AND company     = '"&rs("company")&"'         " 
'Response.write "<pre>"& sql &"</pre><br>"          
    set rs_etc=dbconn.execute(sql)
    if not (rs_etc.eof or rs_etc.bof) then
        saupbu_cost = clng(rs_etc("saupbu_cost")) ' ȸ�纰����κ� �ι������
        if (saupbu_cost=0) then
            charge_per = 0
        else
            charge_per = saupbu_cost / tot_cost * 100.0  ' (6)������ = ȸ�纰����κ� �ι������ / �� �ι������ (%)
        end if
    end if
    rs_etc.close()

  
  charge_per = cost_amt / tot_cost * 100.0  ' (6)������ = ȸ�纰����κ� �ι������ / �� �ι������ (%)
'Response.write "<pre>(6)������ = cost_amt / tot_cost * 100.0 = "& cost_amt &"/"& tot_cost &"*"& 100.0  &"</pre><br>"          

    
    sql = "SELECT * FROM company_asunit WHERE as_month = '"&end_month&"' AND as_company = '"&rs("company")&"'"
'Response.write "<pre>"& sql &"</pre><br>"              
    set rs_etc=dbconn.execute(sql)
    if rs_etc.eof or rs_etc.bof then
        sql="INSERT INTO company_asunit (as_month, as_company, saupbu, as_cnt, divide_amt_1, divide_amt_2, charge_per, cost_amt, reg_id, reg_name, reg_date) "&_
            "     VALUES ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&divide_amt_1&"',"&divide_amt_2&","&charge_per&","&cost_amt&",'"&user_id&"','"&user_name&"',now())"
        dbconn.execute(sql)
    else
        sql = "UPDATE company_asunit                     " & chr(13) &_    
              "   SET saupbu       = '"&saupbu&"'        " & chr(13) &_    
              "     , as_cnt       = '"&rs("as_cnt")&"'  " & chr(13) &_    
              "     , divide_amt_1 = '"&divide_amt_1&"'  " & chr(13) &_    
              "     , divide_amt_2 = '"&divide_amt_2&"'  " & chr(13) &_    
              "     , charge_per   = '"&charge_per&"'    " & chr(13) &_    
              "     , cost_amt     = "&cost_amt&"        " & chr(13) &_    
              " WHERE as_company = '" &rs("company")& "' " & chr(13) &_    
              "   AND as_month   = '" &end_month& "'     "
        dbconn.execute(sql)   
    end if
'Response.write "<pre>"& sql &"</pre><br>"            
    rs_etc.close()
    
    rs.movenext()
loop    
rs.close()

' ���纰 ���� �ڷ� ����
' �ι������ ���
' ó���� zero
sql = "UPDATE company_profit_loss SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"' AND (cost_center ='�ι������') "
dbconn.execute(sql)

sql = "  SELECT as_company as company          " & chr(13) &_    
      "       , sum(charge_per) as charge_per  " & chr(13) &_    
      "    FROM company_asunit                 " & chr(13) &_    
      "   WHERE (as_month = '"&end_month&"')   " & chr(13) &_    
      "GROUP BY as_company                     "
Rs.Open Sql, Dbconn, 1
do until rs.eof
    charge_per = rs("charge_per")

    sql = "SELECT * FROM trade WHERE trade_name = '"&rs("company")&"'"
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        group_name = "Error"
    else
        group_name = rs_trade("group_name")
    end if        

    sql = "  SELECT cost_id, cost_detail, sum(cost_amt_"&cost_month&") as cost " & chr(13) &_    
          "    FROM company_cost                                               " & chr(13) &_    
          "   WHERE (cost_center = '�ι������' )                              " & chr(13) &_    
          "     AND cost_year ='"&cost_year&"'                                 " & chr(13) &_    
          "GROUP BY cost_id, cost_detail                                       "
    rs_etc.Open sql, Dbconn, 1
    do until rs_etc.eof
        
        cost = int(charge_per * clng(rs_etc("cost")))

        sql = "SELECT * FROM company_profit_loss                " & chr(13) &_    
              " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_    
              "   AND company     = '"&rs("company")&"'         " & chr(13) &_    
              "   AND group_name  = '"&group_name&"'            " & chr(13) &_    
              "   AND cost_center = '�ι������'                " & chr(13) &_    
              "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_    
              "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
        set rs_cost=dbconn.execute(sql)
        
        if rs_cost.eof or rs_cost.bof then
            sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&")                       " & chr(13) &_    
                  "     values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','�ι������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&") "
            dbconn.execute(sql)
          else
            sql = "UPDATE company_profit_loss                       " & chr(13) &_    
                  "   SET cost_amt_"&cost_month&"="&cost&"          " & chr(13) &_    
                  " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_    
                  "   AND company     = '"&rs("company")&"'         " & chr(13) &_    
                  "   AND group_name  = '"&group_name&"'            " & chr(13) &_    
                  "   AND cost_center = '�ι������'                " & chr(13) &_    
                  "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_    
                  "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
            dbconn.execute(sql)
        end if      
        
        rs_etc.movenext()
    loop
    rs_etc.close()
    rs.movenext()
loop
rs.close()
' �ι������ ��� ��
























' ����
sql = " SELECT count(*) as tot_cnt FROM as_acpt                             " & chr(13) &_    
      "  WHERE acpt_man in ('���μ�','�ֿ���','�Ѽ���','����ȯ')            " & chr(13) &_    
      "    AND (Cast(visit_date as date) >= '" + from_date + "'             " & chr(13) &_    
      "    AND  Cast(visit_date as date) <= '" + to_date + "'  )            " & chr(13) &_    
      "    AND company not in ('�ڿ���','������ũ��','������ǰ','�Ե���Ż') "  
Set rs=DbConn.Execute(SQL)
won_cnt = clng(rs("tot_cnt")) ' �������� �� �Ǽ�
if won_cnt = "" or isnull(tot_cnt) then
    won_cnt = 0
end if
rs.close()

sql = "   SELECT company, count(*) as as_cnt FROM as_acpt                    " & chr(13) &_    
      "    WHERE acpt_man in ('���μ�','�ֿ���','�Ѽ���','����ȯ')           " & chr(13) &_    
      "      AND (Cast(visit_date as date) >= '" + from_date + "'            " & chr(13) &_    
      "      AND  Cast(visit_date as date) <= '" + to_date + "'  )           " & chr(13) &_    
      "      AND company not in('�ڿ���','������ũ��','������ǰ','�Ե���Ż') " & chr(13) &_    
      " GROUP BY company                                                     " & chr(13) &_    
      " ORDER BY company Asc                                                 "
Rs.Open Sql, Dbconn, 1

do until rs.eof 

    sql = "SELECT saupbu FROM trade WHERE trade_name = '"&rs("company")&"'" 
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        saupbu = "Error"
      else
        saupbu = rs_trade("saupbu")
    end if
    rs_trade.close()
    
    charge_per = clng(rs("as_cnt")) / won_cnt * won_per / 100 ' (����κ� ����AS�Ǽ� / �� ����AS�Ǽ�) * 5%
    cost_amt = int(charge_per * tot_cost)     ' ������ * �� �ι������

    sql = "select * from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
    set rs_etc=dbconn.execute(sql)
    if not (rs_etc.eof or rs_etc.bof) then
        sql = "delete from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
        dbconn.execute(sql)   
    end if

    sql="INSERT INTO company_as (as_month,as_company,saupbu,remote_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date)  " & chr(13) &_    
        "     VALUES ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now()) "
    dbconn.execute(sql)
    
    rs.movenext()
loop    
rs.close()

' ���ݿ� (�湮)

'  �� �湮AS �Ǽ��� ���Ѵ�
sql = "    SELECT count(*) as tot_cnt                               " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_type    <> '����ó��'                     " & chr(13) &_    
      "            and as_process <> '���'                         " & chr(13) &_    
      "            and as_type    <> '��Ư��' )                     " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='�ι������'                        "
'sql = sql & " and ( team not  like '%����%') and (company <> '��Ÿ' and company <> '���̿��������') "
'sql = sql & " and  mg_ce not in ('�Ķ�','�ھ���','��ÿ�','������','������','������') "

Set rs=DbConn.Execute(SQL)
bang_cnt = clng(rs("tot_cnt")) ' �� �湮AS �Ǽ�
if bang_cnt = "" or isnull(tot_cnt) then
	bang_cnt = 0
end if
rs.close()

' ���纰 �湮AS �Ǽ��� ���Ѵ�
sql = "    SELECT company, count(*) as as_cnt                       " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_type    <> '����ó��'                     " & chr(13) &_    
      "            AND as_process <> '���'                         " & chr(13) &_    
      "            AND as_type    <> '��Ư��'  )                    " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='�ι������'                        " & chr(13) &_    
      "  GROUP BY company                                           " & chr(13) &_    
      "  ORDER BY company Asc                                        "
'sql = sql & " and ( team not  like '%����%') and (company <> '��Ÿ' and company <> '���̿��������') "
'sql = sql & " and  mg_ce not in ('�Ķ�','�ھ���','��ÿ�','������','������','������') 
Rs.Open Sql, Dbconn, 1

do until rs.eof 

    sql = "select saupbu from trade where trade_name = '"&rs("company")&"'" 
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        saupbu = "Error"
      else
        saupbu = rs_trade("saupbu")
    end if
    rs_trade.close()
    
    sql = "select * from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
    set rs_etc=dbconn.execute(sql)
    if rs_etc.eof or rs_etc.bof then
        charge_per = clng(rs("as_cnt")) / bang_cnt * bang_per / 100  ' ���纰 �湮AS �Ǽ� /  �� �湮AS �Ǽ� * 95%
        cost_amt = int(charge_per * tot_cost) ' ������ * �� �ι������
        sql="INSERT INTO company_as (as_month,as_company,saupbu,visit_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date) "&_
            "     VALUES ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now())"
        dbconn.execute(sql)
    else
        charge_per = clng(rs("as_cnt")) / bang_cnt * bang_per / 100 + rs_etc("charge_per") ' ���纰 �湮AS �Ǽ� /  �� �湮AS �Ǽ� * 95% + ������
        cost_amt = int(charge_per * tot_cost) ' ������ * �� �ι������
        sql = "UPDATE company_as SET visit_cnt='"&rs("as_cnt")&"', charge_per='"&charge_per&"', cost_amt="&cost_amt&_
              " WHERE as_company='" &rs("company")& "' and as_month = '" &end_month& "'"
        dbconn.execute(sql)   
    end if
    
    rs.movenext()
loop    
rs.close()
rs_etc.close()

' ����κ� ���� �ڷ� ����
' �ι������ ���
' ó���� zero
sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='�ι������') "
dbconn.execute(sql)

sql = "   SELECT saupbu, sum(charge_per) as charge_per " & chr(13) &_
      "     FROM company_as                            " & chr(13) &_
      "    WHERE (as_month = '"&end_month&"')          " & chr(13) &_
      " GROUP BY saupbu                                "
Rs.Open Sql, Dbconn, 1
do until rs.eof
    charge_per = rs("charge_per")

    sql = "  SELECT cost_id, cost_detail, sum(cost_amt_"&cost_month&") as cost    " & chr(13) &_
          "    FROM company_cost                                                  " & chr(13) &_
          "   WHERE (cost_center = '�ι������' ) AND cost_year ='"&cost_year&"'  " & chr(13) &_
          "GROUP BY cost_id, cost_detail                                          "
    rs_etc.Open sql, Dbconn, 1
    do until rs_etc.eof
        
        cost = int(charge_per * clng(rs_etc("cost")))

        sql = "SELECT * FROM saupbu_profit_loss                 " & chr(13) &_
              " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_
              "   AND saupbu      = '"&rs("saupbu")&"'          " & chr(13) &_
              "   AND cost_center = '�ι������'                " & chr(13) &_
              "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_
              "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
        set rs_cost=dbconn.execute(sql)
        
        if rs_cost.eof or rs_cost.bof then
            sql = "INSERT INTO saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&")                  " & chr(13) &_
                  "     VALUES ('"&cost_year&"','"&rs("saupbu")&"','�ι������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&") "
            dbconn.execute(sql)
        else
            sql = "UPDATE saupbu_profit_loss                        " & chr(13) &_
                  "   SET cost_amt_"&cost_month&" = "&cost&"        " & chr(13) &_
                  " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_
                  "   AND saupbu      = '"&rs("saupbu")&"'          " & chr(13) &_
                  "   AND cost_center = '�ι������'                " & chr(13) &_
                  "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_
                  "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
            dbconn.execute(sql)
        end if      
        
        rs_etc.movenext()
    loop
    rs_etc.close()
    rs.movenext()
loop
rs.close()
' �ι������ ��� ��

' ���纰 ���� �ڷ� ����
' �ι������ ���
' ó���� zero
'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='�ι������') "
'dbconn.execute(sql)
'
'sql = " select as_company as company, sum(charge_per) as charge_per from company_as Where (as_month = '"&end_month&"') GROUP BY as_company"
'Rs.Open Sql, Dbconn, 1
'do until rs.eof
'   charge_per = rs("charge_per")
'
'   sql = "select * from trade where trade_name = '"&rs("company")&"'"
'   set rs_trade=dbconn.execute(sql)
'   if rs_trade.eof or rs_trade.bof then
'       group_name = "Error"
'     else
'       group_name = rs_trade("group_name")
'   end if        
'
'   sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '�ι������' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
'   rs_etc.Open sql, Dbconn, 1
'   do until rs_etc.eof
'       
'       cost = int(charge_per * clng(rs_etc("cost")))
'
'       sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='�ι������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
'       set rs_cost=dbconn.execute(sql)
'       
'       if rs_cost.eof or rs_cost.bof then
'           sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','�ι������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
'           dbconn.execute(sql)
'         else
'           sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='�ι������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
'           dbconn.execute(sql)
'       end if      
'       
'       rs_etc.movenext()
'   loop
'   rs_etc.close()
'   rs.movenext()
'loop
'rs.close()
' �ι������ ��� ��


' �߰� ����
' ����κ� �ο���,����� ����
sql = " select saupbu from sales_org where sales_year='" & cost_year & "' order by saupbu asc"
Rs.Open Sql, Dbconn, 1
i = 0
tot_person       = 0 ' ����� ��ü �ο�
tot_saupbu_sales = 0 ' ����� ��ü �����

do until rs.eof 

	'sql = "select count(*) from pay_month_give where pmg_id = '1' and pmg_yymm = '"&end_month&"' and mg_saupbu ='"&rs("saupbu")&"'"
	
	''
	'' KDC����� ��� ����ó��
	''
	
	' KDC����ο� �̸��� ���� ���� ���̿�������ſ� �Ҽӵ� ����� cost_except = '2' �� �����Ѵ�.
	sql = "SELECT emp_name, count(*)               "&chr(13)&_
        "  FROM                                  "&chr(13)&_
        "(                                       "&chr(13)&_
        "  SELECT B.*                            "&chr(13)&_
        "    FROM pay_month_give A               "&chr(13)&_
        "        ,emp_master_month B             "&chr(13)&_
        "   WHERE A.pmg_id = '1'                 "&chr(13)&_
        "     AND A.pmg_emp_no = B.emp_no        "&chr(13)&_
        "     AND B.cost_except in ('0','1','2') "&chr(13)&_
        "     AND A.pmg_yymm  = '"&end_month&"'  "&chr(13)&_
        "     AND B.emp_month = '"&end_month&"'  "&chr(13)&_
        "     AND A.mg_saupbu = 'KDC�����'      "&chr(13)&_
        ") A                                     "&chr(13)&_
        "GROUP BY emp_name                       "&chr(13)&_
        "  HAVING count(*) = 2                   "
  'Response.write "<pre>"& sql &"</pre><br>"
  set rs_emp = dbconn.execute(sql)
  do until rs_emp.eof 
    emp_name = rs_emp("emp_name")
    
    sql = "UPDATE emp_master_month              "&chr(13)&_    
          "   SET cost_except = '2'             "&chr(13)&_
          " WHERE emp_name    = '"&emp_name&"'  "&chr(13)&_
          "   AND emp_month   = '"&end_month&"' "&chr(13)&_
          "   AND emp_company = '���̿��������'"
    'Response.write "<pre>"& sql &"</pre><br>"
    dbconn.execute(sql)
     
    rs_emp.movenext()
  loop  
  rs_emp.close()
  
  ' KDC����ο� ���������� �ش��ϴ� ����� cost_except = '2' �� �����Ѵ�.
	sql = "UPDATE emp_master_month                               "&chr(13)&_
        "   SET cost_except = '2'                              "&chr(13)&_
        " WHERE emp_month   = '"&end_month&"'                  "&chr(13)&_
        "   AND cost_center = '����������'                     "&chr(13)&_
        "   AND emp_saupbu  = 'KDC�����'                      "&chr(13)&_
        "   AND emp_no IN ( SELECT pmg_emp_no                  "&chr(13)&_
        "                     FROM pay_month_give              "&chr(13)&_
        "                    WHERE pmg_id    = 1               "&chr(13)&_
        "                      AND pmg_yymm  = '"&end_month&"' "&chr(13)&_
        "                      AND mg_saupbu ='KDC�����'      "&chr(13)&_
        "                 )                                    "
  'Response.write "<pre>"& sql &"</pre><br>"        
  dbconn.execute(sql)

	''
	'' KDC����� ��� ����ó�� _ ��
	''

	
	'����� ��α��� ���� ó��(2016-01-15)
    ' ����κ� �ο��� (�������� / ������Ȳ / '����κ� �ο���Ȳ' �� ��ġ, B.cost_except=2�� ��������)
    sql = "SELECT count(*)                          " & chr(13) &_
          "  FROM pay_month_give  A                 " & chr(13) &_
          "     , emp_master_month B                " & chr(13) &_
          " WHERE A.pmg_id     = '1'                " & chr(13) &_
          "   AND A.pmg_yymm   = '"&end_month&"'    " & chr(13) &_
          "   AND A.mg_saupbu  = '"&rs("saupbu")&"' " & chr(13) &_
          "   AND A.pmg_emp_no =  B.emp_no          " & chr(13) &_
          "   /* AND B.cost_except in ('0','1') */  " & chr(13) &_
          "   AND B.cost_center <> '��������'       " & chr(13) &_
          "   AND B.emp_month  ='"&end_month&"'     "   
    'Response.write "<pre>"&sql&"</pre><br>"
    set rs_emp=dbconn.execute(sql)
    if rs_emp(0) = "" or isnull(rs_emp(0)) then
        saupbu_person = 0
    else
        saupbu_person = clng(rs_emp(0))
    end if
    rs_emp.close()

    ' ����κ� �����
    sql = "SELECT ifnull(sum(cost_amt),0) AS cost_amt          " & chr(13) &_
          "  FROM saupbu_sales                                 " & chr(13) &_
          "   WHERE (    sales_date >= '" + from_date + "'     " & chr(13) &_    
          "          AND sales_date <= '" + to_date + "'  )     " & chr(13) &_    
          "   AND saupbu ='"&rs("saupbu")&"'                   "
'Response.write "<pre>"&sql&"</pre><br>"
    set rs_cost=dbconn.execute(sql)
    if rs_cost(0) = "" or isnull(rs_cost(0)) then
        saupbu_sales = 0
    else
        saupbu_sales = CCur(rs_cost(0))
    end if
    rs_cost.close()


    i = i + 1
    saupbu_tab(i,1) = rs("saupbu")  ' ����� ��
    saupbu_tab(i,2) = saupbu_person ' �ش� ������� �ο���
    saupbu_tab(i,3) = saupbu_sales  ' �ش� ������� �����

    tot_person       = tot_person       + saupbu_person ' ����� ��ü �ο�
    tot_saupbu_sales = tot_saupbu_sales + saupbu_sales  ' ����� ��ü �����
    
    rs.movenext()
loop    
rs.close()

'�������� �Ѿ�
sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '��������'"
'Response.write sql&"<br>"
Set rs=DbConn.Execute(SQL)
tot_cost_amt = clng(rs("tot_cost")) ' (1)�������� = (2) + (3) 
'Response.write "<pre>(1)�������� = "&tot_cost_amt&"</pre><br>"       

tot_cost_amt_h1 = round(tot_cost_amt/2,0)        ' (2)�������� �Ѿ��� 50%
tot_cost_amt_h2 = tot_cost_amt - tot_cost_amt_h1 ' (3)�������� �Ѿ��� 50%
'Response.write "<pre>(2)�������� = "&tot_cost_amt_h1&"</pre><br>"       
'Response.write "<pre>(3)�������� = "&tot_cost_amt_h2&"</pre><br>"       

rs.close()

' ����κ� ���� �ڷ� ����
' ó���� zero
sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='��������') "
dbconn.execute(sql)
sql = "delete from management_cost where cost_month ='"&end_month&"'"
'Response.write sql&"<br>"
dbconn.execute(sql)


' �������� ���

for i = 1 to 10 ' �� ����η� ��ȯ 
    if saupbu_tab(i,1) = "" or isnull(saupbu_tab(i,1)) then
        exit for
    end if

    'saupbu_cost_amt = int(tot_cost_amt * saupbu_per) ' ��ü�ݾ��� �������� (As Is)

    if saupbu_tab(i,2) = 0 then
        saupbu_per       = 0 ' (8)����� �ο� / (4)�� ����� �ο�
        saupbu_cost_amt  = 0 ' (6)��������(�η�) = ((2)��ü�����1 50%) * ((8)����� �ο� / (4)�� ����� �ο�) (To Be)  =E47*C49/C45
        chage_rate_inwon = 0 ' (10) �η� ������ = (6)��������(�η�) / (1)�������� 
    else
        saupbu_per       = saupbu_tab(i,2) / tot_person        ' (8)����� �ο� / (4)�� ����� �ο�
        saupbu_cost_amt  = round(tot_cost_amt_h1 * saupbu_per) ' (6)��������(�η�) = ((2)��ü�����1 50%) * ((8)����� �ο� / (4)�� ����� �ο�) (To Be)  =E47*C49/C45
        chage_rate_inwon = saupbu_cost_amt / tot_cost_amt_h1   ' (10) �η� ������ = (6)��������(�η�) / (1)�������� 
    end if 

    if saupbu_tab(i,3) = 0 then
        saupbu_sale_per = 0 ' (9)����� ���� / (5)�� ����θ��� (To Be)
        saupbu_sale_amt = 0 ' (7)��������(����) = ((3)��ü�����2 50%) * ((9)����� ���� / (5)�� ����θ���) (To Be) =G47*C50/C46
        chage_rate_sale = 0 ' (11) ���� ������ = (7)��������(����) / (1)�������� 
    else
        saupbu_sale_per = saupbu_tab(i,3)  / tot_saupbu_sales      ' (9)����� ���� / (5)�� ����θ��� (To Be)
        saupbu_sale_amt = round(tot_cost_amt_h2 * saupbu_sale_per) ' (7)��������(����) = ((3)��ü�����2 50%) * ((9)����� ���� / (5)�� ����θ���) (To Be) =G47*C50/C46
        chage_rate_sale = saupbu_sale_amt / tot_cost_amt_h2        ' (11) ���� ������ = (7)��������(����) / (1)�������� 
    end if 

    ' ����� �� ���� �հ踦 ���Ѵ�.

    ' �� ������� ���纰 �����հ�
    sql = "  SELECT saupbu, company, sum(cost_amt) as cost  " & chr(13) &_
          "    FROM saupbu_sales                            " & chr(13) &_
          "   WHERE (    sales_date >= '" + from_date + "'  " & chr(13) &_    
          "          AND sales_date <= '" + to_date + "'  ) " & chr(13) &_    
          "     AND saupbu ='"&saupbu_tab(i,1)&"'           " & chr(13) &_
          "GROUP BY saupbu, company                         "
'Response.write "<pre>"&sql&"</pre><br>"
    rs_etc.Open sql, Dbconn, 1
    k = 0
    do until rs_etc.eof
        k = k + 1
        if saupbu_sales = 0 then
            charge_per = 0 ' ex, ��Ÿ�����
        else
            charge_per = rs_etc("cost") / saupbu_sales ' ������� �հ� / ����θ��� �հ�
        end if
        cost_amt = int(charge_per * saupbu_cost_amt) ' ��е� �������� (���纰)
        
        sql = "INSERT INTO management_cost ( cost_month              " & chr(13) &_
              "                            , saupbu                  " & chr(13) &_
              "                            , company                 " & chr(13) &_
              "                            , tot_person              " & chr(13) &_
              "                            , saupbu_person           " & chr(13) &_
              "                            , saupbu_per              " & chr(13) &_
              "                            , tot_cost_amt            " & chr(13) &_
              "                            , saupbu_cost_amt         " & chr(13) &_
              "                            , saupbu_sale             " & chr(13) &_
              "                            , tot_sale                " & chr(13) &_
              "                            , sale_per                " & chr(13) &_
              "                            , saupbu_sale_amt         " & chr(13) &_
              "                            , charge_per              " & chr(13) &_
              "                            , cost_amt                " & chr(13) &_
              "                            , reg_id                  " & chr(13) &_
              "                            , reg_name                " & chr(13) &_
              "                            , reg_date                " & chr(13) &_
              "                            )                         " & chr(13) &_
              "                     VALUES ( '"&end_month&"'         " & chr(13) &_
              "                            , '"&saupbu_tab(i,1)&"'   " & chr(13) &_
              "                            , '"&rs_etc("company")&"' " & chr(13) &_
              "                            , "&tot_person&"          " & chr(13) &_
              "                            , "&saupbu_tab(i,2)&"     " & chr(13) &_
              "                            , "&chage_rate_inwon&"    " & chr(13) &_
              "                            , "&tot_cost_amt&"        " & chr(13) &_
              "                            , "&saupbu_cost_amt&"     " & chr(13) &_
              "                            , "&saupbu_tab(i,3)&"     " & chr(13) &_
              "                            , "&tot_saupbu_sales&"    " & chr(13) &_
              "                            , "&chage_rate_sale&"      " & chr(13) &_
              "                            , "&saupbu_sale_amt&"     " & chr(13) &_
              "                            , "&charge_per&"          " & chr(13) &_
              "                            , "&cost_amt&"            " & chr(13) &_
              "                            , '"&user_Id&"'           " & chr(13) &_
              "                            , '"&user_name&"'         " & chr(13) &_
              "                            , now()                   " & chr(13) &_
              "                            )                         "
'Response.write "<pre>"&sql&"</pre><br>"
        dbconn.execute(sql)

        rs_etc.movenext()
    loop
' ������ ������ ���
    if k = 0 then
        sql = "INSERT INTO management_cost ( cost_month              " & chr(13) &_
              "                            , saupbu                  " & chr(13) &_
              "                            , company                 " & chr(13) &_
              "                            , tot_person              " & chr(13) &_
              "                            , saupbu_person           " & chr(13) &_
              "                            , saupbu_per              " & chr(13) &_
              "                            , tot_cost_amt            " & chr(13) &_
              "                            , saupbu_cost_amt         " & chr(13) &_
              "                            , saupbu_sale             " & chr(13) &_
              "                            , tot_sale                " & chr(13) &_
              "                            , sale_per                " & chr(13) &_
              "                            , saupbu_sale_amt         " & chr(13) &_
              "                            , charge_per              " & chr(13) &_
              "                            , cost_amt                " & chr(13) &_
              "                            , reg_id                  " & chr(13) &_
              "                            , reg_name                " & chr(13) &_
              "                            , reg_date                " & chr(13) &_
              "                            )                         " & chr(13) &_
              "                     VALUES ( '"&end_month&"'         " & chr(13) &_
              "                            , '"&saupbu_tab(i,1)&"'   " & chr(13) &_
              "                            , ''                      " & chr(13) &_
              "                            , "&tot_person&"          " & chr(13) &_
              "                            , "&saupbu_tab(i,2)&"     " & chr(13) &_
              "                            , "&chage_rate_inwon&"    " & chr(13) &_
              "                            , "&tot_cost_amt&"        " & chr(13) &_
              "                            , "&saupbu_cost_amt&"     " & chr(13) &_
              "                            , "&saupbu_tab(i,3)&"     " & chr(13) &_
              "                            , "&tot_saupbu_sales&"    " & chr(13) &_
              "                            , "&chage_rate_sale&"     " & chr(13) &_
              "                            , "&saupbu_sale_amt&"     " & chr(13) &_
              "                            , "&charge_per&"          " & chr(13) &_
              "                            , "&cost_amt&"            " & chr(13) &_
              "                            , '"&user_Id&"'           " & chr(13) &_
              "                            , '"&user_name&"'         " & chr(13) &_
              "                            , now()                   " & chr(13) &_
              "                            )                         " 
'Response.write "<pre>"&sql&"</pre><br>"
        dbconn.execute(sql)
    end if
    rs_etc.close()

    sql = "   SELECT cost_id, cost_detail                  " & chr(13) &_
          "        , sum(cost_amt_"&cost_month&") as cost  " & chr(13) &_
          "     FROM company_cost                          " & chr(13) &_
          "    WHERE (cost_center = '��������' )         " & chr(13) &_
          "      AND cost_year ='"&cost_year&"'            " & chr(13) &_
          " GROUP BY cost_id, cost_detail                  "
'Response.write "<pre>"&sql&"</pre><br>"
    rs_etc.Open sql, Dbconn, 1
    do until rs_etc.eof

        'cost = int(saupbu_per * clng(rs_etc("cost"))) ' (As is)
        if clng(rs_etc("cost")) = 0 then
            cost_h1 = 0 ' �ο���� 50%
            cost_h2 = 0 ' ����׹�� 50%
        else
            cost_h1 = round(clng(rs_etc("cost")) / 2) ' �ο���� 50%
            cost_h2 = clng(rs_etc("cost")) - cost_h1  ' ����׹�� 50%
        end if

        cost1 = int(saupbu_per      * cost_h1)
        cost2 = int(saupbu_sale_per * cost_h2)
        cost  = cost1 + cost2 ' (To Be) �ο���� + ����׹��

        sql = "SELECT * FROM saupbu_profit_loss                " & chr(13) &_
              " WHERE cost_year   ='"&cost_year&"'             " & chr(13) &_
              "   AND saupbu      ='"&saupbu_tab(i,1)&"'       " & chr(13) &_
              "   AND cost_center ='��������'                " & chr(13) &_
              "   AND cost_id     ='"&rs_etc("cost_id")&"'     " & chr(13) &_
              "   AND cost_detail ='"&rs_etc("cost_detail")&"' "
        set rs_cost=dbconn.execute(sql)
            
        if rs_cost.eof or rs_cost.bof then
            sql = "INSERT INTO saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&")                     " & chr(13) &_
                  "     VALUES ('"&cost_year&"','"&saupbu_tab(i,1)&"','��������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&") "
            dbconn.execute(sql)
        else
            sql = "UPDATE saupbu_profit_loss                       " & chr(13) &_
                  "   SET cost_amt_"&cost_month&"="&cost&"         " & chr(13) &_
                  " WHERE cost_year   ='"&cost_year&"'             " & chr(13) &_
                  "   AND saupbu      ='"&saupbu_tab(i,1)&"'       " & chr(13) &_
                  "   AND cost_center ='��������'                " & chr(13) &_
                  "   AND cost_id     ='"&rs_etc("cost_id")&"'     " & chr(13) &_
                  "   AND cost_detail ='"&rs_etc("cost_detail")&"' "
            dbconn.execute(sql)
        end if      

        rs_etc.movenext()
    loop
    rs_etc.close()
next
' �������� ��� ��

' ���纰 ���� �ڷ� ����
' �������� ���
' ó���� zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='��������') "
dbconn.execute(sql)

sql = " select company,saupbu_per, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY company"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	charge_per = rs("charge_per")

	sql = "select * from trade where trade_name = '"&rs("company")&"'"
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		group_name = "Error"
	  else
		group_name = rs_trade("group_name")
	end if		  

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '��������' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * clng(rs_etc("cost")) * rs("saupbu_per"))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='��������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','��������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='��������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()

' ���纰 ������ ���
' ó���� zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='������') "
dbconn.execute(sql)

sql = " select saupbu,company, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY saupbu,company"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	charge_per = rs("charge_per")

	sql = "select * from trade where trade_name = '"&rs("company")&"'"
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		group_name = "Error"
	  else
		group_name = rs_trade("group_name")
	end if		  

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '������' ) and (saupbu = '"&rs("saupbu")&"' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * Cdbl(rs_etc("cost")))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()
' ���纰 ������ ��� ��

if end_yn = "C" then
	sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
	"' and saupbu = '�����/��������'"
  else
	sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
	"','�����/��������','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
end if
dbconn.execute(sql)

if Err.number <> 0 then
	dbconn.RollbackTrans 
	end_msg = emp_msg + "ó���� Error�� �߻��Ͽ����ϴ�...."
else    
	dbconn.CommitTrans
	end_msg = emp_msg + "����ó�� �Ǿ����ϴ�...."
end if

Response.write "<script language=javascript>"
Response.write "alert('"&end_msg&"');"
Response.write "location.replace('cost_end_mg.asp');"
Response.write "</script>"
Response.End

dbconn.Close()
Set dbconn = Nothing

end if
%>
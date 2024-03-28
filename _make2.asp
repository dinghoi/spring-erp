<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
    Dim DbConnect
    DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"

    Set Dbconn = Server.CreateObject("ADODB.Connection")
    Set Rs     = Server.CreateObject("ADODB.Recordset")
    Set Rs1    = Server.CreateObject("ADODB.Recordset")

    Dbconn.open DbConnect


    Sql="SHOW TABLES LIKE 'test_org_name'"
    rs.Open Sql, Dbconn, 1
    Response.write "<pre>"&sql&"</pre><br>"
    
    if not rs.eof then
    
        sql = "drop table  test_org_name"
        dbconn.execute(sql)
        Response.write "<pre>"&sql&"</pre><br>"

    end if 

    rs.close()

    sql = "  create table test_org_name                                                                                          " & chr(13) &_
          "  as                                                                                                                  " & chr(13) &_
          "  select org_name                                                                                                     " & chr(13) &_ 
          "    from test_2018_person_cost p                                                                                      " & chr(13) &_ 
          "   where p.org_name in (                                                                                              " & chr(13) &_ 
          "                          select company                                                                              " & chr(13) &_ 
          "                            from (                                                                                    " & chr(13) &_ 
          "                                    select saupbu, company, sales_date, approve_no, sales_memo, sales_amt, count(*)   " & chr(13) &_ 
          "                                      from saupbu_sales                                                               " & chr(13) &_ 
          "                                     where sales_date >= '2018-01-01'                                                 " & chr(13) &_ 
          "                                       and saupbu <> '' and company <> ''                                             " & chr(13) &_ 
          "                                       and sales_date <= '2018-12-31'                                                 " & chr(13) &_ 
          "                                  group by saupbu, company, sales_date, approve_no, sales_memo, sales_amt             " & chr(13) &_ 
          "                                  order by saupbu, company, sales_date, approve_no, sales_memo, sales_amt             " & chr(13) &_ 
          "                                 ) a                                                                                  " & chr(13) &_ 
          "                        group by company                                                                              " & chr(13) &_ 
          "                       )                                                                                              " & chr(13) &_ 
          " group by org_name                                                                                                    "
          ' "                                       and (company not like '한화생명%'                                              " & chr(13) &_ 
          ' "                                               and company not like '한생%')                                          " & chr(13) &_ 
    dbconn.execute(sql)
    Response.write "<pre>"&sql&"</pre><br>"


    Sql="select org_name from test_org_name "
    rs.Open Sql, Dbconn, 1
    Response.write "<pre>"&sql&"</pre><br>"

    %>
    <table border="1" cellpadding="0" cellspacing="0">
        <tr valign="middle" bgcolor="#EFEFEF" class="style11B">
            <td width="220" height="28"><div align="center">고객사</div></td>
            <td width="155"><div align="center">매출합</div></td>
            <td width="160"><div align="center">매입합</div></td>
            <td width="170"><div align="center">비용합</div></td>
            <td width="235"><div align="center">손익</div></td>
        </tr>
    <%

    do until rs.eof

        sql = "select a,b,c,a-(b+c)  as result                                                                                                                                                                 " & chr(13) &_
              "from                                                                                                                                                                                            " & chr(13) &_
              "(                                                                                                                                                                                               " & chr(13) &_
              "select                                                                                                                                                                                          " & chr(13) &_
              "(                                                                                                                                                                                               " & chr(13) &_
              "  select sum(sales_amt) sumvalue                                                                                                                                                                " & chr(13) &_
              "  from                                                                                                                                                                                          " & chr(13) &_
              "  (                                                                                                                                                                                             " & chr(13) &_
              "    select saupbu, company, sales_date, approve_no, sales_memo, sales_amt, count(*)                                                                                                             " & chr(13) &_
              "      from saupbu_sales                                                                                                                                                                         " & chr(13) &_
              "     where sales_date >= '2018-01-01'                                                                                                                                                           " & chr(13) &_
              "       and saupbu <> '' and company <> ''                                                                                                                                                       " & chr(13) &_
              "       and  sales_date <= '2018-12-31'                                                                                                                                                          " & chr(13) &_
              "       and company = '" & rs("org_name") & "'                                                                                                                                                   " & chr(13) &_
              "  group by saupbu, company, sales_date, approve_no, sales_memo, sales_amt                                                                                                                       " & chr(13) &_
              "  ) a                                                                                                                                                                                           " & chr(13) &_
              ") a                                                                                                                                                                                             " & chr(13) &_
              ",                                                                                                                                                                                               " & chr(13) &_
              "(                                                                                                                                                                                               " & chr(13) &_
              "select sum(cost)  sumvalue                                                                                                                                                                      " & chr(13) &_
              "from                                                                                                                                                                                            " & chr(13) &_
              "(                                                                                                                                                                                               " & chr(13) &_
              "  select p.*                                                                                                                                                                                    " & chr(13) &_
              "    from test_2018_person_cost p                                                                                                                                                                " & chr(13) &_
              "   where p.org_name = '" & rs("org_name") & "'                                                                                                                                                  " & chr(13) &_
              ") b                                                                                                                                                                                             " & chr(13) &_
              ") b                                                                                                                                                                                             " & chr(13) &_
              ",                                                                                                                                                                                               " & chr(13) &_
              "(                                                                                                                                                                                               " & chr(13) &_
              "select sum(cost) sumvalue                                                                                                                                                                       " & chr(13) &_
              "from                                                                                                                                                                                            " & chr(13) &_
              "(                                                                                                                                                                                               " & chr(13) &_
              "  select pmg_emp_no as emp_no, '급여' as gubun, pmg_emp_name,pmg_date as slip_date,pmg_emp_name as user_name,'' as user_grade,'급여' as slip_memo,pmg_give_total as cost, '급여' as cost_detail " & chr(13) &_
              "    from pay_month_give                                                                                                                                                                         " & chr(13) &_
              "   where pmg_yymm like '2018%'                                                                                                                                                                  " & chr(13) &_
              "     and pmg_emp_no in                                                                                                                                                                          " & chr(13) &_
              "                    (                                                                                                                                                                           " & chr(13) &_
              "                      select p.emp_no                                                                                                                                                           " & chr(13) &_
              "                       from test_2018_person_cost p                                                                                                                                             " & chr(13) &_
              "                      where p.org_name  = '" & rs("org_name") & "'                                                                                                                              " & chr(13) &_
              "                    )                                                                                                                                                                           " & chr(13) &_
              ") c                                                                                                                                                                                             " & chr(13) &_
              ") c                                                                                                                                                                                             " & chr(13) &_
              ") minus                                                                                                                                                                                         "
            'Response.write "<pre>"&sql&"</pre><br>"
            rs1.Open Sql, Dbconn, 1
            
            
            if not rs1.eof then
                org_name = rs("org_name")
                a = rs1("a")
                b = rs1("b")
                c = rs1("c")
                result = rs1("result")
                
                %>
                <tr>
                    <td height="27"><div align="center"><%=org_name%></div></td>
                    <td height="27"><div align="center"><%=a%></div></td>
                    <td height="27"><div align="center"><%=b%></div></td>
                    <td hei8ght="27"><div align="center"><%=c%></div></td>
                    <td hei8ght="27"><div align="center"><%=result%></div></td>
                </tr>
                <%
                rs1.movenext()
            end if
            

            rs1.close()
        

        rs.movenext()
  loop

  %>
  </table>
  <%

  rs.close()
  Set rs = Nothing
  
%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
    Dim DbConnect
    DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"

    Set Dbconn = Server.CreateObject("ADODB.Connection")
    Set Rs     = Server.CreateObject("ADODB.Recordset")
    Set Rs1    = Server.CreateObject("ADODB.Recordset")

    Dbconn.open DbConnect

    Sql="SHOW TABLES LIKE 'temp_person_cost'"
    rs.Open Sql, Dbconn, 1
    Response.write "<pre>"&sql&"</pre><br>"

    if not rs.eof then

        sql = "drop table  temp_person_cost"
        dbconn.execute(sql)
        Response.write "<pre>"&sql&"</pre><br>"
    end if 

    rs.close()
    Set rs = Nothing


    sql = "  CREATE TABLE temp_person_cost                                                                 " & chr(13) &_
          "  AS                                                                                            " & chr(13) &_
          "       SELECT '���Լ��ݰ�꼭' as gubun                                                         " & chr(13) &_
          "            , a.emp_no as emp_no                                                                " & chr(13) &_
          "            , a.company as org_name                                                             " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.cost,concat(a.account,' '                                                       " & chr(13) &_
          "            , a.account_item) as cost_detail                                                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM general_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE a.tax_bill_yn = 'Y'                                                               " & chr(13) &_
          "          AND (a.slip_date >= '2018-01-31' AND a.slip_date <= '2018-12-31')                     " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '��Ư��' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.work_date as slip_date                                                          " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , a.work_item as slip_memo                                                          " & chr(13) &_
          "            , a.overtime_amt as cost                                                            " & chr(13) &_
          "            , a.work_gubun as cost_detail                                                       " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM overtime  a                                                                       " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.work_date,1,4), SUBSTRING(a.work_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.work_date >= '2018-01-01' AND a.work_date <= '2018-12-31')                     " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '�Ϲݰ��' as gubun                                                               " & chr(13) &_
          "            , a.emp_no as emp_no                                                                " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.cost                                                                            " & chr(13) &_
          "            , concat(a.account,' ',a.account_item) as cost_detail                               " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM general_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.slip_date >= '2018-01-01' AND a.slip_date <= '2018-12-31')                     " & chr(13) &_
          "          AND a.tax_bill_yn <> 'Y'                                                              " & chr(13) &_
          "          AND a.slip_gubun = '���'                                                             " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '���߱���' as gubun                                                               " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.company,' ',a.run_memo) as slip_memo                                     " & chr(13) &_
          "            , a.fare as cost                                                                    " & chr(13) &_
          "            , a.transit as cost_detail                                                          " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '2018-01-01' AND a.run_date <= '2018-12-31')                       " & chr(13) &_
          "          AND a.car_owner = '���߱���'                                                          " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '����Ÿ�' as gubun                                                               " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.far as cost                                                                     " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '2018-01-01' AND a.run_date <= '2018-12-31')                       " & chr(13) &_
          "          AND a.car_owner = '����'                                                              " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '������' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.oil_price as cost                                                               " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '2018-01-01' AND a.run_date <= '2018-12-31')                       " & chr(13) &_
          "          AND a.car_owner = 'ȸ��'                                                              " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '������' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.parking as cost                                                                 " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '2018-01-01' AND a.run_date <= '2018-12-31')                       " & chr(13) &_
          "          AND a.parking > 0                                                                     " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '�����' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.toll as cost                                                                    " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '2018-01-01' AND a.run_date <= '2018-12-31')                       " & chr(13) &_
          "          AND a.toll > 0                                                                        " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '����������' as gubun                                                             " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.repair_cost as cost                                                             " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '2018-01-01' AND a.run_date <= '2018-12-31')                       " & chr(13) &_
          "          AND a.car_owner = 'ȸ��'                                                              " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT a.emp_no as emp_no                                                                " & chr(13) &_
          "            , '����ī��' as gubun                                                               " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.price as cost                                                                   " & chr(13) &_
          "            , concat(a.account,' ',a.account_item) as cost_detail                               " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM card_slip  a                                                                      " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE a.card_type like '%����%'                                                         " & chr(13) &_
          "          AND (a.slip_date >= '2018-01-01' AND a.slip_date <= '2018-12-31')                     " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '����ī��' as gubun                                                               " & chr(13) &_
          "            , a.emp_no as emp_no                                                                " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.cost as cost                                                                    " & chr(13) &_
          "            , concat(a.account,' ',a.account_item) as cost_detail                               " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM card_slip a                                                                       " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE a.card_type not like '%����%'                                                     " & chr(13) &_
          "          AND (a.slip_date >= '2018-01-01' AND a.slip_date <= '2018-12-31')                     "
    dbconn.execute(sql)  
       
    Response.write "<pre>"&sql&"</pre><br>"
%>
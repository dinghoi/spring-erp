<!--#include virtual="/include/JSON_2.0.4.asp"-->
<!--#include virtual="/include/JSON_UTIL_0.1.1.asp"-->

<% 
    Server.ScriptTimeout = 180  ' 3��

    On Error Resume Next

    Dim DbConnect
    DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"

    ' https://code.google.com/archive/p/aspjson/ 
    
    Set Dbconn = Server.CreateObject("ADODB.Connection")
    Set Rs     = Server.CreateObject("ADODB.Recordset")

    Dbconn.open DbConnect

    start  = request.form("start")
    length = request.form("length")

    sql = "      SELECT  gubun                                                                                       " & chr(13) &_ 
          "            , emp_no                                                                                      " & chr(13) &_ 
          "            , org_name                                                                                    " & chr(13) &_ 
          "            , slip_date                                                                                   " & chr(13) &_ 
          "            , concat(user_name,' ',user_grade) as user                                                    " & chr(13) &_ 
          "            , slip_memo                                                                                   " & chr(13) &_ 
          "            , cost                                                                                        " & chr(13) &_ 
          "            , cost_detail                                                                                  " & chr(13) &_ 
          "            , emp_saupbu                                                                                  " & chr(13) &_ 
          "            , cost_center                                                                                 " & chr(13) &_ 
          "            , CASE WHEN cost_center = '������'     THEN `cost` ELSE 0 END AS cost_a /* '������' */        " & chr(13) &_ 
          "            , CASE WHEN cost_center = '����������' THEN `cost` ELSE 0 END AS cost_b /* '����������' */    " & chr(13) &_ 
          "            , CASE WHEN cost_center = '�ι������' THEN `cost` ELSE 0 END AS cost_c /* '�ι������' */    " & chr(13) &_ 
          "            , CASE WHEN cost_center = '��������' THEN `cost` ELSE 0 END AS cost_d /* '��������' */    " & chr(13) &_ 
          "            , CASE WHEN cost_center = '��������'   THEN `cost` ELSE 0 END AS cost_e /* '��������' */      " & chr(13) &_ 
          "            , CASE WHEN (    cost_center <> '������'                                                      " & chr(13) &_ 
          "                         and cost_center <> '����������'                                                  " & chr(13) &_ 
          "                         and cost_center <> '��������'                                                  " & chr(13) &_ 
          "                         and cost_center <> '�ι������'                                                  " & chr(13) &_ 
          "                         and cost_center <> '��������'                                                    " & chr(13) &_ 
          "                        )  THEN `cost`                                                                    " & chr(13) &_ 
          "                        ELSE 0                                                                            " & chr(13) &_ 
          "              END AS cost_etc /* '�׿�' */                                                                " & chr(13) &_ 
          "         FROM temp_person_cost                                                                            " & chr(13) &_ 
          "     GROUP BY gubun                                                                                       " & chr(13) &_ 
          "            , emp_no                                                                                      " & chr(13) &_ 
          "            , org_name                                                                                    " & chr(13) &_ 
          "            , slip_date                                                                                   " & chr(13) &_ 
          "            , user_name                                                                                   " & chr(13) &_ 
          "            , user_grade                                                                                  " & chr(13) &_ 
          "            , slip_memo                                                                                   " & chr(13) &_ 
          "            , cost                                                                                        " & chr(13) &_ 
          "            , cost_detail                                                                                 " & chr(13) &_ 
          "            , emp_saupbu                                                                                  " & chr(13) &_ 
          "            , cost_center                                                                                 " & chr(13) &_ 
          "     ORDER BY org_name                                                                                    " & chr(13) &_ 
          "            , slip_date                                                                                   " & chr(13) &_ 
          "        LIMIT "&start&", "&length&"                                                                       "           

    'Response.write sql 
    DataTablesQueryToJSON(dbconn, sql,10,10).Flush

    If Err.number <> 0 Then     '���� �߻� �� �� �κ� ����
	    Response.Write "" & Err.Source & "<br>"
	    Response.Write "���� ��ȣ : " & Err.number & "<br>"
	    Response.Write "���� : " & Err.Description & "<br>"
	Else
	    ' Response.Write "������ �����ϴ�."
	End If

%>


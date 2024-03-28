<!--#include virtual="/include/JSON_2.0.4.asp"-->
<!--#include virtual="/include/JSON_UTIL_0.1.1.asp"-->


<%
Dim member
Set member = jsObject()

member("name") = "Tu?rul"
member("surname") = "Topuz"
member("message") = "Hello World"

member.Flush
%>

<p>

    ------------------------------------------------------------------------------

</p>

<%
Dim DbConnect
    DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"

    Set Dbconn = Server.CreateObject("ADODB.Connection")

    Dbconn.open DbConnect


sql = "   select saupbu, company, sales_date, approve_no, sales_memo, sales_amt  " & chr(13) &_ 
          "     from saupbu_sales limit 100                                        "
QueryToJSON(dbconn, sql).Flush
%>

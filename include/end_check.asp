<%
Dim end_saupbu, sql, rs_end

If saupbu = "" Then
	end_saupbu = "사업부외나머지"
Else
  	end_saupbu = saupbu
End If

sql = "SELECT MAX(end_month) as max_month " &_
      "  FROM cost_end                    " &_
     " WHERE saupbu = '"&end_saupbu&"'   " &_
     "   AND end_yn ='Y'                 "

Set rs_end = DBConn.Execute(sql)

If IsNull(rs_end("max_month")) Then
	end_date = "2014-08-31"
Else
	new_date = DateAdd("m", 1, DateValue(Mid(rs_end("max_month"), 1, 4) & "-" & Mid(rs_end("max_month"), 5, 2) & "-01"))
	end_date = DateAdd("d", -1, new_date)
End If

rs_end.Close()
%>
<%
Function QueryToJSON(dbc, sql)
        Dim rs, jsa
        Set rs = dbc.Execute(sql)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rs.MoveNext
        Wend
        Set QueryToJSON = jsa
End Function

Function DataTablesQueryToJSON(dbc, sql, recordsTotal, recordsFiltered)
        Dim rs, jsa, root

        Set root = jsObject()
        'root("draw") = 1
        root("recordsTotal") = recordsTotal
        root("recordsFiltered") = recordsFiltered

        Set rs = dbc.Execute(sql)
        Set jsa = jsArray()

        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rs.MoveNext
        Wend

        Set root("data") = jsa
        Set DataTablesQueryToJSON = root
End Function
%>
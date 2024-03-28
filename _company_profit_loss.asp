<%@LANGUAGE="VBSCRIPT"%>
<%
Dim DbConnect
DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"

Set Dbconn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")

Dbconn.open DbConnect

curr_date = mid(cstr(now()),1,10)

title_line = "ȸ�纰 ������Ȳ"
%>
<!DOCTYPE html>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        
        <title><%=title_line%></title>
        
        <link href="/include/style.css" type="text/css" rel="stylesheet">
        <link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
        <link href="https://cdn.datatables.net/v/dt/dt-1.10.18/datatables.min.css" type="text/css" rel="stylesheet"> 
        
        <script src="/java/jquery-1.9.1.js"></script>
	    <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
        <script src="/java/js_form.js" type="text/javascript"></script>
        <script src="https://cdn.datatables.net/v/dt/dt-1.10.18/datatables.min.js" type="text/javascript"></script>
        
        <script type="text/javascript">

            var table ;

            $(document).ready( function() 
            {
                $( "#datepicker1" ).datepicker({
                    onSelect: function(dateText) { onProfitItems(dateText, $("#datepicker2").val(), $('#saupbu').val()); } 
                });
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
                //$( "#datepicker1" ).datepicker("setDate", "2018-12-31" );
                $( "#datepicker1" ).datepicker("setDate", "<%=curr_date%>" );                
                
                $( "#datepicker2" ).datepicker({
                    onSelect: function(dateText) { onProfitItems($("#datepicker1").val(), dateText, $('#saupbu').val()); } 
                });
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
                $( "#datepicker2" ).datepicker("setDate", "<%=curr_date%>" );
                
                ////////////////


                $('#button').click( function () 
                {
                    
                    $('#loss').DataTable( {
                        "processing" : true,
                        "serverSide" : true,
                        "destroy"    : true,
                        "ajax"       : {
                                         "url"         : "_ajax_company_loss_list.asp",
                                         "type"        : "post",
                                         "dataType"    : "json",
                                         "contentType" : "application/x-www-form-urlencoded; charset=UTF-8"
                                       },
                        "columns": [
                            { data : "gubun" },
                            { data : "emp_no" },
                            { data : "org_name" },
                            { data : "slip_date" },
                            { data : "user" },
                            { data : "slip_memo" },
                            { data : "cost" },
                            { data : "cost_detail" },
                            { data : "emp_saupbu" },
                            { data : "cost_center" },
                            { data : "cost_a" },
                            { data : "cost_b" },
                            { data : "cost_c" },
                            { data : "cost_d" },
                            { data : "cost_e" },
                            { data : "cost_etc" }
                        ]
                    
                    } );
                        
                    console.log("_ajax_company_loss_list.asp");
                } );

                $('#saupbu').on('change', function() 
                {                    
                    onProfitItems($("#datepicker1").val(),$("#datepicker2").val(),$('#saupbu').val()); 
                } );    
            } );            

            function onLossItems() 
            {
                    $.ajax({ url: "_ajax_company_loss_list.asp"
     				        ,async: false    					
      				        ,type: 'post'
      				        // ,data: params
      				        ,dataType: "json"
                            ,contentType: 'application/x-www-form-urlencoded; charset=UTF-8'
      				        ,beforeSend: function(jqXHR){
      				        	jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
      				        }
      				        ,success: function(data)
      				        {    					
                                console.log(data);
                                /***********
                                $("#divSelectedCompany").text('');  

                                table = $('#profit').DataTable( {
                                            "destroy": true,
                                            "data": data,
                                            // "searching": false, // �˻� �Է� �ڽ� ����
                                            // �÷��� �˻���� ����
                                            "columnDefs": [ { "searchable": false, "targets": 0 },
                                                            { "searchable": false, "targets": 2 },
                                                            { "searchable": false, "targets": 3, "className": 'dt-body-right' },
                                                            { "searchable": false, "targets": 4, "className": 'dt-body-center', 
                                                              render: function ( data, type, row ) {
                                                                    if( data != null) 
                                                                    {
                                                                        data = data.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
                                                                        return data;
                                                                    } 
                                                                    else { return data; }
                                                              }
                                                            }
                                                          ],
                                            "columns": [  { data: 'saupbu' },
                                                          { data: 'company' },
                                                          { data: 'sales_memo' },
                                                          { data: 'sales_amt' },
                                                          { data: 'cnt' }
                                                        ]
                                        } );

                                // �����+���縦 �����Ҷ�...
                                $('#profit tbody').off( 'click', 'tr' );
                                $('#profit tbody').on( 'click', 'tr', function () 
                                {
                                    $(this).toggleClass('selected');

                                    if ($(this).hasClass("selected") === true) 
                                    {
                                        console.log('selected !!')
                                    }
                                    else
                                    {
                                        console.log('unselected !!')
                                    }

                                    companys = new Array; // ������� �迭

                                    var sales_amt = 0 ;

                                    table.rows().every( function ( rowIdx, tableLoop, rowLoop ) {
                                        var d = this.data();
                                    
                                        //console.log(d.company);
                                        var company = d.company;                                        
                                        var node = this.node();

                                        if ($(node).hasClass("selected") === true)
                                        {
                                            companys.push(company); 

                                            sales_amt = sales_amt + eval(d.sales_amt);

                                            console.log(d.company);
                                        }
                                    } );
                                    $("#divSelectedCompany").text(companys.join(',')+' \\'+sales_amt.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
                                } );        
                                *********/
      				        }
      				        ,error: function(jqXHR, status, errorThrown){
      				            alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
      				        }
                    });        
            }

            function onProfitItems(date1, date2, saupbu) 
            {
                alert();
                // ���ڿ��� date�������� �˻�..
                function dateCheck(dateString)
                {
                    if  (dateString=='') return false ;

                    var pattern = /[0-9]{4}-[0-9]{2}-[0-9]{2}/;

                    return pattern.test(dateString);
                }

                if (dateCheck(date1) && dateCheck(date2))
                {
                    var saupbu = escape(saupbu);
                    var params = { "date1" : date1, "date2" : date2, "saupbu" :saupbu };
                                       
                    $.ajax({ url: "_ajax_company_profit_list.asp"
     				        ,async: false    					
      				        ,type: 'post'
      				        ,data: params
      				        ,dataType: "json"
                            ,contentType: 'application/x-www-form-urlencoded; charset=UTF-8'
      				        ,beforeSend: function(jqXHR){
      				        	jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
      				        }
      				        ,success: function(data)
      				        {    					
                                $("#divSelectedCompany").text('');  

                                table = $('#profit').DataTable( {
                                            "destroy": true,
                                            "data": data,
                                            // "searching": false, // �˻� �Է� �ڽ� ����
                                            // �÷��� �˻���� ����
                                            "columnDefs": [ { "searchable": false, "targets": 0 },
                                                            { "searchable": false, "targets": 2 },
                                                            { "searchable": false, "targets": 3, "className": "dt-body-right",
                                                              "render": function ( data, type, row ) {
                                                                            if( data != null) 
                                                                            {
                                                                                data = '\\ '+data.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
                                                                                return data;
                                                                            } 
                                                                            else { return data; }
                                                                        }
                                                            },
                                                            { "searchable": false, "targets": 4, "className": "dt-body-center"}
                                                          ],
                                            "columns": [  { data: 'gubun' },
                                                          { data: 'org_name' },
                                                          { data: 'slip_date' },
                                                          { data: 'user' },
                                                          { data: 'slip_memo' },
                                                          { data: 'cost' },
                                                          { data: 'cost_detail' },
                                                          { data: 'emp_saupbu' },
                                                          { data: 'cost_center' },
                                                          { data: 'cost_a' },
                                                          { data: 'cost_b' },
                                                          { data: 'cost_c' },
                                                          { data: 'cost_d' },
                                                          { data: 'cost_e' },
                                                          { data: 'cost_etc' }
                                                        ]
                                        } );

                                // �����+���縦 �����Ҷ�...
                                $('#profit tbody').off( 'click', 'tr' );
                                $('#profit tbody').on( 'click', 'tr', function () 
                                {
                                    $(this).toggleClass('selected');

                                    if ($(this).hasClass("selected") === true) 
                                    {
                                        console.log('selected !!')
                                    }
                                    else
                                    {
                                        console.log('unselected !!')
                                    }

                                    companys = new Array; // ������� �迭

                                    var sales_amt = 0 ;

                                    table.rows().every( function ( rowIdx, tableLoop, rowLoop ) {
                                        var d = this.data();
                                    
                                        //console.log(d.company);
                                        var company = d.company;                                        
                                        var node = this.node();

                                        if ($(node).hasClass("selected") === true)
                                        {
                                            companys.push(company); 

                                            sales_amt = sales_amt + eval(d.sales_amt);

                                            console.log(d.company);
                                        }
                                    } );
                                    $("#divSelectedCompany").text(companys.join(',')+' \\'+sales_amt.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
                                } );        
                                
      				        }
      				        ,error: function(jqXHR, status, errorThrown){
      				            alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
      				        }
                    });        
                    /*
                    $('#profit').DataTable({
                        "ajax": {
                            "url"  : "_ajax_company_profit_loss_get_company_list.asp",
                            "type" : "POST",
                            "data" : { "date1" : date1, "date2" : date2 }
                        }
                    });
                    */
                }
                console.log(date1);
                console.log(date2);
            }
		</script>
	</head>
	<body>

        

        <div id="wrap">
            <div id="container">
                <button id="button">Row count</button>

                <h3 class="tit"><%=title_line%></h3>
                
                <fieldset class="srch">
                    <label>
                        <strong>����Ⱓ : </strong>
                        <input name="date1" type="text" style="width:70px" id="datepicker1"> 
                        ~
                        <input name="date2" type="text" style="width:70px" id="datepicker2">
                    </label>

                    <select name="saupbu" id="saupbu" style="width:150px">
                        <option value="">��ü</option>
                        <%
                        Sql = "   select saupbu             " & chr(13) &_ 
                              "     from saupbu_sales       " & chr(13) &_ 
                              "    where saupbu is not null " & chr(13) &_ 
                              "      and saupbu <>''        " & chr(13) &_ 
                              " group by saupbu             "
                        rs.Open Sql, Dbconn, 1
                        do until rs.eof

                            %><option value='<%=rs("saupbu")%>' ><%=rs("saupbu")%></option><%
                            
                            rs.movenext()  
                        loop 
                        rs.Close()
                        %>
                        </select>

                    <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                </fieldset>
            </div>

            <table id="profit" class="display" style="width:100%">
                <thead>
                    <tr>
                        <th>�����</th>
                        <th>����</th>
                        <th>���⳻��</th>
                        <th>�ݾ�</th>
                        <th>�Ǽ�</th>
                    </tr>
                </thead>
            </table>

            <div id="divSelectedCompany"/>

            <table id="loss" class="display" style="width:100%">
                <thead>
                    <tr>
                        <th>����</th>        <!-- gubun       -->
                        <th>����</th>      <!-- org_name    -->
                        <th>�߻���</th>      <!-- slip_date   -->
                        <th>���</th>        <!-- user        -->
                        <th>�޸�</th>        <!-- slip_memo   -->
                        <th>�ݾ�</th>        <!-- cost        -->
                        <th>����</th>        <!-- cost_detail -->
                        <th>�����</th>      <!-- emp_saupbu  -->
                        <th>��񱸺�</th>    <!-- cost_center -->
                        <th>������</th>      <!-- cost_a      -->
                        <th>����������</th>  <!-- cost_b      -->
                        <th>�ι������</th>  <!-- cost_c      -->
                        <th>��������</th>  <!-- cost_d      -->
                        <th>��������</th>    <!-- cost_e      -->
                        <th>�׿�</th>        <!-- cost_etc    -->
                    </tr>
                </thead>
            </table>

        </div>
    </body>
</html>    
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

        <script type="text/javascript" id="working">

            // Korean
            var lang_kor = {
                        "decimal" : "",
                        "emptyTable" : "�����Ͱ� �����ϴ�.",
                        "info" : "_START_ - _END_ (�� _TOTAL_ ��)",
                        "infoEmpty" : "0��",
                        "infoFiltered" : "(��ü _MAX_ �� �� �˻����)",
                        "infoPostFix" : "",
                        "thousands" : ",",
                        "lengthMenu" : "_MENU_ ���� ����",
                        "loadingRecords" : "�ε���...",
                        "processing" : "ó����...",
                        "search" : "�˻� : ",
                        "zeroRecords" : "�˻��� �����Ͱ� �����ϴ�.",
                        "paginate" : {
                            "first" : "ù ������",
                            "last" : "������ ������",
                            "next" : "����",
                            "previous" : "����"
                        },
                        "aria" : {
                            "sortAscending" : " :  �������� ����",
                            "sortDescending" : " :  �������� ����"
                        },
                        "processing"   : true,
                        "bProcessing"  : true,
		                "sProcessing"  : "<div style='position:relative; align:center; z-index:1000'><img src='image/loading.gif' alt='�ε�..' /></div>"
            };

            var tableProfit ; // ���� Table ..


            $(document).ready( function()
            {
                // ����Ⱓ (����)
                $( "#datepicker1" ).datepicker({
                    onSelect: function(dateText) {
                        onProfitItems(dateText, $("#datepicker2").val(), $('#saupbu').val());
                    }
                });
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
                $( "#datepicker1" ).datepicker("setDate", "2018-12-31" );
                //$( "#datepicker1" ).datepicker("setDate", "<%=curr_date%>" );

                // ����Ⱓ (����)
                $( "#datepicker2" ).datepicker({
                    onSelect: function(dateText) {
                        onProfitItems($("#datepicker1").val(), dateText, $('#saupbu').val());
                    }
                });
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
                $( "#datepicker2" ).datepicker("setDate", "<%=curr_date%>" );

                // �����
                $('#saupbu').on('change', function() {
                    onProfitItems($("#datepicker1").val(),$("#datepicker2").val(),$('#saupbu').val());
                } );

                // �����ڷ� ���ϱ�
                onProfitItems($("#datepicker1").val(), $("#datepicker2").val(), $('#saupbu').val());

                // ����ڷ� ���ϱ�
                onLossItems();

                $("#company").keyup(function( event )
                {
                    var invalid = [  9   // Tab
                                  //,  8   // Backspace
                                  , 13   // Enter
                                  , 16   // Shift
                                  , 17   // Ctrl
                                  , 18   // Alt
                                  , 19   // Pause, Break
                                  , 20   // CapsLock
                                  , 27   // Esc
                                  , 32   // Space
                                  , 33   // Page Up
                                  , 34   // Page Down
                                  , 35   // End
                                  , 36   // Home
                                  , 37   // Left arrow
                                  , 38   // Up arrow
                                  , 39   // Right arrow
                                  , 40   // Down arrow
                                  , 44   // PrntScrn (see below��)
                                  , 45   // Insert
                                  // , 46   // Delete
                                  ,112   // F1
                                  ,113   // F2
                                  ,114   // F3
                                  ,115   // F4
                                  ,116   // F5
                                  ,117   // F6
                                  ,118   // F7
                                  ,119   // F8
                                  ,120   // F9
                                  ,121   // F10
                                  ,122   // F11
                                  ,123   // F12
                                  ,144   // NumLock
                                  ,145   // ScrollLock
                                  ];

                    if (invalid.indexOf(event.keyCode) == -1)
                    {
                        tableProfit.search( escape(this.value) ).draw();
                    }
                })
            } );



            function onProfitItems(date1, date2, saupbu)
            {
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
                    var params = { "date1" : date1, "date2" : date2, "saupbu" : saupbu };

                    tableProfit = $('#profit').DataTable( {
                        "processing" : true,
                        "serverSide" : true,
                        "destroy"    : true,
                        "language"   : lang_kor,
                        //"bInfo"        : false, // �Ʒ� ������ ���������ʴ´�.
                        "ajax"       : {
                                         "url"         : "_ajax_company_profit_list.asp",
                                         "type"        : "post",
                                         "data"        : params,
                                         "dataType"    : "json",
                                         "contentType" : "application/x-www-form-urlencoded; charset=UTF-8"
                                       },
                        "columns": [{ data: 'saupbu' },
                                    { data: 'company' },
                                    { data: 'sales_memo' },
                                    { data: 'sales_amt' },
                                    { data: 'cnt' }
                                   ],
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
                                     ]
                    } );
                }
                console.log(date1);
                console.log(date2);
            }

            function onLossItems()
            {
                $('#loss').DataTable( {
                        "processing" : true,
                        "serverSide" : true,
                        "destroy"    : true,
                        "language"   : lang_kor,
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
                                    { data : "emp_saupbu" }//,
                                    /*
                                    { data : "cost_center" } ,
                                    { data : "cost_a" },
                                    { data : "cost_b" },
                                    { data : "cost_c" },
                                    { data : "cost_d" },
                                    { data : "cost_e" },
                                    { data : "cost_etc" }
                                    */
                                   ],
                        "columnDefs": [ { "searchable": false, "targets": 0 },
                                        { "searchable": false, "targets": 1 },
                                        { "searchable": true,  "targets": 2 },
                                        { "searchable": false, "targets": 3, "className": "dt-body-center"},
                                        { "searchable": false, "targets": 4 },
                                        { "searchable": false, "targets": 5 },
                                        { "searchable": false, "targets": 6, "className": "dt-body-right",
                                          "render": function ( data, type, row ) {
                                                        if( data != null)
                                                        {
                                                            data = '\\ '+data.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
                                                            return data;
                                                        }
                                                        else { return data; }
                                                    }
                                        },
                                        { "searchable": false, "targets": 7 },
                                        { "searchable": false, "targets": 8 }
                                     ]
                } );

                $(".dataTables_filter").hide(); // �˻���ư�� �����.
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
                        <input id="datepicker1" name="date1" type="text" style="width:70px">
                        ~
                        <input id="datepicker2" name="date2" type="text" style="width:70px">
                    </label>

                    <strong>����� : </strong>
                    <select id="saupbu" name="saupbu" style="width:150px">
                        <option value="">��ü</option>
                        <%
                        Sql = "   SELECT saupbu             " & chr(13) &_
                              "     FROM saupbu_sales       " & chr(13) &_
                              "    WHERE saupbu IS NOT NULL " & chr(13) &_
                              "      AND saupbu <>''        " & chr(13) &_
                              " GROUP BY saupbu             "
                        rs.Open Sql, Dbconn, 1
                        do until rs.eof

                            %><option value='<%=rs("saupbu")%>' ><%=rs("saupbu")%></option><%

                            rs.movenext()
                        loop
                        rs.Close()
                        %>
                        </select>

                    <strong>������ : </strong>
                    <input id="company" type="text" value="">

                    <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>

                </fieldset>
            </div>

            <div style="height:30px;">&nbsp;</div>

            <table id="profit" class="display" style="width:100%">
                <thead>
                    <tr>
                        <th>�����</th>
                        <th>������</th>
                        <th>���⳻��</th>
                        <th>�ݾ�</th>
                        <th>�Ǽ�</th>
                    </tr>
                </thead>
            </table>

            <div style="height:30px;" id="divSelectedCompany">&nbsp;</div>

            <table id="loss" class="display" style="width:100%">
                <thead>
                    <tr>
                        <th>����</th>        <!-- gubun       -->
                        <th>���</th>        <!-- emp_no      -->
                        <th>������</th>      <!-- org_name    -->
                        <th>�߻���</th>      <!-- slip_date   -->
                        <th>���</th>        <!-- user        -->
                        <th>�޸�</th>        <!-- slip_memo   -->
                        <th>�ݾ�</th>        <!-- cost        -->
                        <th>����</th>        <!-- cost_detail -->
                        <th>�����</th>      <!-- emp_saupbu  -->
                        <!-- <th>��񱸺�</th>   --> <!-- cost_center -->
                        <!-- <th>������</th>     --> <!-- cost_a      -->
                        <!-- <th>����������</th> --> <!-- cost_b      -->
                        <!-- <th>�ι������</th> --> <!-- cost_c      -->
                        <!-- <th>��������</th> --> <!-- cost_d      -->
                        <!-- <th>��������</th>   --> <!-- cost_e      -->
                        <!-- <th>�׿�</th>       --> <!-- cost_etc    -->
                    </tr>
                </thead>
            </table>
        </div>
    </body>
</html>
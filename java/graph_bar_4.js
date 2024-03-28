			var chart1;
			var chart2;
			var chart3;
			var chart4;
			var chart;
			$(document).ready(function() {
				j = "3";
				s_tab10 = parseFloat(document.frm.s_tab10.value.replace(/,/g,""));
				s_tab11 = parseFloat(document.frm.s_tab11.value.replace(/,/g,""));
				s_tab12 = parseFloat(document.frm.s_tab12.value.replace(/,/g,""));
				s_tab13 = parseFloat(document.frm.s_tab13.value.replace(/,/g,""));
				s_tab20 = parseFloat(document.frm.s_tab20.value.replace(/,/g,""));
				s_tab21 = parseFloat(document.frm.s_tab21.value.replace(/,/g,""));
				s_tab22 = parseFloat(document.frm.s_tab22.value.replace(/,/g,""));
				s_tab23 = parseFloat(document.frm.s_tab23.value.replace(/,/g,""));
				s_tab30 = parseFloat(document.frm.s_tab30.value.replace(/,/g,""));
				s_tab31 = parseFloat(document.frm.s_tab31.value.replace(/,/g,""));
				s_tab32 = parseFloat(document.frm.s_tab32.value.replace(/,/g,""));
				s_tab33 = parseFloat(document.frm.s_tab33.value.replace(/,/g,""));
				s_tab40 = parseFloat(document.frm.s_tab40.value.replace(/,/g,""));
				s_tab41 = parseFloat(document.frm.s_tab41.value.replace(/,/g,""));
				s_tab42 = parseFloat(document.frm.s_tab42.value.replace(/,/g,""));
				s_tab43 = parseFloat(document.frm.s_tab43.value.replace(/,/g,""));
				s_tab50 = parseFloat(document.frm.s_tab50.value.replace(/,/g,""));
				s_tab51 = parseFloat(document.frm.s_tab51.value.replace(/,/g,""));
				s_tab52 = parseFloat(document.frm.s_tab52.value.replace(/,/g,""));
				s_tab53 = parseFloat(document.frm.s_tab53.value.replace(/,/g,""));
				s_tab60 = parseFloat(document.frm.s_tab60.value.replace(/,/g,""));
				s_tab61 = parseFloat(document.frm.s_tab61.value.replace(/,/g,""));
				s_tab62 = parseFloat(document.frm.s_tab62.value.replace(/,/g,""));
				s_tab63 = parseFloat(document.frm.s_tab63.value.replace(/,/g,""));
				s_tab70 = parseFloat(document.frm.s_tab70.value.replace(/,/g,""));
				s_tab71 = parseFloat(document.frm.s_tab71.value.replace(/,/g,""));
				s_tab72 = parseFloat(document.frm.s_tab72.value.replace(/,/g,""));
				s_tab73 = parseFloat(document.frm.s_tab73.value.replace(/,/g,""));
				s_tab80 = parseFloat(document.frm.s_tab80.value.replace(/,/g,""));
				s_tab81 = parseFloat(document.frm.s_tab81.value.replace(/,/g,""));
				s_tab82 = parseFloat(document.frm.s_tab82.value.replace(/,/g,""));
				s_tab83 = parseFloat(document.frm.s_tab83.value.replace(/,/g,""));
				s_tab90 = parseFloat(document.frm.s_tab90.value.replace(/,/g,""));
				s_tab91 = parseFloat(document.frm.s_tab91.value.replace(/,/g,""));
				s_tab92 = parseFloat(document.frm.s_tab92.value.replace(/,/g,""));
				s_tab93 = parseFloat(document.frm.s_tab93.value.replace(/,/g,""));
				s_tab100 = parseFloat(document.frm.s_tab100.value.replace(/,/g,""));
				s_tab101 = parseFloat(document.frm.s_tab101.value.replace(/,/g,""));
				s_tab102 = parseFloat(document.frm.s_tab102.value.replace(/,/g,""));
				s_tab103 = parseFloat(document.frm.s_tab103.value.replace(/,/g,""));
				s_tab110 = parseFloat(document.frm.s_tab110.value.replace(/,/g,""));
				s_tab111 = parseFloat(document.frm.s_tab111.value.replace(/,/g,""));
				s_tab112 = parseFloat(document.frm.s_tab112.value.replace(/,/g,""));
				s_tab113 = parseFloat(document.frm.s_tab113.value.replace(/,/g,""));
				s_tab120 = parseFloat(document.frm.s_tab120.value.replace(/,/g,""));
				s_tab121 = parseFloat(document.frm.s_tab121.value.replace(/,/g,""));
				s_tab122 = parseFloat(document.frm.s_tab122.value.replace(/,/g,""));
				s_tab123 = parseFloat(document.frm.s_tab123.value.replace(/,/g,""));
				sla_month = document.frm.sla_month.value;
				var mon_tab = new Array();
				sla_yy = sla_month.substr(2,2);
				sla_mm = sla_month.substr(4,2);
				mon_tab[12] = sla_mm;
				for (i = 11; i > 0; i--) {
					sla_mm = sla_mm - 1;
					if ( sla_mm == "00" ){
						sla_mm = "12";
						sla_yy = sla_yy - 1;
					}
					mon_tab[i] = sla_mm;
				}
				chart1 = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view1',
						defaultSeriesType: 'column'
					},
//					title: {
//						text: 'Monthly Average Rainfall'
//					},
//					subtitle: {
//						text: 'Source: WorldClimate.com'
//					},
					xAxis: {
						categories: [
							mon_tab[1], 
							mon_tab[2], 
							mon_tab[3], 
							mon_tab[4], 
							mon_tab[5], 
							mon_tab[6], 
							mon_tab[7], 
							mon_tab[8], 
							mon_tab[9], 
							mon_tab[10], 
							mon_tab[11], 
							mon_tab[12]
						]
					},
					yAxis: {
						min: 0,
						title: {
							text: '종합평가점수'
						}
					},
					legend: {
						layout: 'vertical',
						backgroundColor: '#FFFFFF',
						align: 'center',
						verticalAlign: 'top',
						x: 0,
						y: 0,
						floating: true,
						shadow: true
					},
					tooltip: {
						formatter: function() {
							return ''+
								this.x +'월: '+ this.y +' 점';
						}
					},
					plotOptions: {
						column: {
							pointPadding: 0.2,
							borderWidth: 0
						}
					},
				        series: [{
						name: '주전산기',
						color: '#4572A7',
						data: [s_tab10, s_tab20, s_tab30, s_tab40, s_tab50, s_tab60, s_tab70, s_tab80, s_tab90, s_tab100, s_tab110, s_tab120]
					}]
				});

				chart2 = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view2',
						defaultSeriesType: 'column'
					},
//					title: {
//						text: 'Monthly Average Rainfall'
//					},
//					subtitle: {
//						text: 'Source: WorldClimate.com'
//					},
					xAxis: {
						categories: [
							mon_tab[1], 
							mon_tab[2], 
							mon_tab[3], 
							mon_tab[4], 
							mon_tab[5], 
							mon_tab[6], 
							mon_tab[7], 
							mon_tab[8], 
							mon_tab[9], 
							mon_tab[10], 
							mon_tab[11], 
							mon_tab[12]
						]
					},
					yAxis: {
						min: 0,
						title: {
							text: '종합평가점수'
						}
					},
					legend: {
						layout: 'vertical',
						backgroundColor: '#FFFFFF',
						align: 'center',
						verticalAlign: 'top',
						x: 0,
						y: 0,
						floating: true,
						shadow: true
					},
					tooltip: {
						formatter: function() {
							return ''+
								this.x +'월: '+ this.y +' 점';
						}
					},
					plotOptions: {
						column: {
							pointPadding: 0.2,
							borderWidth: 0
						}
					},
				        series: [{
						name: '1지역',
						color: '#AA4643',
						data: [s_tab11, s_tab21, s_tab31, s_tab41, s_tab51, s_tab61, s_tab71, s_tab81, s_tab91, s_tab101, s_tab111, s_tab121]
					}]
				});
				chart3 = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view3',
						defaultSeriesType: 'column'
					},
//					title: {
//						text: 'Monthly Average Rainfall'
//					},
//					subtitle: {
//						text: 'Source: WorldClimate.com'
//					},
					xAxis: {
						categories: [
							mon_tab[1], 
							mon_tab[2], 
							mon_tab[3], 
							mon_tab[4], 
							mon_tab[5], 
							mon_tab[6], 
							mon_tab[7], 
							mon_tab[8], 
							mon_tab[9], 
							mon_tab[10], 
							mon_tab[11], 
							mon_tab[12]
						]
					},
					yAxis: {
						min: 0,
						title: {
							text: '종합평가점수'
						}
					},
					legend: {
						layout: 'vertical',
						backgroundColor: '#FFFFFF',
						align: 'center',
						verticalAlign: 'top',
						x: 0,
						y: 0,
						floating: true,
						shadow: true
					},
					tooltip: {
						formatter: function() {
							return ''+
								this.x +'월: '+ this.y +' 점';
						}
					},
					plotOptions: {
						column: {
							pointPadding: 0.2,
							borderWidth: 0
						}
					},
				        series: [{
						name: '2지역',
						color: '#89A54E',
						data: [s_tab12, s_tab22, s_tab32, s_tab42, s_tab52, s_tab62, s_tab72, s_tab82, s_tab92, s_tab102, s_tab112, s_tab122]
					}]
				});
				chart4 = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view4',
						defaultSeriesType: 'column'
					},
//					title: {
//						text: 'Monthly Average Rainfall'
//					},
//					subtitle: {
//						text: 'Source: WorldClimate.com'
//					},
					xAxis: {
						categories: [
							mon_tab[1], 
							mon_tab[2], 
							mon_tab[3], 
							mon_tab[4], 
							mon_tab[5], 
							mon_tab[6], 
							mon_tab[7], 
							mon_tab[8], 
							mon_tab[9], 
							mon_tab[10], 
							mon_tab[11], 
							mon_tab[12]
						]
					},
					yAxis: {
						min: 0,
						title: {
							text: '종합평가점수'
						}
					},
					legend: {
						layout: 'vertical',
						backgroundColor: '#FFFFFF',
						align: 'center',
						verticalAlign: 'top',
						x: 0,
						y: 0,
						floating: true,
						shadow: true
					},
					tooltip: {
						formatter: function() {
							return ''+
								this.x +'월: '+ this.y +' 점';
						}
					},
					plotOptions: {
						column: {
							pointPadding: 0.2,
							borderWidth: 0
						}
					},
				        series: [{
						name: '3지역',
						color: '#80699B',
						data: [s_tab13, s_tab23, s_tab33, s_tab43, s_tab53, s_tab63, s_tab73, s_tab83, s_tab93, s_tab103, s_tab113, s_tab123]
					}]
				});
				
				
			});

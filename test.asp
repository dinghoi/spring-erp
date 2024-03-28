<!doctype html>
	<html lang="en">
		<head>
		  <meta charset="utf-8">
		  <title>jQuery UI Tabs - Content via Ajax</title>  
		  <link rel="stylesheet" href="//code.jquery.com/ui/1.10.4/themes/smoothness/jquery-ui.css">  
		  <script src="//code.jquery.com/jquery-1.10.2.js"></script>  
		  <script src="//code.jquery.com/ui/1.10.4/jquery-ui.js"></script>  
		  <link rel="stylesheet" href="/resources/demos/style.css">  
		  <script>
		  $(function() {
				$( "#tabs" ).tabs({
				 beforeLoad: function( event, ui ) {
					ui.jqXHR.error(function() {
						ui.panel.html(
	            "Couldn't load this tab. We'll try to fix this as soon as possible. " +
	            "If this wouldn't be a demo." );
						  });
					  }
				 });
			});
			</script>
		</head>
	<body>
	
	<div id="tabs">
	  <ul>
	     <li><a href="#tabs-1">Preloaded</a></li>
	     <li><a href="ajax/content1.html">Tab 1</a></li>
	     <li><a href="ajax/content2.html">Tab 2</a></li>
	  </ul>
		<div id="tabs-1">
			<p>aaaaa</p>
		</div>
  </div>
 </body>
 </html>
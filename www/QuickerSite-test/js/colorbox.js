$(document).ready(function(){
	//Examples of how to assign the ColorBox event to elements
	$("a[rel='example1']").colorbox();
	$("a[rel='example2']").colorbox({transition:"fade"});
	$("a[rel='example3']").colorbox({transition:"none", width:"75%", height:"75%"});
	$("a[rel='example4']").colorbox({slideshow:true});
	$(".example5").colorbox();
	$(".example5IF").colorbox({iframe:true, innerWidth:720, innerHeight:480});
	
	$(".example6").colorbox({iframe:true, innerWidth:425, innerHeight:344});
	$(".example7").colorbox({width:"80%", height:"80%", iframe:true});
	$(".example8").colorbox({width:"50%", inline:true, href:"#inline_example1"});
	$(".example9").colorbox({
		onOpen:function(){ alert('onOpen: colorbox is about to open'); },
		onLoad:function(){ alert('onLoad: colorbox has started to load the targeted content'); },
		onComplete:function(){ alert('onComplete: colorbox has displayed the loaded content'); },
		onCleanup:function(){ alert('onCleanup: colorbox has begun the close process'); },
		onClosed:function(){ alert('onClosed: colorbox has completely closed'); }
	});
	
	//Example of preserving a JavaScript event for inline calls.
	$("#click").click(function(){ 
		$('#click').css({"background-color":"#f00", "color":"#fff", "cursor":"inherit"}).text("Open this window again and this message will still be here.");
		return false;
	});
});
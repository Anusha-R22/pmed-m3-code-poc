<%if (Application("asDEV") <> "true") then%>
	<script language="javascript">
	document.oncontextmenu=fnContextMenu;
	function fnContextMenu(){return false};
	function fnKeyDown()
	{
		if ((event.keyCode<112)||(event.keyCode>123)) return;
		event.returnValue=false;
		event.keyCode=0;
	}
	document.onkeydown=fnKeyDown;
	</script>
<%end if%>
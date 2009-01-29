///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2006
//
//	BufferDataBrowser event handlers.
//	fnM(...) pops context menu
//	fnMClick(...) handles context menu events
//
///////////////////////////////////////////////////////////////////////////////////////////

var gsItem=""; //global clicked item's delimited list
var gsName="";
var gsValue="";
var gsBack="";

function fnM(nBtn,sItem,sScope,sBack,sName,sValue,c,r,ts,tc,go)
{
	gsItem=sItem;
	gsName=sName;
	gsValue=sValue;
	gsBack=sBack;
	var sHtml;

	if (nBtn==1)
	{
		//left mouse button handler
	}
	else
	{
		sHtml="<table width='130' class='clsPopMenuLinkText'>";
		//Commit
		sHtml+=(c)? "<tr height='15'><td><a href='javascript:fnMClick(1,\""+sScope+"\");'>Save Data</a></td></tr>":"<tr height='15'><td>Save Data</td></tr>"
		//Reject
		sHtml+=(r)? "<tr height='15'><td><a href='javascript:fnMClick(2,\""+sScope+"\");'>Discard Data</a></td></tr>":"<tr height='15'><td>Discard Data</td></tr>"
		//separator
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>"
		//Target select
		sHtml+=(ts)? "<tr height='15'><td><a href='javascript:fnMClick(3,\""+sScope+"\");'>Select Target</a></td></tr>":"<tr height='15'><td>Select Target</td></tr>"
		//Target change
		sHtml+=(tc)? "<tr height='15'><td><a href='javascript:fnMClick(3,\""+sScope+"\");'>Change Target</a></td></tr>":"<tr height='15'><td>Change Target</td></tr>"
		//separator
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>"
		//goto eform
		sHtml+=(go)? "<tr height='15'><td><a href='javascript:fnMClick(4);'>Go to eForm</a></td></tr>":"<tr height='15'><td>Go to eForm</td></tr>";
		sHtml+="</table>"
    
		document.all["divPopMenu"].innerHTML=sHtml;
		fnPopMenuShow(document.all['divPopMenu']);
	}
}
function fnMClick(sType,sScope)
{
	var sURL;
	var bSubmit=true;
	fnPopMenuHide();
	switch(sType)
	{
		case 4:
			// go to eform
			var sUrl;
			var aItm = gsItem.split("`");
            window.parent.fnEformUrl(aItm[0],aItm[1],aItm[2],aItm[7],"",undefined,window.sWinState);
            return;
			break;
		case 1:
			// save data
			if(confirm("Are you sure you wish to save the data?")==true)
			{
				// save
				var sUrl="BufferDataSaveResults.asp";
				//var sUrl="BufferDataSaveResults.aspx";
				// store to action
				document.FormDR.action = sUrl;
			}
			else
			{
				bSubmit=false;
			}
			break;
		case 2:
			// reject data
			if(confirm("Are you sure you wish to reject the data?")==true)
			{
				// save
				var sUrl="BufferDataSaveResults.asp";
				// store to action
				document.FormDR.action = sUrl;
			}
			else
			{
				bSubmit=false;
			}
			break;
		case 3:
			// select / change target
			var sUrl="BufferTargetSelection.asp";
			// store to action
			document.FormDR.action = sUrl;
			break;
		case 5:
			// select target data item
			var sUrl="BufferTargetSelectionSave.asp";
			// store to action
			document.FormDR.action = sUrl;
			break;
		default:
			return;
	}
	if(bSubmit)
	{
		document.FormDR.bback.value=gsBack;
		document.FormDR.bidentifier.value=gsItem;
		document.FormDR.btype.value=sType;
		document.FormDR.bscope.value=sScope;
		document.FormDR.submit();
    }
}
function fnSel(nBtn,sItem,sBack,sName,sValue,s,go)
{
	gsItem=sItem;
	gsBack=sBack;
	gsName=sName;
	gsValue=sValue;
	var sHtml;

	if (nBtn==1)
	{
		//left mouse button handler
	}
	else
	{
		sHtml="<table width='130' class='clsPopMenuLinkText'>";
		//Commit
		sHtml+=(s)? "<tr height='15'><td><a href='javascript:fnMClick(5,0);'>Select Data Item</a></td></tr>":"<tr height='15'><td>Select Data Item</td></tr>"
		//separator
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>"
		//goto eform
		sHtml+=(go)? "<tr height='15'><td><a href='javascript:fnMClick(4);'>Go to eForm</a></td></tr>":"<tr height='15'><td>Go to eForm</td></tr>";
		sHtml+="</table>"
    
		document.all["divPopMenu"].innerHTML=sHtml;
		fnPopMenuShow(document.all['divPopMenu']);
	}
}
// go to new buffer target page
function fnTargetNewPage(sBookmark)
{
	var sUrl = "BufferTargetSelection.asp?bookmark=" + sBookmark;
	// set target url for form
	document.FormDR.action = sUrl;
	// submit
	document.FormDR.submit();
}
function fnHideLoader()
{
	document.all["divMsgBox"].style.visibility='hidden';
}       
function fnPopMenuShow(oDiv)
{
	var nPopHeight=oDiv.clientHeight;
	var nPopWidth=oDiv.clientWidth;
	var nPopYPosition=window.event.clientY;
	var nPopXPosition=window.event.clientX;
	var nVisibleScreenHeight=document.body.clientHeight;
	var nVisibleScreenWidth=document.body.clientWidth;
	var nX=0;
	var nY=0;
	
	if((nVisibleScreenHeight-(nPopYPosition+nPopHeight)>0)||((nVisibleScreenHeight-(nPopYPosition+nPopHeight))>(nPopYPosition-nPopHeight)))
	{
		//if space below, display below
		nY=nPopYPosition;
	}
	else
	{
		//display above
		nY=nPopYPosition-nPopHeight;
	}
	if((nVisibleScreenWidth-(nPopXPosition+nPopWidth)>0)||((nVisibleScreenWidth-(nPopXPosition+nPopWidth))>(nPopXPosition-nPopWidth)))
	{
		//if space right, display right
		nX=nPopXPosition;
	}
	else
	{
		//display left
		nX=nPopXPosition-nPopWidth;
	}
    oDiv.style.pixelLeft=document.body.scrollLeft+nX;
    oDiv.style.pixelTop=document.body.scrollTop+nY;
    oDiv.style.visibility='visible';
}
function fnPopMenuHide()
{
	document.all["divPopMenu"].style.visibility='hidden'
}
document.onclick=fnPopMenuHide;
// stop right click
//document.oncontextmenu=fnContextMenu;
//function fnContextMenu(){return false};
//function fnKeyDown()
//{
//	if ((event.keyCode<112)||(event.keyCode>123)) return;
//	event.returnValue=false;
//	event.keyCode=0;
//}
//document.onkeydown=fnKeyDown;

///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2003
//
//	DataBrowser event handlers.
//	fnM(...) pops context menu
//	fnMClick(...) handles context menu events
//
///////////////////////////////////////////////////////////////////////////////////////////

var gsItem=""; //global clicked item's delimited list
var gsName="";
var gsValue="";

function fnM(nBtn,sItem,sMIScope,sLFScope,sName,sValue,l,u,f,z,go,d,s,n)
{
	gsItem=sItem;
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
		//lock
		sHtml+=(l)? "<tr height='15'><td><a href='javascript:fnMClick(3,\""+sLFScope+"\");'>Lock Item</a></td></tr>":"<tr height='15'><td>Lock Item</td></tr>"
		//unlock
		sHtml+=(u)? "<tr height='15'><td><a href='javascript:fnMClick(4,\""+sLFScope+"\");'>Unlock Item</a></td></tr>":"<tr height='15'><td>Unlock Item</td></tr>"
		//freeze
		sHtml+=(f)? "<tr height='15'><td><a href='javascript:fnMClick(5,\""+sLFScope+"\");'>Freeze Item</a></td></tr>":"<tr height='15'><td>Freeze Item</td></tr>"
		//unfreeze
		sHtml+=(z)? "<tr height='15'><td><a href='javascript:fnMClick(6,\""+sLFScope+"\");'>UnFreeze Item</a></td></tr>":"<tr height='15'><td>UnFreeze Item</td></tr>"
		//goto eform
		sHtml+=(go)? "<tr height='15'><td><a href='javascript:fnMClick(7);'>Go to eForm</a></td></tr>":"<tr height='15'><td>Go to eForm</td></tr>";
		//separator
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>"
		//discrepancy
		sHtml+=(d)? "<tr height='15'><td><a href='javascript:fnMClick(0,\""+sMIScope+"\");'>Add Discrepancy</a></td></tr>":"<tr height='15'><td>Add Discrepancy</td></tr>"
		//sdv
		sHtml+=(s)? "<tr height='15'><td><a href='javascript:fnMClick(1,\""+sMIScope+"\");'>Add SDV Mark</a></td></tr>":"<tr height='15'><td>Add SDV Mark</td></tr>"
		//note
		sHtml+=(n)? "<tr height='15'><td><a href='javascript:fnMClick(2,\""+sMIScope+"\");'>Add Note</a></td></tr>":"<tr height='15'><td>Add Note</td></tr>"
		sHtml+="</table>"
    
		document.all["divPopMenu"].innerHTML=sHtml;
		fnPopMenuShow(document.all['divPopMenu']);
		//document.all["divPopMenu"].style.pixelLeft=document.body.scrollLeft+window.event.clientX;
		//document.all["divPopMenu"].style.pixelTop=document.body.scrollTop+window.event.clientY;
		//document.all["divPopMenu"].style.visibility='visible';
	}
}
function fnMClick(sType,sScope)
{
	var sURL;
	fnPopMenuHide();
	switch(sType)
	{
		case 3:
		case 4:
		case 5:
		case 6:
			/* no confirm dialog required for freezing to match windows
			if(sType==5)
			{
				if (!confirm('This will permanently freeze the selected item(s)')) return;
			}
			*/
			break;
		case 7:
			var sUrl;
			var aItm = gsItem.split("`");
            window.parent.fnEformUrl(aItm[0],aItm[1],aItm[2],aItm[8],"",undefined,window.sWinState);
            return;
			break;
		case 0:
		case 1:
		case 2:
			var aArgs=new Array();
			//action,question caption,mimessage text
			aArgs[0]=gsName;aArgs[1]=gsValue;
			switch (sType)
			{
				case 0:sURL="DialogNewDiscrepancy.htm";break;
				case 1:sURL="DialogNewSDV.htm";break;
				default:sURL="DialogNewNote.htm";break;
			}		
			sRtn=window.showModalDialog(sURL,aArgs,'dialogHeight:250px;dialogWidth:500px;center:yes;status:0;dependent,scrollbars');
			if ((sRtn!=undefined)&&(sRtn!=""))
			{
				document.FormDR.badd.value=sRtn;
			}
			else
			{	
				return;
			}
			break;
		default:
			return;
	}
	document.FormDR.bidentifier.value=gsItem;
    document.FormDR.btype.value=sType;
    document.FormDR.bscope.value=sScope;
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
///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2003
//
//	top menu event handlers
//
///////////////////////////////////////////////////////////////////////////////////////////

var bHidden=false;

function fnLogOutUrl()
{
	window.parent.window.frames[4].fnSaveDataFirst('fnLogOutUrl()');
}
function fnHoldUrl()
{
	window.parent.window.frames[4].fnSaveDataFirst('fnHoldUrl("'+top.fnGetAppUrl(true)+'&act=standby")');
}
function fnSwitchUrl()
{
	window.parent.window.frames[4].fnSaveDataFirst('fnSwitchUrl("'+top.fnGetAppUrl(false)+'&act=switch")');
}
function fnPageLoaded()
{
	fnInitialiseButton('topMenu','clsHoverButton clsHoverButtonActive','clsHoverButton clsHoverButtonInactive','',false);
}
function fnOnMouseOut(oImg)
{
	if(bHidden)
	{
		oImg.src='../img/exp_horiz_inactive.gif'
	}
    else
    {
		oImg.src='../img/col_horiz_inactive.gif'
	}
}
function fnOnClick(oImg)
{
	if(bHidden)
	{
//<!--r-->
		window.parent.document.body.all['f2'].cols='230,*';
//<!--r-->		
		oImg.src='../img/col_horiz_active.gif';
		oImg.alt='Hide menu';
		bHidden=false
	}
	else
	{
//<!--r-->
		window.parent.document.body.all['f2'].cols='0,*';
//<!--r-->	
		oImg.src='../img/exp_horiz_active.gif';
        oImg.alt='Show menu';
        bHidden=true;
    }
}
function fnOnMouseOver(oImg)
{
	if(bHidden)
	{
		oImg.src='../img/exp_horiz_active.gif'
	}
	else
	{
		oImg.src='../img/col_horiz_active.gif'
	}
}
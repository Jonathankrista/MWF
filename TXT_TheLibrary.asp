<SCRIPT RUNAT=Server Language="VBSCRIPT">
Function ToSQLDate( dt ) 
	a="" & dt
	a=LCase(a)
	If (a="null") Then
		ToSQLDate = "NULL" 
	ElseIf IsDate(dt) Then 
		dt = CDate(dt) 
		ToSQLDate = "'" & dt & "'" 
	Else 
		ToSQLDate = "NULL" 
	End If 
End Function 

Function ToSQLInt( num ) 
	a="" & num
	a=LCase(a)
	If (a="null") Then
		ToSQLInt = "NULL" 
	ElseIf IsNumeric( num ) Then 
		ToSQLInt=DoubleToString(num,1,0)
	Else 
		ToSQLInt = "0" 
	End If 
End Function 

Function GetNothing()
    Set GetNothing = Nothing
End Function

Function DoFileExists(FileName)
on error resume next
	set o=Server.CreateObject("Scripting.FileSystemObject")

	if o.FileExists(FileName) then
		DoFileExists=true
	else
		DoFileExists=false
	end if
End Function

Function DoRenameFile(FFileName,TFileName)
on error resume next
	set o=Server.CreateObject("Scripting.FileSystemObject")

	if o.FileExists(FFileName) then
		o.MoveFile FFileName,TFileName
	end if

	DoRenameFile=""
End Function

Function DoDeleteFile(FileName)
on error resume next
	set o=Server.CreateObject("Scripting.FileSystemObject")

	if o.FileExists(FileName) then
		o.DeleteFile FileName
	end if

	DoDeleteFile=""
End Function

Function DoDeleteAllFileInFolder(FolderName)
on error resume next
	set o=Server.CreateObject("Scripting.FileSystemObject")
	Set ObjFolder = o.GetFolder(FolderName)
	Set ObjFiles = ObjFolder.Files
	For Each ObjFile In ObjFiles
		o.DeleteFile ObjFile.Path
	Next
	DoDeleteAllFileInFolder=""
End Function

Function DoDeleteAllRemainOneFileInFolder(FolderName)
on error resume next
	set o=Server.CreateObject("Scripting.FileSystemObject")
	Set ObjFolder = o.GetFolder(FolderName)
	Set ObjFiles = ObjFolder.Files
	old=""
	For Each ObjFile In ObjFiles
		if (old<>"") then
			o.DeleteFile old
		end if
		old=ObjFile.Path
	Next
	DoDeleteAllRemainOneFileInFolder=""
End Function
Function UrlDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       UrlDecode = ""
       Exit Function
    End If

    If (sConvert="") Then
       UrlDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    UrlDecode = sOutput
End Function</SCRIPT>
<SCRIPT RUNAT=Server Language="JSCRIPT">
function GetFilesFromFolder(RootFolder)
{
	var FileName=new Array();
	var objFSO = Server.CreateObject("Scripting.FileSystemObject");
	if (objFSO.FolderExists(RootFolder))
	{
		var objFolder = objFSO.GetFolder(RootFolder);
		var fc = new Enumerator(objFolder.Files);
		for (var i=0; !fc.atEnd();i++)
		{
			FileName[i] =""+ fc.item();
			fc.moveNext();
		}
		objFolder=null;
	}
	objFSO=null;
	return FileName;
}
function ToSQL(In,UseSingleQ)
{
	var Input=""+In;
	var s=""+UseSingleQ;
	if (s!="0")
		s="1";
	return MyToSQL(Input,s); 
}
function ToSQLDbl(num,DP) 
{
	var Input=""+num;
	var DecPlace="" + DP;
	if ((DecPlace!="0")&&(DecPlace!="1")&&(DecPlace!="3")&&(DecPlace!="4")&&(DecPlace!="5")&&(DecPlace!="6"))
		DecPlace="2";
	DecPlace=MyparseInt(DecPlace);
	return MyToSQLDbl(Input,DecPlace);
}
function CheckFrenchPriceFormat(xInString)
{
	var InString=""+xInString;
	var T=""+SysLang;
	if (T=="F")
	{
		var Tmp="";
		for (var i=0; i < InString.length;i++)
		{
			var C=""+InString.substring(i,i+1);
			if (((C >= '0') && (C <= '9'))  || (C=='.') || (C=='-'))
				Tmp+=C;
		}
		Tmp=ReplaceAllNow(Tmp,".",",");
		if (InString.indexOf("$",0)>=0)
			Tmp+=" $";
		return Tmp;
	}
	else
		return ""+InString;
}
function RedirectByCountry(Country,CheckFile)
{
	var Tmp=""+Request.ServerVariables("HTTPS");
	Tmp=Tmp.toLowerCase();
	if (Tmp=="on")
		Tmp="https://"; 
	else
		Tmp="http://";
	Tmp+=(""+Request.ServerVariables("SERVER_NAME")).toLowerCase();
	Tmp+="/"+Country;
	if (CheckFile!="default.asp")
	{
		var TTmp="";
		var Tmp1=""+CheckFile.toLowerCase();
		if (Tmp1.indexOf("paypal") >= 0)
			TTmp+="/" + CheckFile;
		else
			TTmp+="/" + CheckFile.toLowerCase();
		if ((""+CheckQuery!="")&&(""+CheckQuery!="undefined"))
		{
			if (Tmp1.indexOf("paypal") >= 0)
				TTmp+="?" + CheckQuery;
			else
				TTmp+="?" + CheckQuery.toLowerCase();
		}
		if (TTmp!="/")
			Tmp+=TTmp;
		while (Tmp.indexOf("%20",0)>=0)
			Tmp=ReplaceNow(Tmp,"%20","+");
	}
	Response.Status = "301 Moved Permanently";
	Response.AddHeader("Location",Tmp);
	Response.end;
}
</SCRIPT>
<SCRIPT RUNAT=Server Language="VBSCRIPT">
Function MyToSQL( txt,UseSingleQ ) 
	if (UseSingleQ = "0") then
		MyToSQL = """" & Replace( txt, """", """""" ) & """" 
	else
		MyToSQL = "'" & Replace( txt, "'", "''" ) & "'" 
	end if
End Function 

Function MyToSQLDbl(num,DP) 
	a="" & num
	a=LCase(a)
	If (a="null") Then
		MyToSQLDbl = "NULL" 
	elseIf IsNumeric( num ) Then 
		MyToSQLDbl=DoubleToString(num,1,DP)&"."&DoubleToString(num,0,DP)
	Else 
		MyToSQLDbl = "0" 
	End If 
End Function 
</SCRIPT>



<SCRIPT RUNAT=Server Language="JSCRIPT">
/*function MatchSysLangAndURL(SysLang)
{
	var CheckQuery="";
	var CheckURL="";
	if (""+Request.ServerVariables("QUERY_STRING")=="")
		CheckURL=""+Request.ServerVariables("SERVER_NAME")+ Request.ServerVariables("URL");
	else if ((""+Request.ServerVariables("QUERY_STRING")).indexOf("404")==0)
	{
		CheckURL=""+Request.ServerVariables("SERVER_NAME");
		var TPos=(""+Request.ServerVariables("QUERY_STRING")).substring(12,(""+Request.ServerVariables("QUERY_STRING")).length).indexOf("/")+12;
		CheckURL+=(""+Request.ServerVariables("QUERY_STRING")).substring(TPos,(""+Request.ServerVariables("QUERY_STRING")).length);
	}
	else
		CheckURL=""+Request.ServerVariables("SERVER_NAME")+ Request.ServerVariables("URL")+ "?" + Request.ServerVariables("QUERY_STRING");
	while (CheckURL!=CheckURL.replace(new RegExp("//", "g"), "/"))
		CheckURL=CheckURL.replace(new RegExp("//", "g"), "/");

	var NeedRedirect=false;
	if (""+SysLang=="F")
	{
		if (CheckURL.indexOf("www.",0) == 0)
		{
			CheckURL=ReplaceNow(CheckURL,"www.","fr.");
			NeedRedirect=true;
		}
		else if (CheckURL.indexOf("softmoc.com",0) == 0)
		{
			CheckURL=ReplaceNow(CheckURL,"softmoc.com","fr.softmoc.com");
			NeedRedirect=true;
		}
		else if (CheckURL.indexOf("new.",0) == 0)
		{
			CheckURL=ReplaceNow(CheckURL,"new.","newfr.");
			NeedRedirect=true;
		}
		else if ((CheckURL.indexOf("test",0) == 0)&&(CheckURL.indexOf("testf",0) != 0))
		{
			CheckURL=ReplaceNow(CheckURL,"test","testf");
			NeedRedirect=true;
		}
	}
	else
	{
		if (CheckURL.indexOf("fr.",0) == 0)
		{
			CheckURL=ReplaceNow(CheckURL,"fr.","www.");
			NeedRedirect=true;
		}
		else if (CheckURL.indexOf("newfr.",0) == 0)
		{
			CheckURL=ReplaceNow(CheckURL,"newfr.","new.");
			NeedRedirect=true;
		}
		else if (CheckURL.indexOf("testf",0) == 0)
		{
			CheckURL=ReplaceNow(CheckURL,"testf","test");
			NeedRedirect=true;
		}
	}
	if (NeedRedirect)
	{
		var TmpHTTPS=""+Request.ServerVariables("HTTPS");
		TmpHTTPS=TmpHTTPS.toLowerCase();
		if (TmpHTTPS=="on")
			TmpHTTPS="https://"; 
		else
			TmpHTTPS="http://";
		Response.Redirect(TmpHTTPS+CheckURL);
		Response.end;
	}
}
*/
function GetURLInputArray(CheckURL)
{
	var Tmp=CheckURL.toLowerCase();
	if (Tmp.indexOf("https://",0)>=0)
		Tmp=CheckURL.substring(Tmp.indexOf("https://",0)+8,CheckURL.length);
	else if (Tmp.indexOf("http://",0)>=0)
		Tmp=CheckURL.substring(Tmp.indexOf("http://",0)+7,CheckURL.length);
	var URLInputArray=Tmp.split("/");

	return URLInputArray;
}
/*function Opus_HTTP_call(URL,XnvpStr)
{
	var objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1");
	var nvpStr="<?xml version='1.0' encoding='ISO-8859-1' ?>\r\n<request>\r\n"+XnvpStr+"</request>";
    Session("Opus_Request")= nvpStr;
    Session("Opus_nvpReqArray") = XML_deformatNVP(nvpStr);
	try
	{
	    objHttp.Open("POST", URL, false);
	    WinHttpRequestOption_SslErrorIgnoreFlags = 4;
	    objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 0x3300;
	    objHttp.Send(nvpStr);
	}
	catch (Err)
	{
	    Session("Opus_ErrorMessage") = "Opus_HTTP_call() Exception calling Opus (" + URL + "): Message=" + Err.message + ", Description=" + Err.description;
	    Session("Opus_nvpReqArray") = null;
	    Response.Redirect("Opus_APIError.asp");
	}
	Session("Opus_Response") = objHttp.ResponseText;
	var nvpResponseCollection = XML_deformatNVP(objHttp.ResponseText);
	if (nvpResponseCollection("responsecode") == null)
	{
	    Session("Opus_ErrorMessage") = "Opus_HTTP_call() invalid response from calling Opus (" + URL + "): " + objHttp.ResponseText;
	    Response.Redirect("Opus_APIError.asp");
	}
	return (nvpResponseCollection);
}
function XML_deformatNVP (nvpstr)
{
	var xml = Server.CreateObject("Microsoft.XMLDOM");
	xml.async = false;
	xml.loadXML(nvpstr);
	var NvpCollection = Server.CreateObject("Scripting.Dictionary");
	if (xml.parseError.errorCode != 0)
	{
		var Val="Invalid XML file ("+xml.parseError.errorCode+")";
    	NvpCollection.Add(unescape("XML_Error"),unescape(Val));
		return NvpCollection;
	}
   	var node = xml.documentElement.firstChild.parentNode;
    while (node != null)
	{
        var node2 = node.firstChild;
        while (node2 != null)
		{
			var val = node2.firstChild;
            while (val != null)
			{
				NvpCollection.Add(unescape(""+node2.nodeName),unescape(""+val.nodeValue));
                val = val.nextSibling;
            }
            node2 = node2.nextSibling;
        }
        node = node.nextSibling;
    }
	return NvpCollection;
}
function XML_AddParam(name, value) {
    if (value == null || value == "") return ("");
	return ("<" + name + ">" + value + "</" + name + ">\r\n");
}
function GoCCPreAuth(Currency,order_id,FirstName,LastName,PaymentMethod,pan,exp_date_M,exp_date_Y,cvd_value,BillAddress,BillPostal,BillCity,BillState,BillCountry,
						BillPhone,ShipAddress,ShipPostal,ShipCity,ShipState,ShipCountry,ShipPhone,Email,amount)
{
	if (Currency=="US$")
		return GoCCPreAuthORPurchase_Opus(order_id,FirstName+" "+LastName,PaymentMethod,pan,exp_date_M,exp_date_Y,cvd_value,BillAddress,BillPostal,BillCity,BillState,BillCountry,
						BillPhone,ShipAddress,ShipPostal,ShipCity,ShipState,ShipCountry,ShipPhone,Email,amount,4);
	else
		return GoCCPreAuth_Moneris(order_id,pan,exp_date_Y+exp_date_M,cvd_value,GetAddressNo(BillAddress),GetAddress(BillAddress),BillPostal,amount,false);
}
function GoCCPreAuthORPurchase_Opus(order_id,CCName,PaymentMethod,pan,exp_date_M,exp_date_Y,cvd_value,BillAddress,BillPostal,BillCity,BillState,BillCountry,BillPhone,
				ShipAddress,ShipPostal,ShipCity,ShipState,ShipCountry,ShipPhone,Email,amount,action)
{
	var nvpstr=XML_AddParam("merchantid",GetOpus_MerchantID());
	nvpstr+=XML_AddParam("password",GetOpus_Password());

	//1=Purchase, 2=Refund, 3=Void Purchase, 4=Authorization, 5=Capture, 7=Void Capture, 9=Void Authorization
	nvpstr+=XML_AddParam("action",""+action);

	nvpstr+=XML_AddParam("bill_currencycode","USD");
	nvpstr+=XML_AddParam("trackid",order_id);
	nvpstr+=XML_AddParam("bill_cardholder",CCName);
	nvpstr+=XML_AddParam("bill_cc_type","CC");

	//AM=American Express, CB=Carte Blanche, DC=Diners Club, DI=Discover, FP=FirePay, JC=JCB, LA=Laser, MC=MasterCard, MD=Maestro, N=Novus,
	//SO=Solo, SW=Switch, VD=Visa Delta, VE=Visa Electron, VC = Visa
	var CCBrand="MC";
	var Tmp=PaymentMethod.toLowerCase();
	if (Tmp.indexOf("amex",0)>=0)
		CCBrand="AM";
	else if (Tmp.indexOf("american",0)>=0)
		CCBrand="AM";
	else if (Tmp.indexOf("carte",0)>=0)
		CCBrand="CB";
	else if (Tmp.indexOf("diner",0)>=0)
		CCBrand="DC";
	else if (Tmp.indexOf("discover",0)>=0)
		CCBrand="DI";
	else if (Tmp.indexOf("fire",0)>=0)
		CCBrand="FP";
	else if (Tmp.indexOf("jcb",0)>=0)
		CCBrand="JC";
	else if (Tmp.indexOf("laser",0)>=0)
		CCBrand="LA";
	else if (Tmp.indexOf("maestro",0)>=0)
		CCBrand="MD";
	else if (Tmp.indexOf("novus",0)>=0)
		CCBrand="N";
	else if (Tmp.indexOf("solo",0)>=0)
		CCBrand="SO";
	else if (Tmp.indexOf("switch",0)>=0)
		CCBrand="SW";
	else if ((Tmp.indexOf("visa",0)>=0)&&(Tmp.indexOf("delta",0)>=0))
		CCBrand="VD";
	else if ((Tmp.indexOf("visa",0)>=0)&&(Tmp.indexOf("electron",0)>=0))
		CCBrand="VE";
	else if (Tmp.indexOf("visa",0)>=0)
		CCBrand="VC";
	else
		CCBrand="MC";
	nvpstr+=XML_AddParam("bill_cc_brand",CCBrand);

	nvpstr+=XML_AddParam("bill_cc",pan);
	nvpstr+=XML_AddParam("bill_expmonth",exp_date_M);
	nvpstr+=XML_AddParam("bill_expyear",exp_date_Y);
	nvpstr+=XML_AddParam("bill_cvv2",cvd_value);
	nvpstr+=XML_AddParam("bill_address",BillAddress);
	nvpstr+=XML_AddParam("bill_postal",BillPostal);
	nvpstr+=XML_AddParam("bill_city",BillCity);
	nvpstr+=XML_AddParam("bill_state",BillState);
	nvpstr+=XML_AddParam("bill_email",Email);
	nvpstr+=XML_AddParam("bill_country",BillCountry);
	nvpstr+=XML_AddParam("bill_amount",amount);
	nvpstr+=XML_AddParam("bill_phone",BillPhone);
	nvpstr+=XML_AddParam("ship_address",ShipAddress);
	nvpstr+=XML_AddParam("ship_email",Email);
	nvpstr+=XML_AddParam("ship_postal",ShipPostal);
	nvpstr+=XML_AddParam("ship_city",ShipCity);
	nvpstr+=XML_AddParam("ship_state",ShipState);
	nvpstr+=XML_AddParam("ship_phone",ShipPhone);
	nvpstr+=XML_AddParam("ship_country",ShipCountry);
//	nvpstr+=XML_AddParam("account_identifier","");nvpstr+=XML_AddParam("bill_address2","");nvpstr+=XML_AddParam("bill_customerip","");
//	nvpstr+=XML_AddParam("bill_merchantip","");nvpstr+=XML_AddParam("bill_fax","");nvpstr+=XML_AddParam("ship_address2","");
//	nvpstr+=XML_AddParam("ship_type","");nvpstr+=XML_AddParam("ship_fax","");nvpstr+=XML_AddParam("udf1","");
//	nvpstr+=XML_AddParam("udf2","");nvpstr+=XML_AddParam("udf3","");nvpstr+=XML_AddParam("udf4","");
//	nvpstr+=XML_AddParam("udf5","");nvpstr+=XML_AddParam("merchantcustomerid","");nvpstr+=XML_AddParam("product_desc","");
//	nvpstr+=XML_AddParam("product_quantity","");nvpstr+=XML_AddParam("product_unitcost","");nvpstr+=XML_AddParam("xid","");
//	nvpstr+=XML_AddParam("ecivalue","");nvpstr+=XML_AddParam("cavv","");
	var resArray=Opus_HTTP_call(GetOpus_URL(),nvpstr);

//Output Format+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//var Tmp="result: "+resArray("result")+"<br>";
//Tmp+="responsecode: "+resArray("responsecode")+"<br>";
//Tmp+="authcode: "+resArray("authcode")+"<br>";
//Tmp+="tranid: "+resArray("tranid")+"<br>";
//Tmp+="trackid: "+resArray("trackid")+"<br>";
//Tmp+="merchantid: "+resArray("merchantid")+"<br>";
//Tmp+="udf1: "+resArray("udf1")+"<br>";
//Tmp+="udf2: "+resArray("udf2")+"<br>";
//Tmp+="udf3: "+resArray("udf3")+"<br>";
//Tmp+="udf4: "+resArray("udf4")+"<br>";
//Output Format----------------------------------------------------------------------

	var CCErrMsg=""
	if ((""+resArray("responsecode")!="0")&&(""+resArray("responsecode")!="11")&&(""+resArray("responsecode")!="77"))
	{
		if (""+SysLang=="F")
			CCErrMsg="Désolé! Il y a un problème avec votre carte de crédit, veuillez corriger les informations de votre carte de crédit.";
		else
			CCErrMsg="Sorry! There's a problem with your credit card, please correct your credit card information. ("+resArray("responsecode")+")";
	}
	var OutString= "@CCErrMsg+@" + CCErrMsg + "@CCErrMsg-@<br>";
	OutString+= "@TransID+@" + resArray("tranid") + "@TransID-@<br>";
	OutString+= "@AuthCode+@" + resArray("authcode") + "@AuthCode-@<br>";						//char8
	OutString+= "@AVSResultCode+@" + "" + "@AVSResultCode-@<br>";
	OutString+= "@CVDResultCode+@" + "" + "@CVDResultCode-@<br>";
	return OutString;
}
function GoCCRefundORCaptureORVoidPreAuthORVoidPurchaseORVoidCapture_Opus(order_id,transid,amount,action)
{
	var nvpstr=XML_AddParam("merchantid",GetOpus_MerchantID());
	nvpstr+=XML_AddParam("password",GetOpus_Password());

	//1=Purchase, 2=Refund, 3=Void Purchase, 4=Authorization, 5=Capture, 7=Void Capture, 9=Void Authorization
	nvpstr+=XML_AddParam("action","2");

	nvpstr+=XML_AddParam("trackid",order_id);
	nvpstr+=XML_AddParam("transid",transid);

	if (""+action!="2")
		nvpstr+=XML_AddParam("bill_currencycode","USD");

	if (""+amount!="")
		nvpstr+=XML_AddParam("bill_amount",amount);

//	nvpstr+=XML_AddParam("account_identifier","");nvpstr+=XML_AddParam("bill_customerip","");
//	nvpstr+=XML_AddParam("bill_merchantip","");nvpstr+=XML_AddParam("udf1","");
//	nvpstr+=XML_AddParam("udf2","");nvpstr+=XML_AddParam("udf3","");nvpstr+=XML_AddParam("udf4","");
//	nvpstr+=XML_AddParam("udf5","");
	var resArray=Opus_HTTP_call(GetOpus_URL(),nvpstr);

//Output Format+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//var Tmp="result: "+resArray("result")+"<br>";
//Tmp+="responsecode: "+resArray("responsecode")+"<br>";
//Tmp+="authcode: "+resArray("authcode")+"<br>";
//Tmp+="tranid: "+resArray("tranid")+"<br>";
//Tmp+="trackid: "+resArray("trackid")+"<br>";
//Tmp+="merchantid: "+resArray("merchantid")+"<br>";
//Tmp+="udf1: "+resArray("udf1")+"<br>";
//Tmp+="udf2: "+resArray("udf2")+"<br>";
//Tmp+="udf3: "+resArray("udf3")+"<br>";
//Tmp+="udf4: "+resArray("udf4")+"<br>";
//Output Format----------------------------------------------------------------------

	var CCErrMsg=""
	if ((""+resArray("responsecode")!="0")&&(""+resArray("responsecode")!="11")&&(""+resArray("responsecode")!="77"))
	{
		if (""+SysLang=="F")
			CCErrMsg="Désolé! Il y a un problème avec votre carte de crédit, veuillez corriger les informations de votre carte de crédit.";
		else
			CCErrMsg="Sorry! There's a problem with your credit card, please correct your credit card information. ("+resArray("responsecode")+")";
	}
	var OutString= "@CCErrMsg+@" + CCErrMsg + "@CCErrMsg-@<br>";
	OutString+= "@TransID+@" + resArray("tranid") + "@TransID-@<br>";
	OutString+= "@AuthCode+@" + resArray("authcode") + "@AuthCode-@<br>";						//char8
	OutString+= "@AVSResultCode+@" + "" + "@AVSResultCode-@<br>";
	OutString+= "@CVDResultCode+@" + "" + "@CVDResultCode-@<br>";
	return OutString;
}
*/

function GetSoftPOS_HeadServer_Files_Path(Code)
{
	var Output=""+Server.MapPath("/");
	var i=0;
	var Pos= -1;
	for (i=Output.length; i > 0; i--)
	{
		if (Output.substring(i-1,i)=='\\')
		{
			Pos=i-1;
			break;
		}
	}
	if (Pos < 0)
		return "";
	Output=Output.substring(0,Pos);
	var Tmp=Code.toUpperCase();
	if (Tmp=="DIR2")
		return Output + "\\SoftPOS_UserBackup\\HeadServer\\Files\\";
	else
		return Output + "\\SoftPOS\\HeadServer\\Files\\";
}

function GetAmazon_OutgoingFolder()
{
	return "d:\\amazontransport\\production\\outgoing";
}
function GetAmazon_IncomingFolder()
{
	return "d:\\amazontransport\\production\\reports";
}
/*
function GetOpus_URL()
{
	var IsDebug=false;

	if (IsDebug)
		return "http://www.egatepay.com/uat.egp.web/gateway.aspx";
	else
		return "http://www.egatepay.com/uat.egp.web/gateway.aspx";
}
function GetOpus_MerchantID()
{
	var IsDebug=false;

	if (IsDebug)
		return "softmoc";
	else
		return "softmoc";
}
function GetOpus_Password()
{
	var IsDebug=false;

	if (IsDebug)
		return "softmoc1234!";
	else
		return "softmoc1234!";
}
*/

function GetPayPal_API_ENDPOINT(IsUS)
{
	var IsDebug=false;

	if (IsDebug)
		return "https://api-3t.sandbox.paypal.com/nvp";
	else
	{
		if (IsUS)
			return "https://api-3t.paypal.com/nvp";
		else
			return "https://api-3t.paypal.com/nvp";
	}
}
function GetPayPal_PAYPAL_EC_URL(IsUS)
{
	var IsDebug=false;
	
	if (IsDebug)
		return "https://www.sandbox.paypal.com/webscr";
	else
	{
		if (IsUS)
			return "https://www.paypal.com/webscr";
		else
			return "https://www.paypal.com/webscr";
	}
}
function GetPayPal_API_USERNAME(IsUS)
{


//testing buyer account:    buyer_1287084110_per@yahoo.com, 28474742847474
//                        https://developer.paypal.com


	var IsDebug=false;
	
	if (IsDebug)
		return "seller_1287083093_biz_api1.yahoo.com";
	else
	{
		if (IsUS)
			return "paypal_api1.softmoc.com";
		else
			return "paypal_api1.softmoc.com";
	}
}	
function GetPayPal_API_PASSWORD(IsUS)
{
	var IsDebug=false;
	
	if (IsDebug)
		return "1287083108";
	else
	{
		if (IsUS)
			return "F7T9A83T6U7R856P";
		else
			return "F7T9A83T6U7R856P";
	}
}
function GetPayPal_API_SIGNATURE(IsUS)
{
	var IsDebug=false;
	
	if (IsDebug)
		return "AFcWxV21C7fd0v3bYYYRCpSSRl31AKRSWFQCrsh7WYJq9U-E8t-dnsxq";
	else
	{
		if (IsUS)
			return "AFcWxV21C7fd0v3bYYYRCpSSRl31ABFTT6ejk.xAQrWvmJAjNl6VuTcS";
		else
			return "AFcWxV21C7fd0v3bYYYRCpSSRl31ABFTT6ejk.xAQrWvmJAjNl6VuTcS";
	}
}



function GetPayPal_API_VERSION(IsUS)
{
	if (IsUS)
		return "64.0";
	else
		return "64.0";
}
function GetPayPal_API_CREDENTIALS(IsUS)
{
	return "&USER=" + GetPayPal_API_USERNAME(IsUS) + "&PWD=" + GetPayPal_API_PASSWORD(IsUS) + "&SIGNATURE=" + GetPayPal_API_SIGNATURE(IsUS) + "&VERSION=" + GetPayPal_API_VERSION(IsUS);
}
function GetCC_store_id(IsUS)
{
	if (IsUS)
		return "monca09775";
	else
		return "monmpg8720";
}
function GetCC_api_token(IsUS)
{
	if (IsUS)
		return "qIBXonmZo6Ag8fH36uvQ";
	else
		return "LsoYZeniFKrCUme3wMn3";
}
function GetCC_crypt_type(IsUS)
{
	if (IsUS)
		return "7";
	else
		return "7";
}
function GetCC_cvd_indicator(IsUS)
{
	if (IsUS)
		return "1";
	else
		return "1";
}
function IsFromSearchEngine()
{
	var UserAgent=""+Request.ServerVariables("HTTP_USER_AGENT");
	UserAgent=UserAgent.toLowerCase();
	if ((UserAgent.indexOf("googlebot",0)>=0)||(UserAgent.indexOf("bingbot",0)>=0)||(UserAgent.indexOf("yahoobot",0)>=0)||(UserAgent.indexOf("facebook",0)>=0))
		return true;
	return false;
}
function IsLocalNetwork(XIP)
{
	var IP=""+XIP;
/*	if ((IP=="68.164.114.137")||(IP=="68.164.114.138")||(IP=="68.164.114.139")||(IP=="68.164.114.140")||(IP=="68.164.114.141")||(IP=="127.0.0.1"))
		return true;
	if (IP.length>12)
	{
		if (IP.substring(0,12)=="100.100.100.")
			return true;
	}
	if (IP.length>6)
	{
		if (IP.substring(0,6)=="10.30.")
			return true;
	}
	if (IP.length>8)
	{
		if (IP.substring(0,8)=="192.168.")
			return true;
	}
	if (IP.length>11)
	{
		if (IP.substring(0,11)=="10.255.255.")
			return true;
	}
	if (IP.length>11)
	{
		if (IP.substring(0,11)=="207.236.61.")
			return true;
	}
	if (IP.length>10)
	{
		if (IP.substring(0,10)=="208.87.35.")
			return true;
	}
	return false;
*/

	if (IP=="127.0.0.1")
		return true;
	if (IP.length>8)
	{
		if (IP.substring(0,8)=="192.168.")
			return true;
	}
	else if (IP.length>10)
	{
		if (IP.substring(0,10)=="38.99.153.")
			return true;
	}
	return false;
}
function IIf(CheckValue, TruePart, FalsePart)
{
	if (CheckValue)
		return TruePart;
	else
		return FalsePart
}
function CheckC_En(In)
{
	if (In=="1")
		return "P";
	else if (In=="2")
		return "K";		
	else if (In=="3")
		return "H";		
	else if (In=="4")
		return "F";		
	else if (In=="5")
		return "R";		
	else if (In=="6")
		return "M";		
	else if (In=="7")
		return "C";		
	else if (In=="8")
		return "N";		
	else if (In=="9")
		return "O";		
	else if (In=="0")
		return "Q";		
	else if (In=="_")
		return "T";		
	else if (In=="@")
		return "S";		
	else if (In==" ")
		return "Y";		
	else
		return "X";
}
function CheckC_De(In)
{
	if ((In=="P")||(In=="p"))
		return "1";		
	else if ((In=="K")||(In=="k"))
		return "2";		
	else if ((In=="H")||(In=="h"))
		return "3";		
	else if ((In=="F")||(In=="f"))
		return "4";		
	else if ((In=="R")||(In=="r"))
		return "5";		
	else if ((In=="M")||(In=="m"))
		return "6";		
	else if ((In=="C")||(In=="c"))
		return "7";		
	else if ((In=="N")||(In=="n"))
		return "8";		
	else if ((In=="O")||(In=="o"))
		return "9";		
	else if ((In=="Q")||(In=="q"))
		return "0";		
	else if ((In=="T")||(In=="t"))
		return "_";		
	else if ((In=="S")||(In=="s"))
		return "@";		
	else if ((In=="Y")||(In=="y"))
		return " ";		
	else
		return "X";
}
function CheckOutHttps()
{
	var DNSName=""+ Request.ServerVariables("HTTP_HOST");
/*	if ((DNSName=="100.100.100.152")||(DNSName=="207.236.61.152"))
	{
		Application("CheckOutHttps")="http";
		return Application("CheckOutHttps");
	}
*/
	var Tmp="" + Application("CheckOutHttps");
	Tmp=Tmp.toLowerCase();
	if (Tmp!="https")
		Application("CheckOutHttps")="http";
	return Application("CheckOutHttps");
}
function MemberAreaHttps()
{
	var DNSName=""+ Request.ServerVariables("HTTP_HOST");
/*	if ((DNSName=="100.100.100.152")||(DNSName=="207.236.61.152"))
	{
		Application("MemberAreaHttps")="http";
		return Application("MemberAreaHttps");
	}
*/
	var Tmp="" + Application("MemberAreaHttps");
	Tmp=Tmp.toLowerCase();
	if (Tmp!="https")
		Application("MemberAreaHttps")="http";
	return Application("MemberAreaHttps");
}
function CheckInt(InStr)
{
	var Out="";
	var Tmp="" + InStr;
	if (Tmp=="")
		return false;
	for (var i=0; i < Tmp.length; i++)
	{
		var theC=Tmp.substring(i,i+1);
		if ((theC!='0')&&(theC!='1')&&(theC!='2')&&(theC!='3')&&(theC!='4')&&(theC!='5')&&(theC!='6')&&(theC!='7')&&(theC!='8')&&(theC!='9'))
			return false;
	}
	return true;
}
function GetAddressNo(InStr)
{
	var Out="";
	var Tmp="" + InStr;
	for (var i=0; i < Tmp.length; i++)
	{
		var theC=Tmp.substring(i,i+1);
		if ((theC=='0')||(theC=='1')||(theC=='2')||(theC=='3')||(theC=='4')||(theC=='5')||(theC=='6')||(theC=='7')||(theC=='8')||(theC=='9'))
			Out+=theC;
		else
			return Out;
	}
}
function GetAddress(InStr)
{
	var Out="";
	var Tmp="" + InStr;
	for (var i=0; i < Tmp.length; i++)
	{
		var theC=Tmp.substring(i,i+1);
		if ((theC=='0')||(theC=='1')||(theC=='2')||(theC=='3')||(theC=='4')||(theC=='5')||(theC=='6')||(theC=='7')||(theC=='8')||(theC=='9'))
			continue;
		else
			break;
	}
	Out=Tmp.substring(i,Tmp.length);
	if (Out.substring(0,1)==' ')
		Out=Out.substring(1,Out.length);
	return Out;
}
function GetShortProvinceName(Province)
{
	var ProvinceN=Province.toLowerCase();
	if ((""+ProvinceN=="alberta")||(""+ProvinceN=="alberta"))
		return "AB";
	else if ((""+ProvinceN=="british columbia")||(""+ProvinceN=="colombie-britannique"))
		return "BC";
	else if ((""+ProvinceN=="manitoba")||(""+ProvinceN=="manitoba"))
		return "MB";
	else if ((""+ProvinceN=="new brunswick")||(""+ProvinceN=="nouveau-brunswick"))
		return "NB";
	else if ((""+ProvinceN=="newfoundland and labrador")||(""+ProvinceN=="terre-neuve-et-labrador"))
		return "NL";
	else if ((""+ProvinceN=="northwest territories")||(""+ProvinceN=="(territoires du) nord-ouest"))
		return "NT";
	else if ((""+ProvinceN=="nova scotia")||(""+ProvinceN=="nouvelle-écosse"))
		return "NS";
	else if ((""+ProvinceN=="nunavut")||(""+ProvinceN=="nunavut"))
		return "NU";
	else if ((""+ProvinceN=="ontario")||(""+ProvinceN=="ontario"))
		return "ON";
	else if ((""+ProvinceN=="prince edward island")||(""+ProvinceN=="l'île du prince-édouard"))
		return "PE";
	else if ((""+ProvinceN=="quebec")||(""+ProvinceN=="québec"))
		return "QC";
	else if ((""+ProvinceN=="saskatchewan")||(""+ProvinceN=="saskatchewan"))
		return "SK";
	else if ((""+ProvinceN=="yukon")||(""+ProvinceN=="yukon"))
		return "YT";


	else if (""+ProvinceN=='alabama')
		return "AL";
	else if (""+ProvinceN=='apo,ae')
		return "AE";
	else if (""+ProvinceN=='alaska')
		return "AK";
	else if (""+ProvinceN=='arkansas')
		return "AR";
	else if (""+ProvinceN=='arizona')
		return "AZ";
	else if (""+ProvinceN=='california')
		return "CA";
	else if (""+ProvinceN=='colorado')
		return "CO";
	else if (""+ProvinceN=='connecticut')
		return "CT";
	else if (""+ProvinceN=='district of columbia')
		return "DC";
	else if (""+ProvinceN=='delaware')
		return "DE";
	else if (""+ProvinceN=='florida')
		return "FL";
	else if (""+ProvinceN=='georgia')
		return "GA";
	else if (""+ProvinceN=='hawaii')
		return "HI";
	else if (""+ProvinceN=='iowa')
		return "IA";
	else if (""+ProvinceN=='idaho')
		return "ID";
	else if (""+ProvinceN=='illinois')
		return "IL";
	else if (""+ProvinceN=='indiana')
		return "IN";
	else if (""+ProvinceN=='kansas')
		return "KS";
	else if (""+ProvinceN=='kentucky')
		return "KY";
	else if (""+ProvinceN=='louisiana')
		return "LA";
	else if (""+ProvinceN=='massachusetts')
		return "MA";
	else if (""+ProvinceN=='maryland')
		return "MD";
	else if (""+ProvinceN=='maine')
		return "ME";
	else if (""+ProvinceN=='michigan')
		return "MI";
	else if (""+ProvinceN=='minnesota')
		return "MN";
	else if (""+ProvinceN=='missouri')
		return "MO";
	else if (""+ProvinceN=='mississippi')
		return "MS";
	else if (""+ProvinceN=='montana')
		return "MT";
	else if (""+ProvinceN=='north carolina')
		return "NC";
	else if (""+ProvinceN=='north dakota')
		return "ND";
	else if (""+ProvinceN=='nebraska')
		return "NE";
	else if (""+ProvinceN=='new hampshire')
		return "NH";
	else if (""+ProvinceN=='new jersey')
		return "NJ";
	else if (""+ProvinceN=='new mexico')
		return "NM";
	else if (""+ProvinceN=='nevada')
		return "NV";
	else if (""+ProvinceN=='new york')
		return "NY";
	else if (""+ProvinceN=='ohio')
		return "OH";
	else if (""+ProvinceN=='oklahoma')
		return "OK";
	else if (""+ProvinceN=='oregon')
		return "OR";
	else if (""+ProvinceN=='pennsylvania')
		return "PA";
	else if (""+ProvinceN=='puerto rico')
		return "PR";
	else if (""+ProvinceN=='rhode island')
		return "RI";
	else if (""+ProvinceN=='south carolina')
		return "SC";
	else if (""+ProvinceN=='south dakota')
		return "SD";
	else if (""+ProvinceN=='tennessee')
		return "TN";
	else if (""+ProvinceN=='texas')
		return "TX";
	else if (""+ProvinceN=='utah')
		return "UT";
	else if (""+ProvinceN=='virginia')
		return "VA";
	else if (""+ProvinceN=='vermont')
		return "VT";
	else if (""+ProvinceN=='washington')
		return "WA";
	else if (""+ProvinceN=='wisconsin')
		return "WI";
	else if (""+ProvinceN=='west virginia')
		return "WV";
	else if (""+ProvinceN=='wyoming')
		return "WY";


	else
		return ProvinceN;
}
function GetLongProvinceName(ProvinceN)
{
	if (""+ProvinceN=="AB")
		return X("Alberta","Alberta");
	else if (""+ProvinceN=="BC")
		return X("British Columbia","Colombie-Britannique");
	else if (""+ProvinceN=="MB")
		return X("Manitoba","Manitoba");
	else if (""+ProvinceN=="NB")
		return X("New Brunswick","Nouveau-Brunswick");
	else if (""+ProvinceN=="NL")
		return X("Newfoundland and Labrador","Terre-Neuve-et-Labrador");
	else if (""+ProvinceN=="NT")
		return X("Northwest Territories","(territoires du) Nord-Ouest");
	else if (""+ProvinceN=="NS")
		return X("Nova Scotia","Nouvelle-Écosse");
	else if (""+ProvinceN=="NU")
		return X("Nunavut","Nunavut");
	else if (""+ProvinceN=="ON")
		return X("Ontario","Ontario");
	else if (""+ProvinceN=="PE")
		return X("Prince Edward Island","l'île du Prince-Édouard");
	else if (""+ProvinceN=="QC")
		return X("Quebec","Québec");
	else if (""+ProvinceN=="SK")
		return X("Saskatchewan","Saskatchewan");
	else if (""+ProvinceN=="YT")
		return X("Yukon","Yukon");


	else if (""+ProvinceN=='AL')
		return "Alabama";
	else if (""+ProvinceN=='AE')
		return "APO,AE";
	else if (""+ProvinceN=='AK')
		return "Alaska";
	else if (""+ProvinceN=='AR')
		return "Arkansas";
	else if (""+ProvinceN=='AZ')
		return "Arizona";
	else if (""+ProvinceN=='CA')
		return "California";
	else if (""+ProvinceN=='CO')
		return "Colorado";
	else if (""+ProvinceN=='CT')
		return "Connecticut";
	else if (""+ProvinceN=='DC')
		return "District of Columbia";
	else if (""+ProvinceN=='DE')
		return "Delaware";
	else if (""+ProvinceN=='FL')
		return "Florida";
	else if (""+ProvinceN=='GA')
		return "Georgia";
	else if (""+ProvinceN=='HI')
		return "Hawaii";
	else if (""+ProvinceN=='IA')
		return "Iowa";
	else if (""+ProvinceN=='ID')
		return "Idaho";
	else if (""+ProvinceN=='IL')
		return "Illinois";
	else if (""+ProvinceN=='IN')
		return "Indiana";
	else if (""+ProvinceN=='KS')
		return "Kansas";
	else if (""+ProvinceN=='KY')
		return "Kentucky";
	else if (""+ProvinceN=='LA')
		return "Louisiana";
	else if (""+ProvinceN=='MA')
		return "Massachusetts";
	else if (""+ProvinceN=='MD')
		return "Maryland";
	else if (""+ProvinceN=='ME')
		return "Maine";
	else if (""+ProvinceN=='MI')
		return "Michigan";
	else if (""+ProvinceN=='MN')
		return "Minnesota";
	else if (""+ProvinceN=='MO')
		return "Missouri";
	else if (""+ProvinceN=='MS')
		return "Mississippi";
	else if (""+ProvinceN=='MT')
		return "Montana";
	else if (""+ProvinceN=='NC')
		return "North Carolina";
	else if (""+ProvinceN=='ND')
		return "North Dakota";
	else if (""+ProvinceN=='NE')
		return "Nebraska";
	else if (""+ProvinceN=='NH')
		return "New Hampshire";
	else if (""+ProvinceN=='NJ')
		return "New Jersey";
	else if (""+ProvinceN=='NM')
		return "New Mexico";
	else if (""+ProvinceN=='NV')
		return "Nevada";
	else if (""+ProvinceN=='NY')
		return "New York";
	else if (""+ProvinceN=='OH')
		return "Ohio";
	else if (""+ProvinceN=='OK')
		return "Oklahoma";
	else if (""+ProvinceN=='OR')
		return "Oregon";
	else if (""+ProvinceN=='PA')
		return "Pennsylvania";
	else if (""+ProvinceN=='PR')
		return "Puerto Rico";
	else if (""+ProvinceN=='RI')
		return "Rhode Island";
	else if (""+ProvinceN=='SC')
		return "South Carolina";
	else if (""+ProvinceN=='SD')
		return "South Dakota";
	else if (""+ProvinceN=='TN')
		return "Tennessee";
	else if (""+ProvinceN=='TX')
		return "Texas";
	else if (""+ProvinceN=='UT')
		return "Utah";
	else if (""+ProvinceN=='VA')
		return "Virginia";
	else if (""+ProvinceN=='VT')
		return "Vermont";
	else if (""+ProvinceN=='WA')
		return "Washington";
	else if (""+ProvinceN=='WI')
		return "Wisconsin";
	else if (""+ProvinceN=='WV')
		return "West Virginia";
	else if (""+ProvinceN=='WY')
		return "Wyoming";


	else
		return "";
}
function GetLongMonth(Month)
{
	if ("" + Month=="1")
		return "January";
	else if ("" + Month=="2")
		return "February";
	else if ("" + Month=="3")
		return "March";
	else if ("" + Month=="4")
		return "April";
	else if ("" + Month=="5")
		return "May";
	else if ("" + Month=="6")
		return "June";
	else if ("" + Month=="7")
		return "July";
	else if ("" + Month=="8")
		return "August";
	else if ("" + Month=="9")
		return "September";
	else if ("" + Month=="10")
		return "October";
	else if ("" + Month=="11")
		return "November";
	else
		return "December";
}
function X(StringE,StringF)
{
	var T=""+SysLang;
	if (T=="F")
	{
		if ((""+StringF=="")||(""+StringF==" ")||(""+StringF=="undefined"))
			return ""+StringE;
		else
			return ""+StringF;
	}
	else
		return ""+StringE;
}
function WriteToCookie(Index,String,TMinToExpires)
{
	Response.Cookies("" + Index) = ""+ String;
	var MinToExpires=""+TMinToExpires;
	if ((MinToExpires!="undefined")&&(MinToExpires!="")&&(MyparseInt(MinToExpires) > 0))
	{
		var CurrentD = new Date;
		CurrentD.setTime(CurrentD.getTime() + (MyparseInt(MinToExpires) * 60 * 1000)); 
		var Tmp=GetLongMonth(CurrentD.getMonth()+1) + " "+CurrentD.getDate()+", " + CurrentD.getYear() + " "+ CurrentD.getHours()+":"+CurrentD.getMinutes()+":00";
		Response.Cookies("" + Index).Expires = Tmp;
	}
	else
	{
		var CurrentD = new Date;
		CurrentD.setYear(CurrentD.getYear()+1);
		var Tmp=GetLongMonth(CurrentD.getMonth()+1) + " 1, " + CurrentD.getYear() + " 11:59:59 PM";
		Response.Cookies("" + Index).Expires = Tmp;
	}
}
function CharReturn()
{
	return "\n";
}
function AddSpace(InString)
{
	if ((InString=="")||(InString==" "))
		return "&nbsp;";
	else
		return Server.HTMLEncode(InString);
}
function GetRealEmail(In)
{
	var Out=""+In;
	var Tmp=Out.toLowerCase();
	if ((Tmp.indexOf("@gmail.",0)<=0) && (Tmp.indexOf("@hotmail.",0)<=0))
		return ""+In;

	var PlusSignPos= Out.indexOf("+",0);
	var AtSignPos= Out.indexOf("@",0);
	if (PlusSignPos >= 0)
	{
		if (PlusSignPos < AtSignPos)
		{
			if (PlusSignPos > 0)
				Out=In.substring(0,PlusSignPos)+In.substring(AtSignPos,In.length+1);
			else
				Out=In.substring(AtSignPos,In.length+1);
		}
	}	

	if (Tmp.indexOf("@gmail.",0)>0)
	{
		var DotSignPos=Out.indexOf(".",0);
		AtSignPos= Out.indexOf("@",0);
		while ((DotSignPos >= 0)&&(DotSignPos < AtSignPos))
		{
			if (DotSignPos > 0)
				Out=Out.substring(0,DotSignPos)+Out.substring(DotSignPos+1,Out.length+1);
			else
				Out=+Out.substring(1,Out.length+1);
			DotSignPos=Out.indexOf(".",0);
			AtSignPos= Out.indexOf("@",0);
		}
	}
	return Out;
}

function EmailOK(Email)
{
	var X=Email;
	if ((X!='')&&(X!=null))
	{
		var A=false;
		var D=false;
		for (var i=0;i < X.length;i=i+1)
		{
			var C=X.substring(i,i+1);
			if (C=='@')
			{
				if ((A)||(i== X.length-1))
				{
					A=false;
					break;
				}
				A=true;
			}
			else if (C=='.')
			{
				if (i== X.length-1)
				{
					D=false;
					break;
				}
				D=true;
			}
		}
		if ((!A)||(!D))
			return false;
		else
			return true;
	}
	else
		return true;
}
function ReplaceNow(InString,From,To)
{
	var In=""+InString;
	var F=""+From;
	var T=""+To;
	if (In.indexOf(From,0)<0)
		return "";
	else
		return ""+In.replace(F, T);
}
function ReplaceAllNow(InSt,From,To)
{
	var InString="" +InSt;
	if (From=="")
		return ""+InString;
	while (InString.indexOf(From,0)>=0)
		InString=ReplaceNow(InString,From,To);
	return ""+InString;
}
function MakeIt(InSt)
{
	var InString="" +InSt;
	while (InString.indexOf("%",0)>=0)
		InString=ReplaceNow(InString,"%","~!@");
	while (InString.indexOf("~!@",0)>=0)
		InString=ReplaceNow(InString,"~!@","%25");
	while (InString.indexOf("~",0)>=0)
		InString=ReplaceNow(InString,"~","%7E");
	while (InString.indexOf("|",0)>=0)
		InString=ReplaceNow(InString,"|","%7C");
	while (InString.indexOf("}",0)>=0)
		InString=ReplaceNow(InString,"}","%7D");
	while (InString.indexOf("{",0)>=0)
		InString=ReplaceNow(InString,"{","%7B");
	while (InString.indexOf("`",0)>=0)
		InString=ReplaceNow(InString,"`","%60");
	while (InString.indexOf("^",0)>=0)
		InString=ReplaceNow(InString,"^","%5E");
	while (InString.indexOf("]",0)>=0)
		InString=ReplaceNow(InString,"]","%5D");
	while (InString.indexOf("\\",0)>=0)
		InString=ReplaceNow(InString,"\\","%5C");
	while (InString.indexOf("[",0)>=0)
		InString=ReplaceNow(InString,"[","%5B");
	while (InString.indexOf("?",0)>=0)
		InString=ReplaceNow(InString,"?","%3F");
	while (InString.indexOf(">",0)>=0)
		InString=ReplaceNow(InString,">","%3E");
	while (InString.indexOf("=",0)>=0)
		InString=ReplaceNow(InString,"=","%3D");
	while (InString.indexOf("<",0)>=0)
		InString=ReplaceNow(InString,"<","%3C");
	while (InString.indexOf(";",0)>=0)
		InString=ReplaceNow(InString,";","%3B");
	while (InString.indexOf(":",0)>=0)
		InString=ReplaceNow(InString,":","%3A");
	while (InString.indexOf(")",0)>=0)
		InString=ReplaceNow(InString,")","%29");
	while (InString.indexOf("(",0)>=0)
		InString=ReplaceNow(InString,"(","%28");
	while (InString.indexOf("'",0)>=0)
		InString=ReplaceNow(InString,"'","%27");
	while (InString.indexOf("&",0)>=0)
		InString=ReplaceNow(InString,"&","%26");
	while (InString.indexOf("$",0)>=0)
		InString=ReplaceNow(InString,"$","%24");
	while (InString.indexOf("#",0)>=0)
		InString=ReplaceNow(InString,"#","%23");
	while (InString.indexOf('"',0)>=0)
		InString=ReplaceNow(InString,'"',"%22");
	while (InString.indexOf("!",0)>=0)
		InString=ReplaceNow(InString,"!","%21");
	while (InString.indexOf("/",0)>=0)
		InString=ReplaceNow(InString,"/","%2F");
	while (InString.indexOf(",",0)>=0)
		InString=ReplaceNow(InString,",","%2C");
	while (InString.indexOf("+",0)>=0)
		InString=ReplaceNow(InString,"+","%2B");
	while (InString.indexOf(" ",0)>=0)
		InString=ReplaceNow(InString," ","%20");
	return InString;
}
function MakeUPS(InSt)
{
	var InString=InSt;
	while (InString.indexOf(" ",0)>=0)
		InString=ReplaceNow(InString," ","+");
	return InString;
}
function UnMakeIt(InSt)
{
	var InString=InSt;
	while (InString.indexOf("+",0)>=0)
		InString=ReplaceNow(InString,"+"," ");
	while (InString.indexOf("%20",0)>=0)
		InString=ReplaceNow(InString,"%20"," ");
	while (InString.indexOf("%2B",0)>=0)
		InString=ReplaceNow(InString,"%2B","+");
	while (InString.indexOf("%2C",0)>=0)
		InString=ReplaceNow(InString,"%2C",",");
	while (InString.indexOf("%2F",0)>=0)
		InString=ReplaceNow(InString,"%2F","/");
	while (InString.indexOf("%21",0)>=0)
		InString=ReplaceNow(InString,"%21","!");
	while (InString.indexOf("%22",0)>=0)
		InString=ReplaceNow(InString,"%22",'"');
	while (InString.indexOf("%23",0)>=0)
		InString=ReplaceNow(InString,"%23","#");
	while (InString.indexOf("%24",0)>=0)
		InString=ReplaceNow(InString,"%24","$");
	while (InString.indexOf("%26",0)>=0)
		InString=ReplaceNow(InString,"%26","&");
	while (InString.indexOf("%27",0)>=0)
		InString=ReplaceNow(InString,"%27","'");
	while (InString.indexOf("%28",0)>=0)
		InString=ReplaceNow(InString,"%28","(");
	while (InString.indexOf("%29",0)>=0)
		InString=ReplaceNow(InString,"%29",")");
	while (InString.indexOf("%3A",0)>=0)
		InString=ReplaceNow(InString,":");
	while (InString.indexOf("%3B",0)>=0)
		InString=ReplaceNow(InString,"%3B",";");
	while (InString.indexOf("%3C",0)>=0)
		InString=ReplaceNow(InString,"%3C","<");
	while (InString.indexOf("%3D",0)>=0)
		InString=ReplaceNow(InString,"%3D","=");
	while (InString.indexOf("%3E",0)>=0)
		InString=ReplaceNow(InString,"%3E",">");
	while (InString.indexOf("%3F",0)>=0)
		InString=ReplaceNow(InString,"%3F","?");
	while (InString.indexOf("%5B",0)>=0)
		InString=ReplaceNow(InString,"%5B","[");
	while (InString.indexOf("%5C",0)>=0)
		InString=ReplaceNow(InString,"%5C","\\");
	while (InString.indexOf("%5D",0)>=0)
		InString=ReplaceNow(InString,"%5D","]");
	while (InString.indexOf("%5E",0)>=0)
		InString=ReplaceNow(InString,"%5E","^");
	while (InString.indexOf("%60",0)>=0)
		InString=ReplaceNow(InString,"%60","`");
	while (InString.indexOf("%7B",0)>=0)
		InString=ReplaceNow(InString,"%7B","{");
	while (InString.indexOf("%7D",0)>=0)
		InString=ReplaceNow(InString,"%7D","}");
	while (InString.indexOf("%7C",0)>=0)
		InString=ReplaceNow(InString,"%7C","|");
	while (InString.indexOf("%7E",0)>=0)
		InString=ReplaceNow(InString,"%7E","~");
	while (InString.indexOf("%0D",0)>=0)
		InString=ReplaceNow(InString,"%0D","\r");
	while (InString.indexOf("%0A",0)>=0)
		InString=ReplaceNow(InString,"%0A","\n");
	while (InString.indexOf("%25",0)>=0)
		InString=ReplaceNow(InString,"%25","~!@");
	while (InString.indexOf("%40",0)>=0)
		InString=ReplaceNow(InString,"%40","@");
	while (InString.indexOf("~!@",0)>=0)
		InString=ReplaceNow(InString,"~!@","%");

	return UrlDecode(InString);
}
function utf8_decode (str_data)//convert %xx  to french char
 {
  // http://kevin.vanzonneveld.net
  // +   original by: Webtoolkit.info (http://www.webtoolkit.info/)
  // +      input by: Aman Gupta
  // +   improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
  // +   improved by: Norman "zEh" Fuchs
  // +   bugfixed by: hitwork
  // +   bugfixed by: Onno Marsman
  // +      input by: Brett Zamir (http://brett-zamir.me)
  // +   bugfixed by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
  // +   bugfixed by: kirilloid
  // *     example 1: utf8_decode('Kevin van Zonneveld');
  // *     returns 1: 'Kevin van Zonneveld'

  var tmp_arr = [],
    i = 0,
    ac = 0,
    c1 = 0,
    c2 = 0,
    c3 = 0,
    c4 = 0;

  str_data += '';

  while (i < str_data.length) {
    c1 = str_data.charCodeAt(i);
    if (c1 <= 191) {
      tmp_arr[ac++] = String.fromCharCode(c1);
      i++;
    } else if (c1 <= 223) {
      c2 = str_data.charCodeAt(i + 1);
      tmp_arr[ac++] = String.fromCharCode(((c1 & 31) << 6) | (c2 & 63));
      i += 2;
    } else if (c1 <= 239) {
      // http://en.wikipedia.org/wiki/UTF-8#Codepage_layout
      c2 = str_data.charCodeAt(i + 1);
      c3 = str_data.charCodeAt(i + 2);
      tmp_arr[ac++] = String.fromCharCode(((c1 & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
      i += 3;
    } else {
      c2 = str_data.charCodeAt(i + 1);
      c3 = str_data.charCodeAt(i + 2);
      c4 = str_data.charCodeAt(i + 3);
      c1 = ((c1 & 7) << 18) | ((c2 & 63) << 12) | ((c3 & 63) << 6) | (c4 & 63);
      c1 -= 0x10000;
      tmp_arr[ac++] = String.fromCharCode(0xD800 | ((c1>>10) & 0x3FF));
      tmp_arr[ac++] = String.fromCharCode(0xDC00 | (c1 & 0x3FF));
      i += 4;
    }
  }

  return tmp_arr.join('');
}
function CanDivBy(Up,Down)
{
	var rr= 0 + Up%Down;
	if (rr==0)
		return true;
	return false;
}
function OpenCon(ServerName,Login,Password,DatabaseName)
{
	var DNSName=""+ Request.ServerVariables("HTTP_HOST");
/*	if ((DNSName=="100.100.100.152")||(DNSName=="207.236.61.152"))
		ServerName="100.100.100.54";
*/	var x="Driver={SQL Server};Server=";
	if (ServerName!="127.0.0.1")
		x+=ServerName;
	x+=";UID=";
	x+=Login;
	x+=";PWD=";
	x+=Password;
	x+=";WSID=";
	x+=DatabaseName;
	x+="Eng;Database=";
	x+=DatabaseName;
	x+=";DSN=;";
	return x;
}

function ToWildSQL(s)
{
	var tmp="";
	for (var i=0;i < s.length;i=i+1)
	{
		var theC=s.substring(i,i+1);
		tmp+=theC;
		if (theC=='\'')
			tmp+=theC;
	}
	tmp="'%" + tmp + "%'";
	return tmp;
}

function MyparseFloat(s)
{
	var Input="" + s;
	var Output="";
	var Start="";
	var CountLengthBeforeDot=0;
	var GoCount=true;
	for (var i=0; i < Input.length;i++)
	{
		var C=""+Input.substring(i,i+1);
		if (((C >= '0') && (C <= '9'))  || (C=='.') || (C=='-'))
		{
			Output+=C;
			if ((Start=="")&&(C!='-'))
				Start=C;
			if (C=='.')
				GoCount=false;
			if (GoCount)
				CountLengthBeforeDot++;
		}
	}
	if ((CountLengthBeforeDot > 1)&&(Start=="0"))
	{
		Input=Output;
		Output="";
		for (i=0; i < Input.length;i++)
		{
			if (""+Input.substring(i,i+1)!="0")
				break;
		}
		for (; i < Input.length;i++)
			Output+=""+Input.substring(i,i+1);
	}
	if (Output=="")
		Output="0";
	return parseFloat(Output);
}
function MyparseInt(s)
{
	var Input="" + s;
	var Output="";
	var Start="";
	var CountLengthBeforeDot=0;
	var GoCount=true;
	for (var i=0; i < Input.length;i++)
	{
		var C=""+Input.substring(i,i+1);
		if (((C >= '0') && (C <= '9'))  || (C=='.') || (C=='-'))
		{
			Output+=C;
			if ((Start=="")&&(C!='-'))
				Start=C;
			if (C=='.')
				GoCount=false;
			if (GoCount)
				CountLengthBeforeDot++;
		}
	}
	if ((CountLengthBeforeDot > 1)&&(Start=="0"))
	{
		Input=Output;
		Output="";
		for (i=0; i < Input.length;i++)
		{
			if (""+Input.substring(i,i+1)!="0")
				break;
		}
		for (; i < Input.length;i++)
			Output+=""+Input.substring(i,i+1);
	}
	if (Output=="")
		Output="0";
	return parseInt(Output);
}
function DoubleToString(realPrice,NeedDollar,DP)
{
	var DecPlace="" + DP;
	if ((DecPlace!="0")&&(DecPlace!="1")&&(DecPlace!="3")&&(DecPlace!="4")&&(DecPlace!="5")&&(DecPlace!="6"))
		DecPlace="2";
	DecPlace=MyparseInt(DecPlace);

	var Input='' + realPrice;
	if ((parseFloat(realPrice) > 0.000000)&&(parseFloat(realPrice) < 0.000001))
		Input="0.000000";
	if ((parseFloat(realPrice) < 0.000000)&&(parseFloat(realPrice) > -0.000001))
		Input="0.000000";
	if ((MyparseFloat(realPrice) > 0.000000)&&(MyparseFloat(realPrice) < 0.000001))
		Input="0.000000";
	if ((MyparseFloat(realPrice) < 0.000000)&&(MyparseFloat(realPrice) > -0.000001))
		Input="0.000000";
	var TT="" + Input;
	Input="";
	for (var xx=0; xx < TT.length; xx++)
	{
		if (((TT.substring(xx,xx+1)>='0') &&(TT.substring(xx,xx+1)<='9')) || (TT.substring(xx,xx+1)=='-')
						|| (TT.substring(xx,xx+1)=='.'))
			Input+=TT.substring(xx,xx+1);
	}

	var IsZeroNeg='';
	if (NeedDollar)
		if (Input.length >= 2)
		{
			if (Input.substring(0,2)=='-.')
				return "-0";
			else if ((Input.substring(0,1)=='-')&&(Input.substring(1,2)=='0'))
				IsZeroNeg='-';
		}

	var Pos=Input.indexOf('.',0);
	if (Pos<0)
	{
		if (!NeedDollar)
		{
			var Output="";
			for (var i=0; i < DecPlace; i++)
				Output+="0";
			return Output;
		}
		Pos=Input.length;
	}
	var DollarOutput=0;
	if (NeedDollar)
	{
		var Output='';
		for (var i=0;((i < Pos)&&(i < Input.length));i=i+1)
			Output+=Input.substring(i,i+1);
		if (Output!='')
			DollarOutput=MyparseInt(Output);
	}
	var Output='';
	for (var i=Pos+1;(i < Input.length);i=i+1)
		Output+=Input.substring(i,i+1);
	if ((Output=='')||(Output.length < 1))
	{
		if (NeedDollar)
			return ''+IsZeroNeg + DollarOutput;
		Output="";
		for (var i=0; i < DecPlace; i++)
			Output+="0";
		return Output;
	}
	if (Output.length < DecPlace)
	{
		if (NeedDollar)
			return ''+IsZeroNeg+DollarOutput;

		for (var i=Output.length; i < DecPlace; i++)
			Output+="0";
		return Output;
	}
	if (Output.length == DecPlace)
	{
		if (NeedDollar)
			return ''+IsZeroNeg+DollarOutput;
		return Output;
	}

	var Tmp=0;
	var CheckNum=0;
	for (var i=0; i < DecPlace; i++)
	{
		Tmp=(Tmp * 10) +MyparseInt(Output.substring(i,i+1));
		CheckNum=(CheckNum * 10) + 9;
	}
	if (MyparseInt(Output.substring(i,i+1))>=5)
	{
		if (Tmp>=CheckNum)
		{
			if (NeedDollar)
			{
				if (DollarOutput < 0)
					return ''+IsZeroNeg+(DollarOutput-1);
				else
					return ''+IsZeroNeg+(DollarOutput+1);
			}
			else
			{
				Output="";
				for (var i=0; i < DecPlace; i++)
					Output+="0";
				return Output;
			}
		}
		else
			Tmp=Tmp+1;
	}
	if (NeedDollar)
		return ''+IsZeroNeg+DollarOutput;
	if (Tmp==0)
	{
		Output="";
		for (var i=0; i < DecPlace; i++)
			Output+="0";
		return Output;
	}
	else
	{
		Output='' + Tmp;
		for (var i=Output.length; i < DecPlace; i++)
			Output="0"+Output;
	}
	return Output;
}
function RoundDoubleUp(Input)
{
	var tt="" + DoubleToString(Input,true)+"."+DoubleToString(Input,false);
	return parseFloat(tt);
}

function TicketNumberEncode(Input,xcORrORwORd)
{
	var CodeString="3675810924";
	var i;
	for (i=0; i < Input.length; i++)
	{
		if ((Input.substring(i,i+1)<'0')||(Input.substring(i,i+1)>'9'))
			return "";
	}
	var Output="";

	var cORrORwORd=""+xcORrORwORd;
	cORrORwORd=cORrORwORd.toLowerCase();

	if ((cORrORwORd!="w")&&(cORrORwORd!="d")&&(cORrORwORd!="c")&&(cORrORwORd!="i"))
		cORrORwORd="r";

	if (cORrORwORd=="r")
		Output+="2";
	else if (cORrORwORd=="w")
		Output+="3";
	else if (cORrORwORd=="d")
		Output+="4";
	else if (cORrORwORd=="i")
		Output+="5";
	else// if (cORrORwORd=="c")
		Output+="1";

	var n=Input.length;
	if (n<10)
		Output+="0";
	Output+=""+n;
	Output+=Input;
	for (i=n+3; i<15; i++)
	{
		if (i-n-2 >=10)
			Output+='0';
		else
			Output+=(i-n-2);
	}
	var Tmp=0;
	for (i=0; i<15; i++)
		Tmp+= Output.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0);
	while (Tmp >= 10)
		Tmp-=10;
	Output+=Tmp;

	var RealOutput="";
	for (i=0; i < Output.length; i++)
		RealOutput+=CodeString.substring(Output.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0), Output.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0)+1);

	return RealOutput;
}

function TicketNumberDecode(xInput)
{
	var CodeString="3675810924";

	var cORrORwORd = "";
	var i
	for (i=0; i < xInput.length; i++)
	{
		if ((xInput.substring(i,i+1)<'0')||(xInput.substring(i,i+1)>'9'))
			return "";
	}
	var Input="";
	var j;
	for (i=0; i < xInput.length; i++)
	{
		for (j=0 ; j < 10; j++)
		{
			if (xInput.substring(i,i+1)==CodeString.substring(j,j+1))
			{
				Input+=j;
				break;
			}
		}
	}

	var Output="";
	if (Input.length!=16)
		return "";
	var Tmp=0;
	for (i=0; i<15; i++)
		Tmp+= Input.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0);
	while (Tmp >= 10)
		Tmp-=10;
	if (""+Tmp!=Input.substring(15,16))
	{
		if ((Tmp=="0")&&(Input.substring(15,16)=="1"))
		{}
		else
			return "";
	}

	if (Input.substring(0,1)=='1')
		cORrORwORd= "c";
	else if (Input.substring(0,1)=='2')
		cORrORwORd= "r";
	else if (Input.substring(0,1)=='3')
		cORrORwORd= "w";
	else if (Input.substring(0,1)=='4')
		cORrORwORd= "d";
	else if (Input.substring(0,1)=='5')
		cORrORwORd= "i";
	else
		cORrORwORd= "";
	if (cORrORwORd== "")//old logic
	{
		Tmp= (10 * (Input.substring(0,1).charCodeAt(0)- '0'.charCodeAt(0))) + (Input.substring(1,2).charCodeAt(0)- '0'.charCodeAt(0));
		if (Tmp >13)
			return "";
		for (i=0;i < Tmp; i++)
			Output+=Input.substring(i+2,i+3);
		j=i+2;
		for (i=j;i < 15; i++)
		{
			if (i-j+1 > 10)
			{
				if (Input.substring(i,i+1)!='0')
					return "";
			}
			else
			{
				if (Input.substring(i,i+1)!=''+(i-j+1))
					return "";
			}
		}
		cORrORwORd="r";
	}
	else
	{
		Tmp= (10 * (Input.substring(1,2).charCodeAt(0)- '0'.charCodeAt(0))) + (Input.substring(2,3).charCodeAt(0)- '0'.charCodeAt(0));
		if (Tmp >12)
			return "";
		for (i=0;i < Tmp; i++)
			Output+=Input.substring(i+3,i+4);
		j=i+3;
		for (i=j;i < 15; i++)
		{
			if (i-j+1 >= 10)
			{
				if (Input.substring(i,i+1)!='0')
					return "";
			}
			else
			{
				if (Input.substring(i,i+1)!=''+(i-j+1))
					return "";
			}
		}
	}	
	return cORrORwORd + Output;
}

function XNumberEncode(Input,Leng)
{
	var CodeString="3675810924";
	var i;
	if (Leng >= 100)
		return "";
	for (i=0; i < Input.length; i++)
	{
		if ((Input.substring(i,i+1)<'0')||(Input.substring(i,i+1)>'9'))
			return "";
	}
	var Output="";

	var n=Input.length;
	if (n<10)
		Output+="0";
	Output+=""+n;
	Output+=Input;
	for (i=n+2; i< Leng-1; i++)
	{
		if (i-n-1 >=10)
			Output+='0';
		else
			Output+=(i-n-1);
	}
	var Tmp=0;
	for (i=0; i< Leng-1; i++)
		Tmp+= Output.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0);
	while (Tmp >= 10)
		Tmp-=10;
	Output+=Tmp;

	var RealOutput="";
	for (i=0; i < Output.length; i++)
		RealOutput+=CodeString.substring(Output.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0), Output.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0)+1);

	return RealOutput;
}

function XNumberDecode(xInput,Leng)
{
	var CodeString="3675810924";

	var i
	for (i=0; i < xInput.length; i++)
	{
		if ((xInput.substring(i,i+1)<'0')||(xInput.substring(i,i+1)>'9'))
			return "";
	}
	var Input="";
	var j;
	for (i=0; i < xInput.length; i++)
	{
		for (j=0 ; j < 10; j++)
		{
			if (xInput.substring(i,i+1)==CodeString.substring(j,j+1))
			{
				Input+=j;
				break;
			}
		}
	}

	var Output="";
	if (Input.length!=Leng)
		return "";
	var Tmp=0;
	for (i=0; i< Leng-1; i++)
		Tmp+= Input.substring(i,i+1).charCodeAt(0)- '0'.charCodeAt(0);
	while (Tmp >= 10)
		Tmp-=10;
	if (""+Tmp!=Input.substring(Leng-1,Leng))
	{
		if ((Tmp=="0")&&(Input.substring(Leng-1,Leng)=="1"))
		{}
		else
			return "";
	}

	Tmp= (10 * (Input.substring(0,1).charCodeAt(0)- '0'.charCodeAt(0))) + (Input.substring(1,2).charCodeAt(0)- '0'.charCodeAt(0));
	if (Tmp > Leng-3)
		return "";
	for (i=0;i < Tmp; i++)
		Output+=Input.substring(i+2,i+3);
	j=i+2;
	for (i=j;i < Leng-1; i++)
	{
		if (i-j+1 >= 10)
		{
			if (Input.substring(i,i+1)!='0')
				return "";
		}
		else
		{
			if (Input.substring(i,i+1)!=''+(i-j+1))
				return "";
		}
	}
	return Output;
}

function EncodeString(Input)
{
	var Output="";
	for (var i=0; i < Input.length; i++)
	{
		var C=""+(Input.substring(i,i+1).charCodeAt(0)+i+1);
		Output+=""+C.length+C;
	}
	Tmp=Output;
	Output="";
	for (i=Tmp.length-1; i >= 0; i--)
	{
		var C=Tmp.substring(i,i+1);
		if (C=="0")
			Output+="S";
		else if (C=="1")
			Output+="a";
		else if (C=="2")
			Output+="m";
		else if (C=="3")
			Output+="U";
		else if (C=="4")
			Output+="e";
		else if (C=="5")
			Output+="L";
		else if (C=="6")
			Output+="C";
		else if (C=="7")
			Output+="h";
		else if (C=="8")
			Output+="i";
		else if (C=="9")
			Output+="Q";
	}
	return Output;
}
function DecodeString(Input)
{
	if (Input=="")
		return "";
	var Output="";
	var i=Input.length-1;
	var Counter=0;
	while (i >= 0)
	{
		var C=Input.substring(i,i+1);
		if (C=="S")
			C="0";
		else if (C=="a")
			C="1";
		else if (C=="m")
			C="2";
		else if (C=="U")
			C="3";
		else if (C=="e")
			C="4";
		else if (C=="L")
			C="5";
		else if (C=="C")
			C="6";
		else if (C=="h")
			C="7";
		else if (C=="i")
			C="8";
		else if (C=="Q")
			C="9";
		C=MyparseInt(C);
		if (C <= 0)
			return "";
		i--;
		var Tmp="";
		for (var j=0; j < C; j++)
		{
			if (i < 0)
				return "";
			var X=Input.substring(i,i+1);
			if (X=="S")
				Tmp+="0";
			else if (X=="a")
				Tmp+="1";
			else if (X=="m")
				Tmp+="2";
			else if (X=="U")
				Tmp+="3";
			else if (X=="e")
				Tmp+="4";
			else if (X=="L")
				Tmp+="5";
			else if (X=="C")
				Tmp+="6";
			else if (X=="h")
				Tmp+="7";
			else if (X=="i")
				Tmp+="8";
			else if (X=="Q")
				Tmp+="9";
			i--;
		}
		Counter++;
		Output+=String.fromCharCode(MyparseInt(Tmp)-Counter);
	}
	return Output;
}
function RandomN(HowMany)
{
	var WhichOne=parseInt(DoubleToString((Math.random()*(HowMany))+1,true));
	if (WhichOne>HowMany)
	{
		WhichOne=HowMany;
	}
	return WhichOne;
}
function CheckPhoneNumber(Phone)
{
	var Tmp= Phone;
	var Output="";
	for (var i=0;i < Tmp.length;i=i+1)
	{
		var theC=Tmp.substring(i,i+1);
		if ((theC=='0')||(theC=='1')||(theC=='2')||(theC=='3')||(theC=='4')||(theC=='5')||(theC=='6')||(theC=='7')||(theC=='8')||(theC=='9'))
			Output=Output + theC;
	}
	return Output;
}
function ReadEmailTextFile(FileName)
{
	var FileName=Server.MapPath("/template") + "\\" +FileName;
	var fileContent = "";
	var objFile = Server.CreateObject("Scripting.FileSystemObject");
	var objStream = objFile.OpenTextFile(FileName);
	while (!objStream.AtEndOfStream)
		fileContent += objStream.ReadLine() + "\r\n";
	objStream.Close();
	objFile = null;
	return fileContent;
}
function AddReturn(Content,ReturnChar)
{
	var NumOfChar=80;
	var NumOfCheckingLen=30;
	var Input="" +Content;
	var Output="";
	while (true)
	{
		if (Input.length <= NumOfChar)
		{
			Output+= Input;
			return Output;
		}
		var Pos=CheckReturnPos(Input,NumOfChar,NumOfCheckingLen);
		Output+= Input.substring(0,Pos)+ReturnChar;
		Input="" + Input.substring(Pos,Input.length);
	}
}
function ChangeCCNumberToStar(Buffer)
{
	var s=""+Buffer;
	var LenS=s.length;
	if (LenS <= 10)
		return s;
	var out="";
	for (var i=1;i <= LenS;i++)
	{
		if ((i <= 4)||(i > LenS-4))
			out+=s.substring(i-1,i);
		else
			out+='*';
	}
	return out;
}
function InternalE(Conn,Buffer)
{
	var OutString="";
	var Tmp="select dbo.InternalE('2847474',";
	Tmp+=ToSQL(Buffer);
	Tmp+=") Out";
	var r=Conn.Execute(Tmp);
	if (!r.EOF)
		OutString=""+r("Out");
	r.Close();
	r=null;
	return OutString;
}

function SendEmailByCDOReal(HTMLOrTEXT,FromEmail,ToEmail,Subject,Content,CcEmail,BccEmail,
					xSMTPServer,xSendUserName,xSendUserPassword,xSendUsingNetwork,xSMTPServerPort,xSMTPAuthenticate)
{
//Samples
//SendEmailByCDOReal("HTML","bk@softmoc.com","bk@softmoc.com","SSS","CCC1<br>CCC2<br>CCC3");
//SendEmailByCDOReal("TEXT","bk@softmoc.com","bk@softmoc.com","SSS","CCC");

//before
//web.softmoc.com for all real time emails.
//smtp1.softmoc.com for email blasts.


	var SMTPServer=""+xSMTPServer;
	if ((SMTPServer=="")||(SMTPServer=="undefined")||(SMTPServer=="null")||(SMTPServer=="NULL"))
		SMTPServer="transmail.mwfgroup.com";
	var SendUserName=""+xSendUserName;
	if ((SendUserName=="")||(SendUserName=="undefined")||(SendUserName=="null")||(SendUserName=="NULL"))
	{
		if (""+SMTPServer=="softmoc1.smtp.com")
			SendUserName="bk@mwfgroupbahamas.com";
		else if (""+SMTPServer=="softmoc.smtp.com")
			SendUserName="bk@softmoc.com";
		else if (""+SMTPServer=="transmail.mwfgroup.com")
			SendUserName="";
		else if (""+SMTPServer!="")
			SendUserName="websiteemail@softmoc.com";
		else
			SendUserName="";
	}
	var SendUserPassword=""+xSendUserPassword;
	if ((SendUserPassword=="")||(SendUserPassword=="undefined")||(SendUserPassword=="null")||(SendUserPassword=="NULL"))
	{
		if (""+SMTPServer=="softmoc1.smtp.com")
			SendUserPassword="60d3bc75";
		else if (""+SMTPServer=="softmoc.smtp.com")
			SendUserPassword="60d3bc75";
		else if (""+SMTPServer=="transmail.mwfgroup.com")
			SendUserPassword="";
		else if (""+SMTPServer!="")
			SendUserPassword="7oiZA5sY";
		else
			SendUserPassword="";
	}
	var SendUsingNetwork=""+xSendUsingNetwork;
	if ((SendUsingNetwork=="")||(SendUsingNetwork=="undefined")||(SendUsingNetwork=="null")||(SendUsingNetwork=="NULL"))
		SendUsingNetwork="2";//network
	var SMTPServerPort=""+xSMTPServerPort;
	if ((SMTPServerPort=="")||(SMTPServerPort=="undefined")||(SMTPServerPort=="null")||(SMTPServerPort=="NULL"))
		SMTPServerPort="25";
	var SMTPAuthenticate=""+xSMTPAuthenticate;
	if ((SMTPAuthenticate=="")||(SMTPAuthenticate=="undefined")||(SMTPAuthenticate=="null")||(SMTPAuthenticate=="NULL"))
		SMTPAuthenticate="1";//cdoBasic


	var cdoConfig;
	var cdoMessage;
	try
	{
		cdoConfig = Server.CreateObject("CDO.Configuration");
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = ""+SMTPServer;
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = MyparseInt(SendUsingNetwork);
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = MyparseInt(SMTPAuthenticate);
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = ""+SendUserName;
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""+SendUserPassword;
/*		else
		{
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "";
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1;//local
		}
*/
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = MyparseInt(SMTPServerPort);
		cdoConfig.Fields.Update();
		cdoMessage = Server.CreateObject("CDO.Message");
		cdoMessage.Configuration = cdoConfig;
		cdoMessage.To = ""+ToEmail;
		cdoMessage.From = ""+FromEmail;

		if ((""+CcEmail!="undefined")&&(""+CcEmail!=""))
			cdoMessage.Cc = ""+CcEmail;

		if ((""+BccEmail!="undefined")&&(""+BccEmail!=""))
			cdoMessage.Bcc = ""+BccEmail;

		cdoMessage.Subject = ""+Subject;
		if (""+HTMLOrTEXT=="HTML")
			cdoMessage.HTMLBody = ""+Content;
		else
			cdoMessage.TextBody = ""+Content;
		cdoMessage.send();
	}
	catch(e)
	{
		var Tmp="Error Message: " + e.message;
/*		Tmp+="\r\n";
		Tmp+="Error Code: ";
		Tmp+=e.number + 0xFFFF;
		Tmp+="\r\n";
		Tmp+="Error Name: " + e.name;
*/
		Tmp=Tmp.replace(/\r/g, "");
		Tmp=Tmp.replace(/\n/g, "");
		return Tmp;

	}
	cdoMessage= GetNothing();
	cdoConfig= GetNothing();
	return "";
}
function CheckURLInputType(Connection,Val)
{
	if (Val=="all-gender")
		return "XGRO-";
	else if (Val=="all-brand")
		return "VEND-";
	else if (Val=="all-department")
		return "DEPA-";
	else if (Val=="all-style")
		return "STYL-";
	Val=UrlDecode(Val);
	var rs=Connection.Execute("CheckURLInputType "+ToSQL(Val));
	var Tmp=""+rs("Result");
	rs.Close();
	rs=null;
	return Tmp;
}
function SendInvoiceEmail(Connection,TicketID,xDOrW,Lang)
{
	var DOrW=""+xDOrW;
	if (DOrW!="D")
		DOrW="W";
	var Email="";
	var CustomerName="";
	var PaymentMethod="";
	var ShippingCustomerName="";
	var Address="";
	var ShippingAddress="";
	var CreateDate="";
	var Currency="";
	var Discount=0.00;
	var SubTotal="";
	var Total="";
	var PST="";
	var GST="";
	var HST="";
	var InStorePickupStoreID="0";
	var Tmp="";
	var TTT="";
	var Tmp1=0;
	var Tmp2=0;
	var UsingNumber="1";
	var rs=Connection.Execute("select UsingNumber from POS_Main with (NOLOCK)");
	if (!rs.EOF)
		UsingNumber=""+rs("UsingNumber");
	rs.Close();
	rs=null;
	if (DOrW=="D")
		rs=Connection.Execute("select * from RealDirectShipTicketHD with (NOLOCK) where TicketID="+TicketID);
	else
		rs=Connection.Execute("select * from WebTicketHD with (NOLOCK) where TicketID="+TicketID);
	if (!rs.EOF)
	{
		Email=""+rs("Email").value;
		CustomerName=""+rs("FirstName").value+" "+rs("LastName").value;
		PaymentMethod=""+rs("PaymentMethod").value;
		ShippingCustomerName=""+rs("ShipFirstName").value+" "+rs("ShipLastName").value;
		Address=""+rs("Address").value;
		if (""+rs("City").value!="")
			Address+=", "+rs("City").value;
		if (""+rs("State").value!="")
			Address+=", "+rs("State").value;
		if (""+rs("Zip").value!="")
			Address+=", "+rs("Zip").value;
		Address+=", " + rs("Country").value;
		ShippingAddress=""+rs("ShipAddress").value;
		if (""+rs("ShipCity").value!="")
			ShippingAddress+=", "+rs("ShipCity").value;
		if (""+rs("ShipState").value!="")
			ShippingAddress+=", "+rs("ShipState").value;
		if (""+rs("ShipZip").value!="")
			ShippingAddress+=", "+rs("ShipZip").value;
		ShippingAddress+=", " + rs("ShipCountry").value;
		CreateDate="" + rs("CreateDate");
		Currency=""+rs("CurrencyID").value;
		SubTotal=DoubleToString(rs("SubTotal").value,true)+"."+DoubleToString(rs("SubTotal").value,false);
		PST=DoubleToString(rs("Tax1").value,true)+"."+DoubleToString(rs("Tax1").value,false);
		GST=DoubleToString(rs("Tax2").value,true)+"."+DoubleToString(rs("Tax2").value,false);
		HST=DoubleToString(rs("Tax3").value,true)+"."+DoubleToString(rs("Tax3").value,false);
		InStorePickupStoreID=""+rs("InStorePickupStoreID").value;
		Total=DoubleToString(rs("SubTotal").value+rs("Tax1").value+rs("Tax2").value+rs("Tax3").value,true)+"."+DoubleToString(rs("SubTotal").value+rs("Tax1").value+rs("Tax2").value+rs("Tax3").value,false);
	}
	rs.Close();
	rs=null;
	if ((Email!="")&&(EmailOK(Email)))
	{
		var SS,Subj;
		if (""+Lang=="F")
		{
			SS=ReadEmailTextFile("invoice_FR.xml");
			Subj="Service à la clientèle SoftMoc";
		}
		else
		{
			SS=ReadEmailTextFile("invoice.xml");
			Subj="Invoice";
		}
		while (SS.indexOf("#TicketID#",0)>=0)
			SS=ReplaceNow(SS,"#TicketID#",TicketID);
		while (SS.indexOf("#CustomerName#",0)>=0)
			SS=ReplaceNow(SS,"#CustomerName#",CustomerName);
		while (SS.indexOf("#Address#",0)>=0)
			SS=ReplaceNow(SS,"#Address#",Address);
		while (SS.indexOf("#PaymentMethod#",0)>=0)
			SS=ReplaceNow(SS,"#PaymentMethod#",PaymentMethod);
		while (SS.indexOf("#ShippingCustomerName#",0)>=0)
			SS=ReplaceNow(SS,"#ShippingCustomerName#",ShippingCustomerName);
		while (SS.indexOf("#ShippingAddress#",0)>=0)
			SS=ReplaceNow(SS,"#ShippingAddress#",ShippingAddress);
		while (SS.indexOf("#CreateDate#",0)>=0)
			SS=ReplaceNow(SS,"#CreateDate#",CreateDate);
		while (SS.indexOf("#Currency#",0)>=0)
			SS=ReplaceNow(SS,"#Currency#",Currency);
		while (SS.indexOf("#PST#",0)>=0)
			SS=ReplaceNow(SS,"#PST#",PST);
		while (SS.indexOf("#GST#",0)>=0)
			SS=ReplaceNow(SS,"#GST#",GST);
		while (SS.indexOf("#HST#",0)>=0)
			SS=ReplaceNow(SS,"#HST#",HST);
		while (SS.indexOf("#SubTotal#",0)>=0)
			SS=ReplaceNow(SS,"#SubTotal#",SubTotal);
		while (SS.indexOf("#Total#",0)>=0)
			SS=ReplaceNow(SS,"#Total#",Total);
		if (Discount=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#Discount_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#Discount_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#Discount_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		if (PST=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#PST_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#PST_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#PST_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		if (GST=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#GST_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#GST_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#GST_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		if (HST=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#HST_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#HST_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#HST_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		var ItemSection="";
		Tmp1=SS.indexOf("<"+"!--#Item_SECTION#+--"+">",0);
		Tmp2=SS.indexOf("<"+"!--#Item_SECTION#---"+">",0);
		if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
			ItemSection=SS.substring(Tmp1,Tmp2+("<"+"!--#Item_SECTION#---"+">").length);
		if ((Tmp!="")&&(SS.indexOf(ItemSection,0)>=0))
			SS=ReplaceNow(SS,ItemSection,"<"+"!--#ItemAddHere_SECTION#--"+">");
		if (DOrW=="D")
			rs=Connection.Execute("select d.*,i.Description from RealDirectShipTicketDT d with (NOLOCK),POS_Inventory"+UsingNumber+" i with (NOLOCK) where d.ItemID=i.ItemID and d.TicketID="+TicketID+" order by d.IKey");
		else
			rs=Connection.Execute("select d.*,i.Description from WebTicketDT d with (NOLOCK),POS_Inventory"+UsingNumber+" i with (NOLOCK) where d.ItemID=i.ItemID and d.TicketID="+TicketID+" order by d.IKey");
		if (!rs.EOF)
		{
			var ShippingFee=DoubleToString(rs("ShippingFee").value,true)+"."+DoubleToString(rs("ShippingFee").value,false);
			if (InStorePickupStoreID!="0")
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#","In-Store Pick Up - Free");
			}
			else if (rs("ShippingMethod").value=="")
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#","FREE");
			}
			else if (ShippingFee=="0.00")
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#",""+rs("ShippingMethod").value+"&nbsp;(FREE)");
			}
			else
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#",""+rs("ShippingMethod").value+"&nbsp;("+Currency+ShippingFee+")");
			}
		}
		while (!rs.EOF)
		{
			Discount+=parseFloat(DoubleToString(rs("Price").value*rs("Qty").value*rs("DiscountPercentage").value/100,true)+"."+DoubleToString(rs("Price").value*rs("Qty").value*rs("DiscountPercentage").value/100,false));
			Tmp=ItemSection+"\r\n<"+"!--#ItemAddHere_SECTION#--"+">";
			while (Tmp.indexOf("#ItemDesc#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemDesc#",""+rs("Description").value);
			TTT=""+rs("ItemID");
			TTT+=" (";
			if (""+rs("Parameter1")!="")
				TTT+=""+rs("Parameter1");
			if (""+rs("Parameter2")!="")
			{
				TTT+=", ";
				TTT+=""+rs("Parameter2");
			}
			TTT+=")";
			while (Tmp.indexOf("#ItemIDAndSize#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemIDAndSize#",TTT);
			while (Tmp.indexOf("#ItemQty#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemQty#",""+rs("Qty"));
			TTT=DoubleToString(rs("Price").value,true)+"."+DoubleToString(rs("Price").value,false);
			while (Tmp.indexOf("#ItemPrice#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemPrice#",TTT);
			if (SS.indexOf("<"+"!--#ItemAddHere_SECTION#--"+">",0)>=0)
				SS=ReplaceNow(SS,"<"+"!--#ItemAddHere_SECTION#--"+">",Tmp);
			rs.MoveNext();
		}
		Discount=DoubleToString(Discount,true)+"."+DoubleToString(Discount,false);
		while (SS.indexOf("#Discount#",0)>=0)
			SS=ReplaceNow(SS,"#Discount#",Discount);
		rs.Close();
		rs=null;
/*		var objNewMail = Server.CreateObject("CDONTS.NewMail");
		objNewMail.From = "softmoc_automail@softmoc.com (SoftMoc)";
		objNewMail.To = Email;
		objNewMail.Subject = "Invoice";
		objNewMail.BodyFormat = 0;
		objNewMail.MailFormat = 0;
		objNewMail.Body = SS;
		objNewMail.Send();
		objNewMail = null;
*/

		SendEmailByCDOReal("HTML","softmoc_automail@softmoc.com",Email,Subj,SS);

	}
}
function GetSecurityApprovalCode(RealTicketID)
{
	var buf=""+RealTicketID;
	var out1="";
	if (buf.length > 7)
	{
		out1=buf;
		out1=out1.substring(out1.length-7,out1.length);
		buf=out1;
	}
	out1="";
	var out2="";
	var tmp=0;
	for (var x=0; x < buf.length; x++)
	{
		buf1=buf.substring(x,x+1);
		var C=parseInt(buf1) + x + 1;
		C= C % 10;
		buf1=""+C;
		tmp+=C;
		if (x % 2 == 0)
			out1+=buf1;
		else
			out2+=buf1;
	}
	var TT="";
	for (x=out2.length; x > 0; x--)
	{
		TT+=out2.substring(x-1,x);
	}
	for (x=out1.length; x > 0; x--)
	{
		TT+=out1.substring(x-1,x);
	}
	return TT;
}
function GetIP()
{
	var IP="" + Request.ServerVariables("HTTP_X_REAL_IP");
	if ((IP=="")||(IP=="undefined"))
		IP="" + Request.ServerVariables("HTTP_INCAP_CLIENT_IP");
	if ((IP=="")||(IP=="undefined"))
		IP="" + Request.ServerVariables("HTTP_X_FORWARDED_FOR");
	if ((IP=="")||(IP=="undefined"))
		IP="" + Request.ServerVariables("REMOTE_ADDR");
	if (IP=="undefined")
		IP="";
	if (IP.indexOf(",",0)>=0)
	{
		var Tmp=IP;
		var i=0;
		while (i < Tmp.length)
		{
			IP="";
			for (; i < Tmp.length; i++)
			{
				if (((Tmp.substring(i,i+1) < '0') || (Tmp.substring(i,i+1) > '9')) && (Tmp.substring(i,i+1)!='.'))
					break;	
				IP+=Tmp.substring(i,i+1);
			}
			if (IP!="127.0.0.1")
				break;
			for (; i < Tmp.length; i++)
			{
				if (((Tmp.substring(i,i+1) >= '0') && (Tmp.substring(i,i+1) <= '9')) || (Tmp.substring(i,i+1)=='.'))
					break;	
			}
		}
	}
	return IP;
}
function SendShippingInvoiceEmail(ConnectionMain,Connection,TicketID)
{
	var ShippingFee="0.00";
	var ShipBy="";
	var IsRealGiftCardPurchase=false;
	var CompanyName="";
	var CompanyAddress="";
	var InvoiceFooterForRealGiftCardPurchase="";
	var InvoiceFooter="";
	var Email="";
	var CustomerName="";
	var PaymentMethod="";
	var ShippingCustomerName="";
	var Address="";
	var Phone="";
	var Fax="";
	var RealTicketID="";
	var ShipPhone="";
	var ShippingAddress="";
	var CreateDate="";
	var Currency="";
	var Discount="0.00";
	var SubTotal="";
	var Total="";
	var PST="";
	var GST="";
	var HST="";
	var InStorePickupStoreID="0";
	var Tmp="";
	var TTT="";
	var Tmp1=0;
	var Tmp2=0;
	var UsingNumber="1";
	var rs=ConnectionMain.Execute("select StoreName,Address,City,State,Zip,Country,InvoiceFooter,InvoiceFooterForRealGiftCardPurchase from WebServer with (NOLOCK)");
	if (!rs.EOF)
	{
		CompanyName=""+rs("StoreName");
		CompanyAddress=""+rs("Address");
		if (""+rs("Address")!="")
			CompanyAddress+=", ";
		CompanyAddress+=""+rs("City");
		if (""+rs("City")!="")
			CompanyAddress+=", ";
		CompanyAddress+=""+rs("State");
		if (""+rs("State")!="")
			CompanyAddress+=", ";
		CompanyAddress+=""+rs("Zip");
		if (""+rs("Zip")!="")
			CompanyAddress+=", ";
		CompanyAddress+=""+rs("Country");
		InvoiceFooterForRealGiftCardPurchase=""+rs("InvoiceFooterForRealGiftCardPurchase");
		InvoiceFooter=""+rs("InvoiceFooter");
	}
	rs.Close();
	rs=null;
	rs=Connection.Execute("select UsingNumber from POS_Main with (NOLOCK)");
	if (!rs.EOF)
		UsingNumber=""+rs("UsingNumber");
	rs.Close();
	rs=null;
	rs=Connection.Execute("select * from WebTicketHD with (NOLOCK) where TicketID="+TicketID);
	var Lang="E";
	if (!rs.EOF)
	{
		Lang=""+rs("SysLang").value;
		RealTicketID=""+DoubleToString(rs("RealTicketID").value,true);
		Phone=""+rs("Phone").value;
		Fax=""+rs("Fax").value;
		ShipPhone=""+rs("ShipPhone").value;
		Email=""+rs("Email").value;
		CustomerName=""+rs("FirstName").value+" "+rs("LastName").value;
		PaymentMethod=""+rs("PaymentMethod").value;
		ShippingCustomerName=""+rs("ShipFirstName").value+" "+rs("ShipLastName").value;
		Address=""+rs("Address").value;
		if (""+rs("City").value!="")
			Address+=", "+rs("City").value;
		if (""+rs("State").value!="")
			Address+=", "+rs("State").value;
		if (""+rs("Zip").value!="")
			Address+=", "+rs("Zip").value;
		Address+=", " + rs("Country").value;
		ShippingAddress=""+rs("ShipAddress").value;
		if (""+rs("ShipCity").value!="")
			ShippingAddress+=", "+rs("ShipCity").value;
		if (""+rs("ShipState").value!="")
			ShippingAddress+=", "+rs("ShipState").value;
		if (""+rs("ShipZip").value!="")
			ShippingAddress+=", "+rs("ShipZip").value;
		ShippingAddress+=", " + rs("ShipCountry").value;
		CreateDate="" + rs("CreateDate");
		Currency=""+rs("CurrencyID").value;
		SubTotal=DoubleToString(rs("SubTotal").value,true)+"."+DoubleToString(rs("SubTotal").value,false);
		PST=DoubleToString(rs("Tax1").value,true)+"."+DoubleToString(rs("Tax1").value,false);
		GST=DoubleToString(rs("Tax2").value,true)+"."+DoubleToString(rs("Tax2").value,false);
		HST=DoubleToString(rs("Tax3").value,true)+"."+DoubleToString(rs("Tax3").value,false);
		InStorePickupStoreID=""+rs("InStorePickupStoreID").value;
		Total=DoubleToString(rs("SubTotal").value+rs("Tax1").value+rs("Tax2").value+rs("Tax3").value,true)+"."+DoubleToString(rs("SubTotal").value+rs("Tax1").value+rs("Tax2").value+rs("Tax3").value,false);
	}
	rs.Close();
	rs=null;
	var FrequentBuyerCredit="";
	rs=Connection.Execute("select ItemID,FrequentBuyerCredit from WebTicketDT with (NOLOCK) where TicketID="+ToSQLInt(TicketID));
	if (!rs.EOF)
	{
		TTT=""+rs("ItemID");
		TTT=TTT.toUpperCase();
		if (TTT.substring(0,8)=="GIFTCARD")
			IsGiftPurchase=true;
		FrequentBuyerCredit=DoubleToString(""+rs("FrequentBuyerCredit"),true)+'.'+DoubleToString(""+rs("FrequentBuyerCredit"),false);
	}
	rs.Close();
	rs=null;
	if ((Email!="")&&(EmailOK(Email)))
	{
		var SS,Subj;
		if (Lang=="F")
		{
			SS=ReadEmailTextFile("shippinginvoice_FR.xml");
			Subj="Facture d’expédition de SoftMoc";
		}
		else
		{
			SS=ReadEmailTextFile("shippinginvoice.xml");
			Subj="Invoice";
		}
		while (SS.indexOf("#CompanyName#",0)>=0)
			SS=ReplaceNow(SS,"#CompanyName#",CompanyName);
		while (SS.indexOf("#CompanyAddress#",0)>=0)
			SS=ReplaceNow(SS,"#CompanyAddress#",CompanyAddress);
		while (SS.indexOf("#Phone#",0)>=0)
			SS=ReplaceNow(SS,"#Phone#",Phone);
		while (SS.indexOf("#ShipPhone#",0)>=0)
			SS=ReplaceNow(SS,"#ShipPhone#",ShipPhone);
		while (SS.indexOf("#Fax#",0)>=0)
			SS=ReplaceNow(SS,"#Fax#",Fax);
		while (SS.indexOf("#RetailTicketID#",0)>=0)
			SS=ReplaceNow(SS,"#RetailTicketID#",RealTicketID);
		while (SS.indexOf("#TicketID#",0)>=0)
			SS=ReplaceNow(SS,"#TicketID#",TicketID);
		while (SS.indexOf("#CustomerName#",0)>=0)
			SS=ReplaceNow(SS,"#CustomerName#",CustomerName);
		while (SS.indexOf("#Address#",0)>=0)
			SS=ReplaceNow(SS,"#Address#",Address);
		while (SS.indexOf("#PaymentMethod#",0)>=0)
			SS=ReplaceNow(SS,"#PaymentMethod#",PaymentMethod);
		while (SS.indexOf("#ShippingCustomerName#",0)>=0)
			SS=ReplaceNow(SS,"#ShippingCustomerName#",ShippingCustomerName);
		while (SS.indexOf("#ShippingAddress#",0)>=0)
			SS=ReplaceNow(SS,"#ShippingAddress#",ShippingAddress);
		while (SS.indexOf("#CreateDate#",0)>=0)
			SS=ReplaceNow(SS,"#CreateDate#",CreateDate);
		while (SS.indexOf("#Currency#",0)>=0)
			SS=ReplaceNow(SS,"#Currency#",Currency);
		while (SS.indexOf("#FrequentBuyerCredit#",0)>=0)
			SS=ReplaceNow(SS,"#FrequentBuyerCredit#",FrequentBuyerCredit);
		while (SS.indexOf("#PST#",0)>=0)
			SS=ReplaceNow(SS,"#PST#",PST);
		while (SS.indexOf("#GST#",0)>=0)
			SS=ReplaceNow(SS,"#GST#",GST);
		while (SS.indexOf("#HST#",0)>=0)
			SS=ReplaceNow(SS,"#HST#",HST);
		while (SS.indexOf("#SubTotal#",0)>=0)
			SS=ReplaceNow(SS,"#SubTotal#",SubTotal);
		while (SS.indexOf("#Total#",0)>=0)
			SS=ReplaceNow(SS,"#Total#",Total);
		while (SS.indexOf("#SecurityApprovalCode#",0)>=0)
			SS=ReplaceNow(SS,"#SecurityApprovalCode#","*"+GetSecurityApprovalCode(RealTicketID)+"*");
		if (FrequentBuyerCredit=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#FrequentBuyerCredit_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#FrequentBuyerCredit_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#FrequentBuyerCredit_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		if (PST=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#PST_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#PST_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#PST_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		if (GST=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#GST_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#GST_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#GST_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		if (HST=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#HST_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#HST_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#HST_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		var ItemSection="";
		Tmp1=SS.indexOf("<"+"!--#Item_SECTION#+--"+">",0);
		Tmp2=SS.indexOf("<"+"!--#Item_SECTION#---"+">",0);
		if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
			ItemSection=SS.substring(Tmp1,Tmp2+("<"+"!--#Item_SECTION#---"+">").length);
		if ((Tmp!="")&&(SS.indexOf(ItemSection,0)>=0))
			SS=ReplaceNow(SS,ItemSection,"<"+"!--#ItemAddHere_SECTION#--"+">");
		rs=Connection.Execute("select d.*,i.Description from WebTicketDT d with (NOLOCK),POS_Inventory"+UsingNumber+" i with (NOLOCK) where d.ItemID=i.ItemID and d.TicketID="+TicketID+" order by d.IKey");
		ShipBy=""+rs("ShippingMethod");
		if (!rs.EOF)
		{
			ShippingFee=DoubleToString(rs("ShippingFee").value,true)+"."+DoubleToString(rs("ShippingFee").value,false);
			if (InStorePickupStoreID!="0")
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#","In-Store Pick Up - Free");
			}
			else if (rs("ShippingMethod").value=="")
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#","FREE");
			}
			else if (ShippingFee=="0.00")
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#",""+rs("ShippingMethod").value+"&nbsp;(FREE)");
			}
			else
			{
				while (SS.indexOf("#ShipBy#",0)>=0)
					SS=ReplaceNow(SS,"#ShipBy#",""+rs("ShippingMethod").value+"&nbsp;("+Currency+ShippingFee+")");
			}
		}
		while (!rs.EOF)
		{
			Discount+=parseFloat(DoubleToString(rs("Price").value*rs("Qty").value*rs("DiscountPercentage").value/100,true)+"."+DoubleToString(rs("Price").value*rs("Qty").value*rs("DiscountPercentage").value/100,false));
			Tmp=ItemSection+"\r\n<"+"!--#ItemAddHere_SECTION#--"+">";
			while (Tmp.indexOf("#ItemDesc#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemDesc#",""+rs("Description").value);
			TTT=""+rs("ItemID");
			TTT=TTT.toUpperCase();
			if (TTT.substring(0,8)=="GIFTCARD")
				IsRealGiftCardPurchase=false;
			TTT=""+rs("ItemID");
			TTT+=" (";
			if (""+rs("Parameter1")!="")
				TTT+=""+rs("Parameter1");
			if (""+rs("Parameter2")!="")
			{
				TTT+=", ";
				TTT+=""+rs("Parameter2");
			}
			TTT+=")";
			while (Tmp.indexOf("#ItemID#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemID#",""+rs("ItemID"));
			while (Tmp.indexOf("#ItemSize#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemSize#",""+rs("Parameter1"));
			while (Tmp.indexOf("#ItemWidth#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemWidth#",""+rs("Parameter2"));
			while (Tmp.indexOf("#ItemIDAndSize#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemIDAndSize#",TTT);
			while (Tmp.indexOf("#ItemQty#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemQty#",""+rs("Qty"));
			TTT=DoubleToString(rs("Price").value,true)+"."+DoubleToString(rs("Price").value,false);
			while (Tmp.indexOf("#ItemPrice#",0)>=0)
				Tmp=ReplaceNow(Tmp,"#ItemPrice#",TTT);
			if (SS.indexOf("<"+"!--#ItemAddHere_SECTION#--"+">",0)>=0)
				SS=ReplaceNow(SS,"<"+"!--#ItemAddHere_SECTION#--"+">",Tmp);
			rs.MoveNext();
		}
		Discount=DoubleToString(Discount,true)+"."+DoubleToString(Discount,false);
		if (Discount=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#Discount_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#Discount_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#Discount_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		if (ShippingFee=="0.00")
		{
			Tmp="";
			Tmp1=SS.indexOf("<"+"!--#ShippingFee_SECTION#+--"+">",0);
			Tmp2=SS.indexOf("<"+"!--#ShippingFee_SECTION#---"+">",0);
			if ((Tmp1 >= 0)&&(Tmp2 >= 0)&&(Tmp2 > Tmp1))
				Tmp=SS.substring(Tmp1,Tmp2+("<"+"!--#ShippingFee_SECTION#---"+">").length);
			if ((Tmp!="")&&(SS.indexOf(Tmp,0)>=0))
				SS=ReplaceNow(SS,Tmp,"");
		}
		while (SS.indexOf("#ShippingFee#",0)>=0)
			SS=ReplaceNow(SS,"#ShippingFee#",ShippingFee);
		while (SS.indexOf("#ShipBy#",0)>=0)
			SS=ReplaceNow(SS,"#ShipBy#",ShipBy);
		while (SS.indexOf("#Discount#",0)>=0)
			SS=ReplaceNow(SS,"#Discount#",Discount);
		if (IsRealGiftCardPurchase)
		{
			while (SS.indexOf("#Footer#",0)>=0)
				SS=ReplaceNow(SS,"#Footer#",InvoiceFooterForRealGiftCardPurchase);
		}
		else
		{
			while (SS.indexOf("#Footer#",0)>=0)
				SS=ReplaceNow(SS,"#Footer#",InvoiceFooter);
		}
		rs.Close();
		rs=null;
/*		var objNewMail = Server.CreateObject("CDONTS.NewMail");
		objNewMail.From = "softmoc_automail@softmoc.com (SoftMoc)";
		objNewMail.To = Email;
		objNewMail.Subject = "Invoice";
		objNewMail.BodyFormat = 0;
		objNewMail.MailFormat = 0;
		objNewMail.Body = SS;
		objNewMail.Send();
		objNewMail = null;
*/
		SendEmailByCDOReal("HTML","softmoc_automail@softmoc.com",Email,Subj,SS);
	}
}
function SendCCEmail(Country,PayText,FirstName,LastName,NewTicketID,PostalCode,ShipAddress,ShipCity,
			ShipProvince,ShipPostalCode,ShipCountry,Currency,Total,CanTotal,Email,HaveGiftCard)
{

//	if ((PayText=="Amex")||(PayText=="Discover")||(PayText=="Mastercard")||
//				(PayText=="Cheque")||(PayText=="Visa")||(PayText=="MoneyOrder")||(PayText=="PayPal")||(PayText=="Interact"))
//	{
		if ((Email!="")&&(EmailOK(Email)))
		{
			var Subj;
			if (SysLang=="F")
			{
				SS=ReadEmailTextFile("emailconfirm_FR.xml");
				Subj="Confirmation d’expédition de SoftMoc";
			}
			else
			{
				SS=ReadEmailTextFile("emailconfirm.xml");
				Subj="Order Confirmation";
			}
			var buf="";
			if (PayText=="Amex")
				buf="American Express";
			else if (PayText=="Discover")
				buf="Discover";
			else if (PayText=="Mastercard")
				buf="Mastercard";
			else if (PayText=="Cheque")
				buf="Personal Cheque";
			else if (PayText=="Visa")
				buf="Visa";
			else if (PayText=="MoneyOrder")
				buf="Money Order";
			else if (PayText=="PayPal")
				buf="Paypal";
			else if (PayText=="Interact")
				buf="Interact";
			else
				buf=PayText;
			if (""+HaveGiftCard=="1")
			{
				if (buf!="")
					buf+=", ";
				buf+="Gift Card";
			}
			while (SS.indexOf("#PaymentMethod#",0)>=0)
				SS=ReplaceNow(SS,"#PaymentMethod#",buf);
			buf=""+FirstName+" "+LastName;
			while (SS.indexOf("#CustomerName#",0)>=0)
				SS=ReplaceNow(SS,"#CustomerName#",buf);
			buf=""+NewTicketID;
			while (SS.indexOf("#OrderNumber#",0)>=0)
				SS=ReplaceNow(SS,"#OrderNumber#",buf);
			buf="http://www.softmoc.com/CustomerJoin.asp?E=";
			buf+=MakeIt(Email);
			while (SS.indexOf("#CustomerJoinLink#",0)>=0)
				SS=ReplaceNow(SS,"#CustomerJoinLink#",buf);
			buf="http://www.softmoc.com/CustomerRemove.asp?E=";
			buf+=MakeIt(Email);
			while (SS.indexOf("#CustomerRemoveLink#",0)>=0)
				SS=ReplaceNow(SS,"#CustomerRemoveLink#",buf);
			buf=CheckOutHttps()+"://www.softmoc.com/" + Country +"/OrderTrackingreal.asp?T="+DoubleToString(NewTicketID,true);
			buf+="&L=";
			buf+=MakeIt(LastName);
			while (SS.indexOf("#OrderTrackingLink#",0)>=0)
				SS=ReplaceNow(SS,"#OrderTrackingLink#",buf);
			buf="" + ShipAddress;
			if (ShipCity!="")
			{
				buf+=", ";
				buf+=ShipCity;
			}
			if (ShipProvince!="")
			{
				buf+=", ";
				buf+=ShipProvince;
			}
			if (ShipPostalCode!="")
			{
				buf+=", ";
				buf+=ShipPostalCode;
			}
			if (ShipCountry!="")
			{
				buf+=", ";
				buf+=ShipCountry;
			}
			while (SS.indexOf("#ShippingAddress#",0)>=0)
				SS=ReplaceNow(SS,"#ShippingAddress#",buf);
			buf="" + Currency+ DoubleToString(Total,true)+ "." +DoubleToString(Total,false);
			while (SS.indexOf("#TotalLocalPurchase#",0)>=0)
				SS=ReplaceNow(SS,"#TotalLocalPurchase#",buf);
			buf="CA$"+ DoubleToString(CanTotal,true)+ "." +DoubleToString(CanTotal,false);
			while (SS.indexOf("#TotalPurchase#",0)>=0)
				SS=ReplaceNow(SS,"#TotalPurchase#",buf);
/*			var objNewMail = Server.CreateObject("CDONTS.NewMail");
			objNewMail.From = "softmoc_automail@softmoc.com (SoftMoc)";
			objNewMail.To = Email;
			objNewMail.Subject = "Order Confirmation";
			objNewMail.BodyFormat = 0;
			objNewMail.MailFormat = 1;
			objNewMail.Body = SS;
			objNewMail.Send();
			objNewMail = null;
*/
			SendEmailByCDOReal("TEXT","softmoc_automail@softmoc.com",Email,Subj,SS);
		}
//	}
}
function BriteCode_ValidateEmail(Email)
{
	var URL="https://bpi.briteverify.com/emails.json?address=";
	URL+=Email;
	URL+="&apikey=";
	
	URL+="122a8367-6796-4407-8a4b-3430bbc7d298";

	var Result="";
	for (var xx=1; xx <= 3; xx++)
	{
		var objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1");
		try
		{
			objHttp.Open("GET", URL, false);
			WinHttpRequestOption_SslErrorIgnoreFlags = 4;
			objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 0x3300;
			objHttp.Send("");
		}
		catch (Err)
		{
			return "Validate Email Exception calling (" + Tmp + "): Message=" + Err.message + ", Description=" + Err.description;
		}
		var Tmp= ""+ objHttp.ResponseText;
		var Pos=Tmp.indexOf("\"status\":",0);
		Result="";
		if (Pos>=0)
		{
			for (var i=Pos+10;i< Tmp.length; i++)
			{
				if (Tmp.substring(i,i+1)=="\"")
					break;
				Result+=Tmp.substring(i,i+1);
			}
		}

		/*result value
		==============
		valid: The email represents a real account / inbox available at the given domain
		invalid: Not a real email
		unknown: For some reason we cannot verify valid or invalid. Most of the time a domain did not respond quickly enough.
		accept_all: These are domains that respond to every verification request in the affirmative, and therefore cannot be fully verified.
		*/

		if (Result!="unknown")
			break;
	}
	return Result;
}
function CheckBarcodeFormat(xBarcode)
{
	var Barcode=""+xBarcode;
	var L=Barcode.length;
	if ((L != 12)&&(L != 13))
		return false;
	var Odd=0;
	var Even=0;
	for (var i=1; i <= L;i++)
	{
		var Tmp=Barcode.substring(i-1,i);
		if (i % 2 == 0)//even
			Even+=parseInt(Tmp);
		else
			Odd+=parseInt(Tmp);
	}
	if ((L==12)&& (((Odd*3) + Even)  %  10 ==0))
		return true;
	else if ((L==13)&& (((Even*3) + Odd)  %  10 ==0))
		return true;
	return false;
}
function CountFile(TicketID,Buffer)
{
	var OutString="";
	var BBB="";
	var V=0;
	var InV=0;
	var VNotGood=0;
	var recordSet=null;
	var TheResult=0;
	var Connection = Server.CreateObject("ADODB.Connection");
	Connection.Open(OpenCon(Application("IP"), Application("Login"), Application("Password"),Application("DatabaseName")));
	Connection.Execute("Begin Tran");
	for (var x=0;((x< Buffer.length)||(BBB!=""));x++)
	{
		if ((x< Buffer.length)&&(Buffer.substring(x,x+1)>='0')&&(BBB.substring(x,x+1)<='9'))
			BBB+=Buffer.substring(x,x+1);
		else
		{
			if (BBB!="")
			{
				if (!CheckBarcodeFormat(BBB))
				{
					BBB="";
					InV++;
					continue; 
				}
				var X="UPCToCount "+TicketID;
				X+=","+ToSQL(BBB);
				recordSet=Connection.Execute(X);
				if (recordSet.EOF)
					VNotGood++;
				else
				{					
					TheResult=recordSet('TheResult').value;
					if (TheResult==1)
						V++;
					else
						VNotGood++;
				}
				recordSet.Close();
				recordSet = null;
			}
			BBB="";
		}
	}
	Connection.Execute("delete CountTicketDetailDT where Qty=0 and OldQty=0 and CountTicketID="+TicketID);
	Connection.Execute("Commit Tran");
	Connection.Close();
	OutString="@1@Valid UPC (with valid item) : "+V;
	OutString+="\r\nValid UPC (without valid item) : "+VNotGood;
	OutString+="\r\nInvalid UPC : "+InV;
	return OutString;
}
function ChangeCCNumberToStar(s)
{
	var LenS=s.length;
	if (LenS <= 10)
		return s;
	var out="";
	for (var i=1;i <= LenS;i++)
	{
		if ((i <= 4)||(i > LenS-4))
			out+=s.substring(i-1,i);
		else
			out+='*';
	}
	return out;
}
/*function DoingD(num)
{
	if (num=='~!@')
		return '.';
	else if (num=='9')
		return '0';
	else
		return 0+parseInt(num)+1;
}
function IsValidEmail(Email)
{
	if (Email=="")
		return false;
	var a=Email.indexOf("@",0);
	var b=Email.indexOf(".",0);
	if (a< 0)
		return false;
	if (b< 0)
		return false;
	if (a>b)
		return false;
	return true;
}
function InternalCall(x)
{
 var y="";
 if (x==1)
  y="&#65;bc";
 else if (x==2)
  y="123";
 else if (x==3)
  y="qqq";
 else if (x==4)
  y="ddd"
 return y;
}
function MakeSearchRString()
{
	var oArgs = MakeSearchRString.arguments;
	var szRet = "";
	if (oArgs.length <= 0)
		return "SearchR.asp";
	var Country=oArgs[0].toString().toLowerCase();
	var Tmp="SearchR";
	if (Country!="us")
		Tmp+=".asp";
	var CASomething=false;
	var i=1;
	while (i < oArgs.length)
	{
		var Tmp1="";
		if (oArgs[i].toString().length!=1)
		{
			i++;
			continue;
		}
		Tmp1="" + oArgs[i].toString();
		i++;
		if (i >= oArgs.length)
			break;
		var Tmp2="" + oArgs[i].toString();
		i++;
		if (Country=="us")
			Tmp+= "_" + Tmp1 + Tmp2;
		else
		{
			if (!CASomething)
			{
				CASomething=true;
				Tmp+="?";
			}
			else
				Tmp+="&";
			Tmp+= Tmp1 + "=" + Tmp2;
		}
	}
	if (Country=="us")
		Tmp+=".asp";
	return Tmp;
}
function URLName()
{
	return "" + Request.ServerVariables("SERVER_NAME");
}
function SaveImage(Type,Parameter1,Buffer)
{
	var x = Server.CreateObject("Scripting.FileSystemObject");
	var y=x.CreateTextFile("c:\\SoftPOS\\MainServer\\Items\\Images" + Parameter1);
	y.WriteLine(Buffer);
	y.Close();
	x=null;
	return true;
}

function SaveVendorImage(Type,Parameter1,Buffer)
{
	var x = Server.CreateObject("Scripting.FileSystemObject");
	var y=x.CreateTextFile("c:\\SoftPOS\\MainServer\\Vendors\\Images" + Parameter1);
	y.WriteLine(Buffer);
	y.Close();
	x=null;
	return true;
}
function MakeTranFileName(Input)
{
	var Tmp = "";
	for (var i=Input.length;i < 12;i=i+1)
		Tmp = Tmp + "0";
	Tmp = Tmp + Input;
	Tmp = Tmp + "_001";
	return Tmp;
}

function LocalSaveImageNow(Type,Parameter1,Buffer,ThePath)
{
	var x = Server.CreateObject("Scripting.FileSystemObject");
	var y=x.CreateTextFile(ThePath + Parameter1,true);
	y.WriteLine(Buffer);
	y.Close();
	x=null;
	return true;
}
*/
function toRad(value) 
{
	// convert degrees to radians
  return value * Math.PI / 180;
}

function CalDistance(lat1, lon1, lat2, lon2)
{
	if ((lat1==0 && lon1==0)||(lat2==0 && lon2==0))
		return -1; //i.e. UNKNOWN

	var R = 6371; // km
	var dLat = toRad(lat2-lat1);
	var dLon = toRad(lon2-lon1); 
	var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
			Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * 
			Math.sin(dLon/2) * Math.sin(dLon/2); 
	var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)); 
	var d = R * c;
	
	return d;
}
function CheckMobile()    //return     "M", "IPAD", "IPHONE"
{  
	var IsMobile=false;

	var H_UserAgent=""+ Request.ServerVariables("HTTP_USER_AGENT");
	H_UserAgent=H_UserAgent.toLowerCase();
	var H_Accept=""+ Request.ServerVariables("HTTP_ACCEPT");
	H_Accept=H_Accept.toLowerCase();
	var H_AllHttp=""+ Request.ServerVariables("ALL_HTTP");
	H_AllHttp=H_AllHttp.toLowerCase();

	if ((H_UserAgent.indexOf("android",0)>=0)
				||(H_UserAgent.indexOf("up.browser",0)>=0)
				||(H_UserAgent.indexOf("up.link",0)>=0)
				||(H_UserAgent.indexOf("mmp",0)>=0)
				||(H_UserAgent.indexOf("symbian",0)>=0)
				||(H_UserAgent.indexOf("smartphone",0)>=0)
				||(H_UserAgent.indexOf("midp",0)>=0)
				||(H_UserAgent.indexOf("wap",0)>=0)
				||(H_UserAgent.indexOf("phone",0)>=0)

				||(H_UserAgent.indexOf("windows ce",0)>=0)
				||(H_UserAgent.indexOf("avantgo",0)>=0)
				||(H_UserAgent.indexOf("mazingo",0)>=0)
				||(H_UserAgent.indexOf("mobile",0)>=0)
				||(H_UserAgent.indexOf("t68",0)>=0)
				||(H_UserAgent.indexOf("syncalot",0)>=0)
				||(H_UserAgent.indexOf("blazer",0)>=0)

				||(H_UserAgent.indexOf("ipad",0)>=0)
				||(H_UserAgent.indexOf("ipod",0)>=0)
				||(H_UserAgent.indexOf("iphone",0)>=0)
				||(H_UserAgent.indexOf("android",0)>=0)
				||(H_UserAgent.indexOf("opera mini",0)>=0)
				||(H_UserAgent.indexOf("blackberry",0)>=0)
 				||(H_UserAgent.indexOf("palm",0)>=0)
 				||(H_UserAgent.indexOf("hiptop",0)>=0)
  				||(H_UserAgent.indexOf("plucker",0)>=0)
 				||(H_UserAgent.indexOf("xiino",0)>=0)
 				||(H_UserAgent.indexOf("blazer",0)>=0)
 				||(H_UserAgent.indexOf("elaine",0)>=0)
 				||(H_UserAgent.indexOf("iris",0)>=0)
 				||(H_UserAgent.indexOf("3g_t",0)>=0)
 				||(H_UserAgent.indexOf("opera mobi",0)>=0)
 				||(H_UserAgent.indexOf("mini 9.5",0)>=0)
 				||(H_UserAgent.indexOf("vx1000",0)>=0)
 				||(H_UserAgent.indexOf("lge ",0)>=0)
 				||(H_UserAgent.indexOf("m800",0)>=0)
 				||(H_UserAgent.indexOf("e860",0)>=0)
 				||(H_UserAgent.indexOf("u940",0)>=0)
 				||(H_UserAgent.indexOf("ux840",0)>=0)
 				||(H_UserAgent.indexOf("compal",0)>=0)
 				||(H_UserAgent.indexOf("wireless",0)>=0)
 				||(H_UserAgent.indexOf(" mobi",0)>=0)
 				||(H_UserAgent.indexOf("ahong",0)>=0)
 				||(H_UserAgent.indexOf("lg380",0)>=0)
 				||(H_UserAgent.indexOf("lgku",0)>=0)
 				||(H_UserAgent.indexOf("lgu900",0)>=0)
 				||(H_UserAgent.indexOf("lg210",0)>=0)
 				||(H_UserAgent.indexOf("lg47",0)>=0)
 				||(H_UserAgent.indexOf("lg920",0)>=0)
 				||(H_UserAgent.indexOf("lg840",0)>=0)
 				||(H_UserAgent.indexOf("lg370",0)>=0)
 				||(H_UserAgent.indexOf("sam-r",0)>=0)
 				||(H_UserAgent.indexOf("mg50",0)>=0)
 				||(H_UserAgent.indexOf("s55",0)>=0)
 				||(H_UserAgent.indexOf("g83",0)>=0)
 				||(H_UserAgent.indexOf("t66",0)>=0)
 				||(H_UserAgent.indexOf("vx400",0)>=0)
 				||(H_UserAgent.indexOf("mk99",0)>=0)
 				||(H_UserAgent.indexOf("d615",0)>=0)
 				||(H_UserAgent.indexOf("d763",0)>=0)
 				||(H_UserAgent.indexOf("el370",0)>=0)
 				||(H_UserAgent.indexOf("sl900",0)>=0)
 				||(H_UserAgent.indexOf("mp500",0)>=0)
 				||(H_UserAgent.indexOf("samu3",0)>=0)
 				||(H_UserAgent.indexOf("samu4",0)>=0)
 				||(H_UserAgent.indexOf("vx10",0)>=0)
 				||(H_UserAgent.indexOf("xda_",0)>=0)
 				||(H_UserAgent.indexOf("samu5",0)>=0)
 				||(H_UserAgent.indexOf("samu6",0)>=0)
 				||(H_UserAgent.indexOf("samu7",0)>=0)
 				||(H_UserAgent.indexOf("samu9",0)>=0)
 				||(H_UserAgent.indexOf("a615",0)>=0)
 				||(H_UserAgent.indexOf("b832",0)>=0)
 				||(H_UserAgent.indexOf("m881",0)>=0)
 				||(H_UserAgent.indexOf("s920",0)>=0)
 				||(H_UserAgent.indexOf("n210",0)>=0)
 				||(H_UserAgent.indexOf("s700",0)>=0)
 				||(H_UserAgent.indexOf("c-810",0)>=0)
 				||(H_UserAgent.indexOf("_h797",0)>=0)
 				||(H_UserAgent.indexOf("mob-x",0)>=0)
 				||(H_UserAgent.indexOf("sk16d",0)>=0)
 				||(H_UserAgent.indexOf("848b",0)>=0)
 				||(H_UserAgent.indexOf("mowser",0)>=0)
 				||(H_UserAgent.indexOf("s580",0)>=0)
 				||(H_UserAgent.indexOf("r800",0)>=0)
 				||(H_UserAgent.indexOf("471x",0)>=0)
 				||(H_UserAgent.indexOf("v120",0)>=0)
 				||(H_UserAgent.indexOf("rim8",0)>=0)
 				||(H_UserAgent.indexOf("c500foma:",0)>=0)
 				||(H_UserAgent.indexOf("160x",0)>=0)
 				||(H_UserAgent.indexOf("x160",0)>=0)
 				||(H_UserAgent.indexOf("480x",0)>=0)
 				||(H_UserAgent.indexOf("x640",0)>=0)
 				||(H_UserAgent.indexOf("t503",0)>=0)
 				||(H_UserAgent.indexOf("w839",0)>=0)
 				||(H_UserAgent.indexOf("i250",0)>=0)
 				||(H_UserAgent.indexOf("sprint",0)>=0)
 				||(H_UserAgent.indexOf("w398samr810",0)>=0)
 				||(H_UserAgent.indexOf("m5252",0)>=0)
 				||(H_UserAgent.indexOf("c7100",0)>=0)
 				||(H_UserAgent.indexOf("mt126",0)>=0)
 				||(H_UserAgent.indexOf("x225",0)>=0)
 				||(H_UserAgent.indexOf("s5330",0)>=0)
 				||(H_UserAgent.indexOf("s820",0)>=0)
 				||(H_UserAgent.indexOf("htil-g1",0)>=0)
 				||(H_UserAgent.indexOf("fly v71",0)>=0)
 				||(H_UserAgent.indexOf("s302",0)>=0)
 				||(H_UserAgent.indexOf("-x113",0)>=0)
 				||(H_UserAgent.indexOf("novarra",0)>=0)
 				||(H_UserAgent.indexOf("k610i",0)>=0)
 				||(H_UserAgent.indexOf("-three",0)>=0)
 				||(H_UserAgent.indexOf("8325rc",0)>=0)
 				||(H_UserAgent.indexOf("8352rc",0)>=0)
 				||(H_UserAgent.indexOf("sanyo",0)>=0)
 				||(H_UserAgent.indexOf("vx54",0)>=0)
 				||(H_UserAgent.indexOf("c888",0)>=0)
 				||(H_UserAgent.indexOf("nx250",0)>=0)
 				||(H_UserAgent.indexOf("n120",0)>=0)
 				||(H_UserAgent.indexOf("mtk ",0)>=0)
 				||(H_UserAgent.indexOf("c5588",0)>=0)
 				||(H_UserAgent.indexOf("s710",0)>=0)
 				||(H_UserAgent.indexOf("t880",0)>=0)
 				||(H_UserAgent.indexOf("c5005",0)>=0)
 				||(H_UserAgent.indexOf("i;458x",0)>=0)
 				||(H_UserAgent.indexOf("p404i",0)>=0)
 				||(H_UserAgent.indexOf("s210",0)>=0)
 				||(H_UserAgent.indexOf("c5100",0)>=0)
 				||(H_UserAgent.indexOf("teleca",0)>=0)
 				||(H_UserAgent.indexOf("s940",0)>=0)
 				||(H_UserAgent.indexOf("c500",0)>=0)
 				||(H_UserAgent.indexOf("s590",0)>=0)
 				||(H_UserAgent.indexOf("foma",0)>=0)
 				||(H_UserAgent.indexOf("samsu",0)>=0)
 				||(H_UserAgent.indexOf("vx8",0)>=0)
 				||(H_UserAgent.indexOf("vx9",0)>=0)
 				||(H_UserAgent.indexOf("a1000",0)>=0)
 				||(H_UserAgent.indexOf("_mms",0)>=0)
 				||(H_UserAgent.indexOf("myx",0)>=0)
 				||(H_UserAgent.indexOf("a700",0)>=0)
 				||(H_UserAgent.indexOf("gu1100",0)>=0)
 				||(H_UserAgent.indexOf("bc831",0)>=0)
 				||(H_UserAgent.indexOf("e300",0)>=0)
 				||(H_UserAgent.indexOf("ems100",0)>=0)
 				||(H_UserAgent.indexOf("me701",0)>=0)
 				||(H_UserAgent.indexOf("me702m-three",0)>=0)
 				||(H_UserAgent.indexOf("sd588",0)>=0)
 				||(H_UserAgent.indexOf("s800",0)>=0)
 				||(H_UserAgent.indexOf("8325rc",0)>=0)
 				||(H_UserAgent.indexOf("ac831",0)>=0)
 				||(H_UserAgent.indexOf("mw200",0)>=0)
 				||(H_UserAgent.indexOf("brew ",0)>=0)
 				||(H_UserAgent.indexOf("d88",0)>=0)
 				||(H_UserAgent.indexOf("htc",0)>=0)
 				||(H_UserAgent.indexOf("355x",0)>=0)
 				||(H_UserAgent.indexOf("m50",0)>=0)
 				||(H_UserAgent.indexOf("km100",0)>=0)
 				||(H_UserAgent.indexOf("d736",0)>=0)
 				||(H_UserAgent.indexOf("p-9521",0)>=0)
 				||(H_UserAgent.indexOf("telco",0)>=0)
 				||(H_UserAgent.indexOf("sl74",0)>=0)
 				||(H_UserAgent.indexOf("ktouch",0)>=0)
 				||(H_UserAgent.indexOf("me702",0)>=0)
 				||(H_UserAgent.indexOf("8325rc",0)>=0)
 				||(H_UserAgent.indexOf("kddi",0)>=0)
 				||(H_UserAgent.indexOf("phone",0)>=0)
 				||(H_UserAgent.indexOf("lg ",0)>=0)
 				||(H_UserAgent.indexOf("sonyericsson",0)>=0)
 				||(H_UserAgent.indexOf("samsung",0)>=0)
 				||(H_UserAgent.indexOf("240x",0)>=0)
 				||(H_UserAgent.indexOf("x320",0)>=0)
 				||(H_UserAgent.indexOf("vx10",0)>=0)
 				||(H_UserAgent.indexOf("nokia",0)>=0)
 				||(H_UserAgent.indexOf("sony cmd",0)>=0)
 				||(H_UserAgent.indexOf("motorola",0)>=0)
 				||(H_UserAgent.indexOf("up.browser",0)>=0)
 				||(H_UserAgent.indexOf("up.link",0)>=0)
 				||(H_UserAgent.indexOf("mmp",0)>=0)
 				||(H_UserAgent.indexOf("symbian",0)>=0)
 				||(H_UserAgent.indexOf("smartphone",0)>=0)
 				||(H_UserAgent.indexOf("midp",0)>=0)
 				||(H_UserAgent.indexOf("wap",0)>=0)
 				||(H_UserAgent.indexOf("vodafone",0)>=0)
 				||(H_UserAgent.indexOf("o2",0)>=0)
 				||(H_UserAgent.indexOf("pocket",0)>=0)
 				||(H_UserAgent.indexOf("kindle",0)>=0)
 				||(H_UserAgent.indexOf("mobile",0)>=0)
 				||(H_UserAgent.indexOf("psp",0)>=0)
 				||(H_UserAgent.indexOf("treo",0)>=0)

				||(H_Accept.indexOf("text/vnd.wap.wml",0)>=0)
				||(H_Accept.indexOf("application/vnd.wap.xhtml+xml",0)>=0)

				||(H_UserAgent.substring(0,3)=="1207")
				||(H_UserAgent.substring(0,3)=="3gso")
				||(H_UserAgent.substring(0,3)=="4thp")
				||(H_UserAgent.substring(0,3)=="501i")
				||(H_UserAgent.substring(0,3)=="502i")
				||(H_UserAgent.substring(0,3)=="503i")
				||(H_UserAgent.substring(0,3)=="504i")
				||(H_UserAgent.substring(0,3)=="505i")
				||(H_UserAgent.substring(0,3)=="506i")
				||(H_UserAgent.substring(0,3)=="6310")
				||(H_UserAgent.substring(0,3)=="6590")
				||(H_UserAgent.substring(0,3)=="770s")
				||(H_UserAgent.substring(0,3)=="802s")
				||(H_UserAgent.substring(0,3)=="a wa")
				||(H_UserAgent.substring(0,3)=="abac")
				||(H_UserAgent.substring(0,3)=="acer")
				||(H_UserAgent.substring(0,3)=="acs-")
				||(H_UserAgent.substring(0,3)=="acoo")
				||(H_UserAgent.substring(0,3)=="aiko")
				||(H_UserAgent.substring(0,3)=="airn")
				||(H_UserAgent.substring(0,3)=="alav")
				||(H_UserAgent.substring(0,3)=="alca")
				||(H_UserAgent.substring(0,3)=="alco")
				||(H_UserAgent.substring(0,3)=="amoi")
				||(H_UserAgent.substring(0,3)=="anex")
				||(H_UserAgent.substring(0,3)=="anny")
				||(H_UserAgent.substring(0,3)=="anyw")
				||(H_UserAgent.substring(0,3)=="aptu")
				||(H_UserAgent.substring(0,3)=="arch")
				||(H_UserAgent.substring(0,3)=="argo")
				||(H_UserAgent.substring(0,3)=="aste")
				||(H_UserAgent.substring(0,3)=="asus")
				||(H_UserAgent.substring(0,3)=="andr")
				||(H_UserAgent.substring(0,3)=="attw")
				||(H_UserAgent.substring(0,3)=="audi")
				||(H_UserAgent.substring(0,3)=="au-m")
				||(H_UserAgent.substring(0,3)=="aur ")
				||(H_UserAgent.substring(0,3)=="aus ")
				||(H_UserAgent.substring(0,3)=="avan")
				||(H_UserAgent.substring(0,3)=="beck")
				||(H_UserAgent.substring(0,3)=="bell")
				||(H_UserAgent.substring(0,3)=="benq")
				||(H_UserAgent.substring(0,3)=="bilb")
				||(H_UserAgent.substring(0,3)=="bird")
				||(H_UserAgent.substring(0,3)=="blac")
				||(H_UserAgent.substring(0,3)=="blaz")
				||(H_UserAgent.substring(0,3)=="brew")
				||(H_UserAgent.substring(0,3)=="brvw")
				||(H_UserAgent.substring(0,3)=="bumb")
				||(H_UserAgent.substring(0,3)=="bw-n")
				||(H_UserAgent.substring(0,3)=="bw-u")
				||(H_UserAgent.substring(0,3)=="c55/")
				||(H_UserAgent.substring(0,3)=="capi")
				||(H_UserAgent.substring(0,3)=="ccwa")
				||(H_UserAgent.substring(0,3)=="cdm-")
				||(H_UserAgent.substring(0,3)=="cell")
				||(H_UserAgent.substring(0,3)=="chtm")
				||(H_UserAgent.substring(0,3)=="cldc")
				||(H_UserAgent.substring(0,3)=="cmd-")
				||(H_UserAgent.substring(0,3)=="cond")
				||(H_UserAgent.substring(0,3)=="craw")
				||(H_UserAgent.substring(0,3)=="dait")
				||(H_UserAgent.substring(0,3)=="dall")
				||(H_UserAgent.substring(0,3)=="dang")
				||(H_UserAgent.substring(0,3)=="dbte")
				||(H_UserAgent.substring(0,3)=="dc-s")
				||(H_UserAgent.substring(0,3)=="devi")
				||(H_UserAgent.substring(0,3)=="dica")
				||(H_UserAgent.substring(0,3)=="dmob")
				||(H_UserAgent.substring(0,3)=="doco")
				||(H_UserAgent.substring(0,3)=="dopo")
				||(H_UserAgent.substring(0,3)=="ds12")
				||(H_UserAgent.substring(0,3)=="ds-d")
				||(H_UserAgent.substring(0,3)=="el49")
				||(H_UserAgent.substring(0,3)=="eric")
				||(H_UserAgent.substring(0,3)=="eml2")
				||(H_UserAgent.substring(0,3)=="emul")
				||(H_UserAgent.substring(0,3)=="elai")
				||(H_UserAgent.substring(0,3)=="eric")
				||(H_UserAgent.substring(0,3)=="erk0")
				||(H_UserAgent.substring(0,3)=="es18")
				||(H_UserAgent.substring(0,3)=="ez40")
				||(H_UserAgent.substring(0,3)=="ez60")
				||(H_UserAgent.substring(0,3)=="ez70")
				||(H_UserAgent.substring(0,3)=="ezos")
				||(H_UserAgent.substring(0,3)=="ezwa")
				||(H_UserAgent.substring(0,3)=="ezze")
				||(H_UserAgent.substring(0,3)=="fake")
				||(H_UserAgent.substring(0,3)=="fetc")
				||(H_UserAgent.substring(0,3)=="fly-")
				||(H_UserAgent.substring(0,3)=="fly_")
				||(H_UserAgent.substring(0,3)=="g-mo")
				||(H_UserAgent.substring(0,3)=="g1 u")
				||(H_UserAgent.substring(0,3)=="g560")
				||(H_UserAgent.substring(0,3)=="gene")
				||(H_UserAgent.substring(0,3)=="gf-5")
				||(H_UserAgent.substring(0,3)=="go.w")
				||(H_UserAgent.substring(0,3)=="good")
				||(H_UserAgent.substring(0,3)=="grad")
				||(H_UserAgent.substring(0,3)=="grun")
				||(H_UserAgent.substring(0,3)=="haie")
				||(H_UserAgent.substring(0,3)=="hcit")
				||(H_UserAgent.substring(0,3)=="hd-m")
				||(H_UserAgent.substring(0,3)=="hd-p")
				||(H_UserAgent.substring(0,3)=="hd-t")
				||(H_UserAgent.substring(0,3)=="hei-")
				||(H_UserAgent.substring(0,3)=="hiba")
				||(H_UserAgent.substring(0,3)=="hita")
				||(H_UserAgent.substring(0,3)=="hipt")
				||(H_UserAgent.substring(0,3)=="hp i")
				||(H_UserAgent.substring(0,3)=="hpip")
				||(H_UserAgent.substring(0,3)=="hs-c")
				||(H_UserAgent.substring(0,3)=="htc ")
				||(H_UserAgent.substring(0,3)=="htc-")
				||(H_UserAgent.substring(0,3)=="htc_")
				||(H_UserAgent.substring(0,3)=="htca")
				||(H_UserAgent.substring(0,3)=="htcg")
				||(H_UserAgent.substring(0,3)=="htcp")
				||(H_UserAgent.substring(0,3)=="htcs")
				||(H_UserAgent.substring(0,3)=="htct")
				||(H_UserAgent.substring(0,3)=="huaw")
				||(H_UserAgent.substring(0,3)=="hutc")
				||(H_UserAgent.substring(0,3)=="i-20")
				||(H_UserAgent.substring(0,3)=="i-go")
				||(H_UserAgent.substring(0,3)=="i-ma")
				||(H_UserAgent.substring(0,3)=="i230")
				||(H_UserAgent.substring(0,3)=="iac-")
				||(H_UserAgent.substring(0,3)=="iac/")
				||(H_UserAgent.substring(0,3)=="ibro")
				||(H_UserAgent.substring(0,3)=="idea")
				||(H_UserAgent.substring(0,3)=="ig01")
				||(H_UserAgent.substring(0,3)=="ikom")
				||(H_UserAgent.substring(0,3)=="im1k")
				||(H_UserAgent.substring(0,3)=="inno")
				||(H_UserAgent.substring(0,3)=="iris")
				||(H_UserAgent.substring(0,3)=="ipaq")
				||(H_UserAgent.substring(0,3)=="jata")
				||(H_UserAgent.substring(0,3)=="java")
				||(H_UserAgent.substring(0,3)=="jbro")
				||(H_UserAgent.substring(0,3)=="jemu")
				||(H_UserAgent.substring(0,3)=="jigs")
				||(H_UserAgent.substring(0,3)=="kddi")
				||(H_UserAgent.substring(0,3)=="keji")
				||(H_UserAgent.substring(0,3)=="kgt/")
				||(H_UserAgent.substring(0,3)=="klon")
				||(H_UserAgent.substring(0,3)=="kpt ")
				||(H_UserAgent.substring(0,3)=="kwc-")
				||(H_UserAgent.substring(0,3)=="kyoc")
				||(H_UserAgent.substring(0,3)=="kyok")
				||(H_UserAgent.substring(0,3)=="leno")
				||(H_UserAgent.substring(0,3)=="lexi")
				||(H_UserAgent.substring(0,3)=="lg/a")
				||(H_UserAgent.substring(0,3)=="lg/b")
				||(H_UserAgent.substring(0,3)=="lg/c")
				||(H_UserAgent.substring(0,3)=="lg/d")
				||(H_UserAgent.substring(0,3)=="lg/f")
				||(H_UserAgent.substring(0,3)=="lg/g")
				||(H_UserAgent.substring(0,3)=="lg/k")
				||(H_UserAgent.substring(0,3)=="lg/l")
				||(H_UserAgent.substring(0,3)=="lg/m")
				||(H_UserAgent.substring(0,3)=="lg/o")
				||(H_UserAgent.substring(0,3)=="lg/p")
				||(H_UserAgent.substring(0,3)=="lg/s")
				||(H_UserAgent.substring(0,3)=="lg/t")
				||(H_UserAgent.substring(0,3)=="lg/u")
				||(H_UserAgent.substring(0,3)=="lg/w")
				||(H_UserAgent.substring(0,3)=="lg50")
				||(H_UserAgent.substring(0,3)=="lg54")
				||(H_UserAgent.substring(0,3)=="lg-a")
				||(H_UserAgent.substring(0,3)=="lg-b")
				||(H_UserAgent.substring(0,3)=="lg-c")
				||(H_UserAgent.substring(0,3)=="lg-d")
				||(H_UserAgent.substring(0,3)=="lg-f")
				||(H_UserAgent.substring(0,3)=="lg-g")
				||(H_UserAgent.substring(0,3)=="lg-k")
				||(H_UserAgent.substring(0,3)=="lg-l")
				||(H_UserAgent.substring(0,3)=="lg-m")
				||(H_UserAgent.substring(0,3)=="lg-o")
				||(H_UserAgent.substring(0,3)=="lg-p")
				||(H_UserAgent.substring(0,3)=="lg-s")
				||(H_UserAgent.substring(0,3)=="lg-t")
				||(H_UserAgent.substring(0,3)=="lg-u")
				||(H_UserAgent.substring(0,3)=="lg-w")
				||(H_UserAgent.substring(0,3)=="lg a")
				||(H_UserAgent.substring(0,3)=="lg b")
				||(H_UserAgent.substring(0,3)=="lg c")
				||(H_UserAgent.substring(0,3)=="lg d")
				||(H_UserAgent.substring(0,3)=="lg f")
				||(H_UserAgent.substring(0,3)=="lg g")
				||(H_UserAgent.substring(0,3)=="lg k")
				||(H_UserAgent.substring(0,3)=="lg l")
				||(H_UserAgent.substring(0,3)=="lg m")
				||(H_UserAgent.substring(0,3)=="lg o")
				||(H_UserAgent.substring(0,3)=="lg p")
				||(H_UserAgent.substring(0,3)=="lg s")
				||(H_UserAgent.substring(0,3)=="lg t")
				||(H_UserAgent.substring(0,3)=="lg u")
				||(H_UserAgent.substring(0,3)=="lg w")
				||(H_UserAgent.substring(0,3)=="lge-")
				||(H_UserAgent.substring(0,3)=="lge/")
				||(H_UserAgent.substring(0,3)=="libw")
				||(H_UserAgent.substring(0,3)=="lynx")
				||(H_UserAgent.substring(0,3)=="m1-w")
				||(H_UserAgent.substring(0,3)=="m3ga")
				||(H_UserAgent.substring(0,3)=="m50/")
				||(H_UserAgent.substring(0,3)=="m-cr")
				||(H_UserAgent.substring(0,3)=="mate")
				||(H_UserAgent.substring(0,3)=="maui")
				||(H_UserAgent.substring(0,3)=="maxo")
				||(H_UserAgent.substring(0,3)=="mc01")
				||(H_UserAgent.substring(0,3)=="mc21")
				||(H_UserAgent.substring(0,3)=="mcca")
				||(H_UserAgent.substring(0,3)=="medi")
				||(H_UserAgent.substring(0,3)=="merc")
				||(H_UserAgent.substring(0,3)=="meri")
				||(H_UserAgent.substring(0,3)=="midp")
				||(H_UserAgent.substring(0,3)=="mio8")
				||(H_UserAgent.substring(0,3)=="mioa")
				||(H_UserAgent.substring(0,3)=="mits")
				||(H_UserAgent.substring(0,3)=="mmef")
				||(H_UserAgent.substring(0,3)=="mo01")
				||(H_UserAgent.substring(0,3)=="mo02")
				||(H_UserAgent.substring(0,3)=="mobi")
				||(H_UserAgent.substring(0,3)=="mode")
				||(H_UserAgent.substring(0,3)=="modo")
				||(H_UserAgent.substring(0,3)=="mot ")
				||(H_UserAgent.substring(0,3)=="mot-")
				||(H_UserAgent.substring(0,3)=="moto")
				||(H_UserAgent.substring(0,3)=="motv")
				||(H_UserAgent.substring(0,3)=="mozz")
				||(H_UserAgent.substring(0,3)=="mt50")
				||(H_UserAgent.substring(0,3)=="mtp1")
				||(H_UserAgent.substring(0,3)=="mtv ")
				||(H_UserAgent.substring(0,3)=="mwbp")
				||(H_UserAgent.substring(0,3)=="mywa")
				||(H_UserAgent.substring(0,3)=="n100")
				||(H_UserAgent.substring(0,3)=="n101")
				||(H_UserAgent.substring(0,3)=="n102")
				||(H_UserAgent.substring(0,3)=="n202")
				||(H_UserAgent.substring(0,3)=="n203")
				||(H_UserAgent.substring(0,3)=="n300")
				||(H_UserAgent.substring(0,3)=="n302")
				||(H_UserAgent.substring(0,3)=="n500")
				||(H_UserAgent.substring(0,3)=="n502")
				||(H_UserAgent.substring(0,3)=="n505")
				||(H_UserAgent.substring(0,3)=="n700")
				||(H_UserAgent.substring(0,3)=="n710")
				||(H_UserAgent.substring(0,3)=="nec-")
				||(H_UserAgent.substring(0,3)=="nem-")
				||(H_UserAgent.substring(0,3)=="neon")
				||(H_UserAgent.substring(0,3)=="netf")
				||(H_UserAgent.substring(0,3)=="newg")
				||(H_UserAgent.substring(0,3)=="newt")
				||(H_UserAgent.substring(0,3)=="noki")
				||(H_UserAgent.substring(0,3)=="nok6")
				||(H_UserAgent.substring(0,3)=="nzph")
				||(H_UserAgent.substring(0,3)=="o2 x")
				||(H_UserAgent.substring(0,3)=="o2-x")
				||(H_UserAgent.substring(0,3)=="o2im")
				||(H_UserAgent.substring(0,3)=="oper")
				||(H_UserAgent.substring(0,3)=="opti")
				||(H_UserAgent.substring(0,3)=="opwv")
				||(H_UserAgent.substring(0,3)=="oran")
				||(H_UserAgent.substring(0,3)=="owg1")
				||(H_UserAgent.substring(0,3)=="p800")
				||(H_UserAgent.substring(0,3)=="palm")
				||(H_UserAgent.substring(0,3)=="pana")
				||(H_UserAgent.substring(0,3)=="pand")
				||(H_UserAgent.substring(0,3)=="pant")
				||(H_UserAgent.substring(0,3)=="pdxg")
				||(H_UserAgent.substring(0,3)=="pg-1")
				||(H_UserAgent.substring(0,3)=="pg-2")
				||(H_UserAgent.substring(0,3)=="pg-3")
				||(H_UserAgent.substring(0,3)=="pg-6")
				||(H_UserAgent.substring(0,3)=="pg-8")
				||(H_UserAgent.substring(0,3)=="pg-c")
				||(H_UserAgent.substring(0,3)=="pg13")
				||(H_UserAgent.substring(0,3)=="phil")
				||(H_UserAgent.substring(0,3)=="pire")
				||(H_UserAgent.substring(0,3)=="play")
				||(H_UserAgent.substring(0,3)=="pluc")
				||(H_UserAgent.substring(0,3)=="pn-2")
				||(H_UserAgent.substring(0,3)=="pock")
				||(H_UserAgent.substring(0,3)=="port")
				||(H_UserAgent.substring(0,3)=="pose")
				||(H_UserAgent.substring(0,3)=="prox")
				||(H_UserAgent.substring(0,3)=="psio")
				||(H_UserAgent.substring(0,3)=="pt-g")
				||(H_UserAgent.substring(0,3)=="qa-a")
				||(H_UserAgent.substring(0,3)=="qc-2")
				||(H_UserAgent.substring(0,3)=="qc-3")
				||(H_UserAgent.substring(0,3)=="qc-5")
				||(H_UserAgent.substring(0,3)=="qc-7")
				||(H_UserAgent.substring(0,3)=="qc07")
				||(H_UserAgent.substring(0,3)=="qc12")
				||(H_UserAgent.substring(0,3)=="qc21")
				||(H_UserAgent.substring(0,3)=="qc32")
				||(H_UserAgent.substring(0,3)=="qc60")
				||(H_UserAgent.substring(0,3)=="qci-")
				||(H_UserAgent.substring(0,3)=="qtek")
				||(H_UserAgent.substring(0,3)=="qwap")
				||(H_UserAgent.substring(0,3)=="r380")
				||(H_UserAgent.substring(0,3)=="r600")
				||(H_UserAgent.substring(0,3)=="raks")
				||(H_UserAgent.substring(0,3)=="rim9")
				||(H_UserAgent.substring(0,3)=="rove")
				||(H_UserAgent.substring(0,3)=="rozo")
				||(H_UserAgent.substring(0,3)=="sage")
				||(H_UserAgent.substring(0,3)=="sama")
				||(H_UserAgent.substring(0,3)=="sams")
				||(H_UserAgent.substring(0,3)=="sany")
				||(H_UserAgent.substring(0,3)=="sava")
				||(H_UserAgent.substring(0,3)=="sch-")
				||(H_UserAgent.substring(0,3)=="scoo")
				||(H_UserAgent.substring(0,3)=="sc01")
				||(H_UserAgent.substring(0,3)=="scp-")
				||(H_UserAgent.substring(0,3)=="sdk/")
				||(H_UserAgent.substring(0,3)=="s55/")
				||(H_UserAgent.substring(0,3)=="sec-")
				||(H_UserAgent.substring(0,3)=="sec0")
				||(H_UserAgent.substring(0,3)=="sec1")
				||(H_UserAgent.substring(0,3)=="se47")
				||(H_UserAgent.substring(0,3)=="semc")
				||(H_UserAgent.substring(0,3)=="send")
				||(H_UserAgent.substring(0,3)=="seri")
				||(H_UserAgent.substring(0,3)=="sgh-")
				||(H_UserAgent.substring(0,3)=="shar")
				||(H_UserAgent.substring(0,3)=="sie-")
				||(H_UserAgent.substring(0,3)=="siem")
				||(H_UserAgent.substring(0,3)=="sk-0")
				||(H_UserAgent.substring(0,3)=="slid")
				||(H_UserAgent.substring(0,3)=="sl45")
				||(H_UserAgent.substring(0,3)=="smal")
				||(H_UserAgent.substring(0,3)=="smar")
				||(H_UserAgent.substring(0,3)=="smit")
				||(H_UserAgent.substring(0,3)=="smb3")
				||(H_UserAgent.substring(0,3)=="smt5")
				||(H_UserAgent.substring(0,3)=="soft")
				||(H_UserAgent.substring(0,3)=="sony")
				||(H_UserAgent.substring(0,3)=="sp01")
				||(H_UserAgent.substring(0,3)=="sph-")
				||(H_UserAgent.substring(0,3)=="spv ")
				||(H_UserAgent.substring(0,3)=="spv-")
				||(H_UserAgent.substring(0,3)=="sy01")
				||(H_UserAgent.substring(0,3)=="symb")
				||(H_UserAgent.substring(0,3)=="t-mo")
				||(H_UserAgent.substring(0,3)=="t218")
				||(H_UserAgent.substring(0,3)=="t250")
				||(H_UserAgent.substring(0,3)=="t600")
				||(H_UserAgent.substring(0,3)=="t610")
				||(H_UserAgent.substring(0,3)=="t618")
				||(H_UserAgent.substring(0,3)=="tagt")
				||(H_UserAgent.substring(0,3)=="talk")
				||(H_UserAgent.substring(0,3)=="tcl-")
				||(H_UserAgent.substring(0,3)=="tdg-")
				||(H_UserAgent.substring(0,3)=="teli")
				||(H_UserAgent.substring(0,3)=="telm")
				||(H_UserAgent.substring(0,3)=="tim-")
				||(H_UserAgent.substring(0,3)=="topl")
				||(H_UserAgent.substring(0,3)=="tosh")
				||(H_UserAgent.substring(0,3)=="treo")
				||(H_UserAgent.substring(0,3)=="ts70")
				||(H_UserAgent.substring(0,3)=="tsm-")
				||(H_UserAgent.substring(0,3)=="tsm3")
				||(H_UserAgent.substring(0,3)=="tsm5")
				||(H_UserAgent.substring(0,3)=="tx-9")
				||(H_UserAgent.substring(0,3)=="up.b")
				||(H_UserAgent.substring(0,3)=="upg1")
				||(H_UserAgent.substring(0,3)=="upsi")
				||(H_UserAgent.substring(0,3)=="utst")
				||(H_UserAgent.substring(0,3)=="v400")
				||(H_UserAgent.substring(0,3)=="v750")
				||(H_UserAgent.substring(0,3)=="veri")
				||(H_UserAgent.substring(0,3)=="vk-v")
				||(H_UserAgent.substring(0,3)=="virg")
				||(H_UserAgent.substring(0,3)=="vite")
				||(H_UserAgent.substring(0,3)=="voda")
				||(H_UserAgent.substring(0,3)=="vulc")
				||(H_UserAgent.substring(0,3)=="vk-v")
				||(H_UserAgent.substring(0,3)=="vk40")
				||(H_UserAgent.substring(0,3)=="vk50")
				||(H_UserAgent.substring(0,3)=="vk52")
				||(H_UserAgent.substring(0,3)=="vk53")
				||(H_UserAgent.substring(0,3)=="vm40")
				||(H_UserAgent.substring(0,3)=="vx52")
				||(H_UserAgent.substring(0,3)=="vx53")
				||(H_UserAgent.substring(0,3)=="vx60")
				||(H_UserAgent.substring(0,3)=="vx61")
				||(H_UserAgent.substring(0,3)=="vx70")
				||(H_UserAgent.substring(0,3)=="vx80")
				||(H_UserAgent.substring(0,3)=="vx81")
				||(H_UserAgent.substring(0,3)=="vx83")
				||(H_UserAgent.substring(0,3)=="vx85")
				||(H_UserAgent.substring(0,3)=="vx98")
				||(H_UserAgent.substring(0,3)=="w3c ")
				||(H_UserAgent.substring(0,3)=="w3c-")
				||(H_UserAgent.substring(0,3)=="wap-")
				||(H_UserAgent.substring(0,3)=="wapa")
				||(H_UserAgent.substring(0,3)=="wapi")
				||(H_UserAgent.substring(0,3)=="wapj")
				||(H_UserAgent.substring(0,3)=="wapm")
				||(H_UserAgent.substring(0,3)=="wapp")
				||(H_UserAgent.substring(0,3)=="wapr")
				||(H_UserAgent.substring(0,3)=="waps")
				||(H_UserAgent.substring(0,3)=="wapt")
				||(H_UserAgent.substring(0,3)=="wapu")
				||(H_UserAgent.substring(0,3)=="wapv")
				||(H_UserAgent.substring(0,3)=="wapy")
				||(H_UserAgent.substring(0,3)=="webc")
				||(H_UserAgent.substring(0,3)=="whit")
				||(H_UserAgent.substring(0,3)=="winc")
				||(H_UserAgent.substring(0,3)=="winw")
				||(H_UserAgent.substring(0,3)=="wig ")
				||(H_UserAgent.substring(0,3)=="wmlb")
				||(H_UserAgent.substring(0,3)=="wonu")
				||(H_UserAgent.substring(0,3)=="x700")
				||(H_UserAgent.substring(0,3)=="xda ")
				||(H_UserAgent.substring(0,3)=="xda-")
				||(H_UserAgent.substring(0,3)=="xda2")
				||(H_UserAgent.substring(0,3)=="xdag")
				||(H_UserAgent.substring(0,3)=="yas-")
				||(H_UserAgent.substring(0,3)=="your")
				||(H_UserAgent.substring(0,3)=="zte-")
				||(H_UserAgent.substring(0,3)=="zeto")

				||(H_AllHttp.indexOf("operamini",0)>=0))
		IsMobile=true;

	if ((""+ Request.ServerVariables("HTTP_X_WAP_PROFILE")!="undefined")
				||(""+ Request.ServerVariables("HTTP_PROFILE")!="undefined"))
		IsMobile=true;

	if ((H_UserAgent.indexOf("windows",0)>=0)&&(H_UserAgent.indexOf("windows ce",0)<0))
		IsMobile=false;

	if (IsMobile)
	{
		if (H_UserAgent.indexOf("ipad",0)>=0)
			return "IPAD";
		else if (H_UserAgent.indexOf("iphone",0)>=0)
			return "IPHONE";
		else
			return "M";
	}
	else
		return "";
}
function GetInBetween(S,Start,End)
{
	var TmpA=S.indexOf(Start,0);
	var TmpB=S.indexOf(End,0);
	if ((TmpA < 0)||(TmpB <=0)||(TmpA > TmpB))
		return "";
	else
		return S.substring(TmpA+Start.length,TmpB);
}

function GetProvinceByPostalCode(xPostalCode)
{
	var Tmp=(""+xPostalCode).toLowerCase();
	var PostalCode="";
	for (var i=0; i < Tmp.length; i++)
	{
		if (Tmp.substring(i,i+1)!=" ")
			PostalCode+=Tmp.substring(i,i+1);
	}
	if (PostalCode.length!=6)
		return "";
	if ((PostalCode.substring(0,1) < 'a')||(PostalCode.substring(0,1) > 'z')
			||(PostalCode.substring(1,2) < '0')||(PostalCode.substring(1,2) > '9')
			||(PostalCode.substring(2,3) < 'a')||(PostalCode.substring(2,3) > 'z')
			||(PostalCode.substring(3,4) < '0')||(PostalCode.substring(3,4) > '9')
			||(PostalCode.substring(4,5) < 'a')||(PostalCode.substring(4,5) > 'z')
			||(PostalCode.substring(5,6) < '0')||(PostalCode.substring(5,6) > '9'))
		return "";
	if (PostalCode.substring(0,1)=="t")
		return "AB";
	else if (PostalCode.substring(0,1)=="v")
		return "BC";
	else if (PostalCode.substring(0,1)=="r")
		return "MB";
	else if (PostalCode.substring(0,1)=="e")
		return "NB";
	else if (PostalCode.substring(0,1)=="a")
		return "NL";
//	else if (PostalCode.substring(0,1)=="x")
//		return "NT";
	else if (PostalCode.substring(0,1)=="b")
		return "NS";
	else if (PostalCode.substring(0,1)=="x")
		return "NU";
	else if ((PostalCode.substring(0,1)=="k")||(PostalCode.substring(0,1)=="l")||(PostalCode.substring(0,1)=="m")||(PostalCode.substring(0,1)=="n")||(PostalCode.substring(0,1)=="p"))
		return "ON";
	else if (PostalCode.substring(0,1)=="c")
		return "PE";
	else if ((PostalCode.substring(0,1)=="g")||(PostalCode.substring(0,1)=="h")||(PostalCode.substring(0,1)=="j"))
		return "QC";
	else if (PostalCode.substring(0,1)=="s")
		return "SK";
	else if (PostalCode.substring(0,1)=="y")
		return "YT";
	else
		return "";
}

function UPS_XMLCall(AccessLicenseNumber,UserId,Password,ShipperNumber,ShipperAddress,ShipperCity,ShipperStateProvinceCode,ShipperPostalCode,ShipperCountryCode,
				ShipToAddress,ShipToCity,ShipToStateProvinceCode,ShipToPostalCode,ShipToCountryCode,Weight)
{
	var URL="https://onlinetools.ups.com/ups.app/xml/Rate";
//	var AccessLicenseNumber="DC93324A68A19D30";
//	var UserId="Mwfgroupbahamas";
//	var Password="1234abcd!";
	
	if (ShipperNumber=="")
		ShipperNumber="33A88F";

	ShipperStateProvinceCode=ShipperStateProvinceCode.toUpperCase();
	ShipToStateProvinceCode=ShipToStateProvinceCode.toUpperCase();
	if (ShipperPostalCode=="")
	{
		ShipperAddress="1400 Hopkins Street";
		ShipperCity="Whitby";
		ShipperStateProvinceCode="ON";
		ShipperPostalCode="L1N2C3";
		ShipperCountryCode="CA";
	}

	ShipperNumber=ShipperNumber.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipperAddress=ShipperAddress.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipperCity=ShipperCity.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipperStateProvinceCode=ShipperStateProvinceCode.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipperPostalCode=ShipperPostalCode.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipperCountryCode=ShipperCountryCode.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipToAddress=ShipToAddress.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipToCity=ShipToCity.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipToStateProvinceCode=ShipToStateProvinceCode.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipToPostalCode=ShipToPostalCode.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
	ShipToCountryCode=ShipToCountryCode.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");

	var strXML="<?xml version='1.0' ?>\r\n";
	strXML+="<AccessRequest xml:lang='en-US'>\r\n";
		strXML+="<AccessLicenseNumber>" + AccessLicenseNumber + "</AccessLicenseNumber>\r\n";
		strXML+="<UserId>" + UserId +"</UserId>\r\n";
		strXML+="<Password>" + Password +"</Password>\r\n";
	strXML+="</AccessRequest>\r\n";
	strXML+="<?xml version='1.0' ?>\r\n";
	strXML+="<RatingServiceSelectionRequest>\r\n";
		strXML+="<Request>\r\n";
			strXML+="<TransactionReference>\r\n";
				strXML+="<CustomerContext>Rating and Service</CustomerContext>\r\n";
				strXML+="<XpciVersion>1.0</XpciVersion>\r\n";
			strXML+="</TransactionReference>\r\n";
			strXML+="<RequestAction>Rate</RequestAction>\r\n";
			strXML+="<RequestOption>Shop</RequestOption>\r\n";
		strXML+="</Request>\r\n";
		strXML+="<PickupType>\r\n";
			strXML+="<Code>01</Code>\r\n";
			strXML+="<Description>Daily Pickup</Description>\r\n";
		strXML+="</PickupType>\r\n";
		strXML+="<Shipment>\r\n";
			strXML+="<Description>Rate Shopping</Description>\r\n";
			strXML+="<RateInformation><NegotiatedRatesIndicator>Y</NegotiatedRatesIndicator></RateInformation>\r\n";
			strXML+="<Shipper>\r\n";
				strXML+="<ShipperNumber>" + ShipperNumber + "</ShipperNumber>\r\n";
				strXML+="<Address>\r\n";
					strXML+="<AddressLine1>" + ShipperAddress + "</AddressLine1>\r\n";
					strXML+="<AddressLine2 />\r\n";
					strXML+="<AddressLine3 />\r\n";
					strXML+="<City>" + ShipperCity + "</City>\r\n";
					strXML+="<StateProvinceCode>" + ShipperStateProvinceCode + "</StateProvinceCode>\r\n";
					strXML+="<PostalCode>" + ShipperPostalCode + "</PostalCode>\r\n";
					strXML+="<CountryCode>" + ShipperCountryCode + "</CountryCode>\r\n";
				strXML+="</Address>\r\n";
			strXML+="</Shipper>\r\n";
			strXML+="<ShipTo>\r\n";
				strXML+="<CompanyName></CompanyName>\r\n";
				strXML+="<AttentionName></AttentionName>\r\n";
				strXML+="<PhoneNumber></PhoneNumber>\r\n";
				strXML+="<Address>\r\n";
					strXML+="<AddressLine1>"+ ShipToAddress +"</AddressLine1>\r\n";
					strXML+="<AddressLine2 />\r\n";
					strXML+="<AddressLine3 />\r\n";
					strXML+="<City>"+ ShipToCity +"</City>\r\n";
					strXML+="<StateProvinceCode>" + ShipToStateProvinceCode + "</StateProvinceCode>\r\n";
					strXML+="<PostalCode>"+ ShipToPostalCode +"</PostalCode>\r\n";
					strXML+="<CountryCode>"+ ShipToCountryCode +"</CountryCode>\r\n";
					strXML+="<ResidentialAddressIndicator>true</ResidentialAddressIndicator>\r\n";
				strXML+="</Address>\r\n";
			strXML+="</ShipTo>\r\n";
			strXML+="<ShipFrom>\r\n";
				strXML+="<CompanyName></CompanyName>\r\n";
				strXML+="<AttentionName></AttentionName>\r\n";
				strXML+="<PhoneNumber></PhoneNumber>\r\n";
				strXML+="<FaxNumber></FaxNumber>\r\n";
				strXML+="<Address>\r\n";
					strXML+="<AddressLine1>" + ShipperAddress + "</AddressLine1>\r\n";
					strXML+="<AddressLine2 />\r\n";
					strXML+="<AddressLine3 />\r\n";
					strXML+="<City>" + ShipperCity + "</City>\r\n";
					strXML+="<StateProvinceCode>" + ShipperStateProvinceCode + "</StateProvinceCode>\r\n";
					strXML+="<PostalCode>" + ShipperPostalCode + "</PostalCode>\r\n";
					strXML+="<CountryCode>" + ShipperCountryCode + "</CountryCode>\r\n";
				strXML+="</Address>\r\n";
			strXML+="</ShipFrom>\r\n";
			

/*++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Valid domestic values:
	14 = Next Day Air Early AM,
	01 = Next Day Air,
	13 = Next Day Air Saver,
	59 = 2nd Day Air AM,
	02 = 2nd Day Air,
	12 = 3 Day Select,
	03 = Ground.
Valid international values:
	11= Standard,
	07 = Worldwide Express,
	54 = Worldwide Express Plus,
	08 = Worldwide Expedited,
	65 = Saver. Required for Rating and Ignored for Shopping.
Valid Poland to Poland Same Day values:
	82 = UPS Today Standard,
	83 = UPS Today Dedicated Courier,
	84 = UPS Today Intercity,
	85 = UPS Today Express,
	86 = UPS Today Express Saver
*/
///////////////			strXML+="<Service><Code>01</Code></Service>\r\n";
//--------------------------------------------------------------------------------------------


			strXML+="<Package>\r\n";
				strXML+="<PackagingType>\r\n";


/*++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Valid values:
00 = UNKNOWN;
01 = UPS Letter;
02 = Package;
03 = Tube;
04 = Pak;
21 = Express Box;
24 = 25KG Box;
25 = 10KG Box;
30 = Pallet;
2a = Small Express Box;
22b = Medium Express Box;
2c = Large Express Box
*/
					strXML+="<Code>02</Code>\r\n";
//--------------------------------------------------------------------------------------------


					strXML+="<Description></Description>\r\n";
				strXML+="</PackagingType>\r\n";
				strXML+="<Description>Rate</Description>\r\n";
				strXML+="<PackageWeight>\r\n";
					strXML+="<UnitOfMeasurement>\r\n";
						strXML+="<Code>LBS</Code>\r\n";
					strXML+="</UnitOfMeasurement>\r\n";
					strXML+="<Weight>" + Weight + "</Weight>\r\n";
				strXML+="</PackageWeight>\r\n";
			strXML+="</Package>\r\n";
			strXML+="<ShipmentServiceOptions />\r\n";
		strXML+="</Shipment>\r\n";
	strXML+="</RatingServiceSelectionRequest>\r\n";

	var objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1");
	try
	{
	    objHttp.Open("POST", URL, false);
	    WinHttpRequestOption_SslErrorIgnoreFlags = 4;
	    objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 0x3300;
	    objHttp.Send(strXML);
	}
	catch (Err)
	{
	    return "0@UPS_XML_call() Exception calling UPS (" + URL + "): Message=" + Err.message + ", Description=" + Err.description;
	}
	var Res= objHttp.ResponseText;

	var Tmp=GetInBetween(Res,"<ResponseStatusCode>","</ResponseStatusCode>");
	if (Tmp!="1")
	{
		Tmp=GetInBetween(Res,"<ErrorDescription>","</ErrorDescription>");
		if (Tmp=="")
			return "0@@ErrMsg+@Error@ErrMsg-@";
		else
			return "0@@ErrMsg+@"+Tmp+"@ErrMsg-@";
	}
	var Out="1@@ErrMsg+@@ErrMsg-@";
	var OutCode="";
	var OutAmount="";
	var i=0;
	while (true)
	{
		Tmp=GetInBetween(Res,"<RatedShipment>","</RatedShipment>");
		if (Tmp=="")
			break;
		Res=Res.substring(Res.indexOf("</RatedShipment>",0)+"</RatedShipment>".length,Res.length);

		OutCode=GetInBetween(Tmp,"<Service>","</Service>");
		if (OutCode=="")
			continue;
		OutCode=GetInBetween(OutCode,"<Code>","</Code>");
		if (OutCode=="")
			continue;

		OutAmount=GetInBetween(Tmp,"<NegotiatedRates>","</NegotiatedRates>");
		if (OutAmount!="")
		{
			OutAmount=GetInBetween(OutAmount,"<NetSummaryCharges>","</NetSummaryCharges>");
			if (OutAmount!="")
			{
				OutAmount=GetInBetween(OutAmount,"<GrandTotal>","</GrandTotal>");
				if (OutAmount!="")
					OutAmount=GetInBetween(OutAmount,"<MonetaryValue>","</MonetaryValue>");
			}
		}
		if (OutAmount=="")
		{
			OutAmount=GetInBetween(Tmp,"<TotalCharges>","</TotalCharges>");
			if (OutAmount!="")
				OutAmount=GetInBetween(OutAmount,"<MonetaryValue>","</MonetaryValue>");
		}
		if (OutAmount=="")
			continue;

		i++;
		Out+="@Code"+i+"+@";
		Out+=OutCode;
		Out+="@Code"+i+"-@";
		Out+="@Amt"+i+"+@";
		Tmp="" + DoubleToString(OutAmount,true)+"."+DoubleToString(OutAmount,false);
		if (parseFloat(Tmp) <= 0.00)
			Tmp="0.00";			
		Out+=Tmp;
		Out+="@Amt"+i+"-@";
	}








	while (true)
	{
		Tmp=GetInBetween(Res,"<RatedShipment>","</RatedShipment>");
		if (Tmp=="")
			break;
		Res=Res.substring(Res.indexOf("</RatedShipment>",0)+"</RatedShipment>".length,Res.length);

		OutCode=GetInBetween(Tmp,"<Service>","</Service>");
		if (OutCode=="")
			continue;
		OutCode=GetInBetween(OutCode,"<Code>","</Code>");
		if (OutCode=="")
			continue;

		OutAmount=GetInBetween(Tmp,"<TotalCharges>","</TotalCharges>");
		if (OutAmount=="")
			continue;
		OutAmount=GetInBetween(OutAmount,"<MonetaryValue>","</MonetaryValue>");
		if (OutAmount=="")
			continue;

		i++;
		Out+="@Code"+i+"+@";
		Out+=OutCode;
		Out+="@Code"+i+"-@";
		Out+="@Amt"+i+"+@";
		Tmp="" + DoubleToString(OutAmount,true)+"."+DoubleToString(OutAmount,false);
		if (parseFloat(Tmp) <= 0.00)
			Tmp="0.00";			
		Out+=Tmp;
		Out+="@Amt"+i+"-@";
	}
	return Out;

/*  Response Example
<?xml version="1.0"?>
<RatingServiceSelectionResponse>
	<Response>
		<TransactionReference>
			<CustomerContext>Rating and Service</CustomerContext>
			<XpciVersion>1.0</XpciVersion>
		</TransactionReference>
		<ResponseStatusCode>1</ResponseStatusCode>
		<ResponseStatusDescription>Success</ResponseStatusDescription>
	</Response>
	<RatedShipment>
		<Service><Code>11</Code></Service>
		<RatedShipmentWarning>Your invoice may vary from the displayed reference rates</RatedShipmentWarning>
		<BillingWeight>
			<UnitOfMeasurement><Code>LBS</Code></UnitOfMeasurement>
			<Weight>1.0</Weight>
		</BillingWeight>
		<TransportationCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>14.09</MonetaryValue>
		</TransportationCharges>
		<ServiceOptionsCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>0.00</MonetaryValue>
		</ServiceOptionsCharges>
		<TotalCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>14.09</MonetaryValue>
		</TotalCharges>
		<GuaranteedDaysToDelivery/>
		<ScheduledDeliveryTime/>
		<RatedPackage>
			<TransportationCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TransportationCharges>
			<ServiceOptionsCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</ServiceOptionsCharges>
			<TotalCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TotalCharges>
			<Weight>1.0</Weight>
			<BillingWeight>
				<UnitOfMeasurement>
					<Code/>
				</UnitOfMeasurement>
				<Weight/>
			</BillingWeight>
		</RatedPackage>
	</RatedShipment>
	<RatedShipment>
		<Service><Code>02</Code></Service>
		<RatedShipmentWarning>Your invoice may vary from the displayed reference rates</RatedShipmentWarning>
		<BillingWeight>
			<UnitOfMeasurement><Code>LBS</Code></UnitOfMeasurement>
			<Weight>1.0</Weight>
		</BillingWeight>
		<TransportationCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>14.09</MonetaryValue>
		</TransportationCharges>
		<ServiceOptionsCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>0.00</MonetaryValue>
		</ServiceOptionsCharges>
		<TotalCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>14.09</MonetaryValue>
		</TotalCharges>
		<GuaranteedDaysToDelivery/>
		<ScheduledDeliveryTime/>
		<RatedPackage>
			<TransportationCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TransportationCharges>
			<ServiceOptionsCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</ServiceOptionsCharges>
			<TotalCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TotalCharges>
			<Weight>1.0</Weight>
			<BillingWeight>
				<UnitOfMeasurement>
					<Code/>
				</UnitOfMeasurement>
				<Weight/>
			</BillingWeight>
		</RatedPackage>
	</RatedShipment>
	<RatedShipment>
		<Service><Code>13</Code></Service>
		<RatedShipmentWarning>Your invoice may vary from the displayed reference rates</RatedShipmentWarning>
		<BillingWeight>
			<UnitOfMeasurement><Code>LBS</Code></UnitOfMeasurement>
			<Weight>1.0</Weight>
		</BillingWeight>
		<TransportationCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>14.44</MonetaryValue>
		</TransportationCharges>
		<ServiceOptionsCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>0.00</MonetaryValue>
		</ServiceOptionsCharges>
		<TotalCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>14.44</MonetaryValue>
		</TotalCharges>
		<GuaranteedDaysToDelivery>1</GuaranteedDaysToDelivery>
		<ScheduledDeliveryTime>12:00 Noon</ScheduledDeliveryTime>
		<RatedPackage>
			<TransportationCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TransportationCharges>
			<ServiceOptionsCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</ServiceOptionsCharges>
			<TotalCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TotalCharges>
			<Weight>1.0</Weight>
			<BillingWeight>
				<UnitOfMeasurement>
					<Code/>
				</UnitOfMeasurement>
				<Weight/>
			</BillingWeight>
		</RatedPackage>
	</RatedShipment>
	<RatedShipment>
		<Service><Code>14</Code></Service>
		<RatedShipmentWarning>Your invoice may vary from the displayed reference rates</RatedShipmentWarning>
		<BillingWeight>
			<UnitOfMeasurement><Code>LBS</Code></UnitOfMeasurement>
			<Weight>1.0</Weight>
		</BillingWeight>
		<TransportationCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>53.48</MonetaryValue>
		</TransportationCharges>
		<ServiceOptionsCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>0.00</MonetaryValue>
		</ServiceOptionsCharges>
		<TotalCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>53.48</MonetaryValue>
		</TotalCharges>
		<GuaranteedDaysToDelivery>1</GuaranteedDaysToDelivery>
		<ScheduledDeliveryTime>8:30 A.M.</ScheduledDeliveryTime>
		<RatedPackage>
			<TransportationCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TransportationCharges>
			<ServiceOptionsCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</ServiceOptionsCharges>
			<TotalCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TotalCharges>
			<Weight>1.0</Weight>
			<BillingWeight>
				<UnitOfMeasurement>
					<Code/>
				</UnitOfMeasurement>
				<Weight/>
			</BillingWeight>
		</RatedPackage>
	</RatedShipment>
	<RatedShipment>
		<Service><Code>01</Code></Service>
		<RatedShipmentWarning>Your invoice may vary from the displayed reference rates</RatedShipmentWarning>
		<BillingWeight>
			<UnitOfMeasurement><Code>LBS</Code></UnitOfMeasurement>
			<Weight>1.0</Weight>
		</BillingWeight>
		<TransportationCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>23.22</MonetaryValue>
		</TransportationCharges>
		<ServiceOptionsCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>0.00</MonetaryValue>
		</ServiceOptionsCharges>
		<TotalCharges>
			<CurrencyCode>CAD</CurrencyCode>
			<MonetaryValue>23.22</MonetaryValue>
		</TotalCharges>
		<GuaranteedDaysToDelivery>1</GuaranteedDaysToDelivery>
		<ScheduledDeliveryTime>10:30 A.M.</ScheduledDeliveryTime>
		<RatedPackage>
			<TransportationCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TransportationCharges>
			<ServiceOptionsCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</ServiceOptionsCharges>
			<TotalCharges>
				<CurrencyCode/>
				<MonetaryValue/>
			</TotalCharges>
			<Weight>1.0</Weight>
			<BillingWeight>
				<UnitOfMeasurement>
					<Code/>
				</UnitOfMeasurement>
				<Weight/>
			</BillingWeight>
		</RatedPackage>
	</RatedShipment>
</RatingServiceSelectionResponse>
*/
}
function UPSTracking_XMLCall(AccessLicenseNumber,UserId,Password,Connection)
{
	var URL="https://onlinetools.ups.com/ups.app/xml/QVEvents";
//	var AccessLicenseNumber="DC93324A68A19D30";
//	var UserId="Mwfgroupbahamas";
//	var Password="1234abcd!";

	var strXML="<?xml version='1.0' ?>\r\n";
	strXML+="<AccessRequest xml:lang='en-US'>\r\n";
		strXML+="<AccessLicenseNumber>" + AccessLicenseNumber + "</AccessLicenseNumber>\r\n";
		strXML+="<UserId>" + UserId +"</UserId>\r\n";
		strXML+="<Password>" + Password +"</Password>\r\n";
	strXML+="</AccessRequest>\r\n";
	strXML+="<?xml version='1.0' ?>\r\n";
	strXML+="<QuantumViewRequest xml:lang='en-US'>\r\n";
		strXML+="<Request>\r\n";
			strXML+="<RequestAction>QVEvents</RequestAction>\r\n";
		strXML+="</Request>\r\n";
	strXML+="</QuantumViewRequest>\r\n";

	var objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1");
	try
	{
	    objHttp.Open("POST", URL, false);
	    WinHttpRequestOption_SslErrorIgnoreFlags = 4;
	    objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 0x3300;
	    objHttp.Send(strXML);
	}
	catch (Err)
	{
	    return "0@UPSTracking_XML_call() Exception calling UPS (" + URL + "): Message=" + Err.message + ", Description=" + Err.description;
	}
	var OrgRes= objHttp.ResponseText;
//return OrgRes;
	var Tmp=GetInBetween(OrgRes,"<ResponseStatusCode>","</ResponseStatusCode>");
	if (Tmp!="1")
	{
		Tmp=GetInBetween(OrgRes,"<ErrorDescription>","</ErrorDescription>");
		if (Tmp=="")
			return "0@@ErrMsg+@Error@ErrMsg-@";
		else
			return "0@@ErrMsg+@"+Tmp+"@ErrMsg-@";
	}
	var FromS="";
	var ToS="";
	var Res="";
	var OutCode="";
	for (var i=1; i <= 3; i++)
	{
		Res=OrgRes;
		if (i==1)
		{
			FromS="<Exception>";
			ToS="</Exception>";
			OutCode="X";
		}
		else if (i==2)
		{
			FromS="<Origin>";
			ToS="</Origin>";
			OutCode="O";
		}
		else if (i==3)
		{
			FromS="<Delivery>";
			ToS="</Delivery>";
			OutCode="D";
		}
		while (true)
		{
			Tmp=GetInBetween(Res,FromS,ToS);
			if (Tmp=="")
				break;
			Res=Res.substring(Res.indexOf(ToS,0)+ToS.length,Res.length);
			Tmp=GetInBetween(Tmp,"<TrackingNumber>","</TrackingNumber>");
			if (Tmp!="")
				Connection.Execute("WriteUPSTrackingStatus "+ToSQL(Tmp)+","+ToSQL(OutCode));
		}
	}
	return "1@@ErrMsg+@@ErrMsg-@";

/*  Response Example
<?xml version="1.0" encoding="ISO-8859-1"?>
<QuantumViewResponse>
    <Response>
        <TransactionReference/>
        <ResponseStatusCode>1</ResponseStatusCode>
        <ResponseStatusDescription>Success</ResponseStatusDescription>       
    </Response>
    <QuantumViewEvents>
        <SubscriberID>SubScriptAll</SubscriberID>
        <SubscriptionEvents>
            <Name>OutboundFull</Name>
            <Number>45BD09FCCEEA27BA</Number>
            <SubscriptionStatus>
                <Code>A</Code>
                <Description>Active</Description>
            </SubscriptionStatus>
            <SubscriptionFile>
                <FileName>070824_141035001</FileName>
                <StatusType>
                    <Code>U</Code>
                    <Description>Unread</Description>
                </StatusType>
                <Manifest>
                    <Shipper>
                        <Name>BILLY CO</Name>
                        <ShipperNumber>QVDR01</ShipperNumber>
                        <Address>
                            <AddressLine1>3406 BUFORD HWY</AddressLine1>
                            <City>DULUTH</City>
                            <StateProvinceCode>GA</StateProvinceCode>
                            <PostalCode>300963551</PostalCode>
                            <CountryCode>US</CountryCode>
                        </Address>
                    </Shipper>
                    <ShipTo>
                        <Address>
                            <ConsigneeName>RSD</ConsigneeName>
                            <AddressLine1>26021 ATLANTIC OCEAN DRIVE</AddressLine1>
                            <City>LAKE FOREST</City>
                            <StateProvinceCode>CA</StateProvinceCode>
                            <PostalCode>926308831</PostalCode>
                            <CountryCode>US</CountryCode>
                        </Address>
                    </ShipTo>
                    <ReferenceNumber>
                        <Code>00</Code>
                        <Value>123456789</Value>
                    </ReferenceNumber>
                    <ReferenceNumber>
                        <Code>PO</Code>
                        <Value>444789</Value>
                    </ReferenceNumber>
                    <Service>
                        <Code>001</Code>
                    </Service>
                    <PickupDate>20070824</PickupDate>
                    <DocumentsOnly>3</DocumentsOnly>
                    <Package>
                        <Activity>
                            <Date>20070824</Date>
                            <Time>132526</Time>
                        </Activity>
                        <Description>MerchandiseDescription</Description>
                        <Dimensions>
                            <Length>00000000</Length>
                            <Width>00000000</Width>
                            <Height>00000000</Height>
                        </Dimensions>
                        <DimensionalWeight>
                            <UnitOfMeasurement>
                                <Code>LBS</Code>
                            </UnitOfMeasurement>
                            <Weight>0000020</Weight>
                        </DimensionalWeight>
                        <PackageWeight>
                            <Weight>+0002.0</Weight>
                        </PackageWeight>
                        <TrackingNumber>1ZQVDR018493830864</TrackingNumber>
                        <ReferenceNumber>
                            <Code>00</Code>
                            <Value>123456789</Value>
                        </ReferenceNumber>
                        <ReferenceNumber>
                            <Code>PO</Code>
                            <Value>444789</Value>
                        </ReferenceNumber>
                        <ReferenceNumber>
                            <Code>00</Code>
                            <Value>A1</Value>
                        </ReferenceNumber>                        
                        <PackageServiceOptions>
                            <COD/>
                        </PackageServiceOptions>
                    </Package>
                    <ShipmentServiceOptions>
                        <CallTagARS>
                            <Code>8</Code>
                        </CallTagARS>
                    </ShipmentServiceOptions>
                    <ShipmentChargeType>P/P</ShipmentChargeType>
                    <BillToAccount>
                        <Option>01</Option>
                        <Number>QVDR01</Number>
                    </BillToAccount>
                </Manifest>                
                <Manifest>
                    <Shipper>
                        <Name>BILLY CO</Name>
                        <ShipperNumber>QVDR01</ShipperNumber>
                        <Address>
                            <AddressLine1>3406 BUFORD HWY</AddressLine1>
                            <City>DULUTH</City>
                            <StateProvinceCode>GA</StateProvinceCode>
                            <PostalCode>300963551</PostalCode>
                            <CountryCode>US</CountryCode>
                        </Address>
                    </Shipper>
                    <ShipTo>
                        <Address>
                            <ConsigneeName>RSD</ConsigneeName>
                            <AddressLine1>26021 ATLANTIC OCEAN DRIVE</AddressLine1>
                            <City>LAKE FOREST</City>
                            <StateProvinceCode>CA</StateProvinceCode>
                            <PostalCode>926308831</PostalCode>
                            <CountryCode>US</CountryCode>
                        </Address>
                    </ShipTo>
                    <ReferenceNumber>
                        <Code>00</Code>
                        <Value>123456789</Value>
                    </ReferenceNumber>
                    <ReferenceNumber>
                        <Code>PO</Code>
                        <Value>444789</Value>
                    </ReferenceNumber>
                    <Service>
                        <Code>001</Code>
                    </Service>
                    <PickupDate>20070824</PickupDate>
                    <DocumentsOnly>3</DocumentsOnly>
                    <Package>
                        <Activity>
                            <Date>20070824</Date>
                            <Time>134527</Time>
                        </Activity>
                        <Description>MerchandiseDescription</Description>
                        <Dimensions>
                            <Length>00000000</Length>
                            <Width>00000000</Width>
                            <Height>00000000</Height>
                        </Dimensions>
                        <DimensionalWeight>
                            <UnitOfMeasurement>
                                <Code>LBS</Code>
                            </UnitOfMeasurement>
                            <Weight>0000020</Weight>
                        </DimensionalWeight>
                        <PackageWeight>
                            <Weight>+0002.0</Weight>
                        </PackageWeight>
                        <TrackingNumber>1ZQVDR018493984663</TrackingNumber>                        
                        <ReferenceNumber>
                            <Code>PO</Code>
                            <Value>444789</Value>
                        </ReferenceNumber>
                       
                        <PackageServiceOptions>
                            <COD/>
                        </PackageServiceOptions>
                    </Package>
                    <ShipmentServiceOptions>
                        <CallTagARS>
                            <Code>8</Code>
                        </CallTagARS>
                    </ShipmentServiceOptions>
                    <ShipmentChargeType>P/P</ShipmentChargeType>
                    <BillToAccount>
                        <Option>01</Option>
                        <Number>QVDR01</Number>
                    </BillToAccount>
                </Manifest>                               
            </SubscriptionFile>
            <SubscriptionFile>
                <FileName>070824_143048001</FileName>
                <StatusType>
                    <Code>U</Code>
                    <Description>Unread</Description>
                </StatusType>
                <Generic>
                    <ActivityType>VM</ActivityType>
                    <TrackingNumber>1ZQVDR010190099432</TrackingNumber>
                    <ShipperNumber>QVDR01</ShipperNumber>
                    <Service>
                        <Code>001</Code>
                    </Service>
                    <Activity>
                        <Date>20070824</Date>
                        <Time>130533</Time>
                    </Activity>
                    <BillToAccount>
                        <Option>03</Option>
                        <Number>QVDR03</Number>
                    </BillToAccount>
                    <ShipTo>
                        <ReceivingAddressName>RSDTCNTRL-LAKEFOREST</ReceivingAddressName>
                    </ShipTo>
                </Generic>
                <Generic>
                    <ActivityType>VM</ActivityType>
                    <TrackingNumber>1ZQVDR010190148307</TrackingNumber>
                    <ShipperNumber>QVDR01</ShipperNumber>
                    <Service>
                        <Code>001</Code>
                    </Service>
                    <Activity>
                        <Date>20070824</Date>
                        <Time>134438</Time>
                    </Activity>
                    <BillToAccount>
                        <Option>03</Option>
                        <Number>QVDR03</Number>
                    </BillToAccount>
                    <ShipTo>                        <ReceivingAddressName>RSDTCNTRL-LAKEFOREST</ReceivingAddressName>
                    </ShipTo>
                </Generic>                
            </SubscriptionFile>
            <SubscriptionFile>
                <FileName>070827_133055001</FileName>
                <StatusType>
                    <Code>U</Code>
                    <Description>Unread</Description>
                </StatusType>
                <Exception>
                    <ShipperNumber>QVDR01</ShipperNumber>
                    <TrackingNumber>1ZQVDR010192675065</TrackingNumber>
                    <Date>20070824</Date>
                    <Time>101000</Time>
                    <UpdatedAddress>
                        <ConsigneeName>MMX</ConsigneeName>
                        <StreetNumberLow>1801</StreetNumberLow>
                        <StreetName>SANDALWOOD</StreetName>
                        <StreetType>DR</StreetType>
                        <PoliticalDivision2>ATLANTA</PoliticalDivision2>
                        <PoliticalDivision1>GA</PoliticalDivision1>
                        <CountryCode>US</CountryCode>
                        <PostcodePrimaryLow>30350</PostcodePrimaryLow>
                    </UpdatedAddress>
                    <ReasonCode>78</ReasonCode>
                    <ReasonDescription>THE SHIPPER HAS REQUESTED A DELIVERY INTERCEPT FOR THIS PACKAGE</ReasonDescription>
                    <Resolution>
                        <Code>AH</Code>
                        <Description>THE ADDRESS HAS BEEN CORRECTED. THE DELIVERY HAS BEEN RESCHEDULED</Description>
                    </Resolution>
                    <BillToAccount>
                        <Option>03</Option>
                        <Number>QVDR03</Number>
                    </BillToAccount>
                </Exception>                
                <Origin>
                    <PackageReferenceNumber>
                        <Code>00</Code>
                        <Value>QAST ROW PKG 01</Value>
                    </PackageReferenceNumber>
                    <PackageReferenceNumber>
                        <Code>33</Code>
                        <Value>12345</Value>
                    </PackageReferenceNumber>
                    <ShipmentReferenceNumber>
                        <Code>00</Code>
                        <Value>04ROW01</Value>
                    </ShipmentReferenceNumber>
                    <ShipperNumber>QVDR01</ShipperNumber>
                    <TrackingNumber>1ZQVDR017890823915</TrackingNumber>
                    <Date>20070824</Date>
                    <Time>101000</Time>
                    <ActivityLocation>
                        <AddressArtifactFormat>
                            <PoliticalDivision2>ROSWELL-ROSWELL</PoliticalDivision2>
                            <PoliticalDivision1>GA</PoliticalDivision1>
                            <CountryCode>US</CountryCode>
                        </AddressArtifactFormat>
                    </ActivityLocation>
                    <BillToAccount>
                        <Option>03</Option>
                        <Number>QVDR03</Number>
                    </BillToAccount>
                </Origin>                              
            </SubscriptionFile>            
        </SubscriptionEvents>
    </QuantumViewEvents>
</QuantumViewResponse>
*/
}
function Amazon_OrderImport(Connection)
{
	var Folder_Order=GetAmazon_IncomingFolder();
	var Folder_Order_Processed=""+Server.MapPath("/");
	Folder_Order_Processed+="\\AmazonOrderFileBackup";
	var FileName_Order=GetFilesFromFolder(Folder_Order);
	var Pos=0;
	var FN="";
	for (var i=0; i < FileName_Order.length; i++)
	{
		Pos=FileName_Order[i].lastIndexOf("\\");
		if (Pos >=0)
			FN=FileName_Order[i].substring(Pos+1,FileName_Order[i].length);
		else
			FN=FileName_Order[i];
		FN=FN.toUpperCase();
		if (FN.indexOf("ORDER",0)<0)
		{
			DoDeleteFile(FileName_Order[i]);
			continue;
		}
		if (DoFileExists(Folder_Order_Processed+"\\"+FN))
		{
			DoDeleteFile(FileName_Order[i]);
			continue;
		}
		var objFSO = Server.CreateObject("Scripting.FileSystemObject");
		var objStream = objFSO.OpenTextFile(FileName_Order[i]);
		for (var j=0;(!objStream.AtEndOfStream);j++)
		{
			if (j==0)
			{		
				objStream.ReadLine();
				continue;
			}
			Connection.Execute("InsertAmazonOrderOriginalContent "+ToSQL(FN)+","+ToSQLInt(j+1)+","+ToSQL(objStream.ReadLine()));
		}
		objStream.Close();
		objFSO.MoveFile(FileName_Order[i],Folder_Order_Processed+"\\"+FN);
		objFSO = null;
	}
	Connection.Execute("ProcessAmazonOrderOriginalContent");
	return true;
}
function TextFileOrScreen_Write(TextFileHandle,Output)
{
	if (TextFileHandle==null)
	{
		Response.write(Output);
		Response.Flush;
	}
	else
		TextFileHandle.write(Output);
}

function Amazon_ExportFeedString(Connection,TextFileHandle)
{
	var i;
	var NumOfCol=0;

	var FlatFile="TemplateType=Shoes\tVersion=2013.0131\tThis row for Amazon.com use only.  Do not modify or delete.\tBasic Product information - These attributes need to be populated for all your items.\t\t\t\t\t\t\t\t\t";
	FlatFile+="Offer Information - These attributes are required to make your item buyable for customers on the site.\t\t\t\t\t\t\t\t\t\t\t";
	FlatFile+="Sales Price information - for a sales promotion you can specify a reduced price with these attributes.\t\t\t";
	FlatFile+="Item discovery information - These attributes have an effect on how customers can find your product on the site.\t\t\t\t\t\t\t\t\t\t";
	FlatFile+="Image Information - see Image Info tab for details.\t\t\t\t\t\t\t\t\t\t";
	FlatFile+="FBA - make use of these columns if you are participating in the \"Fulfillment by Amazon\" program.\t\t\t\t\t\t\t\t";
	FlatFile+="Variation information - populate this section if your Product is available in different variations (size/color).\t\t\t\t";
	FlatFile+="Shoe Product Information - these attributes are specific to certain product types.  Please use associated Valid Values for more detail.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t";
	FlatFile+="Shoe Dimensions\t\t\t\t\t\t\t\t\t\t\t";
	FlatFile+="Infrequently used attributes \r\n";
	TextFileOrScreen_Write(TextFileHandle,FlatFile);
	FlatFile="";
	var rs=Connection.Execute("GetItemForAmazon 1,1");
	if (!rs.EOF)
	{
		NumOfCol=rs.Fields.Count;
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields(i).name;
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		var Tmp=ReplaceAllNow(FlatFile,"material-fabric","material-type");
		TextFileOrScreen_Write(TextFileHandle,Tmp);
		FlatFile="";
	}
	while (!rs.EOF)
	{
		FlatFile="";
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields.Item(i);
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		rs.MoveNext();
	}
	rs.Close();
	rs=null;
}
function Amazon_SubmitFeed(Connection)
{
	var CurrentD = new Date;
	var Output=GetAmazon_OutgoingFolder()+"\\AmazonFeedFile_";

//////////////////	DoDeleteAllFileInFolder(Output);
	Output+=CurrentD.getYear();
	Output+="_";
	if (CurrentD.getMonth()+1 < 10)
		Output+="0";
	Output+=(CurrentD.getMonth()+1);
	Output+="_";
	if (CurrentD.getDate() < 10)
		Output+="0";
	Output+=CurrentD.getDate();
	Output+="_";
	if (CurrentD.getHours() < 10)
		Output+="0";
	Output+=CurrentD.getHours();
	Output+="_";
	if (CurrentD.getMinutes() < 10)
		Output+="0";
	Output+=CurrentD.getMinutes();
	Output+="_";
	if (CurrentD.getSeconds() < 10)
		Output+="0";
	Output+=CurrentD.getSeconds();
/////////////	Output+=".txt";

	var x = Server.CreateObject("Scripting.FileSystemObject");
	var y=x.CreateTextFile(Output+".txt");

	Amazon_ExportFeedString(Connection,y);
	y.Close();
	x=null;
/////////////////////////////////////////////	DoRenameFile(Output+".tmp",Output+".txt");
	return Output+".txt";

/*	var URL="https://mws.amazonservices.com";
	var MerchantID="A1IHHI9MRUKJ91";
	var MarketplaceID="ATVPDKIKX0DER";
	var AWSAccessKeyID="AKIAJQ2PPFKJCIMNKIDQ";
//	var SecretKey="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";
	var Signature="MOTGj0dup2CVxJx%2FmwxmrkRJeXuiJ69nQqqCHrSWK2I%3D";

	var CurrentD = new Date;
	var FlatFileAll="";
	var Output=""+Server.MapPath("/");
	Output+="\\tmp\\Amazon";
	DoDeleteAllFileInFolder(Output);
	Output+="\\AmazonFeedFile_";
	Output+=CurrentD.getYear();
	Output+="_";
	if (CurrentD.getMonth()+1 < 10)
		Output+="0";
	Output+=(CurrentD.getMonth()+1);
	Output+="_";
	if (CurrentD.getDate() < 10)
		Output+="0";
	Output+=CurrentD.getDate();
	Output+="_";
	if (CurrentD.getHours() < 10)
		Output+="0";
	Output+=CurrentD.getHours();
	Output+="_";
	if (CurrentD.getMinutes() < 10)
		Output+="0";
	Output+=CurrentD.getMinutes();
	Output+="_";
	if (CurrentD.getSeconds() < 10)
		Output+="0";
	Output+=CurrentD.getSeconds();
	Output+=".txt";

	var x = Server.CreateObject("Scripting.FileSystemObject");
	var y=x.CreateTextFile(Output);
	
	var FlatFile="sku\tproduct-name\tproduct-id\tproduct-id-type\tbrand\tproduct-description\tbullet-point1\tbullet-point2\tbullet-point3\tbullet-point4\tbullet-point5\t";
	FlatFile+="item-price\tcurrency\tproduct-tax-code\tquantity\tsale-price\tsale-from-date\tsale-through-date\t";
	FlatFile+="search-terms1\tsearch-terms2\tsearch-terms3\tsearch-terms4\tsearch-terms5\titem-type\tmain-image-url\tother-image-url1\tother-image-url2\t";
	FlatFile+="other-image-url3\tparent-child\tparent-sku\t";
	FlatFile+="relationship-type\tvariation-theme\tdepartment\tcolor\tcolor-map\tsize\tmaterial-fabric1\tmaterial-fabric2\tmaterial-fabric3\r\n";
	y.Write(FlatFile);
	FlatFileAll+=FlatFile;
	FlatFile="";

	var OldColorRelatedID= -1;
	var rs=Connection.Execute("GetItemForAmazon");
	var Tmp1,Tmp2,SaleStartDate,SaleEndDate,QuickStringCut;
	var RegPrice,SalePrice;
	while (!rs.EOF)
	{
		SaleStartDate=""+rs("SaleStartDate_Y")+"-";
		if (0+rs("SaleStartDate_M") < 10)
			SaleStartDate+="0"+rs("SaleStartDate_M")+"-";
		else
			SaleStartDate+=""+rs("SaleStartDate_M")+"-";
		if (0+rs("SaleStartDate_D") < 10)
			SaleStartDate+="0"+rs("SaleStartDate_D")+"-";
		else
			SaleStartDate+=""+rs("SaleStartDate_D")+"-";

		SaleEndDate=""+rs("SaleEndDate_Y")+"-";
		if (0+rs("SaleEndDate_M") < 10)
			SaleEndDate+="0"+rs("SaleEndDate_M")+"-";
		else
			SaleEndDate+=""+rs("SaleEndDate_M")+"-";
		if (0+rs("SaleEndDate_D") < 10)
			SaleEndDate+="0"+rs("SaleEndDate_D")+"-";
		else
			SaleEndDate+=""+rs("SaleEndDate_D")+"-";

		Tmp2=""+rs("QuickString");
		QuickStringCut="";
		for (var xx=0; xx < 5; xx++)
		{
			if (Tmp2=="")
				QuickStringCut+="\t";
			else if (Tmp2.length <= 50)
			{
				QuickStringCut+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(Tmp2,"\r\n"," "),"\n"," "),"\r"," ");
				QuickStringCut+="\t";
				Tmp2="";
			}
			else
			{
				Tmp1=Tmp2.substring(0,50);
				while (Tmp1.length > 0)
				{
					if (Tmp1.substring(Tmp1.length-1,Tmp1.length)==' ')
						break;
					Tmp1=Tmp1.substring(0,Tmp1.length-1);
				}
				Tmp2=Tmp2.substring(Tmp1.length,Tmp2.length);
				QuickStringCut+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(Tmp1,"\r\n"," "),"\n"," "),"\r"," ");
				QuickStringCut+="\t";
			}
		}
			
		if (0+rs("ColorRelatedID")!=OldColorRelatedID)
		{
			if (0+rs("NumberOfColorRelated") <=1)
			{
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("ItemID"),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="_[PARENT]";
			}
			else
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("ColorRelatedUserID"),"\r\n"," "),"\n"," "),"\r"," ");
			FlatFile+="\t";

			Tmp1=""+rs("StyleColor");
			Tmp2="";
			for (var xx=0; xx < Tmp1.length; xx++)
			{
				Tmp2+=Tmp1.substring(xx,xx+1);
				if (((Tmp1.substring(xx,xx+1) >= 'a')&&(Tmp1.substring(xx,xx+1) <= 'z'))||((Tmp1.substring(xx,xx+1) >= 'A')&&(Tmp1.substring(xx,xx+1) <= 'Z')))
				{
					if (xx+1 < Tmp1.length)
						Tmp2+=Tmp1.substring(xx+1,Tmp1.length).toLowerCase();
					break;
				}
			}

			Tmp1=""+rs("AmazonName");
			Tmp1+=" ";
			Tmp1+=Tmp2;

			if (""+rs("Size1")!="")
			{
				Tmp1+=" ";
				Tmp1+=""+rs("Size1");
			}
			FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(Tmp1,"\r\n"," "),"\n"," "),"\r"," ");
			FlatFile+="\t";
			FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("UPC1"),"\r\n"," "),"\n"," "),"\r"," ");
			FlatFile+="\tUPC\t";
			FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("BrandName"),"\r\n"," "),"\n"," "),"\r"," ");
			FlatFile+="\t";
			FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("OnLineLongDesc"),"\r\n"," "),"\n"," "),"\r"," ");
			FlatFile+="\t";
			for (var xx=1; xx <= 5; xx++)
			{
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("FeatureBenefits"+xx),"\r\n"," "),"\n"," "),"\r"," "),String.fromCharCode(8226),"");
				FlatFile+="\t";
			}
			FlatFile+="\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tparent\t\t\t";
			
			if (MyparseInt(""+rs("CountStyleColor")) > 1)
			{
				if (""+rs("Size2")!="")
					FlatFile+="SizeColor";
				else
					FlatFile+="Color";
			}
			else
			{
				if (""+rs("Size2")!="")
					FlatFile+="Size";
				else
					FlatFile+="";
			}

			FlatFile+="\t";
			FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("AmazonDepartment"),"\r\n"," "),"\n"," "),"\r"," ");
			FlatFile+="\t\t\t\t\r\n";
			y.Write(FlatFile);
			FlatFileAll+=FlatFile;
			FlatFile="";
			if (0+rs("ColorRelatedID") <= 0)
				OldColorRelatedID= -1;
			else
				OldColorRelatedID=0+rs("ColorRelatedID");
		}
		for (var x=1; x <= 18; x++)
		{
			if (((x==1)||(""+rs("Size"+x)!=""))  &&  (0+rs("Qty"+x) >= 0+rs("FeedDefinitions_MinQty")))
			{
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("ItemID"),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\t";
				Tmp1=""+rs("StyleColor");
				Tmp2="";
				for (var xx=0; xx < Tmp1.length; xx++)
				{
					Tmp2+=Tmp1.substring(xx,xx+1);
					if (((Tmp1.substring(xx,xx+1) >= 'a')&&(Tmp1.substring(xx,xx+1) <= 'z'))||((Tmp1.substring(xx,xx+1) >= 'A')&&(Tmp1.substring(xx,xx+1) <= 'Z')))
					{
						if (xx+1 < Tmp1.length)
							Tmp2+=Tmp1.substring(xx+1,Tmp1.length).toLowerCase();
						break;
					}
				}

				Tmp1=""+rs("AmazonName");
				Tmp1+=" ";
				Tmp1+=Tmp2;

				if (""+rs("Size"+x)!="")
				{
					Tmp1+=" ";
					Tmp1+=""+rs("Size"+x);
				}
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(Tmp1,"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\t";
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("UPC"+x),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\tUPC\t";
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("BrandName"),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\t";
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("OnLineLongDesc"),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\t";
				for (var xx=1; xx <= 5; xx++)
				{
					FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("FeatureBenefits"+xx),"\r\n"," "),"\n"," "),"\r"," "),String.fromCharCode(8226),"");
					FlatFile+="\t";
				}
				RegPrice=SalePrice=0.00;
				if (0.00+rs("OnLineUSAPrice") > 0)
				{
					if (0.00+rs("ListPrice_US") > 0.00+rs("OnLineUSAPrice"))
					{
						RegPrice=0.00+rs("ListPrice_US");
						SalePrice=0.00+rs("OnLineUSAPrice");
					}
					else
					{
						RegPrice=0.00+rs("OnLineUSAPrice");
						SalePrice= -1.00;
					}
				}
				else
				{
					if (0.00+rs("ListPrice_US") > 0.00+rs("OnLinePrice"))
					{
						RegPrice=0.00+rs("ListPrice_US");
						SalePrice=0.00+rs("OnLinePrice");
					}
					else
					{
						RegPrice=0.00+rs("OnLinePrice");
						SalePrice= -1.00;
					}
				}
				FlatFile+=RegPrice;
				FlatFile+="\t";
				FlatFile+="USD\tA_GEN_NOTAX\t";
				FlatFile+="3\t";
				if (SalePrice <= 0.00)
					FlatFile+="\t\t";
				else
				{
					FlatFile+=SalePrice;
					FlatFile+="\t";
					FlatFile+=SaleStartDate;
					FlatFile+="\t";
					FlatFile+=SaleEndDate;
				}
				FlatFile+="\t";
				FlatFile+=QuickStringCut;
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("AmazonItemType"),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\t";
				if (""+rs("Thumbnail")!="")
				{
 					FlatFile+="http://www.softmoc.com/items/images";
					FlatFile+=""+rs("Thumbnail");
				}
				FlatFile+="\t";
				if (""+rs("Thumbnail_X2")!="")
				{
 					FlatFile+="http://www.softmoc.com/items/images";
					FlatFile+=""+rs("Thumbnail_X2");
				}
				FlatFile+="\t";
				if (""+rs("Thumbnail_X3")!="")
				{
 					FlatFile+="http://www.softmoc.com/items/images";
					FlatFile+=""+rs("Thumbnail_X3");
				}
				FlatFile+="\t";
				if (""+rs("Thumbnail_X4")!="")
				{
 					FlatFile+="http://www.softmoc.com/items/images";
					FlatFile+=""+rs("Thumbnail_X4");
				}
				FlatFile+="\tchild\t";
				if (0+rs("NumberOfColorRelated") <=1)
				{
					FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("ItemID"),"\r\n"," "),"\n"," "),"\r"," ");
					FlatFile+="_[PARENT]";
					FlatFile+="\t";
				}
				else
				{
					FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("ColorRelatedUserID"),"\r\n"," "),"\n"," "),"\r"," ");
					FlatFile+="\t";
				}
				FlatFile+="Variation\t";

				if (MyparseInt(""+rs("CountStyleColor")) > 1)
				{
					if (""+rs("Size2")!="")
						FlatFile+="SizeColor";
					else
						FlatFile+="Color";
				}
				else
				{
					if (""+rs("Size2")!="")
						FlatFile+="Size";
					else
						FlatFile+="";
				}

				FlatFile+="\t";
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("AmazonDepartment"),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\t";

				Tmp1=""+ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("StyleColor"),"\r\n"," "),"\n"," "),"\r"," ");
				Tmp2="";
				for (var xx=0; xx < Tmp1.length; xx++)
				{
					Tmp2+=Tmp1.substring(xx,xx+1);
					if (((Tmp1.substring(xx,xx+1) >= 'a')&&(Tmp1.substring(xx,xx+1) <= 'z'))||((Tmp1.substring(xx,xx+1) >= 'A')&&(Tmp1.substring(xx,xx+1) <= 'Z')))
					{
						if (xx+1 < Tmp1.length)
							Tmp2+=Tmp1.substring(xx+1,Tmp1.length).toLowerCase();
						break;
					}
				}
				FlatFile+=Tmp2;

				FlatFile+="\t\t";
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("Size"+x),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\t";
				FlatFile+=ReplaceAllNow(ReplaceAllNow(ReplaceAllNow(""+rs("DesignElementWebDescription"),"\r\n"," "),"\n"," "),"\r"," ");
				FlatFile+="\r\n";
				y.Write(FlatFile);
				FlatFileAll+=FlatFile;
				FlatFile="";
			}
		}
		rs.MoveNext();
	}
	rs.Close();
	rs=null;
	y.Close();
	x=null;


Response.write(Output);
Response.end;


	var objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1");
	try
	{
		var Tmp=URL;
		Tmp+="/?AWSAccessKeyId=" + AWSAccessKeyID;
		Tmp+="&Action=SubmitFeed&FeedType=_POST_FLAT_FILE_LISTINGS_DATA_";
		Tmp+="&MarketplaceIdList.Id.1=" + MarketplaceID;
		Tmp+="&Merchant=" + MerchantID;
		Tmp+="&PurgeAndReplace=true";
		Tmp+="&SignatureMethod=HmacSHA256";
		Tmp+="&SignatureVersion=2";
		Tmp+="&Signature=" + Signature;
		Tmp+="&Timestamp=";
		Tmp+=CurrentD.getYear();
		Tmp+="-";
		if (CurrentD.getMonth()+1 < 10)
			Tmp+="0";
		Tmp+=(CurrentD.getMonth()+1);
		Tmp+="-";
		if (CurrentD.getDate() < 10)
			Tmp+="0";
		Tmp+=CurrentD.getDate();
		Tmp+="T";
		if (CurrentD.getHours() < 10)
			Tmp+="0";
		Tmp+=CurrentD.getHours();
		Tmp+="%3A";
		if (CurrentD.getMinutes() < 10)
			Tmp+="0";
		Tmp+=CurrentD.getMinutes();
		Tmp+="%3A00Z";
		Tmp+="&Version=2009-01-01";
	    objHttp.Open("POST", Tmp, false);
	    WinHttpRequestOption_SslErrorIgnoreFlags = 4;
	    objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 0x3300;
	    objHttp.Send(FlatFileAll);
	}
	catch (Err)
	{
	    return "0@Amazon_SubmitFeed Exception calling (" + Tmp + "): Message=" + Err.message + ", Description=" + Err.description;
	}
	var Res= objHttp.ResponseText;



	return Res;
*/}

function Amazon_SubmitPriceQty(Connection)
{
	var CurrentD = new Date;
	var Output=GetAmazon_OutgoingFolder()+"\\AmazonPriceQtyFile_";

//////////////////	DoDeleteAllFileInFolder(Output);
	Output+=CurrentD.getYear();
	Output+="_";
	if (CurrentD.getMonth()+1 < 10)
		Output+="0";
	Output+=(CurrentD.getMonth()+1);
	Output+="_";
	if (CurrentD.getDate() < 10)
		Output+="0";
	Output+=CurrentD.getDate();
	Output+="_";
	if (CurrentD.getHours() < 10)
		Output+="0";
	Output+=CurrentD.getHours();
	Output+="_";
	if (CurrentD.getMinutes() < 10)
		Output+="0";
	Output+=CurrentD.getMinutes();
	Output+="_";
	if (CurrentD.getSeconds() < 10)
		Output+="0";
	Output+=CurrentD.getSeconds();
/////////////	Output+=".txt";

	var x = Server.CreateObject("Scripting.FileSystemObject");
	var y=x.CreateTextFile(Output+".txt");

	Amazon_ExportPriceQtyString(Connection,y);
	y.Close();
	x=null;
////////////////////////////////////	DoRenameFile(Output+".tmp",Output+".txt");
	return Output+".txt";
}

function Amazon_ExportPriceQtyString(Connection,TextFileHandle)
{
	var i;
	var NumOfCol=0;

	var FlatFile="";
	var rs=Connection.Execute("GetItemForAmazon 1,0");
	if (!rs.EOF)
	{
		NumOfCol=rs.Fields.Count;
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields(i).name;
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		FlatFile="";
	}
	while (!rs.EOF)
	{
		FlatFile="";
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields.Item(i);
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		rs.MoveNext();
	}
	rs.Close();
	rs=null;

	rs=Connection.Execute("GetItemForAmazon 0,0");
/*	FlatFile="";
	if (!rs.EOF)
	{
		NumOfCol=rs.Fields.Count;
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields(i).name;
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		FlatFile="";
	}
*/	while (!rs.EOF)
	{
		FlatFile="";
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields.Item(i);
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		rs.MoveNext();
	}
	rs.Close();
	rs=null;
}
function GenericFeed_ExportFeedString(Country,Lang,IsMainWebSite,Connection,TextFileHandle)
{
	var i;
	var NumOfCol=0;

	var FlatFile="";
	var rs=Connection.Execute("GoSearchResult_GenericFeed 0,"+ToSQL(Country)+","+ToSQL(Lang)+","+IsMainWebSite);
	if (!rs.EOF)
	{
		NumOfCol=rs.Fields.Count;
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields(i).name;
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		FlatFile="";
	}
	while (!rs.EOF)
	{
		FlatFile="";
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields.Item(i);
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		rs.MoveNext();
	}
	rs.Close();
	rs=null;
}
function GenericFeed_SubmitFeed(Connection)
{
	var CurrentD = new Date;
	var OutputR=""+Server.MapPath("/")+"\\GenericFeed_FileExport";
	DoDeleteAllFileInFolder(OutputR);
	var Output;
	for (var i=1; i <= 4; i++)
	{
		Output=OutputR;
		if (i==1)
			Output+="\\GenericFeed_CA";
		else if (i==2)
			Output+="\\GenericFeed_CA-FR";
		else if (i==3)
			Output+="\\GenericFeed_US";
		else
			Output+="\\GenericFeed_MOBILE_CA";

		var x = Server.CreateObject("Scripting.FileSystemObject");
		var y=x.CreateTextFile(Output+".tmp");

		if (i==1)
			GenericFeed_ExportFeedString('ca','E',1,Connection,y);
		else if (i==2)
			GenericFeed_ExportFeedString('ca','F',1,Connection,y);
		else if (i==3)
			GenericFeed_ExportFeedString('us','E',1,Connection,y);
		else
			GenericFeed_ExportFeedString('ca','E',0,Connection,y);
		y.Close();
		x=null;
		DoRenameFile(Output+".tmp",Output+".txt");
	}
	return "1";
}

function Criteo_ExportFeedString(Connection,TextFileHandle,Lang)
{
	var i;
	var NumOfCol=0;

//id|name|smallimage|bigimage|producturl|description|price|retailprice|discount|recommendable|instock

	var FlatFile="";
	var rs=Connection.Execute("GetItemForCriteo 1,"+ToSQL(Lang));
	if (!rs.EOF)
	{
		NumOfCol=rs.Fields.Count;
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields(i).name;
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		FlatFile="";
	}
	while (!rs.EOF)
	{
		FlatFile="";
		for (i=0; i < NumOfCol; i++)
		{
			if (FlatFile!="")
				FlatFile+="\t";
			FlatFile+=rs.Fields.Item(i);
		}
		FlatFile+="\r\n";
		TextFileOrScreen_Write(TextFileHandle,FlatFile);
		rs.MoveNext();
	}
	rs.Close();
	rs=null;
}
function Criteo_SubmitFeed(Connection)
{
	var CurrentD = new Date;
	var OutputR=""+Server.MapPath("/")+"\\Criteo_FileExport";
	DoDeleteAllFileInFolder(OutputR);
	var Output;
	for (var i=1; i <= 2; i++)
	{
		Output=OutputR;
		if (i==1)
			Output+="\\CriteoCatalogFeed_E";
		else
			Output+="\\CriteoCatalogFeed_F";

		var x = Server.CreateObject("Scripting.FileSystemObject");
		var y=x.CreateTextFile(Output+".tmp");

		if (i==1)
			Criteo_ExportFeedString(Connection,y,'E');
		else
			Criteo_ExportFeedString(Connection,y,'F');
		y.Close();
		x=null;
		DoRenameFile(Output+".tmp",Output+".txt");
	}
	return "1";
}
</SCRIPT>

<SCRIPT RUNAT=Server LANGUAGE="VBScript">
Function GetCCPreAuth_Real(order_id,pan,ED_Y,ED_M,cvd_value,avs_street_number,avs_street_name,avs_zipcode,amount,IsUS)
	Dim store_id,api_token,crypt_type,cvd_indicator,OutString,CCErrMsg,exp_date
	exp_date=ED_Y & ED_M
	store_id = GetCC_store_id(IsUS)
	api_token = GetCC_api_token(IsUS)
	crypt_type = GetCC_crypt_type(IsUS)
	cvd_indicator=GetCC_cvd_indicator(IsUS)
	Set out = server.CreateObject("Moneris.Request")
	out.initRequest store_id, api_token, "https://www3.moneris.com/gateway2/servlet/MpgRequest"
	Set myTran = server.CreateObject("Moneris.Preauth")
	myTran.setAvsInfo avs_street_number, avs_street_name, avs_zipcode
	if ((cvd_indicator<>"undefined") and (cvd_value<>"undefined") and (cvd_value<>"")) then myTran.setCvdInfo cvd_indicator, cvd_value
	out.setRequest myTran.formatRequest(order_id, amount, pan, exp_date, crypt_type)
	out.sendRequest
	OutString=""
	CCErrMsg=""
	if (out.getResponseCode = null) or ("" & out.getResponseCode = "null") or ("" & out.getResponseCode = "") then
		CCErrMsg="Error: " & out.getMessage
	elseif (out.getResponseCode >= 50) then
		if (""&SysLang="F") then
			CCErrMsg="Désolé! Il y a un problème avec votre carte de crédit, veuillez corriger les informations de votre carte de crédit."
		else
			CCErrMsg="Sorry! There's a problem with your credit card, please correct your credit card information."
		end if
	end if
	OutString=OutString & "@MPIInlineForm+@@MPIInlineForm-@<br>"
	OutString=OutString & "@CCErrMsg+@" & CCErrMsg & "@CCErrMsg-@<br>"
	OutString=OutString & "@TransID+@" & out.getTransID & "@TransID-@<br>"
	OutString=OutString & "@AuthCode+@" & out.getAuthCode & "@AuthCode-@<br>"						'char8
	OutString=OutString & "@AVSResultCode+@" & out.getAvsResultCode & "@AVSResultCode-@<br>"		'char1
	OutString=OutString & "@CVDResultCode+@" & out.getCvdResultCode & "@CVDResultCode-@<br>"		'char2
	GetCCPreAuth_Real = OutString
End Function

Function GetCCPreAuth_Secure_Real(order_id,pan,ED_Y,ED_M,cvd_value,avs_street_number,avs_street_name,avs_zipcode,amount,IsUS,ReturnURL)
	Dim store_id,api_token,crypt_type,cvd_indicator,OutString,CCErrMsg,md,accept,useragent, exp_date, mpiSucc, cavv,mpiMsg,i
	exp_date=ED_Y & ED_M
	store_id = GetCC_store_id(IsUS)
	api_token = GetCC_api_token(IsUS)
	crypt_type = GetCC_crypt_type(IsUS)
	cvd_indicator=GetCC_cvd_indicator(IsUS)
	OutString=""
	CCErrMsg=""

'FOR TEST
'store_id="moneris"
'api_token="hurgle"

	if len( order_id ) < 20 then
		order_id = order_id & "-"
		d = 20 - len( order_id )
		for i = 1 to d
			order_id = order_id & "0"
		next
	end if

	md = order_id & ";" & pan & ";" & ED_Y & ";" & ED_M & ";" & cvd_value & ";"
	Set out = server.CreateObject("Moneris.Request")

'out.initMpiRequest "https://esqa.moneris.com/mpi/servlet/MpiServlet"          'TEST
out.initMpiRequest "https://www3.moneris.com/mpi/servlet/MpiServlet"          'PRODUCTION

	accept = Request.ServerVariables( "HTTP_ACCEPT" )
	useragent = Request.ServerVariables( "HTTP_USER_AGENT" )
	Set purreq = Server.CreateObject( "Moneris.MPIReq" )
	out.setRequest purreq.formatRequest(store_id,api_token,purreq.formatTxnRequest(order_id,amount,pan,exp_date,md,ReturnURL,accept,useragent))
	out.sendRequest

	mpiSucc = "" & out.getMPISuccess
	mpiMsg = "" & out.getMessage
	if mpiSucc = "true" then
		'Creates VBV PIN prompt on customer's browser.
		OutString=OutString & "@MPIInlineForm+@" & out.getMPIInlineForm &"@MPIInlineForm-@<br>"
		OutString=OutString & "@CCErrMsg+@@CCErrMsg-@<br>"
		OutString=OutString & "@TransID+@@TransID-@<br>"
		OutString=OutString & "@AuthCode+@@AuthCode-@<br>"						'char8
		OutString=OutString & "@AVSResultCode+@@AVSResultCode-@<br>"		'char1
		OutString=OutString & "@CVDResultCode+@@CVDResultCode-@<br>"		'char2
	else
		if (""&SysLang="F") then
			CCErrMsg="Désolé! Il y a un problème avec votre carte de crédit, veuillez corriger les informations de votre carte de crédit."
		else
			CCErrMsg="Sorry! There's a problem with your credit card, please correct your credit card information.(1)"
		end if
		OutString=OutString & "@MPIInlineForm+@@MPIInlineForm-@<br>"
		OutString=OutString & "@CCErrMsg+@" & CCErrMsg & "@CCErrMsg-@<br>"
		OutString=OutString & "@TransID+@@TransID-@<br>"
		OutString=OutString & "@AuthCode+@@AuthCode-@<br>"						'char8
		OutString=OutString & "@AVSResultCode+@@AVSResultCode-@<br>"		'char1
		OutString=OutString & "@CVDResultCode+@@CVDResultCode-@<br>"		'char2
	end if
	GetCCPreAuth_Secure_Real = OutString
End Function

Function GetCCPreAuth_Secure_Final_Real(order_id,pan,ED_Y,ED_M,cvd_value,avs_street_number,avs_street_name,avs_zipcode,amount,IsUS,MPI_MD)
	Dim store_id,api_token,crypt_type,cvd_indicator,OutString,CCErrMsg,accept,useragent, xorder_id,exp_date,mpiSucc,cavv,MPIMessage
	exp_date=ED_Y & ED_M
	store_id = GetCC_store_id(IsUS)
	api_token = GetCC_api_token(IsUS)
	crypt_type = GetCC_crypt_type(IsUS)
	cvd_indicator=GetCC_cvd_indicator(IsUS)
	OutString=""
	CCErrMsg=""

	Set out = Server.CreateObject( "Moneris.Request" )

'FOR TEST
'store_id="moneris"
'api_token="hurgle"

'out.initMpiRequest "https://esqa.moneris.com/mpi/servlet/MpiServlet"          'TEST
out.initMpiRequest "https://www3.moneris.com/mpi/servlet/MpiServlet"          'PRODUCTION


	Set purreq = Server.CreateObject( "Moneris.MPIReq" )
	out.setRequest purreq.formatRequest( store_id, api_token, purreq.formatAcsRequest( Request.Form( "PaRes" ), MPI_MD ) )
	out.sendRequest
	mpiSucc = out.getMPISuccess

	if mpiSucc = "true" then
		'Send transaction to host using CAVV purchase or CAVV preauth, refer to sample
		'code for 'eSELECTplus. Call ‘getMPICavv’ to obtain the CAVV value.
		'If you are using preauth/capture model, be sure to call getMPIMessage() so that
		'the value can be stored and used in the capture transaction after on to protect
		'your chargeback liability. (e.g. getMPIMessage()= A = crypt type of 6 for
		'follow on transaction and getMPIMessage() = Y = crypt type of 5 for follow on
		'transaction.
		cavv = out.getMPICavv
		MPIMessage=out.getMPIMessage

		Set out = Server.CreateObject( "Moneris.Request" )

'out.initRequest store_id, api_token,"https://esqa.moneris.com/gateway2/servlet/MpgRequest"          'TEST
out.initRequest store_id, api_token, "https://www3.moneris.com/gateway2/servlet/MpgRequest"         'PRODUCTION

		Set purreq = Server.CreateObject( "Moneris.cavvpreauth" )

		'Recurring setup.
		'???????????????????????????purreq.setRecur "week", "true", "2007/11/30", "3", "8", "13.50"

		out.setRequest purreq.formatRequest( order_id, amount, pan, expiry_date,cavv)
		out.sendRequest
		'Display financial transaction result.

		OutString=OutString & "@MPIInlineForm+@@MPIInlineForm-@<br>"
		OutString=OutString & "@CCErrMsg+@@CCErrMsg-@<br>"
		OutString=OutString & "@TransID+@" & out.getTransID & "@TransID-@<br>"
		OutString=OutString & "@AuthCode+@" & out.getAuthCode & "@AuthCode-@<br>"						'char8
		OutString=OutString & "@AVSResultCode+@@AVSResultCode-@<br>"		'char1
		OutString=OutString & "@CVDResultCode+@@CVDResultCode-@<br>"		'char2
		OutString=OutString & "@MPIMessage+@" & MPIMessage & "@MPIMessage-@<br>"
		GetCCPreAuth_Secure_Final_Real=OutString
	else
		'Do not send transaction as the cardholder failed authentication.
		OutString=OutString & "@MPIInlineForm+@@MPIInlineForm-@<br>"
		OutString=OutString & "@CCErrMsg+@Error, Credit Card transaction has been declined.(2)@CCErrMsg-@<br>"
		OutString=OutString & "@TransID+@@TransID-@<br>"
		OutString=OutString & "@AuthCode+@@AuthCode-@<br>"						'char8
		OutString=OutString & "@AVSResultCode+@@AVSResultCode-@<br>"		'char1
		OutString=OutString & "@CVDResultCode+@@CVDResultCode-@<br>"		'char2
		GetCCPreAuth_Secure_Final_Real=OutString
	end if
End Function

Function GetInteractPurchase_Secure_Final_Real(order_id,amount,IsUS,IDEBIT_TRACK2)
	Dim store_id,api_token,OutString,CCErrMsg
	store_id = GetCC_store_id(IsUS)
	api_token = GetCC_api_token(IsUS)
	Set out = server.CreateObject("Moneris.Request")



'FOR TEST
'store_id = "store3"
'api_token = "yesguy"

'out.initRequest store_id , api_token , "https://esqa.moneris.com/gateway2/servlet/MpgRequest"  'TEST
out.initRequest store_id, api_token, "https://www3.moneris.com/gateway2/servlet/MpgRequest"         'PRODUCTION


	Set myTran = server.CreateObject("Moneris.IDebitPurchase")
	out.setRequest myTran.formatRequest(order_id, amount, IDEBIT_TRACK2)
	out.sendRequest
	OutString=""
	CCErrMsg=""
	if (out.getResponseCode = null) or ("" & out.getResponseCode = "null") or ("" & out.getResponseCode = "") then
		if (""&SysLang="F") then
			CCErrMsg="Désolé! Il y a un problème avec votre carte de crédit, veuillez corriger les informations de votre carte de crédit."
		else
			CCErrMsg="Sorry! There's a problem with your debit card, purchase has been declined."
		end if
'		CCErrMsg="Error: " & out.getMessage
	elseif (out.getResponseCode >= 50) then



'FOR TEST
'		tmp="ERROR<br><br>Response Code: " & out.getResponseCode & "<br>Message: " & out.getMessage & "<br>"
'		tmp=tmp & "Trans ID: " & out.getTransID & "<br>"
'		tmp=tmp & "Auth Code: " & out.getAuthCode & "<br>"
'		Response.write tmp
'		Response.end

'FOR PRODUCTION
		if (""&SysLang="F") then
			CCErrMsg="Désolé! Il y a un problème avec votre carte de crédit, veuillez corriger les informations de votre carte de crédit."
		else
			CCErrMsg="Sorry! There's a problem with your debit card, purchase has been declined."
		end if
	


	end if
	OutString=OutString & "@CCErrMsg+@" & CCErrMsg & "@CCErrMsg-@<br>"
	OutString=OutString & "@TransID+@" & out.getTransID & "@TransID-@<br>"
	OutString=OutString & "@AuthCode+@" & out.getAuthCode & "@AuthCode-@<br>"						'char8
	GetInteractPurchase_Secure_Final_Real = OutString
End Function
</SCRIPT>
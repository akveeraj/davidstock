<%@LANGUAGE = "VBSCRIPT" 
@ENABLESESSIONSTATE = FALSE%>
<% 
Option Explicit
Response.buffer = True
Dim strDB
'------------------If your site is hosted by another site then change your path in the DBQ value below i.e. in place of Server.MapPath("/SmartReferrer.mdb") type in Server.MapPath("/Your_site_path/SmartReferrer.mdb")-----------------


strDB =  "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("SmartReferrer.mdb") & ";DefaultDir=" & Server.MapPath(".") & ";DriverId=25;FIL=MS Access;MaxBufferSize=512;PageTimeout=5"


'------------------End of Database connection string -----------------
%> 
<HTML>
<head>
<title>Smart Referrer Admin</title>
<META content="" name="Description">
<META content="" name="Keywords">
<META content="noindex" name="robots">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<STYLE>
#Activate {
	LEFT: 111px; POSITION: absolute; TOP: 252px; VISIBILITY: hidden; Z-INDEX: 1
}
#Edit {
	LEFT: 196px; POSITION: absolute; TOP: 252px; VISIBILITY: hidden; Z-INDEX: 1
}
#Delete {
	LEFT: 280px; POSITION: absolute; TOP: 252px; VISIBILITY: hidden; Z-INDEX: 1
}
#Deactivate {
	LEFT: 196px; POSITION: absolute; TOP: 252px; VISIBILITY: hidden; Z-INDEX: 1
}
</STYLE>
<SCRIPT language=javascript>
<!--
window.onerror = null;
 var bName = navigator.appName;
 var bVer = parseInt(navigator.appVersion);
 var NS4 = (bName == "Netscape" && bVer >= 4);
 var IE4 = (bName == "Microsoft Internet Explorer" && bVer >= 4);
 var NS3 = (bName == "Netscape" && bVer < 4);
 var IE3 = (bName == "Microsoft Internet Explorer" && bVer < 4);
 var menuActive = 0
 var menuOn = 0
 var onLayer
 var timeOn = null// LAYER SWITCHING CODE
if (NS4 || IE4) {
 if (navigator.appName == "Netscape"){
 layerStyleRef="layer.";
 layerRef="document.layers";
 styleSwitch="";
layerVis="show";
layerHid="hide";
 }else
{
 layerStyleRef="layer.style.";
 layerRef="document.all";
 styleSwitch=".style";
layerVis="visible";
layerHid="hidden";
 }
}
 
// SHOW MENU
function showLayer(layerName){
if (NS4 || IE4) {
 if (timeOn != null) {
 clearTimeout(timeOn)
 hideLayer(onLayer)
 }
 if (NS4 || IE4) {
 eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="'+layerVis+'"');
 } 
 onLayer = layerName
 }
}// HIDE MENU
function hideLayer(layerName){
 if (menuActive == 0) {
 if (NS4 || IE4) {
 eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="'+layerHid+'"');
 }
 }
}// TIMER FOR BUTTON MOUSE OUT
function btnTimer() {
 timeOn = setTimeout("btnOut()",1000)
}// BUTTON MOUSE OUT
function btnOut(layerName) {
 if (menuActive == 0) {
 hideLayer(onLayer)
 }
}// MENU MOUSE OVER 
function menuOver(itemName) {
 clearTimeout(timeOn)
 menuActive = 1
}// MENU MOUSE OUT 
function menuOut(itemName) {
 menuActive = 0 
 timeOn = setTimeout("hideLayer(onLayer)", 400)

 }// SET BACKGROUND COLOR 
function setBgColor(layer, color) {
  if (NS4)
    eval('document.all.'+layer+'.bgColor="'+color+'"');
  if (IE4)
    eval('document.all.'+layer+'.style.backgroundColor="'+color+'"');
}
// -->
</SCRIPT>
<script language = "Javascript">
<!--

function ValidateForm(){
	var URL=frmSmartReferrerAdmin.txtURL.value
    	if ((URL==null)||(URL=="")){
		alert("Please enter the monitored page URL")
		frmSmartReferrerAdmin.txtURL.focus()
		return false
	}
	return true
 }
//-->
</script>
<style type="text/css">
<!--
.smartreflink {  font-family: Arial, Helvetica, sans-serif; font-size: 8pt; color: #FFFFAE; text-decoration: none}
a:link {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt}
.arial {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; color: #000066}
.title {  font-family: Arial, Helvetica, sans-serif; font-size: 10pt; font-weight: bold; color: #000066}
.subtitle {  font-family: Arial, Helvetica, sans-serif; font-size: 10pt; font-weight: bold; color: #990000}
.font {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt}
.smartreflinkBlack { font-family: Arial, Helvetica, sans-serif; font-size: 8pt; color: #660000}
--> </style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" link="#003399" vlink="#003399" alink="#003399">
<map name="SmartReferrer">
  <area shape="rect" coords="91,6,298,49" href="http://www.smartwebby.com" target="_blank" alt="Smart Web Solutions">
  <area shape="rect" coords="1,0,90,49" href="http://www.smartwebby.com/web_products/smart_referrer/default.asp" target="_blank" alt="Free Referrer monitoring mini statistics">
</map> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#990000"> 
    <td valign="bottom" width="70%"><img src="smart_referrer_header.gif" width="344" height="50" usemap="#SmartReferrer" border="0"></td>
    <td> <%
dim intRID,strReferrerPage,strDate,strTD,strCTD,intRefID,intTotalCount,arrDiff,intDiff,strDiff,strQuery,strReport,intToday,iCnt,strDel,strEdit,submit,strHidURL,strBtnAdd,strBtnUpdate,strBtnDelete,strEditDel,strURL,rsCheck
submit=request.form("hidSubmit")
intRID=request.QueryString("rid")
intRefID=request.QueryString("refid")
strReferrerPage=request.QueryString("ref")
strDel=request.querystring("del")
strEdit=request.querystring("edit")
intNav=request.QueryString("NAV")
strReport=request.QueryString("report")
If strReport="gen" and intRefID<>"" then strReport="nohits"
If strReport="nohits" and intRefID="" then intRefID="0"
strDiff=request.QueryString("dif")
arrDiff=Split(strDiff,",",-1,1)
strQuery="report=" & strReport & "&rid=" & intRID & "&ref=" & strReferrerPage
strDate=DisplayDate(CDate(Date()))
strTD="<td align='center'><font class='font'>"
strCTD="</font></td>"
	response.write "<div align='right'><font face='Arial, Helvetica, sans-serif' size='2'>"
	if intRID<>""  and strDel="" and strEdit="" then 
		response.write "<a href='SmartReferrerAdmin.asp'><font class='smartreflink'>Back to display of monitored pages</font></a><br>"
		iCnt=0
		If strReport<>"gen" then 
			Response.write " <a href='SmartReferrerAdmin.asp?report=gen" & "&rid=" & intRID & "&ref=" & strReferrerPage & "'><font class='smartreflink'>General Report of All Referrers</font></a><br>"
			iCnt=1	
		End if
		If strReport<>"nohits" then Response.write " <a href='SmartReferrerAdmin.asp?report=nohits" & "&rid=" & intRID & "&ref=" & strReferrerPage & "'><font class='smartreflink'>Report of Referrers with Zero Hits Today</font></a>"
		If iCnt=0 then response.write "<br>"
		If strReport<>"" then Response.write " <a href='SmartReferrerAdmin.asp?rid=" & intRID & "&ref=" & strReferrerPage & "'><font class='smartreflink'>Report of Referrers that gave Hits Today</font></a>"
response.write "</font></div>"
	End if
	%></td>
    <td width="20">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3" height="10" background="smart_referrer_fill.gif"></td>
   </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="10">&nbsp;</td>
    <td width="99%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="18"> 
            <div id=Activate style="left: 356px; top: 3px; z-index: 2; width: 418px; height: 37px; visibility: hidden"> 
              <table border=0 cellpadding=0 cellspacing=2 bgcolor="#660000">
                <tr> 
                  <td><font class="font" color="#F9EFA3">Activating a page to 
                    be monitored: </font><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">First 
                    add the URL of the new page by using the <b>'Add URL' function. 
                    Now </b><b>include the </b>smart_referrer.htm page in the 
                    monitored page (Note: extn .asp). Recommended at the bottom 
                    of the page for better loading time.</font></td>
                </tr>
              </table>
            </div>
            <div id=Edit style="left: 352px; top: 3px; width: 420px; height: 37px; z-index: 3; visibility: hidden"><table border=0 cellpadding=0 cellspacing=2 bgcolor="#660000">
                <tr> 
                  <td><font class="font" color="#F9EFA3">Using 'Edit Monitored 
                    page':</font><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">This function should be used if you have mispelled a URL or would 
                    like to archive it to preserve the existing data. To archive 
                    change the URL to something like <b>'(URL)-(some date)'</b> 
                    and add the same URL as a new monitored page. </font></td>
                </tr>
              </table>
            </div>
            <div id=Delete style="left: 356px; top: 3px; width: 418px; height: 33px; z-index: 4; visibility: hidden"> 
              <table border=0 cellpadding=0 cellspacing=2 bgcolor="#660000">
                <tr> 
                  <td><font class="font" color="#F9EFA3">Using 'Delete Monitored 
                    page': </font> <font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">This 
                    function should be used only if you wish to <b>permanently 
                    remove</b> the URL and all existing data related to it. Caution: 
                    This action cannot be reversed.</font></td>
                </tr>
              </table>
            </div>
            <div id=Deactivate style="left: 356px; top: 3px; width: 418px; height: 37px; z-index: 5; visibility: hidden"> 
              <table border=0 cellpadding=0 cellspacing=2 bgcolor="#660000">
                <tr> 
                  <td> <font class="font" color="#F9EFA3">Deactivating a Monitored 
                    page: </font> <font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Two 
                    ways to do this. Remove the Include file smart_referrer.htm 
                    from the monitored page code (We recommend this for decreasing 
                    loading time) OR archive the respective monitored page through 
                    the edit function.</font></td>
                </tr>
              </table>
            </div>
            <font class="subtitle">Date: </font><font class="font"><%=strDate%></font><font class="subtitle"> 
            Time:</font><font class="font"> <%=Time()%></font></td>
          <td align="right" height="18"><% 
if intRID<>"" and strDel="" and strEdit="" then 
	response.write "<b><font class='subtitle'>Monitored Page URL: </font></b><font class='font'>" & strReferrerPage & "</font><br>" 
else 
	response.write "<a href='SmartReferrerAdmin.asp?#admin'><font class='font' color='#CC6600'> Add a new URL</font></a>&nbsp; &nbsp; &nbsp; "
%> <font class="subtitle">Help: </font><font class="font"><a href="SmartReferrerAdmin.asp" onMouseOver="showLayer('Activate')"><font class="font" color="#CC6600">Activate</font></a> 
            , <a href="SmartReferrerAdmin.asp" onMouseOut="btnTimer(),showLayer('Activate')" onMouseOver="showLayer('Edit')"><font class="font" color="#CC6600">Edit</font></a> 
            , <a href="SmartReferrerAdmin.asp" onMouseOut="btnTimer(),showLayer('Activate')" onMouseOver="showLayer('Delete')"><font class="font" color="#CC6600">Delete</font></a> 
            , <a href="SmartReferrerAdmin.asp" onMouseOut="btnTimer(),showLayer('Activate')" onMouseOver="showLayer('Deactivate')"><font class="font" color="#CC6600">Deactivate</font></a> 
            </font><% 
End if
%></td>
        </tr>
      </table>
<%
intTotalCount=0
For each intDiff in arrDiff
	intTotalCount=intTotalCount+intDiff
Next
intDiff=0
if intNav = "" then
	intPage = 1
else
	intPage = cint(intNav)
end if
Dim DBConn,rsReferrer,fldReferrer,fldHits,fldLastHit,fldSum,fldURL,fldRID,strSQL,fldWeekHits,fldDayHits,fldRefID
dim intPage,intStart,intFinish,intCount,intPageCount,intRecord,intNav
if intRID="" or strDel<>"" or strEdit<>"" then 
	strSQL="SELECT tblReferrerPages.fldURL, tblReferrerPages.fldRID, Sum(tblReferrer.fldDayHits) AS SumOffldDayHits, tblReferrerPages.fldArchive FROM tblReferrer, tblReferrerPages WHERE tblReferrer.fldRID=tblReferrerPages.fldRID and (((DateDiff('d',[fldLastHit],Date()))=0)) GROUP BY tblReferrerPages.fldURL, tblReferrerPages.fldRID, tblReferrerPages.fldArchive UNION SELECT tblReferrerPages.fldURL, tblReferrerPages.fldRID,0 AS SumOffldDayHits, tblReferrerPages.fldArchive FROM tblReferrer ,tblReferrerPages WHERE tblReferrer.fldRID=tblReferrerPages.fldRID and tblReferrer.fldRID not in (select fldRID from tblReferrer where DateDiff('d',[fldLastHit],Date())=0) GROUP BY tblReferrerPages.fldURL, tblReferrerPages.fldRID, tblReferrerPages.fldArchive ORDER BY tblReferrerPages.fldArchive, SumOffldDayHits DESC"
else
	If strReport="gen" then strSQL="exec selsp_referrer '" & intRID & "'"
	If strReport="nohits" then strSQL="exec selsp_referrer_nohits '" & intRID & "'" 
	If strReport="" then strSQL="exec selsp_referrer_today '" & intRID & "'"
end if
	Set DBConn = Server.CreateObject("ADODB.Connection") 
	DBConn.Open strDB
	If intRID<>"" and Instr(Application("SRUpdate"),"," & CStr(intRID) & ",")=0 then 
		dim rsDay,varDay,varWeek,rsDate,varDate
		set rsDate=DBConn.execute("SELECT Max(fldLastHit) FROM tblReferrer WHERE fldRID=" & intRID)
		If not rsDate.eof then 
			varDate=rsDate(0)
			set rsDay=DBConn.execute("exec selsp_update_referrer '" & intRID & "','" & varDate & "'")
			If not rsDay.eof then 
				varDay=rsDay(0)
				If (Datediff("d",varDate,CDate(Date()))>=7 OR DateDiff("ww",varDate,CDate(Date()))>0) then
					If not rsDay.eof then rsDay.movenext
					If not rsDay.eof then 
						varWeek=rsDay(0)
						rsDay.close
						set rsDay=Nothing
						DBConn.execute ("update tblReferrerPages set fldLastDay=" & varDay & ", fldLastWeek=" & varWeek & ", fldLastDate='" & varDate & "',fldLastWeekDate='" & varDate & "' where fldRID=" & intRID)		
					End if
					DBConn.execute ("update tblReferrerPages set fldLastDay=" & varDay & ",fldLastDate='" & varDate & "' where fldRID=" & intRID)
				End if
			Else
				rsDay.close
				set rsDay=Nothing

			End if
		End if
		rsDate.close
		set rsDate=Nothing
		
		Application("SRUpdate")=Application("SRUpdate") & "," & intRID & ","
	End If
	If submit<>"" then Call DisplaySmartReferrerAdmin()
	if intRID<>"" and strDel="" and strEdit="" then Call DisplayReferrerReport()
	set rsReferrer=Server.CreateObject("ADODB.Recordset")
	rsReferrer.ActiveConnection = DBConn
	rsReferrer.Source = strSQL
	rsReferrer.CursorType = 0
	rsReferrer.CursorLocation = 3
	rsReferrer.LockType = 1
	rsReferrer.Open
	rsReferrer.PageSize =20
	rsReferrer.CacheSize = rsReferrer.PageSize
	intPageCount = rsReferrer.PageCount
	intCount = rsReferrer.RecordCount
	Set rsReferrer.ActiveConnection = Nothing	
	If (not rsReferrer.EOF) then
		rsReferrer.AbsolutePage = intPage 
		intStart = rsReferrer.AbsolutePosition
	End if
	if CInt(intPage) = CInt(intPageCount) then
		intFinish = intCount
	else
		intFinish = intStart + (rsReferrer.PageSize - 1)
	end if
	intCount=rsReferrer.recordCount
	If not rsReferrer.eof then 
		response.write "<br><font class='font' color='#CC6600'>Displaying Records <b>" & intStart & "</b> to <b>" & intFinish & "</b> Of <b>" & intCount & "</b></font>"
		if intRID="" or strDel<>"" or strEdit<>"" then 
			response.write "<br>"
			Call Navigate()
			Call DisplayReferrerPages()
			If submit="" then Call DisplaySmartReferrerAdmin()
		else	
			If intRefID="" then 
				If strReport="" then 
					response.write "<br>"
					Call Navigate()
				End if
				Call DisplayReferrers() 
			else 
				response.write "<br>"
				Call Navigate()
				Call DisplayNoHits() 
			end if
		end if
		rsReferrer.close
		set rsReferrer=Nothing
		DBConn.close
		set DBConn=Nothing	
		Call Navigate()	
	Else
		if intRID="" or strDel<>"" or strEdit<>"" then Call DisplaySmartReferrerAdmin()
		if strReport="nohits" then Response.write "<p>&nbsp;</p><p align='center'><font face='Arial' size='3' color='#CC6600'><b>There are no records under this report </b></font><font class='title'><br> (Note: Take a look at the 'general report for all referrers' or the 'today's hits report')</font></p>"
	End if

'--- Navigation of Results ----
Sub Navigate()
	if CInt(intPage) > 1 then 
		Response.write "<a href='SmartReferrerAdmin.asp?NAV=" & intPage - 1 & "&" & strQuery
		if intRefID="" and InStrRev(strDiff,",")<>0 then response.write "&dif=" & Left(strDiff,InStrRev(strDiff,",")-1)
		if intRefID<>"" then response.write "&refid=" & intRefID
		response.write "'><font color='#CC3333'><b>Previous</b></font></a>&nbsp; &nbsp; &nbsp;"
	End if
	if CInt(intPage) < CInt(intPageCount) then
		Response.write "<a href='SmartReferrerAdmin.asp?NAV=" & intPage + 1 & "&" & strQuery
		if intRefID="" then 
			if strDiff<>"" then response.write "&dif=" & strDiff & "," & intDiff else response.write "&dif=" & intDiff
		else
			response.write "&refid=" & intRefID
		End if
		response.write "'><font color='#CC3333'><b>Next</b></font></a>"
	End if
	
End Sub

'--- Display of the Referrers for General Hits and Today's Hits report----
Sub DisplayReferrers()
	set fldReferrer=rsReferrer(0)
	set fldHits=rsReferrer(1)
	set fldWeekHits=rsReferrer(2)
	set fldDayHits=rsReferrer(3)
	set fldLastHit=rsReferrer(4)
	set fldRefID=rsReferrer(5)
%><br>
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr height="25"> 
          <td><font class="title">Referrer</font></td>
          <td align="center"><font class="title">Today's Hits</font> </td>
          <td align="center"><font class="title">Week's Hits </font> </td>
          <td align="center"><font class="title"> Total Hits</font> </td>
        </tr>
        <tr bgcolor="#000000"> 
          <td height="2" colspan="4"></td>
        </tr>
        <% 
	For intCount=1 to rsReferrer.Pagesize 
		If not rsReferrer.eof then 
			response.write "<tr height='22' bgColor=" 
			If intCount mod 2=0 then response.write "'#FFEAEA'>" else response.write "'#FFFFF2'>"
          		response.write "<td><font class='font'><a href='" & fldReferrer & "' target='_blank'>" & fldReferrer & "</a>" & strCTD 
			If DateDiff("d",rsReferrer(4),Date())=0 then 
				response.write strTD & fldDayHits & strCTD 
			Else 
				response.write strTD & "0 <a href='SmartReferrerAdmin.asp?refid=" & fldRefID & "&" & strQuery & "&NAV=" & (intTotalCount\rsReferrer.PageSize)+1 & "'><font size='1'>Last Hit</font></a>" & strCTD
				intTotalCount=intTotalCount+1
				intDiff=intDiff+1
			End if
			If DateDiff("ww",rsReferrer(4),Date())=0 then 
				response.write strTD & fldWeekHits & strCTD
			Else 
				response.write strTD & "0" & strCTD
			End if
			response.write strTD & fldHits & strCTD & "</tr>"
			rsReferrer.movenext
		End if
	Next
	set fldReferrer=Nothing
	set fldHits=Nothing
	set fldWeekHits=Nothing
	set fldDayHits=Nothing
	set fldLastHit=Nothing
	set fldRefID=Nothing
%> 
      </table>
      <%
End Sub

'--- Display of the Monitored pages ----
Sub DisplayReferrerPages()
	set fldURL=rsReferrer(0)
	set fldRID=rsReferrer(1)
	set fldSum=rsReferrer(2)
%> 
      <div align="right"><br>
        <font face="Arial, Helvetica, sans-serif" color="#CC0000" size="2"></font> 
      </div>
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr height="25"> 
          <td><font class="title">Monitored Page URL </font></td>
          <td align="center"><font class="title">Today's Hits</font></td>
          <td align="center"><font class="title">View Log?</font> </td>
          <td align="center"><font class="title">Edit?</font></td>
          <td align="center"><font class="title">Delete?</font></td>
        </tr>
        <tr bgcolor="#000000"> 
          <td colspan="5" height="2"></td>
        </tr>
        <%
	dim intArchived
	intArchived=0
	For intCount=1 to rsReferrer.pagesize 
		if not rsReferrer.eof then 
			If rsReferrer(3)<>"0" and intArchived=0 then 
				response.write "<tr bgcolor='#000000'><td colspan='5' height='2'></td></tr><tr height='22'><td align='right'><font class='font'>Total Hits: </font></td><td align='center'><font class='font'><b>" & intTotalCount & "</b></font></td><td colspan='3'></td></tr>"
				response.write "<tr height='22'><td colspan='5'><font class='title'>Archived Records - the following URLs have been Deactivated</font></td></tr><tr bgcolor='#000000'><td colspan='5' height='2'></td></tr>"
				intArchived=1
			End if
			response.write "<tr height='22' bgColor=" 
			If intCount mod 2=0 then response.write "'#FFEAEA'>" else response.write "'#FFFFF2'>"
          		response.write "<td><font class='font'>" & fldURL & strCTD & strTD 
			If fldSum<>"" and rsReferrer(3)="0" then 
				response.write fldSum 
				intTotalCount=intTotalCount+fldSum
			ElseIf rsReferrer(3)<>"0" then 
				response.write "-" 
			else 
				response.write "0" 
			End if
			response.write strCTD & strTD & "<a href='SmartReferrerAdmin.asp?rid=" & fldRID & "&ref=" & fldURL & "'>View Log</a>" & strCTD & strTD & "<a href='SmartReferrerAdmin.asp?edit=true&rid=" & fldRID & "&ref=" & fldURL & "&arc=" & rsReferrer(3) & "#admin'>Edit</a>" & strCTD & strTD & "<a href='SmartReferrerAdmin.asp?del=true&rid=" & fldRID & "&ref=" & fldURL & "#admin'>Delete</a>" & strCTD & "</tr>"
			If intCount=rsReferrer.pagesize then 
				response.write "<tr bgcolor='#000000'><td colspan='5' height='2'></td></tr><tr height='22'><td align='right'><font class='font'>Total Hits: </font></td><td align='center'><font class='font'><b>" & intTotalCount & "</b></font></td><td colspan='3'></td></tr>"
			End if
			rsReferrer.movenext
		Else
			If intArchived=0 then 
				response.write "<tr bgcolor='#000000'><td colspan='5' height='2'></td></tr><tr height='22'><td align='right'><font class='font'>Total Hits: </font></td><td align='center'><font class='font'><b>" & intTotalCount & "</b></font></td><td colspan='3'></td></tr>"
				Exit For
			End if
		End if
	Next
	If intArchived=1 then response.write "<tr bgcolor='#000000'><td colspan='5' height='2'></td></tr>"
	set fldURL=Nothing
	set fldRID=Nothing
	set fldSum=Nothing
%> 
      </table>
      <%
End Sub

'--- Administration for Monitored pages (Add, Delete and Update) ----
Sub DisplaySmartReferrerAdmin()
strHidURL=request.form("hidURL")
strBtnAdd=request.form("btnAdd")
strBtnUpdate=request.form("btnUpdate")
strBtnDelete=request.form("btnDelete")
if submit<>"" then
	strURL = request.form("txtURL") 
	If strBtnAdd<>"" then 
		dim rsNewRefRID
		set rsCheck=DBConn.Execute("Select fldRID from tblReferrerPages where fldURL='" & strURL & "'")
		If rsCheck.eof then 
			DBConn.Execute("INSERT INTO tblReferrerPages(fldURL) VALUES ('" & strURL & "')")
			set rsNewRefRID=DBConn.Execute("Select fldRID from tblReferrerPages where fldURL='" & strURL & "'")
			DBConn.Execute("INSERT INTO tblReferrer(fldRID,fldReferrer) VALUES (" & rsNewRefRID(0) & ",'Unknown/Direct')")
			rsNewRefRID.close
			set rsNewRefRID=Nothing
			response.write "<p><font class='font' color='#000066'>The new URL <b>" & strURL & "</b> has been added to the database. <br><font class='font' color='#990000'>Note: Do not forget to add the server side include in the entered URL file to activate it.</font></font></p>"
		Else
			response.write "<p><font class='font' color='#990000'>The URL <b>" & strURL & "</b> cannot be added to the database since it already exists.</font><br><font class='font' color='#000066'>Note: Edit the existing URL to preserve the data (for example call it <b>'" & strURL & " till (some date)'</b> and then add the same URL for a fresh record of hits). You can also delete it to remove all previous data stored related to the URL and the URL.</font></p>"
		End if
		rsCheck.close
		set rsCheck=Nothing
		strEditDel="False"
		strURL=""
		Call DisplayAdminForm()
	Elseif strBtnUpdate<>"" then
		dim strArchive,strHidArchive
		strArchive=request.form("optArchive")
		strHidArchive=request.form("hidArchive")
		set rsCheck=DBConn.Execute("Select fldRID from tblReferrerPages where fldURL='" & strURL & "' and fldRID<>" & intRID)
		If rsCheck.eof then
			If strURL<>strHidURL and strArchive<>strHidArchive then 
				DBConn.Execute("update tblReferrerPages set fldURL ='" & strURL & "', fldArchive='" & strArchive & "' where fldRID=" & intRID) 
				response.write "<p><font class='font' color='#000066'>The URL <b>" & strHidURL & "</b> has been updated to <b>" & strURL & "</b> in the database And has been <b>"
				If strArchive="0" then response.write "Activated" else response.write "Archived"
				response.write "</b>.</font></p>"
			ElseIf strURL<>strHidURL then 
				DBConn.Execute("update tblReferrerPages set fldURL ='" & strURL & "' where fldRID=" & intRID) 
				response.write "<p><font class='font' color='#000066'>The URL <b>" & strHidURL & "</b> has been updated to <b>" & strURL & "</b> in the database.</font></p>"
			ElseIf strArchive<>strHidArchive then 
				DBConn.Execute("update tblReferrerPages set fldArchive='" & strArchive & "' where fldRID=" & intRID) 
				response.write "<p><font class='font' color='#000066'>The URL <b>" & strHidURL & "</b> has been <b>"
			If strArchive="0" then response.write "Activated" else response.write "Archived"
			response.write "</b>.</font></p>"
			Else 
				response.write "<p><font class='font' color='#000066'>No changes were requested to the existing record.</font></p>"
			End if
		Else 
			response.write "<p><font class='font' color='#990000'><b>The URL cannot be updated to the database since it already exists.</b><br></font><font class='font' color='#000066'>Note: Edit the existing URL to preserve the data (for example call it 'URL till some date' and then add the same URL for a fresh record of hits). You can also delete it to remove all previous data. </font></p>"
			strURL=strReferrerPage
			Call DisplayAdminForm()
		End if
		rsCheck.close
		set rsCheck=Nothing
	Elseif strBtnDelete<>"" then
		If intRID<>"" then 
			DBConn.Execute("Delete from tblReferrer where fldRID=" & intRID)
			DBConn.Execute("Delete from tblReferrerPages where fldRID=" & intRID)
			response.write "<p><font class='font' color='#000066'>The URL <b>" & strHidURL & "</b> has been deleted from the database.</font></p>"
		else
			response.write "<p><font class='font' color='#000066'>No record specified to be deleted from the database.</font></p>"
		End if
		
	End if
Else
	strURL=strReferrerPage
	Call DisplayAdminForm()
End if
End Sub

'--- Display of Admin form----
Sub DisplayAdminForm()
%> <a name="admin"></a> 
      <p><font class="title"><%
	If submit="" then 
		if strEdit<>"" then 
			response.write "Modify the Monitored Page URL" 
		elseif strDel<>"" then 
			response.write "Are you sure you want to delete the Monitored Page and all the data related to it?" 
		Else
			response.write "Add a New Monitored Page to Smart Referrer"
		End if
	End if
	If submit<>"" then 
		If strEditDel="False" then response.write "Add a New Monitored Page to Smart Referrer" else response.write "Modify the Monitored Page URL" 
	End if
	
	%></font></p>
      <form method="post" action="" name="frmSmartReferrerAdmin">
        <table width="100%" border="0" cellspacing="0" cellpadding="4">
          <tr> 
            <td width="14%" valign="middle"><font class="font">Monitored page 
              URL:&nbsp;&nbsp;<%
		dim strArchive
		strArchive=request.querystring("arc")
		if strURL="" then strURL="http://"
		if strDel="" then 
			response.write "<input type='text' name='txtURL' maxlength='100' size='60' value='" & strURL & "'> &nbsp; &nbsp; "
			If strArchive="0" then response.write "<br>Archive the page?<input type='radio' name='optArchive' value='0' checked> No <input type='radio' name='optArchive' value='1'> Yes, I would like to Archive the monitored page and all its current data <br><input type='hidden' name='hidArchive' value='" & strArchive & "'><br>" 
			If strArchive="1" then response.write "<br>Activate the page?<input type='radio' name='optArchive' value='1' checked> No <input type='radio' name='optArchive' value='0'> Yes, I would like to Activate the archived page with all its stored data <br><input type='hidden' name='hidArchive' value='" & strArchive & "'><br>" 
		else 
			response.write "<b>" & strURL & "</b> &nbsp; &nbsp; "
		End if
		If submit<>"" then 
			response.write "<input type='submit' name='"
			if strEditDel="True" or strEdit<>"" then 
				response.write "btnUpdate' value='Update Now!' onclick='return ValidateForm()'>" 
			Else
				response.write "btnAdd' value='Add Now!' onclick='return ValidateForm()'>"
			End if
		else 
			response.write "<input type='submit' name='"
			if strEdit<>"" then 
				response.write "btnUpdate' value='Update Now!' onclick='return ValidateForm()'>" 
			elseif strDel<>"" then 
				response.write "btnDelete' value='Yes Delete Now!'>" 
			Else
				response.write "btnAdd' value='Add Now!' onclick='return ValidateForm()'>"
			End if
		End if
		%> 
              <input type="hidden" name="hidSubmit" value="true">
              <input type="hidden" name="hidURL" value="<%=strURL%>">
              </font></td>
          </tr>
        </table>
      </form>
      <%
End Sub

'--- Display Monitored page brief report ----
Sub DisplayReferrerReport()
	dim strWeek,strPrevWeek,dtWeekDate
	strWeek=DisplayWeek(Date())
	Dim rsReport
	set rsReport=DBConn.execute("selsp_referrer_report '" & intRID & "'")
	intToday=rsReport(1)
	dtWeekDate=rsReport(6)
	strPrevWeek=DisplayWeek(dtWeekDate)
%><br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center"><font class="title">Today's Date </font> </td>
          <td align="center"><font class="title">Today's Hits</font> </td>
          <td align="center"><font class="title">Day Previously Hit</font> </td>
          <td align="center"><font class="title"> That Day's Hits</font> </td>
          <td align="center"><font class="title"> This Week</font><br><font size="1" face="Arial, Helvetica, sans-serif" color="#990000">(<%=strWeek%>)</font> 
          </td>
          <td align="center"><font class="title">Last Recorded Week<br>
            </font><font size="1" face="Arial, Helvetica, sans-serif" color="#990000">(<%If strPrevWeek<>strWeek then response.write strPrevWeek else response.write "No Previous Hits" %>)</font> 
          </td>
          <td align="center"><font class="title">Total Hits </font> </td>
        </tr>
        <tr> 
          <td colspan="7" height="2" bgcolor="#000000"><img src="/images/114pixels.gif" width="114" height="1" alt="smart referrer"></td>
        </tr>
        <% 
		dim i
		If not rsReport.eof then 
			response.write "<tr height='22' bgColor='#FFEAEA'>"
			response.write  strTD & strDate & strCTD
			For i=1 to 5
				If rsReport(i)<>"" then
					response.write  strTD & DisplayDate(rsReport(i)) & strCTD 
				else 
					If rsReport(0)=intToday then response.write strTD & "No Previous Hits" & strCTD else response.write  strTD & "0" & strCTD
				End if
			Next
			response.write  strTD & rsReport(0) & strCTD
			response.write "</tr>"
		end if
	rsReport.close
	set rsReport=Nothing	
%> 
      </table>
      <br>
      <% 
response.write "<div align='center'><font class='subtitle'>"
If strReport="nohits" then Response.write "<b>Report of Referrers that gave zero hits today </b>(Note: Details of the last hit to the monitored page included)"
If strReport="gen" then Response.write "<b>General Report of all Referrers recorded till date </b>(Note: Displayed in the order of total hits recorded)"
If strReport="" then 
	If intToday<>"" then
		Response.write "<b>Report of all Referrers that accessed the monitored page Today</b> (Note: Displayed in the order of total hits recorded today)"
	Else
		Response.write "</font><p>&nbsp</p><font class='font' size='3' color='#CC6600'><b>There were no hits to the monitored page today</b></font><font class='title'><br> (Note: Take a look at the 'general report for all referrers' or the 'zero hits today report')</font>"
	End if
End if
response.write "</font></div>"

End Sub 

'--- Display of the Referrers for NoHits Report report----
Sub DisplayNoHits()
	set fldReferrer=rsReferrer(0)
	set fldHits=rsReferrer(1)
	set fldWeekHits=rsReferrer(2)
	set fldDayHits=rsReferrer(3)
	set fldLastHit=rsReferrer(4)
	set fldRefID=rsReferrer(5)
%> 
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr height="22"> 
          <td><font class="title">Referrer</font></td>
          <td align="center"><font class="title">Total Hits</font> </td>
          <td align="center"><font class="title">Date Last Referred</font> </td>
          <td align="center"><font class="title"> Date Hits</font> </td>
          <td align="center"><font class="title">Last Recorded Week</font> </td>
          <td align="center"><font class="title">Week Hits </font> </td>
        </tr>
        <tr> 
          <td colspan="7" height="2" bgcolor="#000000"></td>
        </tr>
        <% 
		For intCount=1 to rsReferrer.Pagesize 
		If not rsReferrer.eof then 
			response.write "<tr height='22' bgColor=" 
			If fldRefID=cint(intRefID) then 
				response.write "'#FFCCCC'>"
			Else
				If intCount mod 2=0 then response.write "'#FFEAEA'>" else response.write "'#FFFFF2'>"
			End if
			response.write "<td><a href='" & fldReferrer & "' target='_blank'><font class='font'>" & fldReferrer & "</font></a></td>"
			response.write strTD & fldHits & strCTD 
			response.write strTD & "<b>" & DisplayDate(rsReferrer(4)) & "</b>" & strCTD 
			response.write strTD & fldDayHits & strCTD 
			response.write strTD & DisplayWeek(rsReferrer(4)) & strCTD 
			response.write strTD & fldWeekHits & strCTD & "</tr>"
			rsReferrer.movenext
		End if
	Next
	set fldReferrer=Nothing
	set fldHits=Nothing
	set fldWeekHits=Nothing
	set fldDayHits=Nothing
	set fldLastHit=Nothing
	set fldRefID=Nothing
%> 
      </table>
      <% End Sub 

function DisplayDate(strEntryDate)
	If IsDate(strEntryDate) then 
		strEntryDate=FormatDateTime(CDate(strEntryDate),1)
		strEntryDate=trim(right(strEntryDate,len(strEntryDate)-Instr(strEntryDate,",")))
		strEntryDate=left(strEntryDate,3) & " " & right(strEntryDate,len(strEntryDate)-Instr(strEntryDate," "))
	End If
	DisplayDate=strEntryDate
end function

function DisplayShortDate(strEntryDate)
	If IsDate(strEntryDate) then 
		strEntryDate=FormatDateTime(CDate(strEntryDate),1)
		strEntryDate=trim(right(strEntryDate,len(strEntryDate)-Instr(strEntryDate,",")))
		strEntryDate=left(strEntryDate,3) & " " & right(strEntryDate,len(strEntryDate)-Instr(strEntryDate," "))
		strEntryDate=left(strEntryDate,len(strEntryDate)-6)
	End If
	DisplayShortDate=strEntryDate
end function

function DisplayWeek(strEntryDate)
	If IsDate(strEntryDate) then 
		strEntryDate=DisplayShortDate(DateAdd("d",1-Weekday(strEntryDate),strEntryDate)) & " - " & DisplayShortDate(DateAdd("d",7-Weekday(strEntryDate),strEntryDate))
	End If
	DisplayWeek=strEntryDate
end function
%><p>&nbsp;</p></td>
    <td width="10">&nbsp;</td>
  </tr>
</table>
</body>
</HTML>

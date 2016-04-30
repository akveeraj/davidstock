<%
Dim oGeoIP,strErrMsg
Dim strIP,strCountryName,strCountryCode

Set oGeoIP = New CountryLookup
oGeoIP.GeoIPDataBase = Server.MapPath("GeoIP.dat")
If oGeoIP.ErrNum(strErrMsg) <> 0 Then
	Response.Write(strErrMsg)
Else
	strIP = request.ServerVariables("REMOTE_ADDR")
	strCountryName = oGeoIP.lookupCountryName(strIP)
	strCountryCode = oGeoIP.lookupCountryCode(strIP)
End If
Set oGeoIP = Nothing
%>
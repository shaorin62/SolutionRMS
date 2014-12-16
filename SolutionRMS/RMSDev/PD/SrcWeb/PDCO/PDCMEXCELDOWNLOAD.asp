<%


Response.ContentType ="application/x-msexcel"

temp_filename = request("temp_filename")
temp_filename = temp_filename & ".xls"


Response.AddHeader "Content-Disposition","attachment; filename=" & temp_filename 
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type=1
objStream.LoadFromFile Server.MapPath("\Excel") & "\" & temp_filename

download = objStream.Read
Response.BinaryWrite download 
Set objStream = nothing
%>

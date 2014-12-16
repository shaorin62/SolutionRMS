<%

set objConn = server.createobject("ADODB.Connection")
strCon  ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev; user id=advsa; password = advsa1234"
objConn.open strCon          
temp_filename = request("temp_filename")
temp_no = temp_filename
temp_filename = "D:\IF\"&temp_filename

Set objfso = Server.CreateObject("Scripting.FileSystemObject")
if not objfso.FileExists(temp_filename) Then 
%>
<script>
alert("미생성 파일 입니다.");
</script>
<%       
Else
	Dim strSQL 

	' ReadFileToQuery 함수로 만들어진 쿼리문을 INSERT 구문과 결합하여
	' 쿼리 한번 실행으로 파일 내용을 DB에 모두 넣게됨
	
	
	strSQL = "INSERT INTO PD_SAP_VOCHNO " & ReadFileToQuery( temp_filename ) 
    strSQL = strSQL & " ;update PD_SAP_VOCHNO set fileseqno = '" & temp_no & "'"
	strSQL = strSQL & " ;update pd_voch_mst "
	strSQL = strSQL & " set pd_voch_mst.vochno = rtrim(b.vochno),"
	strSQL = strSQL & " pd_voch_mst.errcode = b.errcode,"
	strSQL = strSQL & " pd_voch_mst.errmsg = case b.errcode when '0' then '' else b.errmsg end"
	strSQL = strSQL & " from PD_SAP_VOCHNO b"
	strSQL = strSQL & " where pd_voch_mst.taxyearmon = b.taxyearmon"
	strSQL = strSQL & " and pd_voch_mst.taxno = b.taxno"

	strSQL = strSQL & " ;update pd_tax_hdr set pd_tax_hdr.vochno = b.vochno "
	strSQL = strSQL & " from (select taxyearmon,taxno,vochno from PD_SAP_VOCHNO where errcode = 0 and isnull(vochno,'') <> '') b "
	strSQL = strSQL & " where pd_tax_hdr.taxyearmon = b.taxyearmon and pd_tax_hdr.taxno = b.taxno"


	strSQL = strSQL & " ;delete from PD_SAP_VOCHNO where fileseqno = '" & temp_no & "'"

	strSQL = strSQL & " ;update MD_VOCHFILE_MST set endflag = 'Y' where rmsno = '" & temp_no & "'"

	objConn.Execute strSQL
	objConn.close
	Set objConn = Nothing
End If

function ReadFileToQuery(fileName)
   dim fso, file
   dim strData, arrData
   dim sql

   Set fso = Server.CreateObject("Scripting.FileSystemObject")
   Set file = fso.OpenTextFile(fileName, 1)
	
	
   sql = ""
   Do Until file.AtEndOfStream
       strData = file.ReadLine
       arrData = Split(strData , "|")
  
       ' 쿼리문 조립
       if "" <> sql then sql = sql & " UNION "
       sql = sql & " SELECT '" & Join(arrData, "','") & "'"
   Loop

   ReadFileToQuery = sql
end function
%>
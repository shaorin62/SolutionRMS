<%
on  error resume next
set objConn = server.createobject("ADODB.Connection")
set objConn2 = server.createobject("ADODB.Connection")

strCon ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"
strCon2 ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"

objConn.open strCon    

Sql = Sql & " SELECT "
Sql = Sql & " A.DEPT_CODE CC_CODE,"
Sql = Sql & " B.ORGANIZATION CC_NAME,"
Sql = Sql & " 'MC' SC_BU_CODE,"
Sql = Sql & " '' CC_TYPE_CODE,"
Sql = Sql & " A.DEPT_NAME OC_NAME,"
Sql = Sql & " 'Y' PC, "
Sql = Sql & " 'N' USE_YN,"
Sql = Sql & " B.CREATEDDT SDATE, "
Sql = Sql & " B.TERMINATEDDT EDATE "
Sql = Sql & " FROM HRSERVER.SKMNCHR.DBO.TB_HPW_EMP_MAP A LEFT JOIN HRSERVER.SKMNCHR.DBO.vw_RMS_dept B"
Sql = Sql & " ON A.DEPT_CODE = B.ORGCODE"
Sql = Sql & " WHERE 1=1"
Sql = Sql & " AND LEFT(A.LOGONID,1) <> 'P'"
Sql = Sql & " GROUP BY A.DEPT_CODE,B.ORGCODE,B.ORGANIZATION,"
Sql = Sql & " A.DEPT_NAME,B.CREATEDDT,B.TERMINATEDDT"

Set rs=Createobject("adodb.recordset")
rs.Open Sql, strCon,1
Sqltxt0 = " DELETE FROM SC_CC_EHR; "

Do until rs.Eof
strCC_CODE = rs("CC_CODE")
strCC_NAME = rs("CC_NAME")
strSC_BU_CODE = rs("SC_BU_CODE")
strOC_NAME = rs("OC_NAME")
strPC = rs("PC")
strUSE_YN = rs("USE_YN")
strSDATE = rs("SDATE")
strEDATE = rs("EDATE")
strDATE = Date
Sqltxt1 = sqltxt1 & "INSERT INTO SC_CC_EHR (CC_CODE,CC_NAME,SC_BU_CODE,CC_TYPE_CODE,OC_NAME,PC,USE_YN,SDATE,EDATE,CDATE) VALUES ('" & strCC_CODE & "','" & strCC_NAME & "','" & strSU_BUCODE & "','" & CC_TYPE_CODE & "','" & strOC_NAME & "','" & strPC & "','" & strUSE_YN & "','" & strSDATE & "','" & strEDATE & "','" & strDATE & "');"
rs.movenext
loop

Sqltxt2 = sqltxt2 & " UPDATE SC_CC SET USE_YN = 'N',CDATE = '" & strDATE & "' FROM"
Sqltxt2 = sqltxt2 & " (SELECT A.CC_CODE,A.CDATE FROM SC_CC A WHERE CC_CODE NOT IN (SELECT CC_CODE FROM SC_CC_EHR B WHERE A.CC_CODE = B.CC_CODE)) SUBDATA WHERE SC_CC.CC_CODE = SUBDATA.CC_CODE; "

Sqltxt3 = sqltxt3 & " INSERT INTO SC_CC (CC_CODE,CC_NAME,SC_BU_CODE,CC_TYPE_CODE,LOC_CODE,LOC_NAME,SC_MU_CODE,PU_CODE,PU_NAME,OC_CODE,OC_NAME,PC,CC_HLEVEL,USE_YN,SDATE,EDATE,CDATE)"
Sqltxt3 = sqltxt3 & " SELECT CC_CODE,CC_NAME,SC_BU_CODE,CC_TYPE_CODE,LOC_CODE,LOC_NAME,SC_MU_CODE,PU_CODE,PU_NAME,OC_CODE,OC_NAME,PC,CC_HLEVEL,USE_YN,SDATE,EDATE,CDATE "
Sqltxt3 = sqltxt3 & " FROM SC_CC_EHR A "
Sqltxt3 = sqltxt3 & " WHERE CC_CODE NOT IN (SELECT CC_CODE FROM SC_CC B WHERE A.CC_CODE = B.CC_CODE); "

Sqltxt4 = sqltxt4 & " UPDATE SC_CC SET CC_NAME = B.CC_NAME,OC_NAME = B.OC_NAME,SDATE = B.SDATE,EDATE = B.EDATE,CDATE = B.CDATE "
Sqltxt4 = sqltxt4 & " FROM SC_CC_EHR B "
Sqltxt4 = sqltxt4 & " WHERE SC_CC.CC_CODE = B.CC_CODE; "
Sqltxt4 = sqltxt4 & " DELETE FROM SC_CC_EHR; "


objConn2.open strCon2 
objConn2.Execute Sqltxt0
objConn2.Execute Sqltxt1
objConn2.Execute Sqltxt2
objConn2.Execute Sqltxt3
objConn2.Execute Sqltxt4

if err.number <> 0 then 
	'response.write "System Failed!!!!!"
	response.end
else 
	'Response.write "OK "
end if 


objConn2.Close 
objConn.Close 

Set objConn = Nothing
Set objConn2 = Nothing
response.write sqltxt
%>


<%
on  error resume next
set objConn = server.createobject("ADODB.Connection")
set objConn2 = server.createobject("ADODB.Connection")

strCon ="provider=sqloledb; data source=10.110.10.91\MCEHRDB; initial catalog=EHR; user id=ehrviewer; password = ehrviewer"
strCon2 ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"

objConn.open strCon    

Sql = " SELECT "
Sql = Sql & " EMPNUM EMPNO, "
Sql = Sql & " EMPNAME EMP_NAME, "
Sql = Sql & " 'MC' SC_BU_CODE, "
Sql = Sql & " '' SC_BU_NAME, "
Sql = Sql & " DEPCOD CC_CODE, "
Sql = Sql & " WORKTYPECODE SC_EMP_STATUS, "
Sql = Sql & " OTHERMAIL E_MAIL, "
Sql = Sql & " TELEPHONENUM TEL, "
Sql = Sql & " MOBILE CELLPHONE, "
Sql = Sql & " '' CUSER, "
Sql = Sql & " '' CDATE, "
Sql = Sql & " '' UUSER, "
Sql = Sql & " '' UDATE, "
Sql = Sql & " '' PASSWORD "
Sql = Sql & " FROM tb_RMS_EMPLOYEE "

Set rs=Createobject("adodb.recordset")
rs.Open Sql, strCon,1

Sqltxt0 = " DELETE FROM SC_EMPLOYEE_EHR;"


Do until rs.Eof
strEMPNO = rs("EMPNO")
strEMP_NAME = rs("EMP_NAME")
strSC_BU_CODE = rs("SC_BU_CODE")
strCC_CODE = rs("CC_CODE")
strSC_EMP_STATUS = rs("SC_EMP_STATUS")
strE_MAIL = rs("E_MAIL")
strTEL = rs("TEL")
strCELLPHONE = rs("CELLPHONE")

strDATE = Date
Sqltxt = sqltxt & "INSERT INTO SC_EMPLOYEE_EHR (EMPNO,EMP_NAME,SC_BU_CODE,SC_BU_NAME,CC_CODE,SC_EMP_STATUS,E_MAIL,TEL,CELLPHONE,CUSER,CDATE,UUSER,UDATE,PASSWORD,USE_YN) VALUES "
Sqltxt = sqltxt & " ('" & strEMPNO & "','" & strEMP_NAME & "','" & strSC_BU_CODE & "','','" & strCC_CODE & "','" & strSC_EMP_STATUS & "','" & strE_MAIL & "','" & strTEL & "','" & strCELLPHONE & "','','" & strDATE & "','','','','N');"
rs.movenext
loop

Sqltxt1 = " UPDATE SC_EMPLOYEE_MST SET USE_YN = 'N',MANAGER= 'N',UDATE = '" & strDATE & "' FROM "
Sqltxt1 = Sqltxt1 & " (SELECT A.EMPNO,A.CDATE FROM SC_EMPLOYEE_MST A WHERE EMPNO NOT IN (SELECT EMPNO FROM SC_EMPLOYEE_EHR B WHERE A.EMPNO = B.EMPNO)) SUBDATA WHERE SC_EMPLOYEE_MST.EMPNO = SUBDATA.EMPNO ;"

Sqltxt2 = " INSERT INTO SC_EMPLOYEE_MST (EMPNO,EMP_NAME,SC_BU_CODE,SC_BU_NAME,CC_CODE,SC_EMP_STATUS,E_MAIL,TEL,CELLPHONE,CUSER,CDATE,UUSER,UDATE,PASSWORD,USE_YN,MANAGER) "
Sqltxt2 = Sqltxt2 & "  SELECT EMPNO,EMP_NAME,SC_BU_CODE,SC_BU_NAME,CC_CODE,SC_EMP_STATUS,E_MAIL,TEL,CELLPHONE,CUSER,CDATE,UUSER,UDATE,PASSWORD,USE_YN,'N' MANAGER"
Sqltxt2 = Sqltxt2 & "  FROM SC_EMPLOYEE_EHR A "
Sqltxt2 = Sqltxt2 & "  WHERE EMPNO NOT IN (SELECT EMPNO FROM SC_EMPLOYEE_MST B WHERE A.EMPNO = B.EMPNO)"

Sqltxt3 = "  UPDATE SC_EMPLOYEE_MST SET CC_CODE = B.CC_CODE,SC_EMP_STATUS = B.SC_EMP_STATUS,E_MAIL = B.E_MAIL,TEL = B.TEL,CELLPHONE = B.CELLPHONE,UDATE = B.CDATE "
Sqltxt3 = Sqltxt3 & "  FROM SC_EMPLOYEE_EHR B "
Sqltxt3 = Sqltxt3 & "  WHERE SC_EMPLOYEE_MST.EMPNO = B.EMPNO; "

Sqltxt4 = "  DELETE FROM SC_EMPLOYEE_EHR;"



objConn2.open strCon2 
objConn2.Execute Sqltxt0
objConn2.Execute Sqltxt
objConn2.Execute Sqltxt1
objConn2.Execute Sqltxt2
objConn2.Execute Sqltxt3
objConn2.Execute Sqltxt4 

if err.number <> 0 then 

else 

end if 



objConn2.Close 
objConn.Close 



Set objConn = Nothing
Set objConn2 = Nothing
%>


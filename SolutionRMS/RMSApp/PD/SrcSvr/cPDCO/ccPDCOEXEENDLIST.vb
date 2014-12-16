'****************************************************************************************
'Generated By  : Kim Tae Ho 
'�ý��۱���    : RMS/PD/Server Control Class
'����   ȯ��   : COM+ Service Server Package
'���α׷���    : ccPDCMPREESTLIST.vb
'��         �� : - ����������
'Ư��  ����    : - CE �� Query ���� ���
'                -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008.11.12 Kim Tae Ho
'            2) 
'****************************************************************************************

Imports System.Xml                  ' XMLó��
Imports SCGLControl                 ' ControlClass�� Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '����ó�� Ŭ����
Imports SCGLUtil.cbSCGLXml          'XMLó�� Ŭ����
Imports SCGLUtil.cbSCGLUtil         '��Ÿ��ƿ��Ƽ Ŭ����
Imports ePDCO                       '����Ƽ �߰�

' ��ƼƼ Ŭ���� ���� �ش� ��ƼƼ Ŭ������ ������Ʈ�� ������ �� Imports �Ͻʽÿ�. 
' Imports ��ƼƼ������Ʈ
Public Class ccPDCOEXEENDLIST
    Inherits ccControl
#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccPDCOEXEENDLIST"                  '�ڽ��� Ŭ������
    Private mobjcePD_EXE_HDR As ePDCO.cePD_EXE_HDR            '�������� �⺻����
    Private mobjcePD_EXE_DTL As ePDCO.cePD_EXE_DTL            '�������� �󼼳���
    Private mobjcePD_ACC_MST As ePDCO.cePD_ACC_MST            'ȸ������� �󼼳��� ����
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password"
#End Region

#Region "GROUP BLOCK : Function Section"
    ' =============== ���긶���� �ش� JOB ��ȸ
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strFROM As String, _
                              ByVal strTO As String, _
                              ByVal strPROJECTNO As String, _
                              ByVal strPROJECTNM As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCLIENTNAME As String, _
                              ByVal strGUBN As String) As Object

        Dim strCols As String         '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strXMLData As String    'XML  Return ����(XML  �� ����� �� ����)
        Dim Con1, Con2, Con3, Con4, Con5, Con6
        'Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,cmbSEARCHENDFLAG.value
        strCols = " A.PROJECTNO, "
        strCols = strCols & " A.JOBNO,"
        strCols = strCols & " A.JOBNAME,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,"
        strCols = strCols & " A.CLIENTSUBCODE,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(A.CLIENTSUBCODE) CLIENTSUBNAME,"
        strCols = strCols & " A.SUBSEQ,"
        strCols = strCols & " DBO.PD_JOBCUST_NAME_FUN(A.SUBSEQ) SUBSEQNAME,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.ENDFLAG) ENDFLAG,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.JOBGUBN) JOBGUBN,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.CREPART) CREPART,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.CREGUBN) CREGUBN,"
        strCols = strCols & " REQDAY,DBO.PD_COMMITION_FUN(A.CLIENTCODE) COMMITION,A.CLIENTCODE,A.PREESTNO,"
        strCols = strCols & " CASE B.DIVAMT-B.AMT WHEN 0 THEN 'û���Ϸ�' ELSE 'û���̿Ϸ�' END DIVFLAG,"
        strCols = strCols & " C.ENDDAY,B.DIVAMT,B.AMT, A.DEMANDYEARMON"

        If strFROM <> "" And strTO <> "" Then
            Con1 = String.Format(" AND (A.CREDAY BETWEEN '{0}' AND  '{1}')", strFROM, strTO)
        End If
        If strFROM <> "" And strTO = "" Then
            Con1 = String.Format(" AND (A.CREDAY > '{0}')", strFROM)
        End If
        If strFROM = "" And strTO <> "" Then
            Con1 = String.Format(" AND (A.CREDAY < '{0}')", strTO)
        End If


        If strPROJECTNM <> "" Then Con2 = String.Format(" AND (A.PROJECTNM like '%{0}%')", strPROJECTNM)
        If strPROJECTNO <> "" Then Con3 = String.Format(" AND (A.PROJECTNO = '{0}')", strPROJECTNO)
        If strCLIENTCODE <> "" Then Con4 = String.Format(" AND (A.CLIENTCODE = '{0}')", strCLIENTCODE)
        If strCLIENTNAME <> "" Then Con5 = String.Format(" AND (LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)) like '%{0}%')", strCLIENTNAME)
        If strGUBN <> "" Then Con6 = String.Format(" AND (CASE ISNULL(C.ENDDAY,'') WHEN '' THEN 'F' ELSE 'T' END = '{0}')", strGUBN)

        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6)

        strFormat = "SELECT {0} FROM PD_EXE_HDR C,V_JOBNO A LEFT JOIN (SELECT X.JOBNO,ISNULL(X.DIVAMT,0) DIVAMT,ISNULL(Y.AMT,0) AMT FROM (SELECT JOBNO,SUM(DIVAMT) DIVAMT FROM PD_DIVAMT GROUP BY JOBNO) X LEFT JOIN (SELECT JOBNO,SUM(AMT) AMT FROM PD_TAX_MST GROUP BY JOBNO) Y ON X.JOBNO = Y.JOBNO) B ON A.JOBNO = B.JOBNO WHERE A.JOBNO = C.JOBNO {1} ORDER BY A.JOBNO DESC"

        SetConfig(strInfoXML) '�⺻���� Setting
        With mobjSCGLConfig '�⺻���� Config ��ü
            strSQL = String.Format(strFormat, strCols, strWhere)
            Try
                ' DB ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array ������ ��ȸ (True �϶� ������� ���� ��ȸ(Sheet Data Binding �� ��� ���), False �϶� �����͸� ��ȸ)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
                ' ------ XML ������ ��ȸ
                'strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                'Return strXMLData
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    ' =============== ���긶�� JOB�� ���곻��
    Public Function SelectRtn_DTL(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strCODE As String) As Object

        Dim strCols As String         '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strXMLData As String    'XML  Return ����(XML  �� ����� �� ����)
        Dim Con1

        strCols = " DBO.PD_ESTJOBNO_FUN(A.PREESTNO) JOBNO,"
        strCols = strCols & " A.PREESTNO,"
        strCols = strCols & " CASE ISNULL(B.SORTSEQ,0) WHEN 0 THEN rank() OVER (ORDER BY A.ITEMCODESEQ) ELSE B.SORTSEQ END AS SORTSEQ,"
        strCols = strCols & " A.ITEMCODESEQ,"
        strCols = strCols & " A.ITEMCODE,"
        strCols = strCols & " DBO.PD_ITEMDIVNAME_FUN(A.ITEMCODE) ITEMCLASS,"
        strCols = strCols & " DBO.PD_ITEMCODENAME_FUN(A.ITEMCODE) ITEMNAME,"
        strCols = strCols & " CASE ISNULL(B.ADDFLAG,'') WHEN '' THEN A.QTY ELSE Null END AS QTY,"
        strCols = strCols & " CASE ISNULL(B.ADDFLAG,'') WHEN '' THEN A.PRICE ELSE Null END AS PRICE,"
        strCols = strCols & " CASE ISNULL(B.ADDFLAG,'') WHEN '' THEN A.AMT ELSE Null END AS AMT,"
        strCols = strCols & " B.OUTSCODE,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(B.OUTSCODE) OUTSNAME,"
        strCols = strCols & " B.ADJAMT,"
        strCols = strCols & " B.STD,"
        strCols = strCols & " B.VOCHNO,"
        strCols = strCols & " B.ADJDAY,B.ADDFLAG,B.SEQ,B.PURCHASENO"


        If strCODE <> "" Then Con1 = String.Format(" AND (DBO.PD_ESTJOBNO_FUN(A.PREESTNO) = '{0}')", strCODE)
        strWhere = BuildFields(" ", Con1)
        strFormat = "SELECT {0} FROM PD_PREEST_HDR C,PD_PREEST_DTL A LEFT JOIN PD_EXE_DTL B ON A.PREESTNO = B.PREESTNO AND A.ITEMCODESEQ = B.ITEMCODESEQ AND A.ITEMCODE = B.ITEMCODE "
        strFormat = strFormat & "WHERE  C.PREESTNO = A.PREESTNO AND C.JOBNO = '" & strCODE & "' AND ISNULL(C.CONFIRMFLAG,'') <> '' {1} ORDER BY CASE ISNULL(B.SORTSEQ,0) WHEN 0 THEN rank() OVER (ORDER BY A.ITEMCODESEQ) ELSE B.SORTSEQ END"

        SetConfig(strInfoXML) '�⺻���� Setting
        With mobjSCGLConfig '�⺻���� Config ��ü
            strSQL = String.Format(strFormat, strCols, strWhere)
            Try
                ' DB ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array ������ ��ȸ (True �϶� ������� ���� ��ȸ(Sheet Data Binding �� ��� ���), False �϶� �����͸� ��ȸ)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData

            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_DTL")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    '���긶�� ��������ó��
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer '������ INSERT/UPDATE
        Dim intRtn As Integer '����� ����
        Dim i, intColCnt, intRows As Integer '����, �÷�Cnt, �ο�Cnt ����

        SetConfig(strInfoXML) '�⺻���� Setting
        With mobjSCGLConfig '�⺻������ ������ �ִ� Config ��ü
            Try
                'XML Element ���� ���� (strMasterXML�� ��ȯ)
                Dim xmlRoot As XmlElement


                'DB���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                'Master ������ ó��


                'Detail ������ ó��
                If IsArray(vntData) Then
                    intRtn = ProcessRtn_DTL(vntData)
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ�����RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                'Resource����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    '============== ���긶��ó�� 
    Public Function ProcessRtn_DTL(ByVal vntData As Object) As Integer '������ INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim lngSEQ As Double
        Dim strPURCHASENO As String
        Dim strPRECODE As String
        Dim strENDDAY As String
        Dim strGROUPGBN
        Dim strJOBNO As String
        Dim strSEQ As Double
        Dim strSQLEND As String
        Dim strSQL As String
        Dim strENDSQL As String
        Dim intRtn2 As Integer
        With mobjSCGLConfig
            Try

                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjcePD_EXE_HDR = New cePD_EXE_HDR(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)


                    For i = 1 To intRows
                        '�μ�Ʈ
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then
                            If GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strENDDAY = GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)

                            strJOBNO = GetElement(vntData, "JOBNO", intColCnt, i, OPTIONAL_STR)

                            '���� ����� �������ڸ� ������Ʈ �Ѵ�.
                            strSQL = "UPDATE PD_EXE_HDR SET ENDDAY = '" & strENDDAY & "' WHERE JOBNO = '" & strJOBNO & "'"
                            '[����]
                            'intRtn = mobjcePD_EXE_HDR.UpdateRtn_Endday(strSQL)

                            '��� �ڷḦ ���� �Ѵ�. << - ���긶�� �ڷ� ������ �ϰ� ó�� �� ���� (PD_CLOSING_MST)
                            'intRtn2 = mobjcePD_EXE_HDR.InsertClosing(strJOBNO, strENDDAY)

                            '����ó���� JOB ��ȣ �� STATUS(����) ���� PF04[������] �� ����
                            strSQLEND = "UPDATE PD_JOBNO SET ENDFLAG = 'PF04',SETYEARMON = '" & strENDDAY & "' WHERE JOBNO = '" & strJOBNO & "'"
                            intRtn = mobjcePD_EXE_HDR.ENDFLAG_Update(strSQLEND)
                        End If
                    Next
                End If

                Return intRows
            Catch err As Exception

                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_DTL")
            Finally

                mobjcePD_EXE_HDR.Dispose()
            End Try
        End With
    End Function
    '���긶����� ��������ó��
    Public Function ProcessRtn_Cancel(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer '������ INSERT/UPDATE
        Dim intRtn As Integer '����� ����
        Dim i, intColCnt, intRows As Integer '����, �÷�Cnt, �ο�Cnt ����

        SetConfig(strInfoXML) '�⺻���� Setting
        With mobjSCGLConfig '�⺻������ ������ �ִ� Config ��ü
            Try
                'XML Element ���� ���� (strMasterXML�� ��ȯ)
                Dim xmlRoot As XmlElement


                'DB���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                'Master ������ ó��

                'Detail ������ ó��
                If IsArray(vntData) Then
                    intRtn = ProcessRtn_Cancel_DTL(vntData)
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ�����RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_Cancel")
            Finally
                'Resource����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    '============== ���긶����� ó��  
    Public Function ProcessRtn_Cancel_DTL(ByVal vntData As Object) As Integer '������ INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim lngSEQ As Double
        Dim strPURCHASENO As String
        Dim strPRECODE As String
        Dim strADJDAY As String
        Dim strGROUPGBN
        Dim strJOBNO As String
        Dim strSEQ As Double
        Dim strSQL As String
        Dim strSQLEND As String
        With mobjSCGLConfig
            Try

                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjcePD_EXE_HDR = New cePD_EXE_HDR(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        '�μ�Ʈ
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then

                            strJOBNO = GetElement(vntData, "JOBNO", intColCnt, i, OPTIONAL_STR)
                            strSQL = "UPDATE PD_EXE_HDR SET ENDDAY = '' WHERE JOBNO = '" & strJOBNO & "'"
                            '[����]
                            'intRtn = mobjcePD_EXE_HDR.UpdateRtn_Endday(strSQL)
                            '�������ó�� �� JOB ��ȣ �� STATUS(����) ���� PF03 [û�����·�]�� ����
                            strSQLEND = "UPDATE PD_JOBNO SET ENDFLAG = 'PF03',SETYEARMON = '' WHERE JOBNO = '" & strJOBNO & "'"
                            intRtn = mobjcePD_EXE_HDR.ENDFLAG_Update(strSQLEND)
                        End If
                    Next
                End If

                Return intRows
            Catch err As Exception

                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_Cancel_DTL")
            Finally

                mobjcePD_EXE_HDR.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : Entity Function Section"




#End Region
End Class

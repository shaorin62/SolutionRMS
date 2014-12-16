'****************************************************************************************
'Generated By  : Kim Tae Ho 
'�ý��۱���    : RMS/PD/Server Control Class
'����   ȯ��   : COM+ Service Server Package
'���α׷���    : ccPDCOPREESTSUB.vb
'��         �� : - �󼼰�������
'Ư��  ����    : - CE �� Query ���� ���
'                -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009.10.19 Kim Tae Ho
'            2) 
'****************************************************************************************
Imports System.Xml                  ' XMLó��
Imports SCGLControl                 ' ControlClass�� Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '����ó�� Ŭ����
Imports SCGLUtil.cbSCGLXml          'XMLó�� Ŭ����
Imports SCGLUtil.cbSCGLUtil         '��Ÿ��ƿ��Ƽ Ŭ����
Imports ePDCO
Public Class ccPDCOCREPART
    Inherits ccControl
#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccPDCOCFINPUT"                  '�ڽ��� Ŭ������
    Private mobjcePD_OUTLIST_MST As ePDCO.cePD_OUTLIST_MST            '������� SqlExe
#End Region

#Region "GROUP BLOCK : Function Section"
    ' =============== ��ü�з� ���� ��ȸ
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strCLASSCODE As String) As Object

        Dim strCols As String         '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strXMLData As String    'XML  Return ����(XML  �� ����� �� ����)

        Dim Con1 As String


        If strCLASSCODE <> "" Then Con1 = String.Format(" AND (ATTR01 = '{0}')", strCLASSCODE)

        strWhere = BuildFields(" ", Con1)

        SetConfig(strInfoXML) '�⺻���� Setting
        With mobjSCGLConfig '�⺻���� Config ��ü
            strFormat = " SELECT  "
            strFormat = strFormat & "CLASS_CODE,CODE,'MC' SC_BU_CODE,CODE_NAME,SORT_SEQ,USE_YN,UPDATE_YN,ATTR02,DEBTOR,ACCOUNT,ATTR01,'N' INSERTYN "
            strFormat = strFormat & "FROM SC_CODE WHERE ATTR02 = 'K' AND LEN(ATTR01) = 4 {0} ORDER BY ATTR01,SORT_SEQ "

            strSQL = String.Format(strFormat, strWhere)

            Try

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    '��ü�з� ���� ����
    '============== û����û �ϴܱ׸��� �ۼ� ����
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer '������ INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        '�Է� ����
        Dim strCLASS_CODE As String
        Dim strCODE As String
        Dim strCODE_NAME As String
        Dim dblSORT_SEQ As Double
        Dim strATTR01 As String
        Dim strDEBTOR As String
        Dim strACCOUNT As String

        Dim strSQL As String



        'CLASS_CODE|CODE|SC_BU_CODE = MC |CODE_NAME|SORT_SEQ|USE_YN = "Y" |UPDATE_YN = "N" |ATTR02 = "K" |DEBTOR|ACCOUNT|ATTR01
        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjcePD_OUTLIST_MST = New cePD_OUTLIST_MST(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        '�μ�Ʈ
                        strCLASS_CODE = GetElement(vntData, "CLASS_CODE", intColCnt, i)
                        strCODE_NAME = GetElement(vntData, "CODE_NAME", intColCnt, i)
                        strATTR01 = GetElement(vntData, "ATTR01", intColCnt, i)
                        strDEBTOR = GetElement(vntData, "DEBTOR", intColCnt, i)
                        strACCOUNT = GetElement(vntData, "ACCOUNT", intColCnt, i)
                        strCODE = SelectRtn_CODE(strATTR01)
                        dblSORT_SEQ = SelectRtn_SORT(strATTR01)
                        If GetElement(vntData, "INSERTYN", intColCnt, i) = "Y" Then

                            strSQL = "INSERT INTO SC_CODE (CLASS_CODE,CODE,SC_BU_CODE,CODE_NAME,SORT_SEQ,USE_YN,UPDATE_YN,ATTR02,DEBTOR,ACCOUNT,ATTR01) "
                            strSQL = strSQL & " VALUES('" & strCLASS_CODE & "','" & strCODE & "','MC','" & strCODE_NAME & "'," & dblSORT_SEQ & ",'Y','N','K','" & strDEBTOR & "','" & strACCOUNT & "','" & strATTR01 & "')"

                            intRtn = mobjcePD_OUTLIST_MST.SqlExe(strSQL)

                        End If
                    Next
                End If


                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjcePD_OUTLIST_MST.Dispose()
            End Try
        End With
    End Function

    '============== ���۸�ü�з� �ڵ� ����
    Public Function SelectRtn_CODE(ByVal strCODE As String) As String
        '������� �ܼ���ȸ
        Dim strSQL, strFormat, strRtn As String
        'SetConfig(strInfoXML) '�⺻���� Setting

        Dim strPRECODE As String

        If strCODE = "PA01" Then
            strPRECODE = "PG"
        ElseIf strCODE = "PA02" Then
            strPRECODE = "PC"
        ElseIf strCODE = "PA05" Then
            strPRECODE = "PS"
        ElseIf strCODE = "PA07" Then
            strPRECODE = "PI"
        ElseIf strCODE = "PA08" Then
            strPRECODE = "PO"
        End If
        With mobjSCGLConfig '�⺻���� Config ��ü

            Try

                strSQL = "SELECT '" & strPRECODE & "'+DBO.LPAD(CAST(CAST(SUBSTRING(MAX(CODE),3,2) AS NUMERIC)+1 AS VARCHAR(2)),2,'0') FROM SC_CODE WHERE ATTR01 = '" & strCODE & "' AND ATTR02 = 'K' AND LEN(ATTR01) = 4 AND SUBSTRING(CODE,1,2) = '" & strPRECODE & "' "
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQNO")
            Finally
            End Try
        End With
        '������� �ܼ���ȸ
    End Function
    '============== ���۸�ü�з� �ڵ� ����
    Public Function SelectRtn_SORT(ByVal strCODE As String) As String
        '������� �ܼ���ȸ
        Dim strSQL, strFormat, strRtn As String
        'SetConfig(strInfoXML) '�⺻���� Setting


        With mobjSCGLConfig '�⺻���� Config ��ü

            Try

                strSQL = "SELECT MAX(SORT_SEQ) +1 FROM SC_CODE WHERE ATTR01 = '" & strCODE & "' AND LEN(ATTR01) = 4"
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQNO")
            Finally
            End Try
        End With
        '������� �ܼ���ȸ
    End Function

    '�޺�Ÿ�԰�������
    Public Function GetDataType(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strCombo As String) As Object



        Dim strSQL, strFormat, strSelFields As String
        Dim vntData As Object
        Dim strWhere
        'Combo ������ [strCombo] --------------------------------
        'JOBGUBN : �ű���, JOBBASE : û������, CREGUBN : ���۱���, CREPART : ��������, 
        Select Case strCombo
            Case ("JOBGUBN")
                strWhere = "PD_JOBKIND"
            Case ("JOBBASE")
                strWhere = "PD_JOBBASE"
            Case ("CREGUBN")
                strWhere = "PD_CREGUBN"
            Case ("CREPART")
                strWhere = "CREPART"
            Case ("ENDFLAG")
                strWhere = "PD_ENDFLAG"
            Case ("PONOGUBN")
                strWhere = "PD_PONOGROUP"
        End Select


        SetConfig(strInfoXML)   '�⺻���� ����

        '��ȸ �ʵ� ����
        strSelFields = "CODE,CODE_NAME"

        'SQL�� ����
        If strWhere = "CREPART" Then
            strFormat = "SELECT {0} " & _
                   "FROM SC_CODE " & _
                   "WHERE ATTR02 = 'K' AND LEN(ATTR01) > 3 AND USE_YN = 'Y' ORDER BY CLASS_CODE,SORT_SEQ"

        Else
            strFormat = "SELECT {0} " & _
                               "FROM SC_CODE " & _
                               "WHERE CLASS_CODE = '" & strWhere & "' AND USE_YN = 'Y' " & _
                               "ORDER BY SORT_SEQ "
        End If


        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields)

            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDataType")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function GetDataType_class(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer) As Object

        Dim strSQL As String
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����

        '��ȸ �ʵ� ����



        With mobjSCGLConfig
            strSQL = " SELECT  CLASS_CODE, "
            strSQL = strSQL & "CASE CLASS_CODE  "
            strSQL = strSQL & "WHEN 'PD_GRAPHICKIND' THEN '�μ�' "
            strSQL = strSQL & "WHEN 'PD_ELECKIND' THEN 'CF' "
            strSQL = strSQL & "WHEN 'PD_PROMOTIONKIND' THEN '���θ��' "
            strSQL = strSQL & "WHEN 'PD_INTERNETKIND' THEN '���ͳ�' "
            strSQL = strSQL & "WHEN 'PD_OTHERSKIND' THEN '��Ÿ' END AS CLASS_CODENAME "
            strSQL = strSQL & "FROM SC_CODE WHERE ATTR02 = 'K' AND LEN(ATTR01) = 4 "
            strSQL = strSQL & "GROUP BY CLASS_CODE,ATTR01 "
            strSQL = strSQL & "ORDER BY ATTR01 "


            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDataType_class")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function GetDataType_debtor(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, _
                                       ByRef intColCnt As Integer) As Object

        Dim strSQL As String
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����

        '��ȸ �ʵ� ����



        With mobjSCGLConfig
            strSQL = "SELECT DEBTOR,'('+DEBTOR+') '+DEBTORNAME FROM SC_DEBTOR"

            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDataType_debtor")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function GetDataType_account(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, _
                                        ByRef intColCnt As Integer) As Object
        Dim strSQL As String
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����

        '��ȸ �ʵ� ����



        With mobjSCGLConfig
            strSQL = "SELECT ACCOUNT,'('+ACCOUNT+') '+ACCOUNTNAME FROM SC_ACCOUNT"

            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDataType_account")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : Entity Function Section"
    'strDATE,strMEETINGDATE,strSHOOTDATE 
   

#End Region
End Class
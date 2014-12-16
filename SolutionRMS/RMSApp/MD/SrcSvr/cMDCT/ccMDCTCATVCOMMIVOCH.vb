'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - ��Ʈ�� Ŭ���� ����Ŀ
'�ý��۱���    : �ַ�Ǹ� /�ý��۸�/Server Control Class
'����   ȯ��    : COM+ Service Server Package
'���α׷���    : ccMDCMDEPTMST.vb
'��         ��    : - ����� ���� �մϴ�.
'Ư��  ����     : - Ư�̻��׿� ���� ǥ��
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 
'            2) 
'****************************************************************************************

Imports System.Xml                  ' XMLó��
Imports SCGLControl                 ' ControlClass�� Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '����ó�� Ŭ����
Imports SCGLUtil.cbSCGLXml          'XMLó�� Ŭ����
Imports SCGLUtil.cbSCGLUtil         '��Ÿ��ƿ��Ƽ Ŭ����
Imports eMDCO                       '����Ƽ �߰�

' ��ƼƼ Ŭ���� ���� �ش� ��ƼƼ Ŭ������ ������Ʈ�� ������ �� Imports �Ͻʽÿ�. 
' Imports ��ƼƼ������Ʈ

Public Class ccMDCTCATVCOMMIVOCH
    Inherits ccControl

#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccMDCMCATVCOMMIVOCH"                  '�ڽ��� Ŭ������
    Private mobjceMD_COMMIVOCH_MST As eMDCO.ceMD_COMMIVOCH_MST             '����� Entity ���� ����
    Private mobjceMD_VOCHFILE_MST As eMDCO.ceMD_VOCHFILE_MST
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password" 'Ŀ�ؼ�Setting
#End Region

#Region "GROUP BLOCK : �ܺο������"
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strREAL_MED_CODE As String, _
                              ByVal strVOCHFLAG As String, _
                              ByVal strFILENO As String) As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2, Con3, Con4, Con5 As String
        Con1 = ""
        Con2 = ""
        Con3 = ""
        Con4 = ""
        Con5 = ""
        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '�ѱ��� ���
            If strYEARMON <> "" Then Con1 = String.Format(" AND (A.TAXYEARMON = '{0}')", strYEARMON)
            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND A.CLIENTCODE = '{0}'", strCLIENTCODE)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND A.REAL_MED_CODE = '{0}'", strREAL_MED_CODE)
            If strVOCHFLAG <> "A" Then
                If strVOCHFLAG = "Y" Then
                    Con4 = String.Format(" AND (CASE ISNULL(A.VOCHNO,'N') WHEN 'N' THEN 'N' WHEN '' THEN 'N' ELSE 'Y' END   = '{0}')", strVOCHFLAG)
                ElseIf strVOCHFLAG = "N" Then
                    Con4 = String.Format(" AND ISNULL(B.RMSNO,'') = '' AND (CASE ISNULL(A.VOCHNO,'N') WHEN 'N' THEN 'N' WHEN '' THEN 'N' ELSE 'Y' END   = '{0}')", strVOCHFLAG)
                ElseIf strVOCHFLAG = "M" Then
                    Con4 = String.Format(" AND ISNULL(B.RMSNO,'') <> '' AND (CASE ISNULL(A.VOCHNO,'N') WHEN 'N' THEN 'N' WHEN '' THEN 'N' ELSE 'Y' END   = '{0}')", "N")
                End If
            End If

            If strFILENO <> "" Then Con5 = String.Format(" AND B.RMSNO = '{0}'", strFILENO)
            '��ȸ �ʵ� ����

            strSelFields = " A.DEMANDDAY POSTINGDATE,"
            strSelFields = strSelFields & " replace(A.real_med_bisno,'-','') CUSTOMERCODE, A.REAL_MED_NAME, "
            'strSelFields = strSelFields & " case isnull(b.summ,'') when '' then convert(char(12),RTRIM(LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.REAL_MED_CODE))))+' ���������' else b.summ end as  SUMM,"
            strSelFields = strSelFields & " case isnull(b.summ,'') when '' then convert(char(12),RTRIM(LTRIM(DBO.md_get_taxmedname_fun(A.taxyearmon, a.taxno))))+' ���������' else b.summ end as  SUMM,"
            strSelFields = strSelFields & " '3000' BA,"
            strSelFields = strSelFields & " '53105' COSTCENTER,"
            strSelFields = strSelFields & " A.SUMAMT,"
            strSelFields = strSelFields & " A.VAT,"
            strSelFields = strSelFields & " 'B5' SEMU,"
            strSelFields = strSelFields & " '7040' BP,"
            'strSelFields = strSelFields & " A.DEMANDDAY,"
            strSelFields = strSelFields & " convert(char(8) , DATEADD(mm, 3,A.DEMANDDAY),112) DEMANDDAY, "
            strSelFields = strSelFields & " A.TAXYEARMON,"
            strSelFields = strSelFields & " A.TAXNO,"
            strSelFields = strSelFields & " 'S' GBN,"
            strSelFields = strSelFields & " A.VOCHNO,B.RMSNO,A.MEDFLAG,B.ERRCODE,B.ERRMSG"
            strFormat = "SELECT {0} FROM MD_COMMITAX_HDR  A LEFT JOIN MD_COMMIVOCH_MST B ON A.TAXYEARMON = B.TAXYEARMON AND A.TAXNO = B.TAXNO" & _
                                     " WHERE A.MEDFLAG IN ('A2')  {1} {2} {3} {4} {5} "
            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2, Con3, Con4, Con5)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    Public Function GetFILENO(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strFILENO As String) As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2 As String
        Con1 = ""
        Con2 = ""

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '�ѱ��� ���
            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
            If strFILENO <> "" Then Con2 = String.Format(" AND RMSNO = '{0}'", strFILENO)

            '��ȸ �ʵ� ����

            strSelFields = " RMSNO,CASE ENDFLAG WHEN 'N' THEN 'ó����' ELSE 'ó���Ϸ�' END ENDFLAG,DBO.SC_EMPNAME_FUN(CUSER) CUSER,CDATE,YEARMON"

            strFormat = "SELECT {0} FROM MD_VOCHFILE_MST WHERE 1=1 {1} {2}  "
            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetFILENO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object, _
                               ByVal strYEARMON As String, _
                               ByVal strSAVEYEARMON As String, _
                               ByVal strSAVESEQ As Double, _
                               ByVal strSAVERMSNO As String) As Integer
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblID As Double '�ڵ� ID ������� ���� ���
        Dim strSC_EMP_STATUS As String
        Dim vntData2 As Object
        Dim intSEQ As Double
        Dim strRMSNO As String
        Dim strPOSTINGDATE As String
        Dim strDEMANDDAY As String

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    'File ��������

                    'POSTINGDATE,DEMANDDAY


                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjceMD_COMMIVOCH_MST = New ceMD_COMMIVOCH_MST(mobjSCGLConfig)
                    mobjceMD_VOCHFILE_MST = New ceMD_VOCHFILE_MST(mobjSCGLConfig)

                    mobjceMD_VOCHFILE_MST.FileInsertDo(strSAVEYEARMON, strSAVESEQ, strSAVERMSNO, "N")
                    '''vntData�� �÷���, �ο���� �����Է�

                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''�ش��ϴ�Row ��ŭ Loop
                    strSC_EMP_STATUS = ""
                    For i = 1 To intRows
                        If Trim(GetElement(vntData, "CHK", intColCnt, i)) = "" Then
                        Else
                            If GetElement(vntData, "CHK", intColCnt, i) = 1 Then
                                If GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strPOSTINGDATE = GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)
                                If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)
                                intRtn = UpdateRtn(vntData, intColCnt, i, strSAVERMSNO, strPOSTINGDATE, strDEMANDDAY)
                            End If
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
                mobjceMD_COMMIVOCH_MST.Dispose()
                mobjceMD_VOCHFILE_MST.Dispose()

            End Try
        End With
    End Function
    Public Function SelectRtn_SEQNO(ByVal strYEARMON As String) As Object
        '������� �ܼ���ȸ
        Dim strSQL, strFormat, strRtn As String
        Dim intRowCnt As Double
        Dim intColCnt As Double
        Dim vntData As Object
        'SetConfig(strInfoXML) '�⺻���� Setting
        With mobjSCGLConfig '�⺻���� Config ��ü


            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strSQL = "select '" & strYEARMON & "' yearmon,isnull(max(seq),0)+1 seq,'" & strYEARMON & "'+dbo.lpad(isnull(max(seq),0)+1,4,'0')+'_S' RMSNO from md_vochfile_mst where yearmon = '" & strYEARMON & "'"
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
        '������� �ܼ���ȸ
    End Function
    Public Function VOCHDELL(ByVal strInfoXML As String, _
                             ByVal strYEAR As String, _
                             ByVal strVOCHNO As String, _
                             ByVal strTAXYEARMON As String, _
                             ByVal strTAXNO As Double) As Integer
        Dim intRtnDell

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()


                'File ��������

                'POSTINGDATE,DEMANDDAY


                '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                mobjceMD_COMMIVOCH_MST = New ceMD_COMMIVOCH_MST(mobjSCGLConfig)

                '��ǥ���� ����
                mobjceMD_COMMIVOCH_MST.Delete(strYEAR, strVOCHNO)
                '���ݰ�꼭�� ��ǥ��ȣ '' �� ������Ʈ
                mobjceMD_COMMIVOCH_MST.UpdateDelete(strTAXYEARMON, strTAXNO)

                mobjceMD_COMMIVOCH_MST.Update_vochno(strTAXYEARMON, strTAXNO, "CATV")



                .mobjSCGLSql.SQLCommitTrans()
                Return intRtnDell
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".VOCHDELL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_COMMIVOCH_MST.Dispose()

            End Try
        End With
    End Function
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                                   ByVal vntData As Object) As Integer
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblID As Double '�ڵ� ID ������� ���� ���
        Dim strSC_EMP_STATUS As String
        Dim vntData2 As Object
        Dim intSEQ As Double

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    'File ��������

                    'POSTINGDATE,DEMANDDAY


                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjceMD_COMMIVOCH_MST = New ceMD_COMMIVOCH_MST(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�

                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''�ش��ϴ�Row ��ŭ Loop
                    For i = 1 To intRows
                        If Trim(GetElement(vntData, "CHK", intColCnt, i)) = "" Then
                        Else
                            If GetElement(vntData, "CHK", intColCnt, i) = 1 And GetElement(vntData, "ERRCODE", intColCnt, i) = 1 Then
                                intRtn = DeleteRtn(vntData, intColCnt, i)
                            End If
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".DeleteRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_COMMIVOCH_MST.Dispose()


            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : �ܺο� ����� Method"
    Private Function UpdateRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strRMSNO As String, _
                               ByVal strPOSTINGDATE As String, _
                               ByVal strDEMANDDAY As String) As Integer
        'strPOSTINGDATE,strDEMANDDAY
        Dim intRtn As Integer
        'POSTINGDATE,CUSTOMERCODE,SUMM,BA,SUMAMT,VAT,SEMU,BP,DEMANDDAY,VENDOR,TAXYEARMON,TAXNO,GBN,VOCHNO,RMSNO
        intRtn = mobjceMD_COMMIVOCH_MST.InsertDo( _
                                       strPOSTINGDATE, _
                                       GetElement(vntData, "CUSTOMERCODE", intColCnt, intRow), _
                                       GetElement(vntData, "SUMM", intColCnt, intRow), _
                                       GetElement(vntData, "BA", intColCnt, intRow), _
                                       GetElement(vntData, "COSTCENTER", intColCnt, intRow), _
                                       GetElement(vntData, "SUMAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "VAT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "SEMU", intColCnt, intRow), _
                                       GetElement(vntData, "BP", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "TAXYEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "TAXNO", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "GBN", intColCnt, intRow), _
                                       GetElement(vntData, "VOCHNO", intColCnt, intRow), _
                                       strRMSNO, _
                                       GetElement(vntData, "MEDFLAG", intColCnt, intRow), _
                                       "523201", _
                                       strPOSTINGDATE)
        'GetElement(vntData, "COMMI_RATE", intColCnt, intRow, NULL_NUM, True), _
        Return intRtn
    End Function
    Private Function DeleteRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer) As Integer
        'strPOSTINGDATE,strDEMANDDAY
        Dim intRtn As Integer
        'POSTINGDATE,CUSTOMERCODE,SUMM,BA,SUMAMT,VAT,SEMU,BP,DEMANDDAY,VENDOR,TAXYEARMON,TAXNO,GBN,VOCHNO,RMSNO
        intRtn = mobjceMD_COMMIVOCH_MST.DeleteDo( _
                                       GetElement(vntData, "TAXYEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "TAXNO", intColCnt, intRow, NULL_NUM, True))
        'GetElement(vntData, "COMMI_RATE", intColCnt, intRow, NULL_NUM, True), _
        Return intRtn
    End Function
#End Region
End Class
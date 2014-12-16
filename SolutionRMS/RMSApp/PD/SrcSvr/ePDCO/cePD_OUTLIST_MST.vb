'****************************************************************************************
'�ý��۱��� : RMS/PD/Server Entity Class
'����  ȯ�� : 
'���α׷��� : cePD_OUTLIST_MST.vb (PD_PD_OUTLIST_MST Entity ó�� Class)
'��      �� : PD_PD_OUTLIST_MST Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-10-19 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class cePD_OUTLIST_MST
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_OUTLIST_MST"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'PREESTNO,JOBNO,PRODUCTIONNAME,DIRECTORNAME,EDIT,CG,TELECINE,RECORDING,CMSONG,STUDIO,MODELAGENCY,DATE,MEETINGDATE,SHOOTDATE,
    'DAYS,HOURS,TITLE,LENGTHS,LENGTHS

    Public Function InsertDo(ByVal strPREESTNO As String, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCTIONNAME As String = OPTIONAL_STR, _
            Optional ByVal strDIRECTORNAME As String = OPTIONAL_STR, _
            Optional ByVal strEDIT As String = OPTIONAL_STR, _
            Optional ByVal strCG As String = OPTIONAL_STR, _
            Optional ByVal strTELECINE As String = OPTIONAL_STR, _
            Optional ByVal strRECORDING As String = OPTIONAL_STR, _
            Optional ByVal strCMSONG As String = OPTIONAL_STR, _
            Optional ByVal strSTUDIO As String = OPTIONAL_STR, _
            Optional ByVal strMODELAGENCY As String = OPTIONAL_STR, _
            Optional ByVal strDATE As String = OPTIONAL_STR, _
            Optional ByVal strMEETINGDATE As String = OPTIONAL_STR, _
            Optional ByVal strSHOOTDATE As String = OPTIONAL_STR, _
            Optional ByVal strDAYS As String = OPTIONAL_STR, _
            Optional ByVal strHOURS As String = OPTIONAL_STR, _
            Optional ByVal strTITLE As String = OPTIONAL_STR, _
            Optional ByVal strLENGTHS As String = OPTIONAL_STR, _
            Optional ByVal strCOMMENTS As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCT As String = OPTIONAL_STR, _
            Optional ByVal strPROJECT As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "PREESTNO", strPREESTNO, strFields, strValues)
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "PRODUCTIONNAME", strPRODUCTIONNAME, strFields, strValues)
            BuildNameValues(",", "DIRECTORNAME", strDIRECTORNAME, strFields, strValues)
            BuildNameValues(",", "EDIT", strEDIT, strFields, strValues)
            BuildNameValues(",", "CG", strCG, strFields, strValues)
            BuildNameValues(",", "TELECINE", strTELECINE, strFields, strValues)
            BuildNameValues(",", "RECORDING", strRECORDING, strFields, strValues)
            BuildNameValues(",", "CMSONG", strCMSONG, strFields, strValues)
            BuildNameValues(",", "STUDIO", strSTUDIO, strFields, strValues)
            BuildNameValues(",", "MODELAGENCY", strMODELAGENCY, strFields, strValues)
            BuildNameValues(",", "DATE", strDATE, strFields, strValues)
            BuildNameValues(",", "MEETINGDATE", strMEETINGDATE, strFields, strValues)
            BuildNameValues(",", "SHOOTDATE", strSHOOTDATE, strFields, strValues)
            BuildNameValues(",", "DAYS", strDAYS, strFields, strValues)
            BuildNameValues(",", "HOURS", strHOURS, strFields, strValues)
            BuildNameValues(",", "TITLE", strTITLE, strFields, strValues)
            BuildNameValues(",", "LENGTHS", strLENGTHS, strFields, strValues)
            BuildNameValues(",", "COMMENTS", strCOMMENTS, strFields, strValues)
            BuildNameValues(",", "PRODUCT", strPRODUCT, strFields, strValues)
            BuildNameValues(",", "PROJECT", strPROJECT, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", dblATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", dblATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", dblATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", dblATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", dblATTR10, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)

            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    '
    Public Function UpdateDo(ByVal strPREESTNO As String, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCTIONNAME As String = OPTIONAL_STR, _
            Optional ByVal strDIRECTORNAME As String = OPTIONAL_STR, _
            Optional ByVal strEDIT As String = OPTIONAL_STR, _
            Optional ByVal strCG As String = OPTIONAL_STR, _
            Optional ByVal strTELECINE As String = OPTIONAL_STR, _
            Optional ByVal strRECORDING As String = OPTIONAL_STR, _
            Optional ByVal strCMSONG As String = OPTIONAL_STR, _
            Optional ByVal strSTUDIO As String = OPTIONAL_STR, _
            Optional ByVal strMODELAGENCY As String = OPTIONAL_STR, _
            Optional ByVal strDATE As String = OPTIONAL_STR, _
            Optional ByVal strMEETINGDATE As String = OPTIONAL_STR, _
            Optional ByVal strSHOOTDATE As String = OPTIONAL_STR, _
            Optional ByVal strDAYS As String = OPTIONAL_STR, _
            Optional ByVal strHOURS As String = OPTIONAL_STR, _
            Optional ByVal strTITLE As String = OPTIONAL_STR, _
            Optional ByVal strLENGTHS As String = OPTIONAL_STR, _
            Optional ByVal strCOMMENTS As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCT As String = OPTIONAL_STR, _
            Optional ByVal strPROJECT As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("JOBNO", strJOBNO), _
                        GetFieldNameValue("PRODUCTIONNAME", strPRODUCTIONNAME), _
                        GetFieldNameValue("DIRECTORNAME", strDIRECTORNAME), _
                        GetFieldNameValue("EDIT", strEDIT), _
                        GetFieldNameValue("CG", strCG), _
                        GetFieldNameValue("TELECINE", strTELECINE), _
                        GetFieldNameValue("RECORDING", strRECORDING), _
                        GetFieldNameValue("CMSONG", strCMSONG), _
                        GetFieldNameValue("STUDIO", strSTUDIO), _
                        GetFieldNameValue("MODELAGENCY", strMODELAGENCY), _
                        GetFieldNameValue("DATE", strDATE), _
                        GetFieldNameValue("MEETINGDATE", strMEETINGDATE), _
                        GetFieldNameValue("SHOOTDATE", strSHOOTDATE), _
                        GetFieldNameValue("DAYS", strDAYS), _
                        GetFieldNameValue("HOURS", strHOURS), _
                        GetFieldNameValue("TITLE", strTITLE), _
                        GetFieldNameValue("LENGTHS", strLENGTHS), _
                        GetFieldNameValue("COMMENTS", strCOMMENTS), _
                        GetFieldNameValue("PRODUCT", strPRODUCT), _
                        GetFieldNameValue("PROJECT", strPROJECT), _
                        GetFieldNameValue("ATTR01", strATTR01), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("ATTR03", strATTR03), _
                        GetFieldNameValue("ATTR04", strATTR04), _
                        GetFieldNameValue("ATTR05", strATTR05), _
                        GetFieldNameValue("ATTR06", dblATTR06), _
                        GetFieldNameValue("ATTR07", dblATTR07), _
                        GetFieldNameValue("ATTR08", dblATTR08), _
                        GetFieldNameValue("ATTR09", dblATTR09), _
                        GetFieldNameValue("ATTR10", dblATTR10), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("PREESTNO", strPREESTNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function


    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    'strPREESTNO,dblSEQ,dblSUBITEMCODESEQ,dblITEMCODESEQ,strITEMCODE
    Public Function DeleteDo(ByVal strPREESTNO As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_OUTLIST_MST WHERE PREESTNO = '" & strPREESTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Select ó��
    '*****************************************************************
    Public Function SelectDo(ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                            Optional ByVal strSEQNO As String = OPTIONAL_STR, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("SEQNO", strSEQNO))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : strSQL �ش繮�� �״�� ó�� û������ ����� ���ֺ� ���� ��� ������ ������Ʈ
    '*****************************************************************
    Public Function SqlExe(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SqlExe")
        End Try
    End Function

#End Region

#Region "��ü ����/����"
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻������ ���� XML
    'objSCGLSql = DB ó�� ��ü �ν��Ͻ� ����    '��ȯ : ����
    '��� : DB ó���� ���� ����⺻���� ����
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "PD_OUTLIST_MST"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class


'****************************************************************************************
'�ý��۱��� : RMS/PD/Server Entity Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : cePD_JOBNO.vb (PD_JOBNO Entity ó�� Class)
'��      �� : PD_JOBNO Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-11-07 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class cePD_CONTRACT_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "cePD_CONTRACT_HDR"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    Public Function InsertDo(ByVal strCONTRACTNO As String, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
            Optional ByVal strLOCALAREA As String = OPTIONAL_STR, _
            Optional ByVal strDELIVERYDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblPREAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCEAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblDELIVERYGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal dblFAULTGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal strPAYMENTGBN As String = OPTIONAL_STR, _
            Optional ByVal strSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strTESTDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTMENT As String = OPTIONAL_STR, _
            Optional ByVal dblTESTAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblLOSTDAY As String = OPTIONAL_STR, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCONFLAG As String = OPTIONAL_STR, _
            Optional ByVal strDIVFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMENT As String = OPTIONAL_STR, _
            Optional ByVal strAMTFLAG As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "CONTRACTNO", strCONTRACTNO, strFields, strValues)
            BuildNameValues(",", "OUTSCODE", strOUTSCODE, strFields, strValues)
            BuildNameValues(",", "CONTRACTNAME", strCONTRACTNAME, strFields, strValues)
            BuildNameValues(",", "LOCALAREA", strLOCALAREA, strFields, strValues)
            BuildNameValues(",", "DELIVERYDAY", strDELIVERYDAY, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "PRERATE", dblPRERATE, strFields, strValues)
            BuildNameValues(",", "PREAMT", dblPREAMT, strFields, strValues)
            BuildNameValues(",", "ENDRATE", dblENDRATE, strFields, strValues)
            BuildNameValues(",", "ENDAMT", dblENDAMT, strFields, strValues)
            BuildNameValues(",", "THISRATE", dblTHISRATE, strFields, strValues)
            BuildNameValues(",", "THISAMT", dblTHISAMT, strFields, strValues)
            BuildNameValues(",", "BALANCERATE", dblBALANCERATE, strFields, strValues)
            BuildNameValues(",", "BALANCEAMT", dblBALANCEAMT, strFields, strValues)
            BuildNameValues(",", "DELIVERYGUARANTY", dblDELIVERYGUARANTY, strFields, strValues)
            BuildNameValues(",", "FAULTGUARANTY", dblFAULTGUARANTY, strFields, strValues)
            BuildNameValues(",", "PAYMENTGBN", strPAYMENTGBN, strFields, strValues)
            BuildNameValues(",", "STDATE", strSTDATE, strFields, strValues)
            BuildNameValues(",", "EDDATE", strEDDATE, strFields, strValues)
            BuildNameValues(",", "CONTRACTDAY", strCONTRACTDAY, strFields, strValues)
            BuildNameValues(",", "MANAGER", strMANAGER, strFields, strValues)
            BuildNameValues(",", "TESTDAY", strTESTDAY, strFields, strValues)
            BuildNameValues(",", "TESTENDDAY", strTESTENDDAY, strFields, strValues)
            BuildNameValues(",", "TESTMENT", strTESTMENT, strFields, strValues)
            BuildNameValues(",", "TESTAMT", dblTESTAMT, strFields, strValues)
            BuildNameValues(",", "LOSTDAY", dblLOSTDAY, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "CONFLAG", strCONFLAG, strFields, strValues)
            BuildNameValues(",", "DIVFLAG", strDIVFLAG, strFields, strValues)
            BuildNameValues(",", "COMENT", strCOMENT, strFields, strValues)
            BuildNameValues(",", "AMTFLAG", strAMTFLAG, strFields, strValues)
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
    Public Function UpdateDo(ByVal strCONTRACTNO As String, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
            Optional ByVal strLOCALAREA As String = OPTIONAL_STR, _
            Optional ByVal strDELIVERYDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblPREAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCEAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblDELIVERYGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal dblFAULTGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal strPAYMENTGBN As String = OPTIONAL_STR, _
            Optional ByVal strSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strTESTDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTMENT As String = OPTIONAL_STR, _
            Optional ByVal dblTESTAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblLOSTDAY As Double = OPTIONAL_NUM, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCONFLAG As String = OPTIONAL_STR, _
            Optional ByVal strDIVFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMENT As String = OPTIONAL_STR, _
            Optional ByVal strAMTFLAG As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("CONTRACTNAME", strCONTRACTNAME), _
                        GetFieldNameValue("LOCALAREA", strLOCALAREA), _
                        GetFieldNameValue("DELIVERYDAY", strDELIVERYDAY), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("PRERATE", dblPRERATE), _
                        GetFieldNameValue("PREAMT", dblPREAMT), _
                        GetFieldNameValue("ENDRATE", dblENDRATE), _
                        GetFieldNameValue("ENDAMT", dblENDAMT), _
                        GetFieldNameValue("THISRATE", dblTHISRATE), _
                        GetFieldNameValue("THISAMT", dblTHISAMT), _
                        GetFieldNameValue("BALANCERATE", dblBALANCERATE), _
                        GetFieldNameValue("BALANCEAMT", dblBALANCEAMT), _
                        GetFieldNameValue("DELIVERYGUARANTY", dblDELIVERYGUARANTY), _
                        GetFieldNameValue("FAULTGUARANTY", dblFAULTGUARANTY), _
                        GetFieldNameValue("PAYMENTGBN", strPAYMENTGBN), _
                        GetFieldNameValue("STDATE", strSTDATE), _
                        GetFieldNameValue("EDDATE", strEDDATE), _
                        GetFieldNameValue("CONTRACTDAY", strCONTRACTDAY), _
                        GetFieldNameValue("MANAGER", strMANAGER), _
                        GetFieldNameValue("TESTDAY", strTESTDAY), _
                        GetFieldNameValue("TESTENDDAY", strTESTENDDAY), _
                        GetFieldNameValue("TESTMENT", strTESTMENT), _
                        GetFieldNameValue("TESTAMT", dblTESTAMT), _
                        GetFieldNameValue("LOSTDAY", dblLOSTDAY), _
                        GetFieldNameValue("CONFIRMFLAG", strCONFIRMFLAG), _
                        GetFieldNameValue("CONFLAG", strCONFLAG), _
                        GetFieldNameValue("DIVFLAG", strDIVFLAG), _
                        GetFieldNameValue("COMENT", strCOMENT), _
                        GetFieldNameValue("AMTFLAG", strAMTFLAG), _
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
                        GetFieldNameValue("CONTRACTNO", strCONTRACTNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������
    '*****************************************************************
    Public Function DeleteDo(ByVal strCONTRACTNO As String) As Integer
        Dim strSQL As String
        Dim strNOW As String
        strNOW = Now

        Try
            strSQL = "DELETE FROM PD_CONTRACT_HDR WHERE CONTRACTNO = '" & strCONTRACTNO & "';"
            strSQL = strSQL & " UPDATE PD_CONTRACT_DTL "
            strSQL = strSQL & " SET CONTRACTNO = '', CONFIRM_USER = '" & mobjSCGLConfig.WRKUSR & "', CONFIRM_DATE = '" & strNOW & "'"
            strSQL = strSQL & " WHERE CONTRACTNO = '" & strCONTRACTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������
    '*****************************************************************
    Public Function DeleteRtn_Confirm_HDR(ByVal strCONTRACTNO As String) As Integer
        Dim strSQL As String
        Dim strNOW As String
        strNOW = Now

        Try
            strSQL = "DELETE FROM PD_CONTRACT_HDR WHERE CONTRACTNO = '" & strCONTRACTNO & "';"
            strSQL = strSQL & " UPDATE PD_CONTRACT_DTL "
            strSQL = strSQL & " SET CONTRACTNO = '', CONFIRM_USER = '" & mobjSCGLConfig.WRKUSR & "', CONFIRM_DATE = '" & strNOW & "'"
            strSQL = strSQL & " WHERE CONTRACTNO = '" & strCONTRACTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateDo_CONFIRM(ByVal strCONTRACTNO As String, _
                                     ByVal strCONFIRMFLAG As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE PD_CONTRACT_HDR SET CONFIRMFLAG = '" & strCONFIRMFLAG & "' WHERE CONTRACTNO = '" & strCONTRACTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
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
        MyBase.EntityName = "PD_CONTRACT_HDR"     'Entity Name ����
    End Sub

#End Region
#End Region
End Class

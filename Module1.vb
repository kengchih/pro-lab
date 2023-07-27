Imports Wellan.Service
Imports System.Web.Util
Imports System.Xml
Imports System.Web.Mail
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
Imports System.Globalization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports PP00SS.Business
Imports Wellan.Ecommerce.StoreFront


Module Module1
    Public g_StartExe As String = ""
    Public g_backupPath As String = ""
    Dim MailFileName As String = ""
    Dim ErrorMessage As String = ""
    Sub main(ByVal CmdArgs() As String)
        g_backupPath = GetConfigValue("BackUpPath")
        If System.IO.Directory.Exists(g_backupPath) = False Then
            System.IO.Directory.CreateDirectory(g_backupPath)
            System.Threading.Thread.Sleep(1000)
        End If

        Select Case CmdArgs.Length
            Case 1
                If UCase(CmdArgs(0).Trim) = "/TSPG_STARTC" Then
                    TSPG_LoadSQLData("", "Capture")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTCRC" Then
                    TSPG_LoadSQLData("", "ClassroomReserveCapture")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTQ" Then
                    TSPG_LoadSQLData("", "AuthorizationQuery")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTQ-Y" Then
                    TSPG_LoadSQLData("", "AuthorizationQuery-Y")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTCRQ" Then
                    TSPG_LoadSQLData("", "ClassroomReserveAuthQuery")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTCRQ-Y" Then
                    TSPG_LoadSQLData("", "ClassroomReserveAuthQuery-Y")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTD" Then
                    TSPG_LoadSQLData("", "Deauthorization")
                    'S230207031、S230215004 新增台新APP錢包排程，授權查詢、取消授權
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTQ_WALLETPAY" AndAlso BwexGetBSNAM("App", "IsWalletPay", "N") = "Y" Then
                    TSPG_WalletPay_LoadSQLData("", "AuthorizationQuery")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTQ-Y_WALLETPAY" AndAlso BwexGetBSNAM("App", "IsWalletPay", "N") = "Y" Then
                    TSPG_WalletPay_LoadSQLData("", "AuthorizationQuery-Y")
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTD_WALLETPAY" AndAlso BwexGetBSNAM("App", "IsWalletPay", "N") = "Y" Then
                    TSPG_WalletPay_LoadSQLData("", "Deauthorization")
                End If
            Case 2
                If UCase(CmdArgs(0).Trim) = "/TSPG_STARTC" Then
                    If UCase(CmdArgs(1).Trim) <> "" Then
                        TSPG_LoadSQLData(CmdArgs(1).Trim, "Capture")
                    End If
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTCRC" Then
                    If UCase(CmdArgs(1).Trim) <> "" Then
                        TSPG_LoadSQLData(CmdArgs(1).Trim, "ClassroomReserveCapture")
                    End If
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTQ" Then
                    If UCase(CmdArgs(1).Trim) <> "" Then
                        TSPG_LoadSQLData(CmdArgs(1).Trim, "AuthorizationQuery")
                    End If
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTCRQ" Then
                    If UCase(CmdArgs(1).Trim) <> "" Then
                        TSPG_LoadSQLData(CmdArgs(1).Trim, "ClassroomReserveAuthQuery")
                    End If
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTD" Then
                    If UCase(CmdArgs(1).Trim) <> "" Then
                        TSPG_LoadSQLData(CmdArgs(1).Trim, "Deauthorization")
                    End If
                    'S230207031、S230215004 新增台新APP錢包排程，授權查詢、取消授權
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTQ_WALLETPAY" AndAlso BwexGetBSNAM("App", "IsWalletPay", "N") = "Y" Then
                    If UCase(CmdArgs(1).Trim) <> "" Then
                        TSPG_WalletPay_LoadSQLData(CmdArgs(1).Trim, "AuthorizationQuery")
                    End If
                ElseIf UCase(CmdArgs(0).Trim) = "/TSPG_STARTD_WALLETPAY" AndAlso BwexGetBSNAM("App", "IsWalletPay", "N") = "Y" Then
                    If UCase(CmdArgs(1).Trim) <> "" Then
                        TSPG_WalletPay_LoadSQLData(CmdArgs(1).Trim, "Deauthorization")
                    End If
                End If
            Case Else
                Dim fm As New MainForm
                fm.ShowDialog()
        End Select

    End Sub


    ''' <summary>
    ''' 抓參數設定值(直接依傳入內容查詢，不分正式測試區)
    ''' </summary>
    ''' <param name="BSNUM"></param>
    ''' <param name="BSITM"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BwexGetBSNAM(ByVal BSNUM As String, ByVal BSITM As String, Optional ByVal DefaultValue As String = "") As String
        Dim strSQL As String
        strSQL = " select BSNAM from BWEX "
        strSQL += " WHERE BSNUM = '" + ReplaceStr(BSNUM) + "'"
        strSQL += " and BSITM='" + ReplaceStr(BSITM) + "'"
        Dim sStr As String = UtilService.GetVal(DataService.WLookup(strSQL, "BSNAM"), ValCollection.文字)

        Dim sReturn As String = IIf(Len(sStr) = 0, DefaultValue, sStr)
        Return sReturn
    End Function

    ''' <summary>
    ''' 抓參數設定值(依config的IsFormal判斷為正式/測試區，決定是否抓_TEST的內容)
    ''' </summary>
    ''' <param name="BSNUM"></param>
    ''' <param name="BSITM"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BwexGetBSNAM2(ByVal BSNUM As String, ByVal BSITM As String) As String
        If System.Configuration.ConfigurationManager.AppSettings("IsFormal") = "N" Then
            BSNUM += "_TEST"
        End If
        Dim strSQL As String
        strSQL = " select BSNAM from BWEX "
        strSQL += " WHERE BSNUM = '" + ReplaceStr(BSNUM) + "'"
        strSQL += " and BSITM='" + ReplaceStr(BSITM) + "'"
        Return UtilService.GetVal(DataService.WLookup(strSQL, "BSNAM"), ValCollection.文字)
    End Function
    Public Function ReplaceStr(ByVal mValue) As String
        If Trim(mValue) <> "" Then
            Return Replace(mValue, "'", "''")
        Else
            Return ""
        End If
    End Function
    Public Sub WriteLogData(ByVal LogData As DataSet, ByVal FunctionID As String, ByVal LogKeyValue As String, ByVal UpdateKind As String, Optional ByVal UpdateUser As String = "")
        Dim _Transaction As SqlClient.SqlTransaction
        Dim mIP, mXMLDataString As String
        Dim SerialNo As String = Guid.NewGuid.ToString
        mIP = String.Empty
        Dim mUpdateUser = String.Empty
        If UpdateUser = "" Then
            mUpdateUser = "EC"
        Else
            mUpdateUser = UpdateUser
        End If
        If LogData Is Nothing Then
            mXMLDataString = "<ROOT></ROOT>"
        Else
            Dim i As Integer
            For i = 0 To LogData.Tables.Count - 1
                AddSystemColumn(LogData.Tables(i), SerialNo)
            Next
            AddRelation(LogData)
            mXMLDataString = LogData.GetXml
        End If

        Dim sSQL As String = "Insert into LogData (UpdateUser,UpdateDate,UpdateKind," _
                  + "FunctionID,XMLData,IP,LogKeyValue) " _
                  + " Values('{0}',dbo.udf_GetSysDatetime(),'{1}','{2}','{3}','{4}','{5}')"
        sSQL = String.Format(sSQL, mUpdateUser, UpdateKind, FunctionID, Replace(mXMLDataString, "'", "''"), mIP, LogKeyValue)
        Try
            DataService.ExecuteSQL(_Transaction, sSQL)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub AddRelation(ByRef LogDS As DataSet)
        'Add Relationship           
        If LogDS.Tables.Count > 1 Then
            Dim j As Integer
            Dim parentCol As DataColumn
            '預設第一個Table為主檔，其他的為明細檔
            parentCol = LogDS.Tables(0).Columns("SerialNo")
            For j = 1 To LogDS.Tables.Count - 1
                Dim childCol As DataColumn
                childCol = LogDS.Tables(j).Columns("SerialNo")
                Dim relLog As DataRelation
                relLog = New DataRelation("Log" + j.ToString, parentCol, childCol)
                LogDS.Relations.Add(relLog)
            Next
        End If
    End Sub
    Private Sub AddSystemColumn(ByRef dt As DataTable, ByVal SerialNo As String)
        '加入SerialNo欄位
        Dim MyCol As New DataColumn
        MyCol.ColumnName = "SerialNo"
        MyCol.DataType = System.Type.GetType("System.String")
        MyCol.DefaultValue = SerialNo
        dt.Columns.Add(MyCol)
    End Sub

    ''' <summary>
    ''' 依設定決定回傳「格林威治時間」或「主機時間」
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSysDatetime() As DateTime
        If System.Configuration.ConfigurationManager.AppSettings("IsUseUTC") = "Y" Then
            Return DateTime.UtcNow
        Else
            Return DateTime.Now
        End If
    End Function


    ''' <summary>
    ''' 依訂單編號至ResponseCodeWebAtm取得訂單傳送序號，丟至合庫取得訂單付款資訊(是否付款成功)
    ''' </summary>
    ''' <param name="SALNO">訂單編號</param>
    ''' <remarks></remarks>
    Private Sub SendTcbWebAtmQueryWebRequest(ByVal SALNO As String)

        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TCBQuery.Log"

        '依訂單編號至ResponseCodeWebAtm取得訂單傳送序號
        Dim sSql As String = "select SendSeqNo from ResponseCodeWebAtm where od_sob='" + SALNO + "' and CommandType='PayRequest' and Isnull(SendSeqNo,'')<>''"
        Dim DsSendSeqNo As DataSet = DataService.CreateDataSet(sSql)
        For i As Integer = 0 To DsSendSeqNo.Tables(0).Rows.Count - 1
            Dim DebugData As String = String.Empty
            Dim AuthResultFalg As Boolean = False

            Dim PostURL As String = BwexGetBSNAM2("Store", "TCBPayURL") '合庫WebAtm交易網址
            'PostURL = BwexGetBSNAM("Store_TEST", "TCBPayURL") 'localhost測試用
            Dim postData As String = "CardBillInqRq={0}"
            Dim TxType As String = BwexGetBSNAM2("Store", "TCBTxType") '交易類別
            Dim PAYNO As String = BwexGetBSNAM2("Store", "TCBPAYNO") '特約商店代碼
            Dim RsURL As String = BwexGetBSNAM2("Store", "TCBInqReturnURL") '合庫WebAtm查詢交易狀態回傳網址
            Dim SendSeqNo = DsSendSeqNo.Tables(0).Rows(i)("SendSeqNo") '傳送序號

            'Hex2Base64(DES3(KEY,SHA -1( 交易類別 +特約商店+傳送序號)))
            Dim MacKey As String = NewTCBMacKey(TxType & PAYNO & SendSeqNo) '押碼
            Dim QueryXML As String = String.Empty
            QueryXML += "<?xml version='1.0' encoding='Big5'?><CardBillInqRq><SendSeqNo>" + SendSeqNo + "</SendSeqNo>" +
                "<TxType>" + TxType + "</TxType><PAYNO>" + PAYNO + "</PAYNO><MAC>" + MacKey + "</MAC><RsURL>" + RsURL + "</RsURL></CardBillInqRq>"

            '將XML轉為Base64後，放在CardBillInqRq中
            postData = String.Format(postData, System.Web.HttpUtility.UrlEncode(Convert.ToBase64String(Encoding.Default.GetBytes(QueryXML))))
            Dim byteData As Byte() = Encoding.Default.GetBytes(postData)

            '紀錄查詢時，丟入合庫之參數內容
            DebugData = "SendWebAtmQueryRequest of POST data : " + PostURL + ";Post Data String: " + postData + ";QueryXML: " + QueryXML
            SaveLogValue("WebAtmQueryPostData-" & SALNO & ":" & DebugData, sFilename)


            Dim webRequest As Net.HttpWebRequest = Net.HttpWebRequest.Create(PostURL)
            webRequest.Method = "POST"
            webRequest.ContentType = "application/x-www-form-urlencoded"
            webRequest.Timeout = 5000
            webRequest.KeepAlive = True
            webRequest.ContentLength = byteData.Length
            Dim postreqstream As Stream = webRequest.GetRequestStream()
            postreqstream.Write(byteData, 0, byteData.Length)
            postreqstream.Close()

            Dim postresponse As Net.HttpWebResponse
            postresponse = DirectCast(webRequest.GetResponse(), Net.HttpWebResponse)
            Dim postreqreader As New StreamReader(postresponse.GetResponseStream())


            Dim ReturnPage As String = postreqreader.ReadToEnd
            SaveLogValue("WebAtmQueryPostData ResponseData-" & SALNO & ":" & ReturnPage, sFilename)

            '實例化HtmlAgilityPack.HtmlDocument對像
            Dim Htmldoc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument()
            '載入HTML
            Htmldoc.LoadHtml(ReturnPage)
            Dim mRqNode As HtmlAgilityPack.HtmlNode = Htmldoc.GetElementbyId("CardBillInqRs")
            Dim endCardBillInqRs As String = mRqNode.Attributes("value").Value
            '紀錄接收到的回傳結果  
            SaveLogValue("WebAtmQueryPostData-" & SALNO & ",SendSeqNo:" & SendSeqNo & ",InqResponseXMLinBase64: " + endCardBillInqRs, sFilename)

            Dim CardBillInqRs As String = System.Text.Encoding.GetEncoding("Big5").GetString(Convert.FromBase64String(endCardBillInqRs))
            '紀錄接收到的回傳結果做Base64解譯後的內容
            SaveLogValue("WebAtmQueryPostData-" & SALNO & ",SendSeqNo:" & SendSeqNo & ",InqResponseXML: " + CardBillInqRs, sFilename)


            Dim Flag As Boolean = False
            Dim RC As String = String.Empty '回覆碼
            Dim MSG As String = String.Empty '回覆碼說明
            Dim rSendSeqNo As String = String.Empty '交易序號
            Dim ONO As String = String.Empty '訂單編號
            Dim CurAmt As String = String.Empty
            Dim TransDate As Date '交易時間
            Dim AcctIdTo As String = String.Empty '銷帳編號(虛擬帳號)
            Dim BankIdFrom As String = String.Empty '轉出銀行別
            Dim AcctIdFrom As String = String.Empty '轉出帳號
            Dim TxnSeqNo As String = String.Empty '金流回傳交易序號
            Dim errorMsg As String = String.Empty
            Dim isError As Boolean = False

            Dim doc As New System.Xml.XmlDocument
            doc.LoadXml(CardBillInqRs.Replace(ControlChars.Quote, "'"))

            Dim mXmlNode As System.Xml.XmlNode

            '回覆碼
            mXmlNode = doc.SelectSingleNode("/CardBillInqRs/RC")
            RC = Trim(mXmlNode.InnerText)

            '回覆碼說明
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/MSG")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/MSG")
                MSG = Trim(mXmlNode.InnerText)
            End If
            If getTCBRcDescription(RC) <> Trim(MSG) Then '合庫回傳的回覆碼說明如果與文件註明的回覆碼說明不一樣，則也加入記錄
                MSG += "(" & getTCBRcDescription(RC) & ")"
            End If

            '交易序號
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/SendSeqNo")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/SendSeqNo")
                rSendSeqNo = Trim(mXmlNode.InnerText)
            End If

            '訂單編號
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/ONO")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/ONO")
                ONO = Trim(mXmlNode.InnerText)
            End If

            If SALNO <> ONO OrElse String.IsNullOrEmpty(SALNO) OrElse String.IsNullOrEmpty(ONO) Then
                isError = True
                errorMsg += "合庫回傳之訂單編號(" + ONO + ")與在送出查詢之訂單編號(" + SALNO + ")不符 "
            End If
            If SendSeqNo <> rSendSeqNo OrElse String.IsNullOrEmpty(SendSeqNo) OrElse String.IsNullOrEmpty(rSendSeqNo) Then
                isError = True
                errorMsg += "合庫回傳之傳送序號(" + rSendSeqNo + ")與在送出查詢之傳送序號(" + SendSeqNo + ")不符 "
            End If

            '金額
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/CurAmt")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/CurAmt")
                CurAmt = Trim(mXmlNode.InnerText)
            End If

            '交易時間
            Dim mTranDate As String = String.Empty
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/TrnDt")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/TrnDt") '交易日期yyyymmdd
                mTranDate += Trim(mXmlNode.InnerText)
            End If
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/TrnTime")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/TrnTime") '交易時間HHmmss
                mTranDate += Trim(mXmlNode.InnerText)
            End If
            If Not String.IsNullOrEmpty(mTranDate) Then
                TransDate = DateTime.ParseExact(mTranDate, "yyyyMMddHHmmss", CultureInfo.InvariantCulture)
            End If

            '銷帳編號(虛擬帳號)
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/AcctIdTo")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/AcctIdTo")
                AcctIdTo = Trim(mXmlNode.InnerText)
            End If

            '轉出銀行別
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/BankIdFrom")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/BankIdFrom")
                BankIdFrom = Trim(mXmlNode.InnerText)
            End If

            '轉出帳號
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/AcctIdFrom")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/AcctIdFrom")
                AcctIdFrom = Trim(mXmlNode.InnerText)
            End If

            '金流回傳交易序號
            If Not IsNothing(doc.SelectSingleNode("/CardBillInqRs/TxnSeqNo")) Then
                mXmlNode = doc.SelectSingleNode("/CardBillInqRs/TxnSeqNo")
                TxnSeqNo = Trim(mXmlNode.InnerText)
            End If


            Dim mResponseCode As String
            mResponseCode = "INSERT INTO [ResponseCodeWebAtm]  ([od_sob],[xCartID],[amount],[CreateTime],[TransTime],[xresponsecode],[ResponseCodeDesc],[SendSeqNo],[TxnSeqNo],[CommandType],[OrderInfo],[AccIdTo],[BandIdFrom],[AccIdFrom],[errormsg])"
            mResponseCode += "VALUES('{0}','{1}','{2}',getdate(),{3},'{4}',N'{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}',N'{13}')"
            mResponseCode = String.Format(mResponseCode, SALNO, "", CurAmt, If(String.IsNullOrEmpty(mTranDate), "NULL", "'" + UtilService.GetDateTime(TransDate, 3) + "'"), RC, MSG, SendSeqNo, TxnSeqNo, "InqResult-PPOpenIE", ONO, AcctIdTo, BankIdFrom, AcctIdFrom, errorMsg)
            Wellan.Service.DataService.ExecuteSQL(mResponseCode)

            If isError Then
                '合庫回傳之訂單編號與在送出查詢之訂單編號不符 或 合庫回傳之傳送序號與在送出查詢之傳送序號不符 
                Dim sText As String = String.Format("WebAtmQueryPostData- SalesOrderNo:{0}, ONO:{1}, SendSeqNo:{2}, TcbSendSeqNo:{3} SalesOrderNo/SenqSeqNo is not equal ONO/TcbSendSeqNo which from TCB", SALNO, ONO, SendSeqNo, rSendSeqNo)
                SaveLogValue(sText, sFilename)
            ElseIf UtilService.GetVal(DataService.ExecuteSQLScale("select OrderStatus from SalesOrderA where SalesOrderNo='" & SALNO & "'"), ValCollection.文字) = "1" Then
                '若查詢付款結果之訂單狀態為已完成，則不更新交易結果
                '重覆付款情形下，同一筆訂單會有多筆傳送序號，先前的傳送序號查詢為付款成功，但後來的傳送序號查詢為付款失敗
                '為避免後面的查詢將原本訂單已完成狀態改為失敗，故若訂單狀態為已完成則不更新交易結果
                Dim sText As String = String.Format("WebAtmQueryPostData- SalesOrderNo:{0}, OrderStatus is Completed", SALNO)
                SaveLogValue(sText, sFilename)
            Else
                Dim mUpdateBillB As String = String.Empty

                '‘000’、’4001’、’4002’收單交易成功
                '其它表示收單交易失敗
                If RC = "000" OrElse RC = "4001" OrElse RC = "4002" Then
                    mUpdateBillB = " update salesordera set Orderstatus='1' where salesorderno='{0}'"
                    mUpdateBillB += " update billB set BankNo='{1}',CheckNo='{2}',AuthDate='{3}' where BillNo='{0}'"
                    mUpdateBillB = String.Format(mUpdateBillB, SALNO, BankIdFrom, AcctIdFrom, UtilService.GetDateTime(TransDate, 1))
                    DataService.ExecuteSQL(mUpdateBillB)
                    Flag = True
                Else
                    '付款失敗
                    mUpdateBillB = " update salesordera set Orderstatus='3' where salesorderno='{0}'"
                    mUpdateBillB = String.Format(mUpdateBillB, SALNO)
                    DataService.ExecuteSQL(mUpdateBillB)
                    Flag = False
                End If

                'Demo code for returnURL 寫入訂單記錄
                Dim MyConnection As SqlClient.SqlConnection = Wellan.Service.SqlConnection.GetConnection()
                Dim MyTrans As SqlClient.SqlTransaction

                If Not (MyConnection.State = ConnectionState.Open) Then
                    MyConnection.Open()
                End If
                MyTrans = MyConnection.BeginTransaction()
                Try

                    Dim LogData As DataSet
                    Dim LogDataSQL As String = ""

                    '新增LogData
                    LogDataSQL = "select * from SalesOrderA where SalesOrderNO = '" & SALNO & "'"
                    LogData = Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "SalesOrderA")
                    LogDataSQL = "select * from SalesOrderB where SalesOrderNO = '" & SALNO & "'"
                    LogData.Tables.Add(Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "SalesOrderB").Tables(0).Copy)
                    LogDataSQL = "select * from BillB where BillNo = '" & SALNO & "'"
                    LogData.Tables.Add(Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "BillB").Tables(0).Copy)
                    WriteLogData(LogData, "SalesOrder", SALNO, "Update", "EC-PPOpenIE")

                    MyTrans.Commit()
                Catch ex As Exception
                    MyTrans.Rollback()
                Finally
                    MyTrans = Nothing
                    MyConnection.Close()
                End Try
            End If
        Next
    End Sub

    Private Function NewTCBMacKey(ByVal dataString As String) As String
        Dim keyString As String = "12345678ABCDEFGH87654321"
        Dim ivString As String = "00000000"

        Dim MacKey As String = String.Empty
        Dim Data As Byte() = Encoding.Default.GetBytes(dataString)
        Dim sha As SHA1 = New SHA1CryptoServiceProvider()
        Dim sha1Byte As Byte() = sha.ComputeHash(Data)

        Dim tdes As TripleDES = TripleDES.Create()
        tdes.IV = Encoding.Default.GetBytes(ivString)
        tdes.Key = Encoding.Default.GetBytes(keyString)
        tdes.Mode = CipherMode.CBC
        tdes.Padding = PaddingMode.Zeros

        Dim ict As ICryptoTransform = tdes.CreateEncryptor()
        Dim enc As Byte() = ict.TransformFinalBlock(sha1Byte, 0, sha1Byte.Length)
        Dim encBase64 As String = Convert.ToBase64String(enc)
        Return encBase64
    End Function

    Private Function getTCBRcDescription(RC As String) As String
        Dim mDescription As String = String.Empty
        Select Case RC
            Case "V211"
                mDescription = "傳送序號重複，訊息丟棄"
            Case "V199"
                mDescription = "主機交易逾時，Pending"
            Case "V198"
                mDescription = "傳送MAC不符，訊息丟棄"
            Case "V197"
                mDescription = "發生不明原因錯誤，請稍後查詢本筆交易是否成功"
            Case "V196"
                mDescription = "無此特約商店"
            Case "V195"
                mDescription = "發生不明原因錯誤，主機未處理"
            Case "V200"
                mDescription = "客戶取消交易"
            Case "000", "4001", "4002"
                mDescription = "收單交易成功"
            Case Else
                mDescription = "收單交易失敗"
        End Select
        Return mDescription
    End Function

    '======TSPG 台新金流

    Private Sub TSPG_SendWebRequest(ByVal SALNO As String, ByVal sKind As String, Optional ByVal OrderAmt As Decimal = 0)
        If sKind = "Capture" Then '台新金流 請款
            TSPG_SendCaptureWebRequest(SALNO, OrderAmt)
        ElseIf sKind = "ClassroomReserveCapture" Then '預約教室請款
            TSPG_SendCRCaptureWebRequest(SALNO, OrderAmt)
        ElseIf sKind = "AuthorizationQuery" OrElse sKind = "AuthorizationQuery-Y" Then '台新查詢授權
            TSPG_SendAuthQueryWebRequest(SALNO)
        ElseIf sKind = "ClassroomReserveAuthQuery" OrElse sKind = "ClassroomReserveAuthQuery-Y" Then '預約教室台新查詢授權
            TSPG_SendCRAuthQueryWebRequest(SALNO)
        ElseIf sKind = "Deauthorization" Then 'S190524016 取消授權
            TSPG_Deauthorization(SALNO)
        End If
    End Sub

    Public Sub TSPG_LoadSQLData(ByVal SALNO As String, Optional ByVal sKind As String = "Capture", Optional ByVal SALDATE As String = "")

        Dim StartTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss.fff")
        Dim OrderDate As String = SetWorkDate("01", 1)

        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_" & sKind & ".Log"
        SaveLogValue(sKind, "Start_Load", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")

        Try
            If SALNO <> "" Then
                If sKind = "ClassroomReserveCapture" Then
                    Dim Amt As Decimal = UtilService.GetVal(DataService.ExecuteSQLScale(" SELECT TotalAmt*100 as TotalAmt FROM ClassroomPreOrder WHERE ClassroomPreOrderNo = '" & SALNO & "' "), ValCollection.數值)
                    TSPG_SendWebRequest(SALNO, sKind, Amt)
                ElseIf sKind = "Capture" Then
                    Dim Amt As Decimal = UtilService.GetVal(DataService.ExecuteSQLScale("SELECT Amt*100 as TransAmt FROM BillB where BillNo = '" & SALNO & "' "), ValCollection.數值)
                    TSPG_SendWebRequest(SALNO, sKind, Amt)
                ElseIf sKind = "Deauthorization" Then
                    Dim IsCancelOrder As Decimal = UtilService.GetVal(DataService.ExecuteSQLScale(" SELECT COUNT(*) FROM SalesOrderA WHERE OrderStatus='4' AND SalesOrderNo='" & SALNO & "' "), ValCollection.數值)
                    If IsCancelOrder > 0 Then
                        TSPG_SendWebRequest(SALNO, sKind)
                    End If
                Else
                    TSPG_SendWebRequest(SALNO, sKind)
                End If
            ElseIf SALDATE <> "" Then
                If sKind = "Deauthorization" Then
                    'PaymentType	19	線上刷卡(台新)
                    'S190524016 判斷"前1天"&"已作廢"且"刷卡交易已授權但未請款"的訂單
                    '台灣區訂單狀態作廢，付款方式線上刷卡，有刷卡授權碼，無請款授權碼
                    'SalesOrderStatus    1	已完成
                    'SalesOrderStatus    2	未完成
                    'SalesOrderStatus    3	付款失敗
                    'SalesOrderStatus    4	作廢
                    Dim sInternetSalesType As String = Replace(BwexGetBSNAM("SalesOrder", "InterentSalesType"), ",", "','")
                    If String.IsNullOrEmpty(sInternetSalesType) Then sInternetSalesType = "INTERNET','APP"
                    Dim strSQL As String = "Select  a.SalesOrderNo,b.Amt*100 as TransAmt,b.AuthDate,a.TotalAmt "
                    strSQL += "from SalesOrderA a WITH (NOLOCK) "
                    strSQL += " inner join BillB b WITH (NOLOCK) on a.SalesOrderNo=b.BillNo "
                    strSQL += " left join OrganizationObject o WITH (NOLOCK) on a.OrgNo = o.OrgNo "
                    strSQL += " where o.RegionNo='01' and a.OrderStatus='4' and isnull(b.PaymentType,'')='19' and isnull(b.AuthCode,'')<>'' and isnull(b.CaptureAuthCode,'')=''  and b.Authdate >= '" & System.Configuration.ConfigurationManager.AppSettings("checkDate") & "' "
                    strSQL += " and a.SalesType IN  ( '" & sInternetSalesType & "' ) "

                    strSQL += " AND a.SalesDate = '" & UtilService.GetDateTime(CDate(SALDATE), 1) & "' "
                    strSQL += " order by a.SalesOrderNo "

                    OrderDate = UtilService.GetDateTime(CDate(SALDATE), 1)
                    SaveLogValue(sKind, " SQL: " & strSQL, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")

                    Dim ExecutedOrderNo As String = ""
                    '非預約教室依訂單單號查詢 => 網路訂單
                    Dim dsSalesOrderA As DataSet = DataService.CreateDataSet(strSQL)
                    SaveLogValue(sKind, "Count:" & dsSalesOrderA.Tables(0).Rows.Count, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
                    If dsSalesOrderA Is Nothing = False AndAlso dsSalesOrderA.Tables(0).Rows.Count > 0 Then
                        For i As Integer = 0 To dsSalesOrderA.Tables(0).Rows.Count - 1
                            TSPG_SendWebRequest(dsSalesOrderA.Tables(0).Rows(i)("SalesorderNo"), sKind, dsSalesOrderA.Tables(0).Rows(i)("TransAmt"))
                            ExecutedOrderNo += "','" & dsSalesOrderA.Tables(0).Rows(i)("SalesorderNo")
                        Next
                    End If

                    SaveLogValue(sKind & "ByDate_ExecutedOrderNo", ExecutedOrderNo & "'  ", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecutedOrderNo.Log", "info")
                    '等待10秒
                    System.Threading.Thread.Sleep(20000)

                End If
            Else
                Dim strSQL As String = String.Empty

                If sKind = "Capture" Then '請款
                    'PaymentType	19	線上刷卡(台新)
                    '台灣區訂單狀態已完成，付款方式線上刷卡，有刷卡授權碼，無請款授權碼
                    'S170927013 請款排程條件增加訂單類別為網路訂單 
                    Dim sInternetSalesType As String = Replace(BwexGetBSNAM("SalesOrder", "InterentSalesType"), ",", "','")
                    If String.IsNullOrEmpty(sInternetSalesType) Then sInternetSalesType = "INTERNET','APP"
                    strSQL = "Select  a.SalesOrderNo,b.Amt*100 as TransAmt,b.AuthDate,a.TotalAmt "
                    strSQL += "from SalesOrderA a WITH (NOLOCK) "
                    strSQL += " inner join BillB b WITH (NOLOCK) on a.SalesOrderNo=b.BillNo "
                    strSQL += " left join OrganizationObject o WITH (NOLOCK) on a.OrgNo = o.OrgNo "
                    strSQL += " where o.RegionNo='01' and a.OrderStatus='1' and isnull(b.PaymentType,'')='19' and isnull(b.AuthCode,'')<>'' and isnull(b.CaptureAuthCode,'')=''  and b.Authdate >= '" & System.Configuration.ConfigurationManager.AppSettings("checkDate") & "' "
                    strSQL += " and a.SalesType IN  ( '" & sInternetSalesType & "' ) "

                    'S180903032 . 調整台新請款時間 7:30請款前一天的帳款 
                    strSQL += " AND a.SalesDate <= '" & UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1) & "' "

                    strSQL += " order by a.SalesOrderNo "
                    OrderDate = UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1)

                ElseIf sKind = "ClassroomReserveCapture" Then
                    '預約單狀態已完成，付款狀態刷卡成功，有刷卡授權碼，無請款授權碼，AddUser=EC
                    'S180608011 將教室預約的請款排程納入"預約單狀態=預約取消"的預約單也要請款 ，需要以轉正式訂單功能同步
                    'ReserveStatus   1	已預約, 3	取消預約
                    'S180326017 前台預約單成立、付款完成時，就產生正式訂單 
                    Dim OpenReserveStatus As String = " '1' "
                    If BwexGetBSNAM("Store", "IsAutoGenClassSalesOrder", "Y") = "Y" Then
                        OpenReserveStatus = " '1','3' "
                    End If
                    strSQL = " Select ClassroomPreOrderNo,TotalAmt*100 as TotalAmt  from ClassroomPreOrder where ReserveStatus IN (" & OpenReserveStatus & ") and PaymentStatus='0' and isnull(AuthCode,'')<>'' and isnull(CaptureAuthCode,'')='' and AddUser='EC' and isnull(PaymentType,'')='19' "

                    'S190122047 教室租借請款"前一天 00:00~23:59"+"已預約"的教室預約單
                    strSQL += " AND OrderDate <= '" & UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1) & "' "

                    OrderDate = UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1)

                ElseIf sKind = "AuthorizationQuery-Y" Then '查詢授權-只查昨天以前的訂單

                    '訂單狀態未完成，付款方式線上刷卡，訂單日期為昨天以前
                    'S201230048 將付款失敗條件移除
                    strSQL = "select * from udf_errAuthCode ('Y','{0}','19')"
                    strSQL = String.Format(strSQL, UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1))
                    OrderDate = UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1)

                ElseIf sKind = "AuthorizationQuery" Then '查詢授權-查N分鐘以前的訂單

                    '訂單狀態未完成，付款方式線上刷卡，新增時間為N分鐘以前的訂單
                    'S201230048 將付款失敗條件移除
                    strSQL = "select * from udf_errAuthCode ('M','{0}','19')"
                    strSQL = String.Format(strSQL, UtilService.GetDateTime(GetSysDatetime().AddMinutes(-CInt(BwexGetBSNAM("Store", "AuthQueryChkMins"))), 3))

                ElseIf sKind = "ClassroomReserveAuthQuery" Then '預約教室-查詢授權-查N分鐘以前的預約單

                    '預約單狀態為非刷卡失敗、預約狀態為已預約，AddUser=EC，新增時間為N分鐘以前的預約單
                    strSQL = "select * from udf_errClassroomReserve ('M','{0}','19')"
                    strSQL = String.Format(strSQL, UtilService.GetDateTime(GetSysDatetime().AddMinutes(-CInt(BwexGetBSNAM("Store", "ClassroomAuthQueryChkMins"))), 3))

                ElseIf sKind = "ClassroomReserveAuthQuery-Y" Then '預約教室-查詢授權-只查昨天以前的預約單

                    '預約單狀態為非刷卡失敗、預約狀態為已預約，AddUser=EC，預約單日期為昨天以前的預約單
                    strSQL = "select * from udf_errClassroomReserve ('Y','{0}','19')"
                    strSQL = String.Format(strSQL, UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1))
                    OrderDate = UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1)

                ElseIf sKind = "Deauthorization" Then
                    'PaymentType	19	線上刷卡(台新)
                    'S190524016 判斷"前1天"&"已作廢"且"刷卡交易已授權但未請款"的訂單
                    '台灣區訂單狀態作廢，付款方式線上刷卡，有刷卡授權碼，無請款授權碼
                    'SalesOrderStatus    1	已完成
                    'SalesOrderStatus    2	未完成
                    'SalesOrderStatus    3	付款失敗
                    'SalesOrderStatus    4	作廢
                    Dim sInternetSalesType As String = Replace(BwexGetBSNAM("SalesOrder", "InterentSalesType"), ",", "','")
                    If String.IsNullOrEmpty(sInternetSalesType) Then sInternetSalesType = "INTERNET','APP"
                    strSQL = "Select  a.SalesOrderNo,b.Amt*100 as TransAmt,b.AuthDate,a.TotalAmt "
                    strSQL += "from SalesOrderA a WITH (NOLOCK) "
                    strSQL += " inner join BillB b WITH (NOLOCK) on a.SalesOrderNo=b.BillNo "
                    strSQL += " left join OrganizationObject o WITH (NOLOCK) on a.OrgNo = o.OrgNo "
                    strSQL += " where o.RegionNo='01' and a.OrderStatus='4' and isnull(b.PaymentType,'')='19' and isnull(b.AuthCode,'')<>'' and isnull(b.CaptureAuthCode,'')=''  and b.Authdate >= '" & System.Configuration.ConfigurationManager.AppSettings("checkDate") & "' "
                    strSQL += " and a.SalesType IN  ( '" & sInternetSalesType & "' ) "
                    'S190524016 判斷"前1天"&"已作廢"且"刷卡交易已授權但未請款"的訂單
                    strSQL += " AND a.SalesDate = '" & UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1) & "' "
                    strSQL += " order by a.SalesOrderNo "

                    OrderDate = UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1)

                End If

                SaveLogValue(sKind, " SQL: " & strSQL, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")

                Dim ExecutedOrderNo As String = ""

                If sKind = "ClassroomReserveAuthQuery" OrElse sKind = "ClassroomReserveAuthQuery-Y" OrElse sKind = "ClassroomReserveCapture" Then
                    '預約教室依預約單號查詢
                    Dim dsClassroomPreOrder As DataSet = DataService.CreateDataSet(strSQL)
                    SaveLogValue(sKind, "Count:" & dsClassroomPreOrder.Tables(0).Rows.Count, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
                    If dsClassroomPreOrder Is Nothing = False AndAlso dsClassroomPreOrder.Tables(0).Rows.Count > 0 Then
                        For i As Integer = 0 To dsClassroomPreOrder.Tables(0).Rows.Count - 1
                            TSPG_SendWebRequest(dsClassroomPreOrder.Tables(0).Rows(i)("ClassroomPreOrderNo"), sKind, dsClassroomPreOrder.Tables(0).Rows(i)("TotalAmt"))
                            ExecutedOrderNo += "','" & dsClassroomPreOrder.Tables(0).Rows(i)("ClassroomPreOrderNo")
                        Next
                    End If
                Else
                    '非預約教室依訂單單號查詢 => 網路訂單
                    Dim dsSalesOrderA As DataSet = DataService.CreateDataSet(strSQL)
                    SaveLogValue(sKind, "Count:" & dsSalesOrderA.Tables(0).Rows.Count, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
                    If dsSalesOrderA Is Nothing = False AndAlso dsSalesOrderA.Tables(0).Rows.Count > 0 Then
                        For i As Integer = 0 To dsSalesOrderA.Tables(0).Rows.Count - 1
                            TSPG_SendWebRequest(dsSalesOrderA.Tables(0).Rows(i)("SalesorderNo"), sKind, dsSalesOrderA.Tables(0).Rows(i)("TransAmt"))
                            ExecutedOrderNo += "','" & dsSalesOrderA.Tables(0).Rows(i)("SalesorderNo")
                        Next
                    End If
                End If

                SaveLogValue(sKind & "_ExecutedOrderNo", ExecutedOrderNo & "'  ", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecutedOrderNo.Log", "info")
                '等待10秒
                System.Threading.Thread.Sleep(20000)
            End If

        Catch ex As Exception
            SaveLogValue(sKind, "Error:" & ex.Message, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
            SaveLogValue("TSPG_LoadSQLDataError", SALNO & ":" & ex.Message, sFilename, "info")
            ErrorMessage &= "執行中發生錯誤，請洽資訊人員!  " & Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "<BR>" & ex.Message
        End Try

        SaveLogValue(sKind, "END", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
        If System.Configuration.ConfigurationManager.AppSettings("SendMail") = "Y" AndAlso SALNO = "" Then

            Dim EndTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss.fff")
            Dim MailBody As String = "執行單據日期:" & OrderDate & "  開始時間:" & StartTime & "<BR>執行完成時間:" & EndTime & "<BR>" & ErrorMessage
            Dim MailBodytoStaff As String = "執行時間:" & StartTime & "<BR>排程執行完成"

            Try
                SendingMail(MailBody, sKind, MailBodytoStaff)
            Catch ex As Exception
                SaveLogValue("TSPG_SendMailError", ex.Message, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
            End Try
        End If

    End Sub
    Public Sub TSPG_WalletPay_LoadSQLData(ByVal SALNO As String, Optional ByVal sKind As String = "Capture", Optional ByVal SALDATE As String = "")
        'S230207031、S230215004 新增台新APP錢包排程，授權查詢、取消授權
        Dim StartTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss.fff")
        Dim OrderDate As String = SetWorkDate("01", 1)

        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_" & sKind & ".Log"
        SaveLogValue(sKind & "_WalletPay", "Start_Load", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")

        Try
            Dim ExecutedOrderNo As String = ""
            Dim dtGetECOrderInfo As DataTable
            Dim dtGetTS_RefundInfo As DataTable
            Dim strSQL As String = String.Empty
            If SALNO <> "" Then
                If sKind = "Deauthorization" Then
                    Dim IsCancelOrder As Decimal = UtilService.GetVal(DataService.ExecuteSQLScale(" SELECT COUNT(*) FROM SalesOrderA WHERE OrderStatus='4' AND SalesOrderNo='" & SALNO & "' "), ValCollection.數值)
                    If IsCancelOrder > 0 Then
                        'dtGetTS_RefundInfo = SB_Order.TS_Refund("zh-TW", "01", UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1))
                    End If
                Else
                    dtGetECOrderInfo = SB_Order.TS_GetECOrderInfo("zh-TW", "01", SALNO, "", "")
                    If dtGetECOrderInfo IsNot Nothing AndAlso dtGetECOrderInfo.Rows.Count > 0 Then
                        SaveLogValue(sKind & "_WalletPay", "TSPG_WalletPay_LoadSQLData:SalesOrderNo:" & SALNO & ",result:" & dtGetECOrderInfo.Rows(0)("Result") & ", ResultDesc:" & dtGetECOrderInfo.Rows(0)("ResultDesc") & ",FailCount:" & dtGetECOrderInfo.Rows(0)("FailCount") & ",SuccessCount:" & dtGetECOrderInfo.Rows(0)("SuccessCount") & ",RtnCode:" & dtGetECOrderInfo.Rows(0)("RtnCode") & ",RtnMessage:" & dtGetECOrderInfo.Rows(0)("RtnMessage"), g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_WalletPay.Log")
                    Else
                        SaveLogValue(sKind & "_WalletPay", "TSPG_WalletPay_LoadSQLData_API TS_GetECOrderInfo_DataTable Nothing:SalesOrderNo:" & SALNO, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_WalletPay.Log")
                    End If
                End If
            Else
                Select Case sKind
                    Case "AuthorizationQuery-Y"
                        '訂單狀態未完成，付款方式線上刷卡，訂單日期為昨天以前
                        'S201230048 將付款失敗條件移除
                        strSQL = "select * from udf_errAuthCode ('Y','{0}','21')"
                        strSQL = String.Format(strSQL, UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1))
                        OrderDate = UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1)
                        SaveLogValue(sKind & "_WalletPay", " SQL: " & strSQL, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
                    Case "AuthorizationQuery"
                        '訂單狀態未完成，付款方式線上刷卡，新增時間為N分鐘以前的訂單
                        'S201230048 將付款失敗條件移除
                        strSQL = "select * from udf_errAuthCode ('M','{0}','21')"
                        strSQL = String.Format(strSQL, UtilService.GetDateTime(GetSysDatetime().AddMinutes(-CInt(BwexGetBSNAM("Store", "AuthQueryChkMins"))), 3))
                        SaveLogValue(sKind & "_WalletPay", " SQL: " & strSQL, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
                    Case "Deauthorization"
                        dtGetTS_RefundInfo = SB_Order.TS_Refund("zh-TW", "01", UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1))
                        If dtGetTS_RefundInfo IsNot Nothing AndAlso dtGetTS_RefundInfo.Rows.Count > 0 Then
                            SaveLogValue(sKind & "_WalletPay", "TSPG_WalletPay_LoadSQLData:OrderDate:" & UtilService.GetDateTime(CDate(SetWorkDate("01", 1)).AddDays(-1), 1) & ",result:" & dtGetTS_RefundInfo.Rows(0)("Result") & ", ResultDesc:" & dtGetTS_RefundInfo.Rows(0)("ResultDesc") & ",RtnCode:" & dtGetTS_RefundInfo.Rows(0)("RtnCode") & ",RtnMessage:" & dtGetTS_RefundInfo.Rows(0)("RtnMessage"), g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_WalletPay.Log")
                        Else
                            SaveLogValue(sKind & "_WalletPay", "TSPG_WalletPay_LoadSQLData_API TS_Refund_DataTable Nothing:SalesOrderNo:" & SALNO, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_WalletPay.Log")
                        End If
                End Select

                If sKind = "AuthorizationQuery-Y" OrElse sKind = "AuthorizationQuery" Then

                    '非網路訂單
                    Dim dsSalesOrderA As DataSet = DataService.CreateDataSet(strSQL)
                    SaveLogValue(sKind & "_WalletPay", "Count:" & dsSalesOrderA.Tables(0).Rows.Count, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
                    If dsSalesOrderA Is Nothing = False AndAlso dsSalesOrderA.Tables(0).Rows.Count > 0 Then
                        For i As Integer = 0 To dsSalesOrderA.Tables(0).Rows.Count - 1
                            dtGetECOrderInfo = SB_Order.TS_GetECOrderInfo("zh-TW", "01", dsSalesOrderA.Tables(0).Rows(i)("SalesorderNo"), "", "")
                            If dtGetECOrderInfo IsNot Nothing AndAlso dtGetECOrderInfo.Rows.Count > 0 Then
                                SaveLogValue(sKind & "_WalletPay", "TSPG_WalletPay_LoadSQLData:SalesorderNo:" & dsSalesOrderA.Tables(0).Rows(i)("SalesorderNo") & ",result:" & dtGetECOrderInfo.Rows(0)("Result") & ", ResultDesc:" & dtGetECOrderInfo.Rows(0)("ResultDesc") & ",FailCount:" & dtGetECOrderInfo.Rows(0)("FailCount") & ",SuccessCount:" & dtGetECOrderInfo.Rows(0)("SuccessCount") & ",RtnCode:" & dtGetECOrderInfo.Rows(0)("RtnCode") & ",RtnMessage:" & dtGetECOrderInfo.Rows(0)("RtnMessage"), g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_WalletPay.Log")
                                ExecutedOrderNo += "','" & dsSalesOrderA.Tables(0).Rows(i)("SalesorderNo")
                            Else
                                SaveLogValue(sKind & "_WalletPay", "TSPG_WalletPay_LoadSQLData_API TS_GetECOrderInfo_DataTable Nothing:SalesOrderNo:" & SALNO, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_WalletPay.Log")
                            End If
                        Next
                    End If

                    SaveLogValue(sKind & "_WalletPay_" & "ByDate_ExecutedOrderNo", ExecutedOrderNo & "'  ", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecutedOrderNo.Log", "info")
                    '等待10秒
                    System.Threading.Thread.Sleep(20000)
                End If
            End If

        Catch ex As Exception
            SaveLogValue(sKind & "_WalletPay", "Error:" & ex.Message, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
            SaveLogValue("TSPG_WalletPay_LoadSQLDataError", SALNO & ":" & ex.Message, sFilename, "info")
            ErrorMessage &= "執行中發生錯誤，請洽資訊人員!  " & Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "<BR>" & ex.Message
        End Try

        SaveLogValue(sKind & "_WalletPay", "END", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
        If System.Configuration.ConfigurationManager.AppSettings("SendMail") = "Y" AndAlso SALNO = "" Then

            Dim EndTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss.fff")
            Dim MailBody As String = "執行單據日期:" & OrderDate & "  開始時間:" & StartTime & "<BR>執行完成時間:" & EndTime & "<BR>" & ErrorMessage
            Dim MailBodytoStaff As String = "執行時間:" & StartTime & "<BR>排程執行完成"

            Try
                SendingMail(MailBody, sKind, MailBodytoStaff)
            Catch ex As Exception
                SaveLogValue("TSPG_WalletPay_SendMailError", ex.Message, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
            End Try
        End If

    End Sub

    Public Sub TSPG_SendCaptureWebRequest(ByVal mSalesOrderNo As String, ByVal mAmt As Decimal)
        '台新金流 請款
        Dim DebugData As String = String.Empty
        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_Capture.Log"
        Dim mUpdateBillB As String = String.Empty
        Dim mResponseCode As String = String.Empty

        'S180808021 現場自取使用的金流刷卡的商店代號皆要獨立
        Dim TSPG_mid As String = SalesOrder_GetTSPG_mid(mSalesOrderNo) '特店代號

        Dim TSPG_APIROOT As String = BwexGetBSNAM2("Store", "TSPG_APIROOT") '路徑
        Dim TSPG_tid As String = BwexGetBSNAM2("Store", "TSPG_tid") '端末代號
        Dim TSPG_post_back_url As String = BwexGetBSNAM2("Store", "TSPG_post_back_url")
        Dim TSPG_result_url As String = BwexGetBSNAM2("Store", "TSPG_tid")

        Dim TSPG_URL As String = "https://" & TSPG_APIROOT & "/other.ashx"
        'S180604019  升級TLS 1.2
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim GetResponseData As JObject
        Dim CaptureAmt As Decimal = 0
        Dim txnResponseCode As String = ""
        Dim msg As String = ""
        Dim auth_id_resp As String = ""

        Try

            Dim json As String = "{""sender"":""rest""," + """ver"":""1.0.0""," + """mid"":""" & TSPG_mid & """," + """tid"":""" & TSPG_tid & """," + """pay_type"":1," + """tx_type"":3," + """params"":" + "{ ""order_no"":""" & mSalesOrderNo & """,""amt"":""" & mAmt & """,""result_flag"":""1""}}"

            SaveLogValue("Capture", mSalesOrderNo & "  " & json & " TSPG_APIROOT " & TSPG_APIROOT, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_JSON.Log")

            Dim webClient As System.Net.WebClient = New System.Net.WebClient()
            webClient.Headers.Add("Content-Type", "application/json")

            Dim reply As String = webClient.UploadString("https://" & TSPG_APIROOT & "/other.ashx", "POST", json)

            GetResponseData = CType(JsonConvert.DeserializeObject(reply), JObject)
            SaveLogValue("Capture", mSalesOrderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_GetResponseData.Log")
            txnResponseCode = GetResponseData("params")("ret_code")

            If txnResponseCode = "00" Then

                msg = ReturnTSPG_OrderStatus(GetResponseData("params")("order_status"))
                CaptureAmt = Convert.ToInt32(GetResponseData("params")("settle_amt").ToString) / 100
                auth_id_resp = GetResponseData("params")("auth_id_resp")

                mUpdateBillB = " update billB set CaptureAuthDate='{1}',CaptureAuthCode='{2}' where BillNo='{0}'"
                mUpdateBillB = String.Format(mUpdateBillB, mSalesOrderNo, UtilService.GetDateTime(Now, 3), GetResponseData("params")("settle_seq"))
                DataService.ExecuteSQL(mUpdateBillB)

                Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & mSalesOrderNo & "'")
                mResponseCode = "exec usp_ResponseCode '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','Capture','19'"
                mResponseCode = String.Format(mResponseCode, mRegionNo, mSalesOrderNo, "", CaptureAmt, auth_id_resp, txnResponseCode, msg, "", msg)
                Wellan.Service.DataService.ExecuteSQL(mResponseCode)

                Dim LogData As DataSet
                Dim LogDataSQL As String = String.Empty
                '新增LogData
                LogDataSQL = "select * from BillB where BillNo = '" & mSalesOrderNo & "'"
                LogData = Wellan.Service.DataService.CreateDataSet(LogDataSQL, "BillB")
                WriteLogData(LogData, "BillB", mSalesOrderNo, "Update")
            Else
                ' msg = GetResponseData("params")("ret_msg")
                SaveLogValue("Capture_Fail ", mSalesOrderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_Capture_Fail.Log", "info")
            End If


        Catch ex As Exception
            '訂單地區別
            Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & mSalesOrderNo & "'")
            mResponseCode = "exec usp_ResponseCode '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','Capture','19'"

            mResponseCode = String.Format(mResponseCode, mRegionNo, mSalesOrderNo, "", CaptureAmt, "", txnResponseCode, msg, "", ex.Message + ";" + msg)
            Wellan.Service.DataService.ExecuteSQL(mResponseCode)
            SaveLogValue("CaptureError", mSalesOrderNo & ":(51) Exception encountered. " + ex.Message, sFilename, "info")
        End Try
    End Sub
    ''' <summary>
    ''' 依訂單編號，丟至台新取得訂單授權資訊(是否授權成功)
    ''' </summary>
    ''' <param name="SALNO">訂單編號</param>
    ''' <remarks></remarks>
    Public Sub TSPG_SendAuthQueryWebRequest(ByVal mSalesOrderNo As String)
        Dim DebugData As String = String.Empty
        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthorizationQuery.Log"
        Dim mUpdateBillB As String = String.Empty
        Dim mResponseCode As String = String.Empty

        'S180808021 現場自取使用的金流刷卡的商店代號皆要獨立
        Dim TSPG_mid As String = SalesOrder_GetTSPG_mid(mSalesOrderNo) '特店代號

        Dim TSPG_APIROOT As String = BwexGetBSNAM2("Store", "TSPG_APIROOT") '路徑
        Dim TSPG_tid As String = BwexGetBSNAM2("Store", "TSPG_tid") '端末代號
        Dim TSPG_post_back_url As String = BwexGetBSNAM2("Store", "TSPG_post_back_url")
        Dim TSPG_result_url As String = BwexGetBSNAM2("Store", "TSPG_tid")
        Dim TSPG_URL As String = "https://" & TSPG_APIROOT & "/other.ashx"

        'S180604019  升級TLS 1.2
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim GetResponseData As JObject
        Dim CaptureAmt As Decimal = 0
        Dim txnResponseCode As String = ""
        Dim msg As String = ""

        Dim sPaymentSubmit As DataTable
        Dim ds As DataSet = DataService.CreateDataSet(" SELECT MemberID,vMemberID,TotalAmt FROM SalesOrderA WHERE SalesOrderNo = '" & mSalesOrderNo & "' ")
        Dim MemberID As String = UtilService.GetVal(ds.Tables(0).Rows(0).Item("MemberID"), ValCollection.文字)
        Dim TotalAmt As String = UtilService.GetVal(ds.Tables(0).Rows(0).Item("TotalAmt"), ValCollection.文字)
        Dim vMemberID As String = UtilService.GetVal(ds.Tables(0).Rows(0).Item("vMemberID"), ValCollection.文字)

        Try

            Dim json As String = "{""sender"":""rest""," + """ver"":""1.0.0""," + """mid"":""" & TSPG_mid & """," + """tid"":""" & TSPG_tid & """," + """pay_type"":1," + """tx_type"":7," + """params"":" + "{ ""order_no"":""" & mSalesOrderNo & """,""result_flag"":""1""}}"

            SaveLogValue("AuthQuery", mSalesOrderNo & "  " & json & " TSPG_APIROOT " & TSPG_APIROOT, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_JSON.Log")

            Dim webClient As System.Net.WebClient = New System.Net.WebClient()
            webClient.Headers.Add("Content-Type", "application/json")

            Dim reply As String = webClient.UploadString("https://" & TSPG_APIROOT & "/other.ashx", "POST", json)

            GetResponseData = CType(JsonConvert.DeserializeObject(reply), JObject)
            SaveLogValue("AuthQuery", mSalesOrderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_GetResponseData.Log")

            txnResponseCode = GetResponseData("params")("ret_code")

            Dim order_status As String = ""
            Dim auth_id_resp As String = "" '授權失敗 無回傳授傳碼
            Dim tx_amt As Decimal
            Dim IsOverQty As String = "" 'S210927023 是否超量，付款成功須等API處理回傳，付款失敗:N

            If txnResponseCode = "00" Then

                order_status = GetResponseData("params")("order_status").ToString() '訂單狀態碼

                tx_amt = CType(GetResponseData("params")("tx_amt").ToString(), Decimal) / 100 '交易金額包含兩位小數，如100 代表1.00 元。
                Dim last_4_digit_of_pan As String = GetResponseData("params")("last_4_digit_of_pan").ToString() '卡號後4 碼

                msg = GetResponseData("params")("order_status").ToString() & ReturnTSPG_OrderStatus(GetResponseData("params")("order_status"))
                CaptureAmt = CType(GetResponseData("params")("tx_amt").ToString(), Decimal) / 100

                '02 已授權
                '03 已請款
                '04 請款已清算
                '06 已退貨
                '08 退貨已清算
                '12 訂單已取消
                'ZP 訂單處理中
                'ZF 授權失敗
                If order_status = "02" OrElse order_status = "03" Then
                    auth_id_resp = GetResponseData("params")("auth_id_resp").ToString() '授權碼

                    'S180723060 金流完成導回時除了異動訂單, 需再判斷APP自取訂單需開發票及出貨單, 此部份改由API統一處理 
                    If BwexGetBSNAM("Store", "IsOpenAPP2", "Y") = "Y" Then

                        sPaymentSubmit = SB_Order.PaymentSubmit("zh-TW", "01", mSalesOrderNo, MemberID, TotalAmt, tx_amt, auth_id_resp, SetWorkDate("01", 1), last_4_digit_of_pan, TSPG_mid, "1", "", vMemberID)
                        'S210927023 超量處理，若PaymentSubmit API回傳Result -1 且 ResultCode = PaymentSubmit32 為此單會員出現購買超量 (API 問題單號S210927024)
                        If sPaymentSubmit.Rows(0)("Result") = "-1" AndAlso sPaymentSubmit.Rows(0)("ResultCode") = "PaymentSubmit32" Then
                            IsOverQty = "Y"
                        Else
                            IsOverQty = "N"
                        End If

                        SaveLogValue("AuthQuery PaymentSubmit API SalesOrderNo:", mSalesOrderNo & " ,result:" & sPaymentSubmit.Rows(0)("Result") & ",ResultDesc:" & sPaymentSubmit.Rows(0)("ResultDesc") & ",ResultCode:" & sPaymentSubmit.Rows(0)("ResultCode"), g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_PaymentSubmitAPI.Log")

                    Else '未啟用APP自取

                        mUpdateBillB = " update salesordera set Orderstatus='1' ,UpdateUser='System',UpdateDatetime= dbo.udf_GetSysDatetime() where salesorderno='{0}' "
                        mUpdateBillB += " update billB set AuthDate='{1}',authCode='{2}',CardType='{3}',StoreID='{4}',CheckNo='{5}' where BillNo='{0}'"
                        mUpdateBillB = String.Format(mUpdateBillB, mSalesOrderNo, GetSysDatetime.ToString("yyyy/MM/dd HH:mm:ss"), auth_id_resp, "", BwexGetBSNAM2("Store", "TSPG_mid"), last_4_digit_of_pan)
                        DataService.ExecuteSQL(mUpdateBillB)

                    End If

                ElseIf order_status = "ZF" Then

                    'S180723060 金流完成導回時除了異動訂單, 需再判斷APP自取訂單需開發票及出貨單, 此部份改由API統一處理 
                    If BwexGetBSNAM("Store", "IsOpenAPP2", "Y") = "Y" Then
                        sPaymentSubmit = SB_Order.PaymentSubmit("zh-TW", "01", mSalesOrderNo, MemberID, TotalAmt, tx_amt, "", SetWorkDate("01", 3), "", TSPG_mid, "-1", "", vMemberID)
                        IsOverQty = "N"
                        SaveLogValue("AuthQuery PaymentSubmit API SalesOrderNo:", mSalesOrderNo & " ,result:" & sPaymentSubmit.Rows(0)("Result") & ",ResultDesc:" & sPaymentSubmit.Rows(0)("ResultDesc") & ",ResultCode:" & sPaymentSubmit.Rows(0)("ResultCode"), g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_PaymentSubmitAPI.Log")
                    Else
                        '付款失敗
                        mUpdateBillB = " update salesordera set Orderstatus='3' ,UpdateDatetime = dbo.udf_GetSysDatetime() ,  UpdateUser = '{1}'  where salesorderno='{0}'"
                        mUpdateBillB = String.Format(mUpdateBillB, mSalesOrderNo, "System")
                        DataService.ExecuteSQL(mUpdateBillB)
                    End If


                End If

                Dim MyConnection As SqlClient.SqlConnection = SqlConnection.GetConnection()
                Dim MyTrans As SqlClient.SqlTransaction

                If Not (MyConnection.State = ConnectionState.Open) Then
                    MyConnection.Open()
                End If
                MyTrans = MyConnection.BeginTransaction()
                Try
                    Dim LogData As DataSet
                    Dim LogDataSQL As String = ""

                    '新增LogData
                    LogDataSQL = "select * from SalesOrderA where SalesOrderNO = '" & mSalesOrderNo & "'"
                    LogData = Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "SalesOrderA")
                    LogDataSQL = "select * from SalesOrderB where SalesOrderNO = '" & mSalesOrderNo & "'"
                    LogData.Tables.Add(Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "SalesOrderB").Tables(0).Copy)
                    LogDataSQL = "select * from BillB where BillNo = '" & mSalesOrderNo & "'"
                    LogData.Tables.Add(Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "BillB").Tables(0).Copy)
                    WriteLogData(LogData, "SalesOrder", mSalesOrderNo, "Update")

                    MyTrans.Commit()
                Catch ex As Exception
                    MyTrans.Rollback()
                Finally
                    MyTrans = Nothing
                    MyConnection.Close()
                End Try

            Else 'ret_code <> 00 亦為交易失敗

                'S180723060 金流完成導回時除了異動訂單, 需再判斷APP自取訂單需開發票及出貨單, 此部份改由API統一處理 
                If BwexGetBSNAM("Store", "IsOpenAPP2", "Y") = "Y" Then
                    sPaymentSubmit = SB_Order.PaymentSubmit("zh-TW", "01", mSalesOrderNo, MemberID, TotalAmt, tx_amt, "", SetWorkDate("01", 3), "", TSPG_mid, "-1", "", vMemberID)
                    IsOverQty = "N"
                    SaveLogValue("AuthQuery PaymentSubmit API SalesOrderNo:", mSalesOrderNo & " ,result:" & sPaymentSubmit.Rows(0)("Result") & ",ResultDesc:" & sPaymentSubmit.Rows(0)("ResultDesc") & ",ResultCode:" & sPaymentSubmit.Rows(0)("ResultCode"), g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_PaymentSubmitAPI.Log")
                Else
                    '付款失敗
                    mUpdateBillB = " update salesordera set Orderstatus='3' ,UpdateDatetime = dbo.udf_GetSysDatetime() ,  UpdateUser = '{1}'  where salesorderno='{0}'"
                    mUpdateBillB = String.Format(mUpdateBillB, mSalesOrderNo, GetSysDatetime.ToString("yyyy/MM/dd HH:mm:ss"), "System")
                    DataService.ExecuteSQL(mUpdateBillB)
                    SaveLogValue("AuthQuery_Fail", mSalesOrderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_AuthQuery_Fail.Log", "info")
                End If



            End If

            Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & mSalesOrderNo & "'")
            mResponseCode = "exec usp_ResponseCodeTW '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','AuthorizationQuery','19','{9}' "
            mResponseCode = String.Format(mResponseCode, mRegionNo, mSalesOrderNo, "", CaptureAmt, auth_id_resp, txnResponseCode & order_status, msg, GetResponseData("params")("rrn"), msg, IsOverQty)
            Wellan.Service.DataService.ExecuteSQL(mResponseCode)

        Catch ex As Exception
            '訂單地區別
            Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & mSalesOrderNo & "'")
            mResponseCode = "exec usp_ResponseCodeTW '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','AuthorizationQuery','19','N' "

            mResponseCode = String.Format(mResponseCode, mRegionNo, mSalesOrderNo, "", CaptureAmt, "", txnResponseCode, msg, "", ex.Message + ";" + msg)
            Wellan.Service.DataService.ExecuteSQL(mResponseCode)
            SaveLogValue("AuthQuery", mSalesOrderNo & ":(51) Exception encountered. " + ex.Message, sFilename, "info")
        End Try

    End Sub
    ''' <summary>
    ''' 依預約教室單號取得請款資料，丟至台新進行請款
    ''' </summary>
    ''' <param name="sClassroomPreorderNo">預約教室單號</param>
    ''' <remarks></remarks>
    Public Sub TSPG_SendCRCaptureWebRequest(ByVal sClassroomPreorderNo As String, ByVal mAmt As Decimal)
        Dim DebugData As String = String.Empty
        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_CRCapture.Log"
        Dim UpdateClassroomPreOrder As String = String.Empty
        Dim mResponseCode As String = String.Empty

        Dim TSPG_APIROOT As String = BwexGetBSNAM2("Store", "TSPG_APIROOT") '路徑
        Dim TSPG_mid As String = BwexGetBSNAM2("Store", "TSPG_mid_CR") '特店代號
        Dim TSPG_tid As String = BwexGetBSNAM2("Store", "TSPG_tid_CR") '端末代號
        Dim TSPG_post_back_url As String = BwexGetBSNAM2("Store", "TSPG_post_back_url")
        Dim TSPG_result_url As String = BwexGetBSNAM2("Store", "TSPG_tid")

        Dim TSPG_URL As String = "https://" & TSPG_APIROOT & "/other.ashx"
        'S180604019  升級TLS 1.2
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim GetResponseData As JObject
        Dim CaptureAmt As Decimal = 0
        Dim txnResponseCode As String = ""
        Dim msg As String = ""
        Dim auth_id_resp As String = ""
        Try

            Dim json As String = "{""sender"":""rest""," + """ver"":""1.0.0""," + """mid"":""" & TSPG_mid & """," + """tid"":""" & TSPG_tid & """," + """pay_type"":1," + """tx_type"":3," + """params"":" + "{ ""order_no"":""" & sClassroomPreorderNo & """,""amt"":""" & mAmt & """,""result_flag"":""1""}}"

            SaveLogValue("TSPG_SendCRCaptureWebRequest", sClassroomPreorderNo & "  " & json & " TSPG_APIROOT " & TSPG_APIROOT, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_SendCRCapture_JSON.Log")

            Dim webClient As System.Net.WebClient = New System.Net.WebClient()
            webClient.Headers.Add("Content-Type", "application/json")

            Dim reply As String = webClient.UploadString("https://" & TSPG_APIROOT & "/other.ashx", "POST", json)

            GetResponseData = CType(JsonConvert.DeserializeObject(reply), JObject)
            SaveLogValue("TSPG_SendCRCaptureWebRequest", sClassroomPreorderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_SendCRCaptureWebRequest_GetResponseData.Log")


            txnResponseCode = GetResponseData("params")("ret_code")

            If txnResponseCode = "00" Then
                msg = ReturnTSPG_OrderStatus(GetResponseData("params")("order_status"))
                CaptureAmt = Convert.ToInt32(GetResponseData("params")("settle_amt").ToString) / 100
                auth_id_resp = GetResponseData("params")("auth_id_resp")

                UpdateClassroomPreOrder = " update ClassroomPreOrder set CaptureAuthDate='{1}',CaptureAuthCode='{2}',UpdateUser='System',UpdateDatetime= dbo.udf_GetSysDatetime() where ClassroomPreOrderNo='{0}'"
                UpdateClassroomPreOrder = String.Format(UpdateClassroomPreOrder, sClassroomPreorderNo, UtilService.GetDateTime(Now, 3), GetResponseData("params")("settle_seq"))
                DataService.ExecuteSQL(UpdateClassroomPreOrder)

                ''S180326017 前台預約單成立、付款完成時，就產生正式訂單
                If BwexGetBSNAM("Store", "IsAutoGenClassSalesOrder", "Y") = "Y" Then
                    Dim ClassSalesOrderNo As String = UtilService.GetVal(DataService.ExecuteSQLScale(" SELECT ClassroomOrderNo FROM ClassroomPreOrder where ClassroomPreOrderNo = '" & sClassroomPreorderNo & "' AND IsFormOrder = 'Y' "), ValCollection.文字)
                    If ClassSalesOrderNo <> "" Then
                        Dim CaptureAuthCodeas As String = GetResponseData("params")("settle_seq")
                        Dim UpdateBillB As String = " UPDATE BillB SET CaptureAuthCode = '" & CaptureAuthCodeas & "' , CaptureAuthDate = '" & UtilService.GetDateTime(Now, 3) & "' WHERE BillNo = '" & ClassSalesOrderNo & "' "
                        DataService.ExecuteSQL(UpdateBillB)
                    End If
                End If

                Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & sClassroomPreorderNo & "'")
                mResponseCode = "exec usp_ResponseCode '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','CRCapture','19'"
                mResponseCode = String.Format(mResponseCode, mRegionNo, sClassroomPreorderNo, "", CaptureAmt, auth_id_resp, txnResponseCode, msg, "", msg)
                Wellan.Service.DataService.ExecuteSQL(mResponseCode)


                Dim LogData As DataSet
                Dim LogDataSQL As String = String.Empty
                '新增LogData
                LogDataSQL = "select * from ClassroomPreOrder where ClassroomPreOrderNo = '" & sClassroomPreorderNo & "'"
                LogData = Wellan.Service.DataService.CreateDataSet(LogDataSQL, "ClassroomPreOrder")
                WriteLogData(LogData, "ClassroomPreOrder", sClassroomPreorderNo, "Update")
            Else
                SaveLogValue("ClassRoom_Capture_Fail ", sClassroomPreorderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_CRCapture_Fail.Log", "info")

            End If

        Catch ex As Exception
            '訂單地區別
            Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & sClassroomPreorderNo & "'")
            mResponseCode = "exec usp_ResponseCode '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','CRCapture','19'"
            mResponseCode = String.Format(mResponseCode, mRegionNo, sClassroomPreorderNo, "", CaptureAmt, "", txnResponseCode, msg, "", ex.Message + ";" + msg)
            Wellan.Service.DataService.ExecuteSQL(mResponseCode)
            SaveLogValue("TSPG_CRCaptureError", sClassroomPreorderNo & ":(51) Exception encountered. " + ex.Message, sFilename, "info")
        End Try

    End Sub
    ''' <summary>
    ''' 依預約教室單號，丟至台新取得預約單授權資訊(是否授權成功)
    ''' </summary>
    ''' <param name="ClassroomPreOrderNo">預約單號</param>
    ''' <remarks></remarks>
    Public Sub TSPG_SendCRAuthQueryWebRequest(ByVal ClassroomPreOrderNo As String)

        Dim DebugData As String = String.Empty
        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ClassroomReserveAuthQuery.Log"
        Dim UpdateClassroomPreOrder As String = String.Empty
        Dim mResponseCode As String = String.Empty

        Dim TSPG_APIROOT As String = BwexGetBSNAM2("Store", "TSPG_APIROOT") '路徑
        Dim TSPG_mid As String = BwexGetBSNAM2("Store", "TSPG_mid_CR") '特店代號
        Dim TSPG_tid As String = BwexGetBSNAM2("Store", "TSPG_tid_CR") '端末代號
        Dim TSPG_post_back_url As String = BwexGetBSNAM2("Store", "TSPG_post_back_url")
        Dim TSPG_result_url As String = BwexGetBSNAM2("Store", "TSPG_tid")


        Dim TSPG_URL As String = "https://" & TSPG_APIROOT & "/other.ashx"
        'S180604019  升級TLS 1.2
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim GetResponseData As JObject
        Dim CaptureAmt As Decimal = 0
        Dim txnResponseCode As String = ""
        Dim msg As String = ""
        Dim ChkPreOrder As Boolean = False

        Dim MyConnection As SqlClient.SqlConnection = SqlConnection.GetConnection()
        Dim MyTrans As SqlClient.SqlTransaction

        If Not (MyConnection.State = ConnectionState.Open) Then
            MyConnection.Open()
        End If
        MyTrans = MyConnection.BeginTransaction()

        Try


            Dim json As String = "{""sender"":""rest""," + """ver"":""1.0.0""," + """mid"":""" & TSPG_mid & """," + """tid"":""" & TSPG_tid & """," + """pay_type"":1," + """tx_type"":7," + """params"":" + "{ ""order_no"":""" & ClassroomPreOrderNo & """,""result_flag"":""1""}}"

            SaveLogValue("TSPG_SendCRAuthQueryWebRequest", ClassroomPreOrderNo & "  " & json & " TSPG_APIROOT " & TSPG_APIROOT, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_SendCRAuthQueryWebRequest_JSON.Log")

            Dim webClient As System.Net.WebClient = New System.Net.WebClient()
            webClient.Headers.Add("Content-Type", "application/json")

            Dim reply As String = webClient.UploadString("https://" & TSPG_APIROOT & "/other.ashx", "POST", json)

            GetResponseData = CType(JsonConvert.DeserializeObject(reply), JObject)

            SaveLogValue("TSPG_SendCRAuthQueryWebRequest", ClassroomPreOrderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_SendCRAuthQueryWebRequest_GetResponseData.Log")

            txnResponseCode = GetResponseData("params")("ret_code")

            Dim auth_id_resp As String = "" '授權失敗 無回傳授傳碼
            Dim order_status As String = ""

            If txnResponseCode = "00" Then

                Dim tx_amt As Decimal = CType(GetResponseData("params")("tx_amt").ToString(), Decimal) / 100 '交易金額包含兩位小數，如100 代表1.00 元。
                Dim last_4_digit_of_pan As String = GetResponseData("params")("last_4_digit_of_pan").ToString() '卡號後4 碼
                order_status = GetResponseData("params")("order_status").ToString() '訂單狀態碼
                CaptureAmt = CType(GetResponseData("params")("tx_amt").ToString(), Decimal) / 100

                '02 已授權
                '03 已請款
                '04 請款已清算
                '06 已退貨
                '08 退貨已清算
                '12 訂單已取消
                'ZP 訂單處理中
                'ZF 授權失敗

                If order_status = "02" OrElse order_status = "03" Then
                    '付款成功
                    auth_id_resp = GetResponseData("params")("auth_id_resp").ToString() '授權碼

                    UpdateClassroomPreOrder = " UPDATE ClassroomPreOrder  "
                    UpdateClassroomPreOrder += " SET ReserveStatus = '1',PaymentStatus='0',PaidAmt='" & CaptureAmt & "', "
                    UpdateClassroomPreOrder += " CheckNo='" & last_4_digit_of_pan & "',CardValidYM='',AuthCode='" & auth_id_resp & "',CardHolder=N'',AuthDate='" & UtilService.GetDateTime(Now, 1) & "',StoreID='" & BwexGetBSNAM2("Store", "TSPG_mid_CR") & "',  "
                    UpdateClassroomPreOrder += " UpdateUser='System',UpdateDatetime= dbo.udf_GetSysDatetime()  "
                    UpdateClassroomPreOrder += " WHERE ClassroomPreOrderNo = '" & ClassroomPreOrderNo & "'"
                    DataService.ExecuteSQL(MyTrans, UpdateClassroomPreOrder)
                    ChkPreOrder = True

                ElseIf order_status = "ZF" Then
                    '付款失敗  PaymentStatus	1	刷卡失敗
                    'ReserveStatus	2	作廢
                    UpdateClassroomPreOrder = " UPDATE ClassroomPreOrder  "
                    UpdateClassroomPreOrder += " SET ReserveStatus = '2',PaymentStatus='1',  "
                    UpdateClassroomPreOrder += " UpdateUser='System',UpdateDatetime= dbo.udf_GetSysDatetime()  "
                    UpdateClassroomPreOrder += " WHERE ClassroomPreOrderNo = '" & ClassroomPreOrderNo & "'  "
                    DataService.ExecuteSQL(MyTrans, UpdateClassroomPreOrder)
                    '開放已鎖定教室預約
                    Dim UpdateClassIsReserve As String = " UPDATE ClassroomReserve SET IsReserve = 'N',UpdateUser='System',UpdateDatetime= dbo.udf_GetSysDatetime() WHERE ReserveNo = '" & ClassroomPreOrderNo & "' "
                    DataService.ExecuteSQL(MyTrans, UpdateClassIsReserve)
                    ChkPreOrder = False
                End If

                Try
                    Dim LogData As DataSet
                    Dim LogDataSQL As String = ""

                    '新增LogData
                    LogDataSQL = "select * from ClassroomPreOrder where ClassroomPreOrderNo = '" & ClassroomPreOrderNo & "'"
                    LogData = Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "ClassroomPreOrder")
                    WriteLogData(LogData, "ClassroomPreOrder", ClassroomPreOrderNo, "Update")

                    MyTrans.Commit()
                Catch ex As Exception
                    MyTrans.Rollback()
                Finally
                    MyTrans = Nothing
                    MyConnection.Close()
                End Try
            Else
                '付款失敗  PaymentStatus	1	刷卡失敗
                'ReserveStatus	2	作廢
                UpdateClassroomPreOrder = " UPDATE ClassroomPreOrder  "
                UpdateClassroomPreOrder += " SET ReserveStatus = '2',PaymentStatus='1',  "
                UpdateClassroomPreOrder += " UpdateUser='System',UpdateDatetime=dbo.udf_GetSysDatetime()  "
                UpdateClassroomPreOrder += " WHERE ClassroomPreOrderNo = '" & ClassroomPreOrderNo & "' "
                DataService.ExecuteSQL(MyTrans, UpdateClassroomPreOrder)

                Dim sReserveNo As String = UtilService.GetVal(DataService.ExecuteSQLScale("Select ReserveNo from ClassroomPreOrder where ClassroomPreorderNo='" & ClassroomPreOrderNo & "'"), ValCollection.文字)
                '開放已鎖定教室預約
                Dim UpdateClassIsReserve As String = " UPDATE ClassroomReserve SET IsReserve = 'N',UpdateUser='System',UpdateDatetime= dbo.udf_GetSysDatetime() WHERE ReserveNo = '" & sReserveNo & "' "
                DataService.ExecuteSQL(MyTrans, UpdateClassIsReserve)
                ChkPreOrder = False

                SaveLogValue("CR_Reserve_AuthQuery_Fail", ClassroomPreOrderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_CR_Reserve_AuthQuery_Fail.Log", "info")

                Try
                    Dim LogData As DataSet
                    Dim LogDataSQL As String = ""

                    '新增LogData
                    LogDataSQL = "select * from ClassroomPreOrder where ClassroomPreOrderNo = '" & ClassroomPreOrderNo & "'"
                    LogData = Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "ClassroomPreOrder")
                    LogDataSQL = "select * from ClassroomReserve where ReserveNo = '" & sReserveNo & "'"
                    LogData.Tables.Add(Wellan.Service.DataService.CreateDataSet(MyTrans, LogDataSQL, "ClassroomReserve").Tables(0).Copy)

                    WriteLogData(LogData, "ClassroomPreOrder", ClassroomPreOrderNo, "Update")

                    MyTrans.Commit()
                Catch ex As Exception
                    MyTrans.Rollback()
                Finally
                    MyTrans = Nothing
                    MyConnection.Close()
                End Try
            End If

            mResponseCode = "exec usp_ResponseCode '{0}',{1},'{2}','{3}','{4}','{5}','{6}','{7}','{8}','CReserveAuthQuery','19' "
            mResponseCode = String.Format(mResponseCode, "01", ClassroomPreOrderNo, "", CaptureAmt, auth_id_resp, txnResponseCode & order_status, msg, "", msg)
            Wellan.Service.DataService.ExecuteSQL(MyTrans, mResponseCode)
        Catch ex As Exception

            mResponseCode = "exec usp_ResponseCode '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','CReserveAuthQuery','19'"
            mResponseCode = String.Format(mResponseCode, "", ClassroomPreOrderNo, "", CaptureAmt, "", txnResponseCode, msg, "", ex.Message + ";" + msg)
            Wellan.Service.DataService.ExecuteSQL(MyTrans, mResponseCode)
            SaveLogValue("CRAuthQueryError", ClassroomPreOrderNo & ":(51) Exception encountered. " + ex.Message, sFilename, "info")
        End Try


        ''S180326017 前台預約單成立、付款完成時，就產生正式訂單
        If BwexGetBSNAM("Store", "IsAutoGenClassSalesOrder", "Y") = "Y" AndAlso ChkPreOrder = True Then
            AutoGenClassSalesOrder(ClassroomPreOrderNo)
        End If

    End Sub
    ''' <summary>
    ''' 依訂單編號，丟至台新執行取消授權
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub TSPG_Deauthorization(ByVal mOrderNo As String)
        Dim DebugData As String = String.Empty
        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_Deauthorization.Log"
        Dim mResponseCode As String = String.Empty

        'S180808021 現場自取使用的金流刷卡的商店代號皆要獨立
        Dim TSPG_mid As String = SalesOrder_GetTSPG_mid(mOrderNo) '特店代號
        Dim TSPG_tid As String = BwexGetBSNAM2("Store", "TSPG_tid") '端末代號

        Dim TSPG_APIROOT As String = BwexGetBSNAM2("Store", "TSPG_APIROOT") '路徑
        Dim TSPG_result_url As String = BwexGetBSNAM2("Store", "TSPG_tid")
        Dim TSPG_URL As String = "https://" & TSPG_APIROOT & "/other.ashx"

        'S180604019  升級TLS 1.2
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim GetResponseData As JObject
        Dim CaptureAmt As Decimal = 0
        Dim txnResponseCode As String = ""
        Dim msg As String = ""

        Try

            '特店欲對已授權但尚未請款的訂單取消授權  
            'tx_type 8 取消授權
            Dim json As String = "{""sender"":""rest""," + """ver"":""1.0.0""," + """mid"":""" & TSPG_mid & """," + """tid"":""" & TSPG_tid & """," + """pay_type"":1," + """tx_type"":8," + """params"":" + "{ ""order_no"":""" & mOrderNo & """,""result_flag"":""1""}}"

            SaveLogValue("Deauthorization", mOrderNo & "  " & json & " TSPG_APIROOT " & TSPG_APIROOT, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_Deauthorization_JSON.Log")

            Dim webClient As System.Net.WebClient = New System.Net.WebClient()
            webClient.Headers.Add("Content-Type", "application/json")

            Dim reply As String = webClient.UploadString("https://" & TSPG_APIROOT & "/other.ashx", "POST", json)

            GetResponseData = CType(JsonConvert.DeserializeObject(reply), JObject)
            SaveLogValue("Deauthorization", mOrderNo & "  " & reply, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_Deauthorization_GetResponseData.Log")

            txnResponseCode = GetResponseData("params")("ret_code")

            Dim order_status As String = ""
            Dim auth_id_resp As String = "" '授權失敗 無回傳授傳碼

            If txnResponseCode = "00" Then
                order_status = GetResponseData("params")("order_status").ToString() '訂單狀態碼
                msg = GetResponseData("params")("order_status").ToString() & ReturnTSPG_OrderStatus(GetResponseData("params")("order_status"))
                '02 已授權
                '03 已請款
                '04 請款已清算
                '06 已退貨
                '08 退貨已清算
                '12 訂單已取消
                'ZP 訂單處理中
                'ZF 授權失敗
                If order_status <> "12" Then
                    SaveLogValue("Deauthorization OrderNo:", mOrderNo & " ,result:" & msg, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_Deauthorization_Fail.Log", "info")
                End If
            Else 'ret_code <> 00 亦為交易失敗
                SaveLogValue("Deauthorization OrderNo:", mOrderNo & ",result:" & txnResponseCode & msg, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_Deauthorization_Fail.Log", "info")
            End If

            Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & mOrderNo & "'")
            mResponseCode = "exec usp_ResponseCode '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','Deauthorization','19'"
            mResponseCode = String.Format(mResponseCode, mRegionNo, mOrderNo, "", CaptureAmt, auth_id_resp, txnResponseCode & order_status, msg, GetResponseData("params")("rrn"), msg)
            Wellan.Service.DataService.ExecuteSQL(mResponseCode)

        Catch ex As Exception
            '訂單地區別
            Dim mRegionNo As String = DataService.ExecuteSQLScale("select o.regionNo from salesordera a left join Organizationobject o on a.Orgno = o.orgno where SalesOrderNo = '" & mOrderNo & "'")
            mResponseCode = "exec usp_ResponseCode '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','Deauthorization','19'"

            mResponseCode = String.Format(mResponseCode, mRegionNo, mOrderNo, "", CaptureAmt, "", txnResponseCode, msg, "", ex.Message + ";" + msg)
            Wellan.Service.DataService.ExecuteSQL(mResponseCode)
            SaveLogValue("DeauthorizationError", mOrderNo & ":(51) Exception encountered. " + ex.Message, sFilename, "info")
        End Try

    End Sub

    Private Function ReturnTSPG_OrderStatus(ByVal mStatus As String) As String

        Dim StatusDesc As String = ""

        Select Case mStatus

            Case "02"
                StatusDesc = "已授權"
            Case "03"
                StatusDesc = "已請款"
            Case "04"
                StatusDesc = "請款已清算"
            Case "06"
                StatusDesc = "已退貨"
            Case "08"
                StatusDesc = "退貨已清算"
            Case "12"
                StatusDesc = "訂單已取消"
            Case "ZP"
                StatusDesc = "訂單處理中"
            Case "ZF"
                StatusDesc = "授權失敗"
            Case Else
                StatusDesc = ""
        End Select

        Return StatusDesc
    End Function

    Public Function AutoGenClassSalesOrder(ByVal mClassroomPreOrderNo As String)

        Dim Result As Boolean = False

        '從教室預約預訂單 抓取 刷卡相關資料
        Dim GetCardInfoSQL As String = "  SELECT *  FROM ClassroomPreOrder Where ClassroomPreOrderNo  = '" & mClassroomPreOrderNo & "' "
        Dim dss As DataSet = DataService.CreateDataSet(GetCardInfoSQL)

        Dim mMemberID As String = ""

        Dim mCheckNo As String = ""
        Dim mCardValidYM As String = ""
        Dim mAuthCode As String = ""
        Dim mCardHolder As String = ""
        Dim mAuthDate As Date
        Dim mStoreID As String = ""
        Dim mClassroomOrgNo As String = ""
        Dim mPaymentType As String = ""

        Dim mTotalAmt As Decimal = 0

        Dim mIsFormOrder As String = ""
        Dim mClassroomOrderNo As String = ""

        If dss.Tables(0).Rows.Count > 0 Then
            Dim dr As DataRow = dss.Tables(0).Rows(0)
            mCheckNo = UtilService.GetVal(dr("CheckNo").ToString, ValCollection.文字)
            mCardValidYM = UtilService.GetVal(dr("CardValidYM").ToString, ValCollection.文字)
            mAuthCode = UtilService.GetVal(dr("AuthCode").ToString, ValCollection.文字)
            mCardHolder = UtilService.GetVal(dr("CardHolder").ToString, ValCollection.文字)
            mAuthDate = UtilService.GetVal(dr("AuthDate"), ValCollection.文字)
            mStoreID = UtilService.GetVal(dr("StoreID").ToString, ValCollection.文字)
            mPaymentType = UtilService.GetVal(dr("PaymentType").ToString, ValCollection.文字)
            mClassroomOrgNo = UtilService.GetVal(dr("OrgNo").ToString, ValCollection.文字)
            mMemberID = UtilService.GetVal(dr("MemberID").ToString, ValCollection.文字)
            mTotalAmt = UtilService.GetVal(dr("TotalAmt").ToString, ValCollection.數值)

            '抓是否有產生過正式訂單
            mIsFormOrder = UtilService.GetVal(dr("IsFormOrder").ToString, ValCollection.文字)
            mClassroomOrderNo = UtilService.GetVal(dr("ClassroomOrderNo").ToString, ValCollection.文字)

            If mIsFormOrder = "Y" AndAlso mClassroomOrderNo <> "" Then '再度判斷是否有產生過正式訂單

                SaveLogValue("AutoGenClassSalesOrder_Error", " AutoGenClassSalesOrder :  已產生過正式訂單 ClassroomPreOrderNo :" & mClassroomPreOrderNo, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_AutoGenClassSalesOrder_Error.Log", "info")
            Else
                'S191114002 付款成功產生訂單時，目前訂單的商品編號為T019，改為參數控管Store	ReserveProductNo	T020
                '商品編號：T019場地收入
                Dim mProductNo As String = BwexGetBSNAM("Store", "ReserveProductNo", "T020")

                'OrderDate 申請日期
                Dim mSalesDate As String = UtilService.GetVal(dr("OrderDate").ToString, ValCollection.文字)

                '轉正式訂單 
                Dim OrgNo As String = IIf(BwexGetBSNAM("Store", "OrgNo") = "", "G", BwexGetBSNAM("Store", "OrgNo")) '依地區別取分公司

                Dim mBonusDate As String = UtilService.GetVal(Wellan.Service.DataService.ExecuteSQLScale(" select BonusDate = Convert(char(10), BonusDate, 111) from BonusDate where BonusType='B' and Locked = 'F' and '" + CDate(mSalesDate).ToString("yyyy/MM/dd ") + "' between Convert(char(10), BonusBeginDate,111) and Convert(char(10), BonusEndDate,111)"), Wellan.Service.ValCollection.文字)

                Dim SelSQL As String = " select PV=isnull(c.PV,0) from Product p"
                SelSQL += " left join udf_GetChangePrice('01','" & CDate(mSalesDate) & "') c on p.ProductNo=c.ProductNo "
                SelSQL &= " where p.ProductNo = '" & mProductNo & "' "
                Dim ProductPV As Decimal = UtilService.GetVal(DataService.ExecuteSQLScale(SelSQL), ValCollection.數值)

                '稅率參考設定檔
                Dim hTaxRate As String = Wellan.Service.UtilService.GetVal(Wellan.Service.DataService.ExecuteSQLScale("select TaxRate = TaxRate / 100 from RegionCurrency where RegionNo = '01' "), Wellan.Service.ValCollection.數值)
                If Trim(hTaxRate) = "" Then hTaxRate = 0
                Dim mTaxRate = hTaxRate * 100
                Dim mTax As Decimal = mTotalAmt - Math.Round(mTotalAmt / (1 + hTaxRate), 0)
                Dim mSubAmtA As Decimal = Math.Round(mTotalAmt / (1 + hTaxRate), 0)

                '訂購批號
                Dim sType As String = "EC"

                'S200819006 改善商業邏輯於windows form取得物件異常狀況
                Dim IsMemberCarrier As String = ""
                Dim BusinessFlag As Boolean = True
                Try
                    Dim objRVSAService As PP00SS.Business.RVSAService = New PP00SS.Business.RVSAService
                    IsMemberCarrier = objRVSAService.IsMemberCarrier()
                Catch ex As Exception
                    BusinessFlag = False
                    ErrorMessage &= "<BR>" & "執行中發生錯誤，請洽資訊人員!  " & Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "<BR>" & "RVSAService.IsMemberCarrier_Error:" & ex.Message
                    SaveLogValue("AutoGenClassSalesOrder_Error", " RVSAService.IsMemberCarrier_Error: ClassroomPreOrderNo :" & mClassroomPreOrderNo & ",ErrorMsg:" & ex.Message, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_AutoGenClassSalesOrder_Error.Log", "info")
                End Try

                Dim mMemberType As String = UtilService.GetVal(DataService.ExecuteSQLScale(" SELECT MemberType FROM Member WHERE MemberID = '" & mMemberID & "' "), ValCollection.文字)

                Dim mLotNo As String = CDate(mSalesDate).ToString("yyyyMMdd") & Format(Hour(SetWorkDate("01", 3)), "00") & sType

                If BusinessFlag = True Then 'S200819006 確認商業邏輯運作正常才產生正式訂單
                    Dim objSalesService As PP00SS.Business.SalesService = New PP00SS.Business.SalesService
                    Dim MyConnection As SqlClient.SqlConnection = SqlConnection.GetConnection()
                    Dim MyTrans As SqlClient.SqlTransaction

                    If Not (MyConnection.State = ConnectionState.Open) Then
                        MyConnection.Open()
                    End If
                    MyTrans = MyConnection.BeginTransaction()

                    Try

                        objSalesService.Transaction = MyTrans
                        Dim mSalesOrderNo As String = objSalesService.GetSalesOrderNo(OrgNo, "SA", mSalesDate)

                        '產生訂單主檔 SalesOrderA
                        Dim cm As New SqlClient.SqlCommand
                        With cm
                            .Connection = MyConnection
                            .Transaction = MyTrans
                            .CommandType = CommandType.StoredProcedure

                            Dim sUSPname = "usp_AddOrder"
                            sUSPname = GetProjectUSP(System.Configuration.ConfigurationManager.AppSettings("BSPRJ"), sUSPname)
                            .CommandText = sUSPname

                            '分公司為"線上購物"、出貨分公司為”教室所屬的分公司”
                            .Parameters.AddWithValue("@SalesOrderNo", mSalesOrderNo)
                            .Parameters.AddWithValue("@MemberID", mMemberID)
                            '分公司為"線上購物"
                            .Parameters.AddWithValue("@OrgNo", OrgNo)
                            .Parameters.AddWithValue("@SalesDate", CDate(mSalesDate))
                            .Parameters.AddWithValue("@BonusDate", CDate(mBonusDate))
                            'Shipment	1	自取
                            .Parameters.AddWithValue("@Shipment", "1")
                            'ShipStatus	2	完全出貨
                            .Parameters.AddWithValue("@ShipStatus", "2")
                            'ShipFlag	Y	立即出貨
                            .Parameters.AddWithValue("@ShipFlag", "Y")
                            .Parameters.AddWithValue("@Paid", "Y")
                            '出貨分公司為”教室所屬的分公司”
                            .Parameters.AddWithValue("@ShipOrgNo", mClassroomOrgNo)
                            .Parameters.AddWithValue("@RACS2", 0)
                            .Parameters.AddWithValue("@LotNo", mLotNo)
                            'SalesType	 SalesType	CLASS	教室
                            .Parameters.AddWithValue("@SalesType", "CLASS")
                            .Parameters.AddWithValue("@TotalPV", ProductPV)

                            .Parameters.AddWithValue("@TotalAmt", mTotalAmt)
                            'SubAmtA + Tax = TotalAmt
                            .Parameters.AddWithValue("@Tax", mTax)
                            .Parameters.AddWithValue("@TaxRate", mTaxRate)
                            .Parameters.AddWithValue("@SubAmtA", mSubAmtA)
                            .Parameters.AddWithValue("@PaidAmt", mTotalAmt) '已付款
                            .Parameters.AddWithValue("@OrderStatus", "1") '訂單類別
                            .Parameters.AddWithValue("@OrdMethod", BwexGetBSNAM("Store", "OrdMethod")) '訂貨方式'S131225005  OrdMethod改抓參數設定
                            .Parameters.AddWithValue("@AddDate", GetSysDatetime.ToString("yyyy/MM/dd HH:mm:ss"))
                            .Parameters.AddWithValue("@AddUser", "EC")

                            '是否產生出貨單 PickingStatus	2	全部產生
                            .Parameters.AddWithValue("@PickingStatus", 2)
                            '揀貨狀態 IsPicking	2	完全揀貨
                            .Parameters.AddWithValue("@IsPicking", 2)

                            .Parameters.AddWithValue("@Receiver", mMemberID)
                            .Parameters.AddWithValue("@IsSendInvoice", "N")

                            .Parameters.AddWithValue("@CouponAmt", 0)
                            .Parameters.AddWithValue("@DueAmt", 0)
                            .Parameters.AddWithValue("@CurCountry", "TW")
                            .Parameters.AddWithValue("@InvCountry", "TW")
                            .Parameters.AddWithValue("@ShipDate", "4")

                            .Parameters.AddWithValue("@CurAddress", Convert.DBNull)
                            .Parameters.AddWithValue("@CurAmt", Convert.DBNull)
                            .Parameters.AddWithValue("@CurCity", Convert.DBNull)
                            .Parameters.AddWithValue("@CurNeighbor", Convert.DBNull)
                            .Parameters.AddWithValue("@CurRoad", Convert.DBNull)
                            .Parameters.AddWithValue("@CurTown", Convert.DBNull)
                            .Parameters.AddWithValue("@CurZipCode", Convert.DBNull)
                            .Parameters.AddWithValue("@InvAddress", Convert.DBNull)
                            .Parameters.AddWithValue("@InvCity", Convert.DBNull)
                            .Parameters.AddWithValue("@InvMemo", Convert.DBNull)
                            .Parameters.AddWithValue("@InvoiceNo", Convert.DBNull)
                            .Parameters.AddWithValue("@InvReceiver", Convert.DBNull)
                            .Parameters.AddWithValue("@InvReceiverTel", Convert.DBNull)
                            .Parameters.AddWithValue("@InvTown", Convert.DBNull)
                            .Parameters.AddWithValue("@InvZipCode", Convert.DBNull)
                            .Parameters.AddWithValue("@Level1", Convert.DBNull)
                            .Parameters.AddWithValue("@Level2", Convert.DBNull)
                            .Parameters.AddWithValue("@Level3", Convert.DBNull)
                            .Parameters.AddWithValue("@Memo", Convert.DBNull)
                            .Parameters.AddWithValue("@NPOBAN", Convert.DBNull)
                            .Parameters.AddWithValue("@OrgMemberID", Convert.DBNull)
                            .Parameters.AddWithValue("@OverPaidAmt", Convert.DBNull)
                            .Parameters.AddWithValue("@PayDate", Convert.DBNull)
                            .Parameters.AddWithValue("@ReceiverMobile", Convert.DBNull)
                            .Parameters.AddWithValue("@ReceiverTel", Convert.DBNull)
                            .Parameters.AddWithValue("@ReceiverTel2", Convert.DBNull)
                            .Parameters.AddWithValue("@UpdateDatetime", Convert.DBNull)
                            .Parameters.AddWithValue("@UpdateUser", Convert.DBNull)
                            .Parameters.AddWithValue("@Package", "N")

                            'S191106038 教室預約轉正式訂單後所開立的發票，除了法人之外，一律都將發票存入"會員載具"=>"發票載具=葡眾會員載具"&"載具號碼:(訂購人會員編號) "
                            'S200215006 "發票寄送方式=電子發票列印紙本"時，需存入後台的"發票型式=索取紙本"
                            ' "法人" 預約教室功能 教室預約單轉正式訂單，電子發票也一同存入
                            If IsMemberCarrier = "B" AndAlso mMemberType <> "2" Then
                                .Parameters.AddWithValue("@CarrierID1", mMemberID)
                                .Parameters.AddWithValue("@CarrierID2", mMemberID)
                                .Parameters.AddWithValue("@CarrierType", BwexGetBSNAM("Store", "MemberEInvoiceType"))
                                .Parameters.AddWithValue("@InvoiceTake", 0)
                            Else
                                .Parameters.AddWithValue("@CarrierID1", Convert.DBNull)
                                .Parameters.AddWithValue("@CarrierID2", Convert.DBNull)
                                .Parameters.AddWithValue("@CarrierType", Convert.DBNull)
                                .Parameters.AddWithValue("@InvoiceTake", 1)
                            End If

                            .ExecuteNonQuery()
                        End With

                        '新增訂單明細 SaleserOrderB
                        Dim cmB As New SqlClient.SqlCommand
                        With cmB
                            .Connection = MyConnection
                            .Transaction = MyTrans
                            .CommandType = CommandType.StoredProcedure

                            Dim sUSPname = "usp_AddOrderDetail"
                            sUSPname = GetProjectUSP(System.Configuration.ConfigurationManager.AppSettings("BSPRJ"), sUSPname)
                            .CommandText = sUSPname
                            .Parameters.AddWithValue("@UnitPrice", mTotalAmt)
                            .Parameters.AddWithValue("@UnitPV", ProductPV)
                            .Parameters.AddWithValue("@SubPV", ProductPV)
                            .Parameters.AddWithValue("@SubAmt", mTotalAmt)
                            .Parameters.AddWithValue("@ItemType", 1)
                            .Parameters.AddWithValue("@SalesOrderNo", mSalesOrderNo)
                            .Parameters.AddWithValue("@ItemNo", 1)
                            .Parameters.AddWithValue("@ProductSKU", mProductNo)
                            .Parameters.AddWithValue("@Qty", 1)
                            .Parameters.AddWithValue("@ShipedQty", 1)
                            .Parameters.AddWithValue("@UpdateDatetime", GetSysDatetime.ToString("yyyy/MM/dd HH:mm:ss"))
                            .Parameters.AddWithValue("@UpdateUser", "EC")
                            .Parameters.AddWithValue("@discount", 100)
                            .Parameters.AddWithValue("@discount2", 0)

                            .Parameters.AddWithValue("@SalesPrice", Convert.DBNull)
                            .Parameters.AddWithValue("@FromProductSet", Convert.DBNull)
                            .Parameters.AddWithValue("@ReturnQty", Convert.DBNull)

                            'S170824042  若ItemType=贈品，寫入促銷編號 & 項次
                            .Parameters.AddWithValue("@RBRBN", "")
                            .Parameters.AddWithValue("@RBBTM", "")

                            .ExecuteNonQuery()
                        End With

                        '更新SalesOrderB UnitPrice2  SubAmt2
                        Dim UpdateSalesOrderB As String = " UPDATE SalesOrderB SET UnitPrice2='" & mSubAmtA & "',SubAmt2='" & mSubAmtA & "' WHERE SalesOrderNo = '" & mSalesOrderNo & "' "
                        DataService.ExecuteSQL(MyTrans, UpdateSalesOrderB)


                        '新增付款明細 BillB
                        Dim InsertBillBSQL As String = " INSERT INTO BillB (BillNo,ItemNo,Amt,PaymentType,BankNo,CheckNo,AuthCode,CardHolder,CardType,ExpirationDateY,ExpirationDateM,InvoiceNo,AuthDate,ReceiveDate,ReceiveFlag,PaidAmt,DiscountAmt,MemberID2,StoreID,CloseDate,ItemNoAR,PayOrgNo,Kind,CardValidYM,AuthorizedCode,IsEncrypted,VerCode,AmtDC,CaptureAuthCode,CaptureAuthDate,PaymentDate,DividePayCnt,Memo)VALUES"
                        InsertBillBSQL += "(N'" & mSalesOrderNo & "',N'1',N'" & mTotalAmt & "',N'" & mPaymentType & " ',NULL,N'" & mCheckNo & "',N'" & mAuthCode & "',N'" & mCardHolder & "',NULL,NULL,NULL,NULL,N'" & mAuthDate & "',NULL,NULL,NULL,NULL,NULL,N'" & mStoreID & "',NULL,NULL,N'" & OrgNo & "',N'1',N'" & mCardValidYM & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)  "
                        DataService.ExecuteSQL(MyTrans, InsertBillBSQL)

                        '回寫教室預約預定單  正式訂購單號  是否轉成正式訂單
                        Dim UpdateClassroomPreOrder As String = " UPDATE ClassroomPreOrder  "
                        UpdateClassroomPreOrder += " SET ClassroomOrderNo = '" & mSalesOrderNo & "',IsFormOrder='Y',  "
                        UpdateClassroomPreOrder += " UpdateUser='EC',UpdateDatetime=N'" & GetSysDatetime.ToString("yyyy/MM/dd HH:mm:ss") & "'  "
                        UpdateClassroomPreOrder += " WHERE ClassroomPreOrderNo = '" & mClassroomPreOrderNo & "' "
                        DataService.ExecuteSQL(MyTrans, UpdateClassroomPreOrder)

                        MyTrans.Commit()
                    Catch ex As Exception
                        MyTrans.Rollback()
                        Dim objErrorLog As PP00SS.Business.ErrorLog = New PP00SS.Business.ErrorLog
                        objErrorLog.WriteErrorLog(" AutoGenClassSalesOrder : FAIL ", "", ex, "ClassroomPreOrderNo: " & mClassroomPreOrderNo)
                        objErrorLog = Nothing
                        SaveLogValue("AutoGenClassSalesOrder_Error", " AutoGenClassSalesOrder : error: " + ex.Message.ToString() & "  ", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_AutoGenClassSalesOrder_Error.Log", "info")
                    Finally
                        MyTrans = Nothing
                        MyConnection.Close()
                    End Try
                End If




            End If



        Else
            SaveLogValue("AutoGenClassSalesOrder_Error", " AutoGenClassSalesOrder :  Not ClassroomPreOrderInfo ", g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_AutoGenClassSalesOrder_Error.Log","info")
        End If


        Return Result
    End Function

#Region "GetProjectUSP"
    Public Function GetProjectUSP(ByVal pProjectName As String, ByVal pUdfName As String) As String
        Dim sUdfName As String
        Dim sProjectUdfName As String
        Dim sSQL As String
        Dim sRetunUdfName As String
        sUdfName = pUdfName
        sProjectUdfName = pUdfName + "_" + pProjectName
        sSQL = "select Count(*) udfCount from sys.objects where object_id = OBJECT_ID(N'[dbo].[{0}]') and   type in (N'P', N'PC')"
        sSQL = String.Format(sSQL, sProjectUdfName)
        If DataService.WLookupF(sSQL, "udfCount") <> 0 Then
            sRetunUdfName = sProjectUdfName
        Else
            sRetunUdfName = sUdfName
        End If

        GetProjectUSP = sRetunUdfName
    End Function
#End Region

    ''' <summary>
    ''' 依地區別判斷時區
    ''' </summary>
    ''' <param name="pRegionNo">地區別</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTimeZoneId(ByVal pRegionNo As String) As String
        Dim sSQL As String = "Select TimeZoneId from RegionCurrency where RegionNo = '" & pRegionNo & "'"
        Dim sTimeZoneId As String = UtilService.GetVal(DataService.ExecuteSQLScale(sSQL), ValCollection.文字)
        If sTimeZoneId = "" Then
            sTimeZoneId = "Taipei Standard Time"
        End If
        Return sTimeZoneId
    End Function

    ''' <summary>
    ''' 設定作業日期
    ''' </summary>
    ''' <remarks></remarks>
    Public Function SetWorkDate(ByVal pRegionNo As String, Optional ByVal Kind As Integer = 1) As String
        Dim gWorkDate As String = ""
        'Dim gBaseUtcOffset As String = GetBaseUtcOffset(pRegionNo) '依地區別取得UTC Time與Local Time的時差
        If System.Configuration.ConfigurationManager.AppSettings("IsUseUTC") = "Y" Then
            gWorkDate = UtilService.GetDateTime(GetTimeZoneInfo(DateTime.UtcNow, pRegionNo), Kind)
        Else
            gWorkDate = UtilService.GetDateTime(Now, Kind)
        End If
        Return gWorkDate
    End Function

    Public Function GetTimeZoneInfo(ByVal pDate As DateTime, ByVal pRegionNo As String) As DateTime
        '格林威治時間
        If System.Configuration.ConfigurationManager.AppSettings("IsUseUTC") = "Y" Then
            Dim sTimeZoneId As String = GetTimeZoneId(pRegionNo)
            Dim hwZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(sTimeZoneId)
            Return TimeZoneInfo.ConvertTimeFromUtc(pDate, hwZone)
        Else
            Return pDate
        End If
    End Function

#Region "S180808021 現場自取使用的金流刷卡的商店代號皆要獨立"
    Public Function SalesOrder_GetTSPG_mid(ByVal SalesOrderNo As String)

        Dim ReStoreID As String = BwexGetBSNAM2("Store", "TSPG_mid") '分公司線上購物 特店代號

        If BwexGetBSNAM("Store", "IsOpenAPP2", "Y") = "Y" Then

            Dim ds As DataSet = DataService.CreateDataSet("SELECT OrgNo,SalesType  FROM SalesOrderA WHERE SalesOrderNo = '" & SalesOrderNo & "' ")

            Dim OrgNo As String = UtilService.GetVal(ds.Tables(0).Rows(0).Item("OrgNo"), ValCollection.文字)
            Dim SalesType As String = UtilService.GetVal(ds.Tables(0).Rows(0).Item("SalesType"), ValCollection.文字)

            If SalesType = "APP2" Then
                ReStoreID = BwexGetBSNAM("Store", "TSPG_mid_APP2_" & OrgNo, "") '抓取該訂單分公司之特店代號
            End If
        End If

        Return ReStoreID

    End Function

#End Region

#Region " Mail "
    Private Sub SendingMail(ByVal sMailBody As String, ByVal Kind As String, ByVal sMailBodyToStaff As String)
        Dim sFilename As String = g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_" & Kind & ".Log"

        Dim fromAddress As String = BwexGetBSNAM("Store", "ServiceMail") ' ServiceMail	inform@pro-partner.com.tw	客服Mail
        Dim toAddress As String = BwexGetBSNAM("Store", "SysAdminMail") '偉盟內部，收件者Mail
        Dim fromName As String = BwexGetBSNAM("Store", "StoreName")


        Try
            If toAddress <> "" Then

                Dim Attachments As String = "" 'S161017029 夾檔
                Dim KindName As String = ""
                Select Case Kind
                    Case "Capture"
                        KindName = "訂單請款"
                    Case "ClassroomReserveCapture" '
                        KindName = "教室預約單請款"
                    Case "AuthorizationQuery"
                        KindName = "查詢授權-查" & BwexGetBSNAM("Store", "AuthQueryChkMins") & "分鐘以前的訂單"
                    Case "AuthorizationQuery-Y"
                        KindName = "查詢授權-昨天訂單"
                    Case "ClassroomReserveAuthQuery"
                        KindName = "查詢授權-查" & BwexGetBSNAM("Store", "ClassroomAuthQueryChkMins") & "分鐘以前教室預約單"
                    Case "ClassroomReserveAuthQuery-Y"  '
                        KindName = "查詢授權-昨天教室預約單"
                    Case "Deauthorization"
                        KindName = "訂單取消授權"
                End Select

                Dim Subject As String = Now.ToString("yyyy/MM/dd") & " 葡眾台新金流排程_" & KindName & " 執行報告 " & System.Configuration.ConfigurationManager.AppSettings("ProjectID")

                Dim xmlNo As String = "SY-" & Guid.NewGuid.ToString '組成mail字串，寄回User
                GetMailXML(xmlNo, fromName, fromAddress, toAddress, toAddress, "", "", "", "", Attachments, Subject, sMailBody.ToString) 'S161017029 夾檔

                'S170222003 
                'Dim objMessageLog As PP00SS.Business.MessageLog = New PP00SS.Business.MessageLog
                'objMessageLog.WriteMessageLog_EMail("PPOpenIE_" & Kind, "", toAddress, sMailBody, Subject)
                'objMessageLog = Nothing

                Dim objMessageLog As Wellan.Ecommerce.StoreFront.MessageLog = New Wellan.Ecommerce.StoreFront.MessageLog
                objMessageLog.WriteMessageLog_EMail("PPOpenIE_" & Kind, "", toAddress, sMailBody, Subject, "1")
                objMessageLog = Nothing

                'S200224022 每天執行的台新請款、解除授權排程，都執行完成結果一並發信給葡眾內部人員
                'S220419002 新增台新教室預約請款排程
                Dim toStaffMailAddress As String = BwexGetBSNAM("Store", "StaffMail", "") '葡眾內部，收件者Mail

                If toStaffMailAddress <> "" AndAlso (Kind = "Deauthorization" OrElse Kind = "Capture" OrElse Kind = "ClassroomReserveCapture") Then

                    Dim SubjectToStaff As String = Now.ToString("yyyy/MM/dd") & " 葡眾台新金流排程_" & KindName

                    Dim sxmlNo As String = "SS-" & Guid.NewGuid.ToString '組成mail字串，寄回User
                    GetMailXML(sxmlNo, fromName, fromAddress, toStaffMailAddress, toStaffMailAddress, "", "", "", "", Attachments, SubjectToStaff, sMailBodyToStaff.ToString) 'S161017029 夾檔

                    'S170222003 
                    'Dim sobjMessageLog As PP00SS.Business.MessageLog = New PP00SS.Business.MessageLog
                    'sobjMessageLog.WriteMessageLog_EMail("PPOpenIE_" & Kind, "", toStaffMailAddress, sMailBodyToStaff, SubjectToStaff)
                    'sobjMessageLog = Nothing

                    Dim sobjMessageLog As Wellan.Ecommerce.StoreFront.MessageLog = New Wellan.Ecommerce.StoreFront.MessageLog
                    sobjMessageLog.WriteMessageLog_EMail("PPOpenIE_" & Kind, "", toStaffMailAddress, sMailBodyToStaff, SubjectToStaff, "1")
                    sobjMessageLog = Nothing

                End If


            End If
        Catch ex As Exception
            SaveLogValue("TSPG_SendMailDetailError", ex.Message, g_backupPath & "\" & Now.ToString("yyyyMMdd") & "_TSPG_ExecRecord.Log", "info")
        End Try

    End Sub

    ''' <summary>
    ''' 產生XML
    ''' </summary>
    ''' <param name="xmlNo">檔名</param>
    ''' <param name="FromName">寄件人姓名</param>
    ''' <param name="FromAddress">寄件人Mail</param>
    ''' <param name="ToName">收件人姓名</param>
    ''' <param name="ToAddress">收件人Mail</param>
    ''' <param name="CcName">副本姓名</param>
    ''' <param name="CcAddress">副本Mail</param>
    ''' <param name="BccName">密件姓名</param>
    ''' <param name="BccAddress">密件Mail</param>
    ''' <param name="Attachments">附件</param>
    ''' <param name="Subject">主旨</param>
    ''' <param name="Body">內文</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetMailXML(ByVal xmlNo As String, ByVal FromName As String, ByVal FromAddress As String, ByVal ToName As String, ByVal ToAddress As String, ByVal CcName As String, ByVal CcAddress As String, ByVal BccName As String, ByVal BccAddress As String, ByVal Attachments As String, ByVal Subject As String, ByVal Body As String) As String
        '檔名 寄件人姓名 寄件人Mail 收件人姓名 收件人Mail 副本姓名 副本Mail 密件姓名 密件Mail 附件 主旨 內文
        Dim Doc As XmlDocument = New XmlDocument
        Dim newEle As XmlElement
        Dim FName As String = BwexGetBSNAM("Store", "MailPath3")  'Mail 寄送路徑

        Dim DocRoot As XmlElement = Doc.CreateElement("EMail")
        Doc.AppendChild(DocRoot)
        '寄件人
        newEle = Doc.CreateElement("FromName")
        newEle.InnerText = FromName
        DocRoot.AppendChild(newEle)
        newEle = Doc.CreateElement("FromAddress")
        newEle.InnerText = FromAddress
        DocRoot.AppendChild(newEle)
        '收件人
        newEle = Doc.CreateElement("ToName")
        newEle.InnerText = ToName
        DocRoot.AppendChild(newEle)
        newEle = Doc.CreateElement("ToAddress")
        newEle.InnerText = ToAddress
        DocRoot.AppendChild(newEle)
        '副本
        newEle = Doc.CreateElement("CcName")
        newEle.InnerText = CcName
        DocRoot.AppendChild(newEle)
        newEle = Doc.CreateElement("CcAddress")
        newEle.InnerText = CcAddress
        DocRoot.AppendChild(newEle)
        '密件副本
        newEle = Doc.CreateElement("BccName")
        newEle.InnerText = BccName
        DocRoot.AppendChild(newEle)
        newEle = Doc.CreateElement("BccAddress")
        newEle.InnerText = BccAddress
        DocRoot.AppendChild(newEle)
        '附件
        newEle = Doc.CreateElement("Attachments")
        newEle.InnerText = Attachments
        DocRoot.AppendChild(newEle)
        '主旨
        newEle = Doc.CreateElement("Subject")
        newEle.InnerText = Subject
        DocRoot.AppendChild(newEle)
        '內文
        newEle = Doc.CreateElement("Body")
        newEle.InnerText = Body
        DocRoot.AppendChild(newEle)
        '產生時間
        newEle = Doc.CreateElement("GenDateTime")
        newEle.InnerText = Now
        DocRoot.AppendChild(newEle)
        'FileName
        FName += "\" + xmlNo + Now.ToString("yyyyMMddHHmmss") + ".xml"

        Doc.Save(FName)
        Return "Y"
    End Function
#End Region
    '台新Header改善預放
    'Public Sub SetHeaderValue(ByVal header As WebHeaderCollection, ByVal name As String, ByVal value As String)
    '    Dim [property] = GetType(WebHeaderCollection).GetProperty("InnerCollection", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)

    '    If [property] IsNot Nothing Then
    '        Dim collection = TryCast([property].GetValue(header, Nothing), NameValueCollection)
    '        collection(name) = value
    '    End If
    'End Sub

End Module

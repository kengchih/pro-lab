Imports System.IO
Imports System.Text
Imports System.Web.HttpContext
Imports System.Net
Imports Wellan.Service
Imports Wellan.Ecommerce.StoreFront
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class CheckOutGoGWTSPG
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        WUserLog1.UserID = Context.User.Identity.Name()
        WUserLog1.MemberID = Context.User.Identity.Name()

        If HttpUtility.HtmlEncode(Request.Params("FromType")) = "APP" Then
            WUserLog1.FunctionID = "APP_CheckOutGoGWTSPG"
        End If
        WUserLog1.WriteLog()

        Dim Flag As Boolean = False
        Dim sMessage As String = ""

        Session("GOGWP") = "TSPG"

        Dim ToTSPG_URL As String = ""

        Dim layoutType = "1"     '客戶端版面類型  1: 一般網頁   2:行動裝置網頁
        Dim TSPG_mid As String = BWEX.GetBSNAM(BWEX.GetBSNUM("Store"), "TSPG_mid", "") '特店代號


        If HttpUtility.HtmlEncode(Request.Params("FromType")) = "APP" Then
            Session("IsOrderFinish") = "N"
            Session("IsSendCompleteMail") = "N" '訂購完成後，系統提示完成訂單並寄出確認信，避免切換語系後又出現一次提示視窗

            '將購買的單號放到準備要傳送到金流系統的訂單編號欄位
            'Session("od_sob") = HttpUtility.HtmlEncode(ReplaceStr(Request.Params("orderNo")))
            Session("PaymentType") = "19"   '傳送付款方式為參數
            layoutType = "2"
            Session("FromType") = "APP"

            'S180808021 現場自取使用的金流刷卡的商店代號皆要獨立
            TSPG_mid = SalesOrder_GetTSPG_mid(Session("od_sob"))

        End If

        'S210615019 PCI DSS CSRF
        If Request.Cookies("CSRFToken").Value Is Nothing OrElse AES.Decrypt(HttpUtility.HtmlEncode(Request.Cookies("CSRFToken").Value)) <> UtilService.GetVal(Session("od_sob"), ValCollection.文字) Then

            NLogToFile.Fatal("CheckOutGoGWTSPG_Fail:CSRF check error:MemberID:" & Context.User.Identity.Name() & ",Cookies(CSRFToken):" & Request.Cookies("CSRFToken").Value & ",Session(od_sob):" & UtilService.GetVal(Session("od_sob"), ValCollection.文字) & ",UserAgents:" & Context.Request.UserAgent())

            If HttpUtility.HtmlEncode(Request.Params("FromType")) <> "APP" Then
                '線上刷卡系統有誤，請稍後再試
                ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alert", "alert('" & Wellan.Web.Application.ResourceString("ShoppingCart_PaymentError1") & "');location.href='" & BWEX.GetBSNAM("Store", "ShoppingCart0") & "';", True)
            Else
                Response.Redirect("wellan://paymentError") 'APP回傳
            End If

        End If

        Dim sSQL As String = "SELECT Amt FROM BillB WHERE BillNo = @BillNo AND PaymentType = '19' "
        Using cn As New System.Data.SqlClient.SqlConnection(ConfigurationManager.AppSettings("ConnectionString"))
            Try
                Dim cmd As New System.Data.SqlClient.SqlCommand(sSQL, cn)

                cmd.Parameters.AddWithValue("@BillNo", Session("od_sob"))
                cn.Open()
                '包含兩位小數，如100 代表1.00 元。
                Session("amount") = Convert.ToInt32(UtilService.GetVal(cmd.ExecuteScalar, ValCollection.數值)) * 100

            Catch ex As Exception
                Session("amount") = 0
                NLogToFile.ErrorInfo("CheckOutGoGWTSPG_Amt:ErrorMsg:" & ex.Message)
                Dim objErrorLog As Wellan.Ecommerce.StoreFront.ErrorLog = New Wellan.Ecommerce.StoreFront.ErrorLog
                objErrorLog.WriteErrorLog("StoreFront_CheckOutGoGWTSPG", "GetAmt", ex, "")
                objErrorLog = Nothing
            Finally
                cn.Close()
                cn.Dispose()
            End Try
        End Using

        NLogToFile.Trace("CheckOutGoGWTSPG:  From " & HttpUtility.HtmlEncode(Request.Params("FromType")) & " SalesOrderNo:" & Session("od_sob") & " ,MemberID: " & Context.User.Identity.Name())

        Dim TSPG_APIROOT As String = BWEX.GetBSNAM(BWEX.GetBSNUM("Store"), "TSPG_APIROOT", "") '路徑

        Dim TSPG_tid As String = BWEX.GetBSNAM(BWEX.GetBSNUM("Store"), "TSPG_tid", "") '端末代號
        Dim TSPG_post_back_url As String = BWEX.GetBSNAM(BWEX.GetBSNUM("Store"), "TSPG_post_back_url", "")
        Dim TSPG_result_url As String = BWEX.GetBSNAM(BWEX.GetBSNUM("Store"), "TSPG_result_url", "")

        If TSPG_APIROOT <> "" AndAlso TSPG_mid <> "" AndAlso Session("amount") <> 0 Then
            Try
                Dim httpWebRequest = DirectCast(WebRequest.Create("https://" & TSPG_APIROOT & "/auth.ashx"), HttpWebRequest)
                'S180604019   升級TLS 1.2
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                'S190314005 (台新建議) 設定KeepAlive屬性設false以避免建立與網際網路資源的持續性連線。
                httpWebRequest.KeepAlive = False
                'S191007015 (台新建議) 用SetHeaderValue()來設定 Connection HTTP 標頭的值並將程式順序調整如下, 將httpWebRequest.KeepAlive = False 放到最上避免setHaderValue後又被蓋掉
                SetHeaderValue(httpWebRequest.Headers, "Connection", "keep-alive")

                httpWebRequest.ContentType = "application/json"
                httpWebRequest.Method = "POST"

                Dim json As String = "{""sender"":""rest""," + """ver"":""1.0.0""," + """mid"":""" & TSPG_mid & """," + """tid"":""" & TSPG_tid & """," + """pay_type"":1," + """tx_type"":1," + """params"":" + "{""layout"":""" & layoutType & """," + """order_no"":""" & Session("od_sob") & """," + """amt"":""" & Session("amount") & """," + """cur"":""NTD""," + """order_desc"":""" & Session("od_sob") & """," + """capt_flag"":""0""," + """result_flag"":""1""," + """post_back_url"":""" & TSPG_post_back_url & """," + """result_url"":""" & TSPG_result_url & """}}"
                NLogToFile.Trace("CheckOutGoGWTSPG_JSON : " & json)

                Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())

                    streamWriter.Write(json)
                    streamWriter.Flush()
                    streamWriter.Close()
                    NLogToFile.Trace("CheckOutGoGWTSPG_httpWebRequest End:" & Session("od_sob"))
                End Using

                Dim webResponse As HttpWebResponse = httpWebRequest.GetResponse()
                If Not httpWebRequest.HaveResponse Then

                    NLogToFile.Trace("CheckOutGoGWTSPG GetResponse :  網站無回應  SalesOrderNo:" & Session("od_sob"))

                    Exit Try
                End If

                If webResponse.StatusCode <> HttpStatusCode.OK And webResponse.StatusCode <> HttpStatusCode.Accepted Then
                    NLogToFile.Trace("CheckOutGoGWTSPG GetResponse :  回應狀況不ok , SalesOrderNo:" & Session("od_sob"))
                    Exit Try
                End If

                Dim GetResponseData As JObject
                Dim result As String = ""
                Dim httpResponse = DirectCast(httpWebRequest.GetResponse(), HttpWebResponse)
                Using streamReader = New StreamReader(httpResponse.GetResponseStream())

                    result = streamReader.ReadToEnd()
                    GetResponseData = CType(JsonConvert.DeserializeObject(result), JObject)

                    NLogToFile.Trace("CheckOutGoGWTSPG  SalesOrderNo: " & Session("od_sob") & " , GetResponse : " & result.ToString())

                    If GetResponseData("params")("ret_code") = "00" Then
                        ToTSPG_URL = GetResponseData("params")("hpp_url")
                    End If
                End Using

                httpResponse.Close()

            Catch ex As Exception
                Dim objErrorLog As ErrorLog = New ErrorLog
                objErrorLog.WriteErrorLog("StoreFront_CheckOutGoGWTSPG", "", ex, "SalesOrderNo: " & Session("od_sob"))
                objErrorLog = Nothing
            End Try

            If ToTSPG_URL = String.Empty Then

                'S190212040 前台需改成判斷有異常時, 需再判斷來源, 若為前台時則依原本做法, 若為APP時則再傳送相關訊息導回APP, 由APP接收後再做其他後續流程
                If HttpUtility.HtmlEncode(Request.Params("FromType")) <> "APP" Then
                    '線上刷卡系統有誤，請稍後再試
                    ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alert", "alert('" & Wellan.Web.Application.ResourceString("ShoppingCart_PaymentError1") & "');location.href='" & BWEX.GetBSNAM("Store", "ShoppingCart0") & "';", True)
                Else
                    Response.Redirect("wellan://paymentError") 'APP回傳
                End If

            Else
                Response.Redirect(ToTSPG_URL)
            End If
        Else

            'S190212040 前台需改成判斷有異常時, 需再判斷來源, 若為前台時則依原本做法, 若為APP時則再傳送相關訊息導回APP, 由APP接收後再做其他後續流程
            If HttpUtility.HtmlEncode(Request.Params("FromType")) <> "APP" Then
                '線上刷卡系統有誤，請稍後再試
                ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alert", "alert('" & Wellan.Web.Application.ResourceString("ShoppingCart_PaymentError1") & "');location.href='" & BWEX.GetBSNAM("Store", "ShoppingCart0") & "';", True)
            Else
                Response.Redirect("wellan://paymentError") 'APP回傳
            End If

        End If
    End Sub

End Class
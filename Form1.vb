Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports System.Globalization
Imports System.Text


' Date: 02/04/2020
' Written By: Chris Dreyer, adapted from previous version
' Purpose: Send HTML email to confirm a new quote, order and shipment confirmation

Public Class Form1
    Dim OrderNumber As String
    Dim CreateUser As String
    Dim SectionError As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Look4QuoteConfirmations()
        'Look4OrderConfirmations()
        Look4DeliveryConfirmations()
        Application.Exit()
    End Sub

    Private Sub Look4QuoteConfirmations()

        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False

        Try
            ' SQL Statement to pull any unsent order confirmations
            Dim SQLStatement As String = "SELECT * FROM PILOT.ZORDCON WHERE STA_0 = 4"
            Dim QuoteConfirmations As DataTable = OpenDataSet(SQLStatement)
            If QuoteConfirmations.Rows.Count > 0 Then
                ' Loops through all the unsent records
                For i = 0 To QuoteConfirmations.Rows.Count - 1
                    ' Sets a variable with the order number
                    OrderNumber = QuoteConfirmations.Rows(i).Item("SOHNUM_0")
                    CreateUser = QuoteConfirmations.Rows(i).Item("CREUSR_0")
                    ' Sends the order confirmation and sets the status to sent
                    SendQuoteConfirmation(OrderNumber)
                    ExecuteSQLQuery("UPDATE PILOT.ZORDCON SET STA_0 = 2 WHERE SOHNUM_0 = '" + OrderNumber + "' AND STA_0 = 4")
                Next
            End If
        Catch ex As Exception
            ' If the code encounters an error this sets the status to 3 and sends an email about the error
            ExecuteSQLQuery("UPDATE PILOT.ZORDCON SET STA_0 = 3 WHERE SOHNUM_0 = '" + OrderNumber + "'")
            SendHTMLEmail("tbailey@packbgr.com;cdreyer@packbgr.com", "BGR <do-not-reply@bgr.us>", "ERROR", SectionError + "<BR>" + vbNewLine + ex.ToString + vbNewLine + OrderNumber)
        End Try

    End Sub


    Private Sub SendQuoteConfirmation(ByVal OrderNumber As String)

        ' SQL Statement to pull the header information of the quote confirmation
        SectionError = "SQL Statement to pull the header information of the quote confirmation"
        Dim SQLStatement As String
        SQLStatement = <Sql><![CDATA[
       SELECT 
                                 s.SQHNUM_0, s.BPCORD_0, s.ORDDAT_0, s.BPCNAM_0, s.BPCNAM_1, s.CUSQUOREF_0,
                                 s.BPCADDLIG_0, s.BPCADDLIG_1, s.BPCADDLIG_2, s.BPCPOSCOD_0, s.BPCCTY_0, 
                                 s.BPCSAT_0, s.BPDNAM_0, s.BPDNAM_1, s.BPDADDLIG_0, s.BPDADDLIG_1, 
                                 s.BPDADDLIG_2, s.BPDPOSCOD_0, s.BPDCTY_0, s.BPDSAT_0, s.BPDCRYNAM_0, 
                                 s.BPCCRYNAM_0, s.INVDTAAMT_1, PILOT.CONTACT.CNTOAEML1_0, PILOT.CONTACT.CNTOAEML2_0, PILOT.CONTACT.CNTOAEML3_0, 
                                 PILOT.CONTACT.CNTOAEML4_0, PILOT.CONTACT.WEB_0, PILOT.TABPAYTERM.LANDESSHO_0 AS PAYTERM, PILOT.ATEXTRA.TEXTE_0, 
                                 s.INVDTAAMT_0 AS MINORD, COALESCE (PILOT.TEXCLOB.TEXTE_0, N'') AS Header, b.REPNAME_0, b.REPPHONE_0, 
                                 b.REPEML_0, a.REPNAME_0 as CSRNAME_0, a.REPPHONE_0 as CSRPHONE_0, a.REPEML_0 as CSREML_0
        FROM            PILOT.TABPAYTERM INNER JOIN
                                 PILOT.CONTACT INNER JOIN
                                 PILOT.SQUOTE as s ON PILOT.CONTACT.CCNCRM_0 = s.ZCONTACT_0 AND PILOT.CONTACT.BPANUM_0 = s.BPCORD_0 ON 
                                 PILOT.TABPAYTERM.PTE_0 = s.PTE_0 INNER JOIN
                                 PILOT.ATEXTRA ON s.EECICT_0 = PILOT.ATEXTRA.IDENT2_0 LEFT OUTER JOIN
                                 PILOT.ZBGRREPS as b ON s.REP_0 = b.USR_0 LEFT OUTER JOIN
                                 PILOT.TEXCLOB ON s.SQHTEX1_0 = PILOT.TEXCLOB.CODE_0 LEFT OUTER JOIN
								 PILOT.ZBGRREPS as a ON s.ZCSR_0 = a.USR_0
        WHERE        (PILOT.ATEXTRA.CODFIC_0 = 'ATABDIV') AND (PILOT.ATEXTRA.LANGUE_0 = 'ENG') AND (PILOT.ATEXTRA.ZONE_0 = 'LNGDES') AND (PILOT.ATEXTRA.IDENT1_0 = '5') 
                                 AND (s.SQHNUM_0 = N'***OrderNumber***')
        ]]></Sql>.Value
        SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
        Dim Order As DataTable = OpenDataSet(SQLStatement)

        SQLStatement = "SELECT REPEML_0,DEPARTMENT_0 FROM PILOT.ZBGRREPS WHERE USR_0 = '" + CreateUser + "'"
        Dim CreateUserDT As DataTable = OpenDataSet(SQLStatement)
        Dim CreateUserEmail As String = CreateUserDT.Rows(0).Item("REPEML_0")
        Dim CreateUserDept As Integer = CreateUserDT.Rows(0).Item("DEPARTMENT_0")

        If Order.Rows.Count > 0 Then
            Dim CustomerRef As String = Order.Rows(0).Item("CUSQUOREF_0")
            ' Opens the html template file and sets the entire file as a string "Body"
            SectionError = "Opens the html template file and sets the entire file as a string Body"
            Dim FileReader As StreamReader = New StreamReader("C:\Program Files (x86)\WebConfirmation\QuoteAckTemplate.html")
            Dim Body As String = FileReader.ReadToEnd
            FileReader.Close()

            ' Replaces the Header information in the template with data from the SQL statement
            SectionError = "Replaces the Header information in the template with data from the SQL statement"
            Body = Replace(Body, "***OrderNumber***", OrderNumber)
            'Body = Replace(Body, "***PONumber***", Order.Rows(0).Item("CUSORDREF_0"))
            Body = Replace(Body, "***PaymentTerms***", Order.Rows(0).Item("PAYTERM").ToString.Substring(InStr(Order.Rows(0).Item("PAYTERM"), "~"), InStr(InStr(Order.Rows(0).Item("PAYTERM"), "~") + 1, Order.Rows(0).Item("PAYTERM"), "~") - InStr(Order.Rows(0).Item("PAYTERM"), "~") - 1))
            Body = Replace(Body, "***Customer***", Order.Rows(0).Item("BPCORD_0") + " - " + Order.Rows(0).Item("BPCNAM_0"))
            Body = Replace(Body, "***FreightTerms***", Order.Rows(0).Item("TEXTE_0"))
            'Body = Replace(Body, "***ShipMethod***", Order.Rows(0).Item("MDL_0") + " (" + Order.Rows(0).Item("BPTNUM_0") + ")")
            Body = Replace(Body, "***SalesRep***", Order.Rows(0).Item("REPNAME_0"))
            Body = Replace(Body, "***RepEmail***", Order.Rows(0).Item("REPEML_0"))
            Body = Replace(Body, "***Phone***", Order.Rows(0).Item("REPPHONE_0"))
            Body = Replace(Body, "***SalesRep2***", Order.Rows(0).Item("CSRNAME_0"))
            Body = Replace(Body, "***RepEmail2***", Order.Rows(0).Item("CSREML_0"))
            Body = Replace(Body, "***Phone2***", Order.Rows(0).Item("CSRPHONE_0"))
            'Body = Replace(Body, "***RevisionNumber***", Order.Rows(0).Item("REVNUM"))
            RichTextBox1.Text = ""
            If Order.Rows(0).Item("Header").ToString().Length > 1 Then

                If Order.Rows(0).Item("Header").ToString().Substring(0, 1) = "{" Then
                    RichTextBox1.Rtf = Order.Rows(0).Item("Header")
                Else
                    RichTextBox1.Text = Order.Rows(0).Item("Header")
                End If
            End If
            Body = Replace(Body, "***Header***", RichTextBox1.Text)

            ' BillTo
            SectionError = "BillTo"
            Dim BillTo As String = ""
            If Len(Order.Rows(0).Item("BPCNAM_0")) > 2 Then BillTo += Order.Rows(0).Item("BPCNAM_0") + "<BR>"
            If Len(Order.Rows(0).Item("BPCADDLIG_0")) > 2 Then BillTo += Order.Rows(0).Item("BPCADDLIG_0") + "<BR>"
            If Len(Order.Rows(0).Item("BPCADDLIG_1")) > 2 Then BillTo += Order.Rows(0).Item("BPCADDLIG_1") + "<BR>"
            If Len(Order.Rows(0).Item("BPCADDLIG_2")) > 2 Then BillTo += Order.Rows(0).Item("BPCADDLIG_2") + "<BR>"
            BillTo += Order.Rows(0).Item("BPCCTY_0") + ", " + Order.Rows(0).Item("BPCSAT_0") + " " + Order.Rows(0).Item("BPCPOSCOD_0")
            Body = Replace(Body, "***BillTo***", BillTo)

            ' ShipTo
            SectionError = "ShipTo"
            Dim ShipTo As String = ""
            If Len(Order.Rows(0).Item("BPDNAM_0")) > 2 Then ShipTo += Order.Rows(0).Item("BPDNAM_0") + "<BR>"
            If Len(Order.Rows(0).Item("BPDADDLIG_0")) > 2 Then ShipTo += Order.Rows(0).Item("BPDADDLIG_0") + "<BR>"
            If Len(Order.Rows(0).Item("BPDADDLIG_1")) > 2 Then ShipTo += Order.Rows(0).Item("BPDADDLIG_1") + "<BR>"
            If Len(Order.Rows(0).Item("BPDADDLIG_2")) > 2 Then ShipTo += Order.Rows(0).Item("BPDADDLIG_2") + "<BR>"
            ShipTo += Order.Rows(0).Item("BPDCTY_0") + ", " + Order.Rows(0).Item("BPDSAT_0") + " " + Order.Rows(0).Item("BPDPOSCOD_0")
            Body = Replace(Body, "***ShipTo***", ShipTo)

            ' SQL Statement to pull the product line information
            SectionError = "SQL Statement to pull the product line information"
            SQLStatement = <Sql><![CDATA[
                SELECT        PILOT.SQUOTE.SQHNUM_0,PILOT.SQUOTED.ITMREF_0, PILOT.SQUOTED.GROPRI_0 as NETPRI_0, 
                                         PILOT.SQUOTED.QTY_0, PILOT.SQUOTED.ITMREFBPC_0,
                                         PILOT.ITMMASTER.ZWEBTITLE_0 ITMDES_0,  
                                         PILOT.SQUOTED.QTY_0, PILOT.SQUOTED.ZQTY2_0, COALESCE (PILOT.TEXCLOB.TEXTE_0, N'') AS LineText, PILOT.SQUOTED.SAU_0 AS UOM, 
                                         PILOT.TABUNIT.UOMDEC_0 AS NUMDEC
                FROM            PILOT.SQUOTE INNER JOIN
                                         PILOT.SQUOTED ON PILOT.SQUOTE.SQHNUM_0 = PILOT.SQUOTED.SQHNUM_0 INNER JOIN
                                         PILOT.ITMMASTER ON PILOT.SQUOTED.ITMREF_0 = PILOT.ITMMASTER.ITMREF_0 INNER JOIN
                                         PILOT.TABUNIT ON PILOT.SQUOTED.SAU_0 = PILOT.TABUNIT.UOM_0 LEFT OUTER JOIN
                                         PILOT.TEXCLOB ON PILOT.SQUOTED.SQDTEX_0 = PILOT.TEXCLOB.CODE_0
                WHERE        (PILOT.SQUOTE.SQHNUM_0 = N'***OrderNumber***')
            ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
            Dim OrderDetails As DataTable = OpenDataSet(SQLStatement)

            ' Filling in the product lines
            SectionError = "Filling in the product lines"
            Dim Details As String = ""
            Dim SubTotal As Decimal = 0
            Dim NUMDEC As Integer = 0
            Dim CustomerItemRef As String = ""
            If OrderDetails.Rows.Count > 0 Then
                For i = 0 To OrderDetails.Rows.Count - 1
                    SectionError = "Filling in the product lines 0"
                    NUMDEC = OrderDetails.Rows(i).Item("NUMDEC")
                    If OrderDetails.Rows(i).Item("ITMREFBPC_0").GetType() Is GetType(DBNull) Then
                        CustomerItemRef = ""
                    Else
                        CustomerItemRef = OrderDetails.Rows(i).Item("ITMREFBPC_0")
                    End If
                    Details += vbTab + "<tr style='height: 50px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbNewLine
                    If CustomerItemRef = "" Then
                        Details += vbTab + vbTab + "<td>" + OrderDetails.Rows(i).Item("ITMREF_0") + "</td>" + vbNewLine
                    ElseIf OrderDetails.Rows(i).Item("ITMREF_0") <> CustomerItemRef Then
                        Details += vbTab + vbTab + "<td>" + OrderDetails.Rows(i).Item("ITMREF_0") + "(" + CustomerItemRef + ")" + "</td>" + vbNewLine
                    Else
                        Details += vbTab + vbTab + "<td>" + OrderDetails.Rows(i).Item("ITMREF_0") + "</td>" + vbNewLine
                    End If
                    Details += vbTab + vbTab + "<td>" + OrderDetails.Rows(i).Item("ITMDES_0") + "</td>" + vbNewLine
                    SectionError = "Filling in the product lines 00"
                    Details += vbTab + vbTab + "<td align='center'>" + OrderDetails.Rows(i).Item("UOM").ToString() + "</center></td>" + vbNewLine
                    If Math.Round(OrderDetails.Rows(i).Item("ZQTY2_0"), NUMDEC) = 999999999 Then
                        Details += vbTab + vbTab + "<td align='center'>" + Math.Round(OrderDetails.Rows(i).Item("QTY_0"), NUMDEC).ToString() + "+</center></td>" + vbNewLine
                    Else
                        Details += vbTab + vbTab + "<td align='center'>" + Math.Round(OrderDetails.Rows(i).Item("QTY_0"), NUMDEC).ToString() + "-" + Math.Round(OrderDetails.Rows(i).Item("ZQTY2_0"), NUMDEC).ToString() + "</center></td>" + vbNewLine
                    End If

                    'SectionError = "Filling in the product lines 1"
                    'Details += vbTab + vbTab + "<td align='center'>" + Math.Round(OrderDetails.Rows(i).Item("SHIQTY"), NUMDEC).ToString() + "</center></td>" + vbNewLine
                    'SectionError = "Filling in the product lines 2"
                    'Details += vbTab + vbTab + "<td align='center'>" + Math.Round(OrderDetails.Rows(i).Item("BOQTY"), NUMDEC).ToString() + "</center></td>" + vbNewLine
                    'SectionError = "Filling in the product lines 3"
                    'Details += vbTab + vbTab + "<td align='center'>" + DateValue(OrderDetails.Rows(i).Item("DEMDLVDAT_0")).ToString("MM/dd/yyyy") + "</center></td>" + vbNewLine
                    SectionError = "Filling in the product lines 4"
                    Details += vbTab + vbTab + "<td align='center'>$" + DirectCast(Math.Round(OrderDetails.Rows(i).Item("NETPRI_0"), 4), Decimal).ToString("#.00##") + "</td>" + vbNewLine
                    SectionError = "Filling in the product lines 5"
                    Details += vbTab + vbTab + "<td align='center'>$" + Math.Round(OrderDetails.Rows(i).Item("NETPRI_0") * OrderDetails.Rows(i).Item("QTY_0"), 2).ToString + "</td>" + vbNewLine
                    SectionError = "Filling in the product lines 6"
                    Details += vbTab + "</tr>" + vbNewLine
                    SubTotal += OrderDetails.Rows(i).Item("NETPRI_0") * OrderDetails.Rows(i).Item("QTY_0")
                    If OrderDetails.Rows(i).Item("LineText") <> "" Then
                        RichTextBox1.Text = ""
                        If OrderDetails.Rows(i).Item("LineText").ToString().Length > 1 Then

                            If OrderDetails.Rows(i).Item("LineText").ToString().Substring(0, 1) = "{" Then
                                RichTextBox1.Rtf = OrderDetails.Rows(i).Item("LineText")
                            Else
                                RichTextBox1.Text = OrderDetails.Rows(i).Item("LineText")
                            End If
                        End If
                        Details += vbTab + "<tr style='height: 50px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbNewLine
                        Details += vbTab + vbTab + "<td colspan = 8>" + RichTextBox1.Text + "</td>" + vbNewLine
                        Details += vbTab + "</tr>" + vbNewLine
                    End If
                Next
            End If

            ' SubTotal
            'SectionError = "SubTotal"
            'Details += vbTab + "<tr style='height: 50px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbNewLine
            'Details += vbTab + vbTab + "<td colspan=""5""></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td align=""leftcolspan=4""><b>SubTotal</b></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td align=""right"">$" + Math.Round(SubTotal, 2).ToString + "</td>" + vbNewLine
            'Details += vbTab + "</tr>" + vbNewLine

            ' Delivery Fee
            ' SectionError = "Delivery Fee"
            'Details += vbTab + "<tr style='height: 50px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbNewLine
            'Details += vbTab + vbTab + "<td colspan=""5""></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td align=""leftcolspan=4""><b>Delivery Fee</b></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td align=""right"">$" + Math.Round(Order.Rows(0).Item("MINORD"), 2).ToString + "</td>" + vbNewLine
            'Details += vbTab + "</tr>" + vbNewLine

            ' Freight
            'SectionError = "Freight"
            'If Order.Rows(0).Item("INVDTAAMT_1") > 0 Then
            '    Details += vbTab + "<tr style='height: 50px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbNewLine
            '    Details += vbTab + vbTab + "<td colspan=""5""></td>" + vbNewLine
            '    Details += vbTab + vbTab + "<td align=""leftcolspan=4""><b>Freight</b></td>" + vbNewLine
            '    Details += vbTab + vbTab + "<td></td>" + vbNewLine
            '    Details += vbTab + vbTab + "<td align=""right"">$" + Math.Round(Order.Rows(0).Item("INVDTAAMT_1"), 2).ToString + "</td>" + vbNewLine
            '    Details += vbTab + "</tr>" + vbNewLine
            'End If

            '' Tax
            'SectionError = "Tax"
            'If Order.Rows(0).Item("ORDINVATI_0") - Order.Rows(0).Item("ORDINVNOT_0") > 0 Then
            '    Details += vbTab + "<tr style='height: 50px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbNewLine
            '    Details += vbTab + vbTab + "<td colspan=""5""></td>" + vbNewLine
            '    Details += vbTab + vbTab + "<td align=""leftcolspan=4""><b>Tax</b></td>" + vbNewLine
            '    Details += vbTab + vbTab + "<td></td>" + vbNewLine
            '    Details += vbTab + vbTab + "<td align=""right"">$" + Math.Round(Order.Rows(0).Item("ORDINVATI_0") - Order.Rows(0).Item("ORDINVNOT_0"), 2).ToString + "</td>" + vbNewLine
            '    Details += vbTab + "</tr>" + vbNewLine
            'End If

            ' Total
            'SectionError = "Total"
            'Details += vbTab + "<tr style='height: 50px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbNewLine
            'Details += vbTab + vbTab + "<td colspan=""5""></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td align=""leftcolspan=4""><b>Total</b></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td></td>" + vbNewLine
            'Details += vbTab + vbTab + "<td align=""right"">$" + Math.Round(Order.Rows(0).Item("INVDTAAMT_0"), 2).ToString + "</td>" + vbNewLine
            'Details += vbTab + "</tr>" + vbNewLine

            ' Add Details to the HTML
            SectionError = "Add Details to the HTML"
            Body = Replace(Body, "***LINES***", Details)

            ' Constructing the ToAddress of the email
            SectionError = "Constructing the ToAddress of the email"
            Dim EmailTo As String = ""
            If Len(Order.Rows(0).Item("WEB_0")) > 1 Then
                EmailTo = EmailTo + Order.Rows(0).Item("WEB_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML1_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML1_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML1_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML2_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML2_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML2_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML3_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML3_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML3_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML4_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML4_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML4_0") + ";"
            End If

            Dim BCC As String = ""
            If Len(CreateUserEmail) > 2 Then
                BCC = CreateUserEmail
            End If
            If Len(CreateUserEmail) > 2 And Len(Order.Rows(0).Item("CSREML_0")) > 2 Then
                BCC = BCC + ";"
            End If
            If Len(Order.Rows(0).Item("CSREML_0")) > 2 Then
                BCC = BCC + Order.Rows(0).Item("CSREML_0")
            End If
            If Len(Order.Rows(0).Item("REPEML_0")) > 2 Then
                BCC = BCC + ";"
                BCC = BCC + Order.Rows(0).Item("REPEML_0")
            End If

            ' Sends the Email "BGR <do-not-reply@bgr.us>"
            SectionError = "Sends the Email"
            If Len(Order.Rows(0).Item("REPEML_0")) > 1 And Len(EmailTo) > 1 Then
                EmailTo = EmailTo.Substring(0, Len(EmailTo) - 1)
                If CreateUserDept = 10 Then
                    If CustomerRef = "" Then
                        SendHTMLEmail(EmailTo, CreateUserEmail, "Quote Confirmation", Body, "", BCC)
                    Else
                        SendHTMLEmail(EmailTo, CreateUserEmail, "Quote Confirmation" + "(" + CustomerRef + ")", Body, "", BCC)
                    End If
                Else
                    If CustomerRef = "" Then
                        SendHTMLEmail(EmailTo, Order.Rows(0).Item("REPEML_0"), "Quote Confirmation", Body, "", BCC)
                    Else
                        SendHTMLEmail(EmailTo, Order.Rows(0).Item("REPEML_0"), "Quote Confirmation" + "(" + CustomerRef + ")", Body, "", BCC)
                    End If
                End If
            End If
            'SendHTMLEmail("msullivan@razorsoft.biz", "BGR <do-not-reply@bgr.us>", "Order Confirmation", EmailTo + "<BR>" + Body)
        Else
            ' This is what happens if the order doesn't exist
            ExecuteSQLQuery("UPDATE PILOT.ZORDCON SET STA_0 = 3 WHERE SOHNUM_0 = '" + OrderNumber + "'")
            SendHTMLEmail("tbailey@bgr.us;bgertner@bgr.us;", "BGR <do-not-reply@bgr.us>", "ERROR - Quote Confirmation cannot Find Order - " + OrderNumber, "ERROR - Quote Confirmation cannot Find Order - " + OrderNumber)
            'SendHTMLEmail("msullivan@razorsoft.biz", "BGR <do-not-reply@bgr.us>", "ERROR - Order Confirmation cannot Find Order - " + OrderNumber, "ERROR - Order Confirmation cannot Find Order - " + OrderNumber)
        End If

    End Sub
    Private Sub Look4OrderConfirmations()
        Try
            ' SQL Statement to pull any unsent order confirmations
            Dim SQLStatement As String = "SELECT * FROM PILOT.ZORDCON WHERE STA_0 = 1"
            Dim OrderConfirmations As DataTable = OpenDataSet(SQLStatement)
            If OrderConfirmations.Rows.Count > 0 Then
                ' Loops through all the unsent records
                For i = 0 To OrderConfirmations.Rows.Count - 1
                    ' Sets a variable with the order number
                    OrderNumber = OrderConfirmations.Rows(i).Item("SOHNUM_0")
                    CreateUser = OrderConfirmations.Rows(i).Item("CREUSR_0")
                    ' Sends the order confirmation
                    SendOrderConfirmation(OrderNumber)
                    ' Sets the status to sent
                    ExecuteSQLQuery("UPDATE PILOT.ZORDCON SET STA_0 = 2 WHERE SOHNUM_0 = '" + OrderNumber + "' AND STA_0 = 1")
                Next
            End If
        Catch ex As Exception
            ' If the code encounters an error this sets the status to 3 and sends an email about the error
            ExecuteSQLQuery("UPDATE PILOT.ZORDCON SET STA_0 = 3 WHERE SOHNUM_0 = '" + OrderNumber + "'")
            SendHTMLEmail("cdreyer@packbgr.com", "BGR <do-not-reply@packbgr.com>", "Testing - ERROR - BIG LINES AT THE TOP", SectionError + "<BR>" + vbCrLf + ex.ToString + vbCrLf + OrderNumber)
            'tbailey@packbgr.com;
        End Try

    End Sub

    Private Sub SendOrderConfirmation(ByVal OrderNumber As String)

        Dim SQLStatement As String
        Dim Order As DataTable
        Dim CreateUserDT As DataTable
        Dim CreateUserEmail As String
        Dim CreateUserDept As Integer
        Dim FileReader As StreamReader
        Dim Body As String
        Dim RichTextBox1 As Object = Nothing
        Dim strbuffer As String
        Dim blnDebugFlag As Boolean

        blnDebugFlag = False

        ' SQL Statement to pull the header information of the order confirmation
        SectionError = "ORDCON - SQL Statement to pull the header information of the order confirmation"

        SQLStatement = <Sql><![CDATA[
           SELECT     s.SOHNUM_0, s.BPCORD_0, s.CUSORDREF_0, s.ORDDAT_0, s.BPINAM_0, 
                                  s.BPINAM_1, s.BPIADDLIG_0, s.BPIADDLIG_1, s.BPIADDLIG_2, s.BPIPOSCOD_0, 
                                  s.BPICTY_0, s.BPISAT_0, s.BPDNAM_0, s.BPDNAM_1, s.BPDADDLIG_0, 
                                  s.BPDADDLIG_1, s.BPDADDLIG_2, s.BPDPOSCOD_0, s.BPDCTY_0, s.BPDSAT_0, 
                                  s.BPDCRYNAM_0, s.BPCCRYNAM_0, s.INVDTAAMT_1, s.ORDINVNOT_0, s.ORDINVATI_0, 
                                  s.MDL_0, s.BPTNUM_0, PILOT.CONTACT.CNTOAEML1_0, PILOT.CONTACT.CNTOAEML2_0, PILOT.CONTACT.CNTOAEML3_0, 
                                  PILOT.CONTACT.CNTOAEML4_0, PILOT.CONTACT.WEB_0, PILOT.TABPAYTERM.LANDESSHO_0 AS PAYTERM, PILOT.ATEXTRA.TEXTE_0, 
                                  s.INVDTAAMT_0 AS MINORD, s.INVDTAAMT_5 AS HANDLFEE, COALESCE (PILOT.TEXCLOB.TEXTE_0, N'') AS Header, b.REPNAME_0, b.REPPHONE_0, 
                                  b.REPEML_0, a.REPNAME_0 as CSRNAME_0, a.REPPHONE_0 as CSRPHONE_0, a.REPEML_0 as CSREML_0, s.REVNUM_0 AS REVNUM, cc.CCL4D_0 AS CC4, cc.CCTYPE_0 AS CCType
            FROM         PILOT.TABPAYTERM INNER JOIN
                                  PILOT.CONTACT INNER JOIN
                                  PILOT.SORDER as s ON PILOT.CONTACT.CCNCRM_0 = s.ZCONTACT_0 and PILOT.CONTACT.BPANUM_0 = s.BPCORD_0 ON PILOT.TABPAYTERM.PTE_0 = s.PTE_0 INNER JOIN
                                  PILOT.ATEXTRA ON s.EECICT_0 = PILOT.ATEXTRA.IDENT2_0 INNER JOIN
                                  PILOT.ZBGRREPS as b ON s.REP_0 = b.USR_0 LEFT OUTER JOIN
                                  PILOT.TEXCLOB ON s.SOHTEX1_0 = PILOT.TEXCLOB.CODE_0 INNER JOIN
								  PILOT.ZBGRREPS as a ON s.ZCSR_0 = a.USR_0 LEFT JOIN
                                  PILOT.XCORDAPP as cc ON s.SOHNUM_0 = cc.VCRNUM_0
            WHERE     (PILOT.ATEXTRA.CODFIC_0 = N'ATABDIV') AND (PILOT.ATEXTRA.LANGUE_0 = N'ENG') AND (PILOT.ATEXTRA.ZONE_0 = N'LNGDES') AND 
                                  (PILOT.ATEXTRA.IDENT1_0 = N'5') AND (s.SOHNUM_0 = N'***OrderNumber***')
        ]]></Sql>.Value
        SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
        Order = OpenDataSet(SQLStatement)

        SQLStatement = "SELECT REPEML_0,DEPARTMENT_0 FROM PILOT.ZBGRREPS WHERE USR_0 = '" + CreateUser + "'"
        CreateUserDT = OpenDataSet(SQLStatement)
        CreateUserEmail = CreateUserDT.Rows(0).Item("REPEML_0")
        CreateUserDept = CreateUserDT.Rows(0).Item("DEPARTMENT_0")

        If Order.Rows.Count > 0 Then

            ' Opens the html template file and sets the entire file as a string "Body"
            SectionError = "ORDCON - Opens the html template file and sets the entire file as a string Body"
            FileReader = New StreamReader("C:\Program Files (x86)\WebConfirmation\OATemplateNew.html")
            Body = FileReader.ReadToEnd
            FileReader.Close()

            ' Replaces the Header information in the template with data from the SQL statement
            SectionError = "ORDCON - Replaces the Header information in the template with data from the SQL statement"
            Body = Replace(Body, "***OrderNumber***", OrderNumber)
            Body = Replace(Body, "***PONumber***", Order.Rows(0).Item("CUSORDREF_0"))
            strbuffer = Order.Rows(0).Item("PAYTERM").ToString
            strbuffer = strbuffer.Substring(InStr(Order.Rows(0).Item("PAYTERM"), "~"), InStr(
                                            InStr(Order.Rows(0).Item("PAYTERM"), "~") + 1,
                                            Order.Rows(0).Item("PAYTERM"), "~") - InStr(Order.Rows(0).Item("PAYTERM"), "~") - 1)
            Body = Replace(Body, "***PaymentTerms***", strbuffer)
            'Body = Replace(Body, "***Customer***", Order.Rows(0).Item("BPCORD_0") + " - " + Order.Rows(0).Item("BPINAM_0"))
            'Body = Replace(Body, "***FreightTerms***", Order.Rows(0).Item("TEXTE_0"))
            Body = Replace(Body, "***ShipMethod***", GetShipText(Order.Rows(0).Item("MDL_0"), Order.Rows(0).Item("BPTNUM_0")))
            Body = Replace(Body, "***REPNAME***", Order.Rows(0).Item("REPNAME_0"))
            Body = Replace(Body, "***REPEMAIL***", Order.Rows(0).Item("REPEML_0"))
            Body = Replace(Body, "***REPPHONE***", Mid(Order.Rows(0).Item("REPPHONE_0"), 1, 3) _
                                           & "-" + Mid(Order.Rows(0).Item("REPPHONE_0"), 4, 3) _
                                           & "-" + Mid(Order.Rows(0).Item("REPPHONE_0"), 7, 4))
            Body = Replace(Body, "***CSRNAME***", Order.Rows(0).Item("CSRNAME_0"))
            Body = Replace(Body, "***CSREMAIL***", Order.Rows(0).Item("CSREML_0"))
            Body = Replace(Body, "***CSRPHONE***", Mid(Order.Rows(0).Item("CSRPHONE_0"), 1, 3) _
                                           & "-" + Mid(Order.Rows(0).Item("CSRPHONE_0"), 4, 3) _
                                           & "-" + Mid(Order.Rows(0).Item("CSRPHONE_0"), 7, 4))
            'Body = Replace(Body, "***RevisionNumber***", Order.Rows(0).Item("REVNUM"))

            Me.RichTextBox1.Text = ""
            If Order.Rows(0).Item("Header").ToString().Length > 1 Then
                If Order.Rows(0).Item("Header").ToString().Substring(0, 1) = "{" Then
                    Me.RichTextBox1.Rtf = Order.Rows(0).Item("Header")
                Else
                    Me.RichTextBox1.Text = Order.Rows(0).Item("Header")
                End If
            End If
            Body = Replace(Body, "***Header***", Me.RichTextBox1.Text)

            ' BillTo
            SectionError = "ORDCON - BillTo"
            Dim BillTo As String = ""
            If Len(Order.Rows(0).Item("BPINAM_0")) > 2 Then BillTo += Order.Rows(0).Item("BPINAM_0")
            If Len(Order.Rows(0).Item("BPCORD_0")) > 2 Then BillTo += " (Cust #: " + Order.Rows(0).Item("BPCORD_0") + ")" + "<BR>"
            If Len(Order.Rows(0).Item("BPIADDLIG_0")) > 2 Then BillTo += Order.Rows(0).Item("BPIADDLIG_0") + "<BR>"
            If Len(Order.Rows(0).Item("BPIADDLIG_1")) > 2 Then BillTo += Order.Rows(0).Item("BPIADDLIG_1") + "<BR>"
            If Len(Order.Rows(0).Item("BPIADDLIG_2")) > 2 Then BillTo += Order.Rows(0).Item("BPIADDLIG_2") + "<BR>"
            BillTo += Order.Rows(0).Item("BPICTY_0") + ", " + Order.Rows(0).Item("BPISAT_0") + " " + Order.Rows(0).Item("BPIPOSCOD_0")
            Body = Replace(Body, "***BillTo***", BillTo)

            ' ShipTo
            SectionError = "ORDCON - ShipTo"
            Dim ShipTo As String = ""
            If Len(Order.Rows(0).Item("BPDNAM_0")) > 2 Then ShipTo += Order.Rows(0).Item("BPDNAM_0") + "<BR>"
            If Len(Order.Rows(0).Item("BPDADDLIG_0")) > 2 Then ShipTo += Order.Rows(0).Item("BPDADDLIG_0") + "<BR>"
            If Len(Order.Rows(0).Item("BPDADDLIG_1")) > 2 Then ShipTo += Order.Rows(0).Item("BPDADDLIG_1") + "<BR>"
            If Len(Order.Rows(0).Item("BPDADDLIG_2")) > 2 Then ShipTo += Order.Rows(0).Item("BPDADDLIG_2") + "<BR>"
            ShipTo += Order.Rows(0).Item("BPDCTY_0") + ", " + Order.Rows(0).Item("BPDSAT_0") + " " + Order.Rows(0).Item("BPDPOSCOD_0")
            Body = Replace(Body, "***ShipTo***", ShipTo)

            ' SQL Statement to pull the product line information
            SectionError = "ORDCON - SQL Statement to pull the product line information"
            SQLStatement = <Sql><![CDATA[
                SELECT        PILOT.SORDER.SOHNUM_0,PILOT.SORDER.STOFCY_0,PILOT.SORDERP.ITMREF_0 as ITMREFNUM, CASE WHEN LEN(PILOT.SORDERP.ITMREFBPC_0) 
                                         < 2 THEN PILOT.SORDERP.ITMREF_0 ELSE PILOT.SORDERP.ITMREFBPC_0 + '(' + PILOT.SORDERP.ITMREF_0 + ')' END AS ITMREF_0, PILOT.SORDERP.NETPRI_0, 
                                         CASE WHEN LEN(PILOT.SORDERQ.CCLREN_0) < 2 THEN QTY_0 ELSE ODLQTY_0 + DLVQTY_0 END AS QTY_0, 
                                         PILOT.ITMMASTER.ITMDES1_0 + N' ' + PILOT.ITMMASTER.ITMDES2_0 + N' ' + PILOT.ITMMASTER.ITMDES3_0 AS ITMDES_0, PILOT.SORDERQ.DEMDLVDAT_0, 
                                         PILOT.SORDERQ.QTY_0 - PILOT.SORDERQ.SHTQTY_0 AS SHIQTY, PILOT.SORDERQ.SHTQTY_0 AS BOQTY, COALESCE (PILOT.TEXCLOB.TEXTE_0, N'') AS LineText, PILOT.SORDERP.SAU_0 AS UOM, 
                                         PILOT.TABUNIT.UOMDEC_0 AS NUMDEC, PILOT.ITMMASTER.ZWEBTITLE_0 AS TITLE
                FROM            PILOT.SORDER INNER JOIN
                                         PILOT.SORDERP ON PILOT.SORDER.SOHNUM_0 = PILOT.SORDERP.SOHNUM_0 INNER JOIN
                                         PILOT.SORDERQ ON PILOT.SORDERP.SOHNUM_0 = PILOT.SORDERQ.SOHNUM_0 AND PILOT.SORDERP.SOPLIN_0 = PILOT.SORDERQ.SOPLIN_0 INNER JOIN
                                         PILOT.ITMMASTER ON PILOT.SORDERP.ITMREF_0 = PILOT.ITMMASTER.ITMREF_0 INNER JOIN
                                         PILOT.TABUNIT ON PILOT.SORDERP.SAU_0 = PILOT.TABUNIT.UOM_0 LEFT OUTER JOIN
                                         PILOT.TEXCLOB ON PILOT.SORDERQ.SOQTEX_0 = PILOT.TEXCLOB.CODE_0
                WHERE        (PILOT.SORDER.SOHNUM_0 = N'***OrderNumber***')
                ORDER BY PILOT.SORDERP.SOPLIN_0
            ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
            Dim OrderDetails As DataTable = OpenDataSet(SQLStatement)

            ' Filling in the product lines
            SectionError = "ORDCON - Filling in the product lines"
            Dim Details As String = ""
            Dim SubTotal As Decimal = 0
            Dim NUMDEC As Integer = 0
            Dim LeadTimes As DataTable
            Dim LeadTimeSupplier As DataTable
            Dim BODate As Date
            Dim LeadTimeDays As Integer = 0
            Dim ITMREFNUM As String
            Dim intTabLevel As Integer = 11

            If OrderDetails.Rows.Count > 0 Then
                For i = 0 To OrderDetails.Rows.Count - 1
                    SectionError = "ORDCON - Filling in the product lines 1"
                    NUMDEC = OrderDetails.Rows(i).Item("NUMDEC")
                    Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                    intTabLevel += 1
                    Details += StrDup(intTabLevel, vbTab) + "<td style=""font-size: 13px; font-family: sans-serif; border-bottom:  1px solid #ccc;"">" + vbCrLf
                    intTabLevel += 1
                    Details += StrDup(intTabLevel, vbTab) + "<strong>" + OrderDetails.Rows(i).Item("TITLE") + "</strong><br><br>" + vbCrLf
                    Details += StrDup(intTabLevel, vbTab) + "Item #: " + OrderDetails.Rows(i).Item("ITMREF_0") + vbCrLf
                    SectionError = "ORDCON - Filling in the product lines 2"
                    Details += StrDup(intTabLevel, vbTab) + "UOM: " + OrderDetails.Rows(i).Item("UOM").ToString() + "<br>" + vbCrLf
                    intTabLevel -= 1
                    Details += StrDup(intTabLevel, vbTab) + "</td>"
                    Details += StrDup(intTabLevel, vbTab) + "<td align=""center"" style=""font-size: 13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">"
                    Details += Math.Round(OrderDetails.Rows(i).Item("QTY_0"), NUMDEC).ToString() + "</td>" + vbCrLf
                    'SectionError = "ORDCON - Filling in the product lines 1"
                    'Details += vbTab + vbTab + "<td align='center'>" + Math.Round(OrderDetails.Rows(i).Item("SHIQTY"), NUMDEC).ToString() + "</center></td>" + vbCrLf
                    'Details += vbTab + vbTab + "<td align=""center"" style=""font-size: 13px; font-family: sans-serif; border-bottom:  1px solid #ccc;"">"

                    SectionError = "ORDCON - Filling in the product lines 3"
                    If Math.Round(OrderDetails.Rows(i).Item("BOQTY")) > 0 Then
                        LeadTimeDays = 0

                        SectionError = "ORDCON - SQL Statement to pull the backorder information "
                        SQLStatement = <Sql><![CDATA[ SELECT PILOT.ITMFACILIT.ITMREF_0,PILOT.ITMFACILIT.OFS_0 FROM PILOT.ITMFACILIT 
                                                      WHERE PILOT.ITMFACILIT.ITMREF_0 ='***ITMREF***' and PILOT.ITMFACILIT.STOFCY_0='***STOFCY***'  ]]></Sql>.Value

                        ITMREFNUM = OrderDetails.Rows(i).Item("ITMREFNUM").ToString

                        SQLStatement = Replace(SQLStatement, "***ITMREF***", ITMREFNUM)
                        SQLStatement = Replace(SQLStatement, "***STOFCY***", OrderDetails.Rows(i).Item("STOFCY_0".ToString()))
                        LeadTimes = OpenDataSet(SQLStatement)
                        LeadTimeDays = LeadTimes.Rows(0).Item("OFS_0")

                        If LeadTimeDays = 0 Then
                            SectionError = "ORDCON - SQL Statement to pull the backorder information for supplier lead "
                            SQLStatement = <Sql><![CDATA[   SELECT        PILOT.BPSUPPLIER.ZAVGLEAD_0
                                                            FROM            PILOT.ITMBPS INNER JOIN
                                                            PILOT.BPSUPPLIER ON PILOT.ITMBPS.BPSNUM_0 = PILOT.BPSUPPLIER.BPSNUM_0
                                                            where PILOT.ITMBPS.PIO_0 ='10' and PILOT.ITMBPS.ITMREF_0='***ITMREF***'  ]]></Sql>.Value
                            SQLStatement = Replace(SQLStatement, "***ITMREF***", ITMREFNUM)
                            LeadTimeSupplier = OpenDataSet(SQLStatement)
                            LeadTimeDays = LeadTimeSupplier.Rows(0).Item("ZAVGLEAD_0")
                        End If

                        SectionError = "ORDCON - Calculating backorder date " + SQLStatement
                        BODate = DateAdd(DateInterval.Day, (1 + LeadTimeDays), Date.Today)
                        If BODate.DayOfWeek = DayOfWeek.Saturday Then
                            BODate = BODate.AddDays(2)
                        ElseIf BODate.DayOfWeek = DayOfWeek.Sunday Then
                            BODate = BODate.AddDays(1)
                        End If
                        If BODate > DateValue(OrderDetails.Rows(i).Item("DEMDLVDAT_0").ToString) Then
                            Details += vbTab + vbTab + "<td width=""75"" align=""center"" style=""font-size: 13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">"
                            Details += "<p style=""background:#fff3cd;color:#856404;padding: 5px 0;"">" + "Estimated Ship<br/>" + BODate.ToString("MM/dd/yyyy") + "</p></td>" + vbCrLf
                        Else
                            Details += vbTab + vbTab + "<td width=""75"" align=""center"" style=""font-size: 13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">"
                            Details += "<p style=""background:#fff3cd;color:#856404;padding: 5px 0;"">" + "Estimated Ship<br/>"
                            Details += DateValue(OrderDetails.Rows(i).Item("DEMDLVDAT_0")).ToString("MM/dd/yyyy") + "</p></td>" + vbCrLf
                        End If
                    Else
                        Details += vbTab + vbTab + "<td width=""75"" align=""center"" style=""font-size: 13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">"
                        Details += "<p style=""background:#cfeddc;color:#298250;padding: 5px 0;"">In Stock<br />Will Ship<br/>"
                        Details += DateValue(OrderDetails.Rows(i).Item("DEMDLVDAT_0")).ToString("MM/dd/yyyy") + "</p></td>" + vbCrLf
                    End If
                    SectionError = "ORDCON - Filling in the product lines 4"
                    Details += vbTab + vbTab + "<td align='center' style=""font-family: arial,sans-serif; font-size: 13px; border-bottom: 1px solid #ccc;"">"
                    Details += CType(Math.Round(OrderDetails.Rows(i).Item("NETPRI_0"), 4), Double).ToString("C2") + "</td>" + vbCrLf
                    SectionError = "ORDCON - Filling in the product lines 5"
                    Details += vbTab + vbTab + "<td align='center' style=""font-family: arial,sans-serif; font-size: 13px; border-bottom: 1px solid #ccc;"">"
                    Details += CType(Math.Round(OrderDetails.Rows(i).Item("NETPRI_0") * OrderDetails.Rows(i).Item("QTY_0"), 2), Double).ToString("C2") + "</td>" + vbCrLf
                    SectionError = "ORDCON - Filling in the product lines 6"
                    Details += vbTab + "</tr>" + vbCrLf
                    SubTotal += OrderDetails.Rows(i).Item("NETPRI_0") * OrderDetails.Rows(i).Item("QTY_0")
                    If OrderDetails.Rows(i).Item("LineText") <> "" Then
                        Me.RichTextBox1.Text = ""
                        If OrderDetails.Rows(i).Item("LineText").ToString().Length > 1 Then
                            If OrderDetails.Rows(i).Item("LineText").ToString().Substring(0, 1) = "{" Then
                                Me.RichTextBox1.Rtf = OrderDetails.Rows(i).Item("LineText")
                            Else
                                Me.RichTextBox1.Text = OrderDetails.Rows(i).Item("LineText")
                            End If
                        End If
                        Details += vbTab + "<tr style='height: 50px; width:75px; line-height: 15px; font-size: 12px; color: #414042; font-family: Arial,sans-serif; font-weight: normal; padding-bottom: 6px;'>" + vbCrLf
                        Details += vbTab + vbTab + "<td colspan = 8>" + Me.RichTextBox1.Text + "</td>" + vbCrLf
                        Details += vbTab + "</tr>" + vbCrLf
                    End If
                Next
            End If

            ' SubTotal
            SectionError = "ORDCON - SubTotal"
            Body = Replace(Body, "***SubTotal***", CType(Math.Round(SubTotal, 2), Double).ToString("C2"))

            ' Delivery Fee
            SectionError = "ORDCON - Delivery Fee"
            If Order.Rows(0).Item("MINORD") > 0 Then
                strbuffer = "<tr> <td align=""left"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += "Delivery Fee</td><td align=""right"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += CType(Math.Round(Order.Rows(0).Item("MINORD"), 2), Double).ToString("C2") + "</td></tr>"
                Body = Replace(Body, "***DeliveryFee***", strbuffer)
            Else
                Body = Replace(Body, "***DeliveryFee***", "")
            End If

            ' Handling Fee
            SectionError = "ORDCON - Handling Fee"
            If Order.Rows(0).Item("HANDLFEE") > 0 Then
                strbuffer = "<tr> <td align=""left"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += "Handling Fee</td><td align=""right"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += CType(Math.Round(Order.Rows(0).Item("HANDLFEE"), 2), Double).ToString("C2") + "</td></tr>"
                Body = Replace(Body, "***HandleFee***", strbuffer)
            Else
                Body = Replace(Body, "***HandleFee***", "")
            End If

            ' Freight
            SectionError = "ORDCON - Freight"
            If Order.Rows(0).Item("INVDTAAMT_1") > 0 Then
                strbuffer = "<tr> <td align=""left"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += "Freight</td><td align=""right"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += CType(Math.Round(Order.Rows(0).Item("INVDTAAMT_1"), 2), Double).ToString("C2") + "</td></tr>"
                Body = Replace(Body, "***Freight***", strbuffer)
            Else
                Body = Replace(Body, "***Freight***", "")
            End If

            ' Tax
            SectionError = "ORDCON - Tax"
            If Order.Rows(0).Item("ORDINVATI_0") - Order.Rows(0).Item("ORDINVNOT_0") > 0 Then
                strbuffer = "<tr> <td align=""left"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += "Tax</td><td align=""right"" style=""font-size:   13px; font-family: sans-serif;"">"
                strbuffer += CType(Math.Round(Order.Rows(0).Item("ORDINVATI_0"), 2) - Math.Round(Order.Rows(0).Item("ORDINVNOT_0"), 2), Double).ToString("C2") + "</td></tr>"
                Body = Replace(Body, "***Tax***", strbuffer)
            Else
                Body = Replace(Body, "***Tax***", "")
            End If

            ' Total
            SectionError = "ORDCON - Total"
            Body = Replace(Body, "***Total***", CType(Math.Round(Order.Rows(0).Item("ORDINVATI_0"), 2), Double).ToString("C2"))

            ' Add Details to the HTML
            SectionError = "ORDCON - Add Details to the HTML"
            Body = Replace(Body, "***LINES***", Details)

            ' Constructing the ToAddress of the email
            SectionError = "ORDCON - Constructing the ToAddress of the email"
            Dim EmailTo As String = ""
            If Len(Order.Rows(0).Item("WEB_0")) > 1 Then
                EmailTo = EmailTo + Order.Rows(0).Item("WEB_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML1_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML1_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML1_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML2_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML2_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML2_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML3_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML3_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML3_0") + ";"
            End If
            If Len(Order.Rows(0).Item("CNTOAEML4_0")) > 1 And Not EmailTo.Contains(Order.Rows(0).Item("CNTOAEML4_0")) Then
                EmailTo = EmailTo + Order.Rows(0).Item("CNTOAEML4_0") + ";"
            End If

            Dim BCC As String = ""
            If Len(CreateUserEmail) > 2 Then
                BCC = CreateUserEmail
            End If
            If Len(CreateUserEmail) > 2 And Len(Order.Rows(0).Item("CSREML_0")) > 2 Then
                BCC = BCC + ";"
            End If
            If Len(Order.Rows(0).Item("CSREML_0")) > 2 Then

                BCC = BCC + Order.Rows(0).Item("CSREML_0")
            End If
            BCC += ";cdreyer@packbgr.com"
            'If Len(Order.Rows(0).Item("REPEML_0")) > 2 Then
            'BCC = BCC + ";"
            'BCC = BCC + Order.Rows(0).Item("REPEML_0")
            'End If

            'Me.RichTextBox1.Text = Body

            ' Sends the Email "BGR <do-not-reply@bgr.us>"
            SectionError = "ORDCON - Sends the Email"
            If Len(Order.Rows(0).Item("CSREML_0")) > 1 And Len(EmailTo) > 1 Then
                EmailTo = EmailTo.Substring(0, Len(EmailTo) - 1)
                If CreateUserDept = 10 Then
                    SendHTMLEmail(EmailTo, CreateUserEmail, "Order Confirmation for Order Number " + OrderNumber, Body, "", BCC)
                    'SendHTMLEmail("cdreyer@packbgr.com", "cdreyer@packbgr.com", "Order Confirmation", Body, "", "cdreyer@packbgr.com")
                    Console.WriteLine("Email sent")
                    If blnDebugFlag Then : Console.WriteLine(Body) : End If
                Else
                    SendHTMLEmail(EmailTo, Order.Rows(0).Item("CSREML_0"), "Order Confirmation For Order Number " + OrderNumber, Body, "", BCC)
                    'SendHTMLEmail("cdreyer@packbgr.com", "cdreyer@packbgr.com", "Order Confirmation", Body, "", "cdreyer@packbgr.com")
                    Console.WriteLine("Email sent")
                    If blnDebugFlag Then : Console.WriteLine(Body) : End If
                End If
            End If
            'SendHTMLEmail("msullivan@razorsoft.biz", "BGR <do-not-reply@bgr.us>", "Order Confirmation", EmailTo + "<BR>" + Body)
        Else
            ' This is what happens if the order doesn't exist
            ExecuteSQLQuery("UPDATE PILOT.ZORDCON SET STA_0 = 3 WHERE SOHNUM_0 = '" + OrderNumber + "'")
            SendHTMLEmail("tbailey@packbgr.com;cdreyer@packbgr.com;", "BGR <do-not-reply@bgr.us>", "ERROR - Order Confirmation cannot Find Order - " + OrderNumber, "ERROR - Order Confirmation cannot Find Order - " + OrderNumber)
            'SendHTMLEmail("msullivan@razorsoft.biz", "BGR <do-not-reply@bgr.us>", "ERROR - Order Confirmation cannot Find Order - " + OrderNumber, "ERROR - Order Confirmation cannot Find Order - " + OrderNumber)
        End If

    End Sub

    Public Sub Look4DeliveryConfirmations()
        Try
            ' SQL Statement to pull any unsent delivery confirmations
            Dim SQLStatement As String = <Sql><![CDATA[
						SELECT DISTINCT
	                        z.SOHNUM_0
                            ,z.SDHNUM_0
	                        ,z.CREUSR_0
	                        ,z.MDL_0
	                        ,z.TRKNUM_0
	                        ,z.SHIDAT_0
	                        ,d.DLVDAT_0
							,z.STAT_0
                        FROM  PILOT.ZDELCON z
                        LEFT JOIN PILOT.SDELIVERY d ON z.SDHNUM_0 = d.SDHNUM_0
                        LEFT JOIN PILOT.SORDER o ON z.SOHNUM_0 = o.SOHNUM_0
						LEFT JOIN PILOT.CONTACT n ON o.ZCONTACT_0 = n.CCNCRM_0
                        WHERE z.STAT_0 = '1' 
                                and z.SHIDAT_0 <= GETDATE() 
                                and d.BETFCY_0 = 1 
                                and o.ZCONTACT_0 <> ''
								and n.WEB_0 <> ''
                                and d.BPCINV_0 not IN('13758', '11488')
		                        and 
		                        ((z.MDL_0 IN('MAIN','COUR'
                                            ,'FF2','FFE','FFP'
                                            ,'LTL','PIN'
                                            ,'SALE'))
			                        OR (z.MDL_0 IN('F3D','FE2','FEP','FFO','FGH','FGR','FIE','FIP','FSO'
                                            ,'GRN'
                                            ,'U2AM','U2DA','U3DA','UGRN','UNAM','UNDA','UNDS','UPS','UST','UWXP') 
				                        and LEN(z.TRKNUM_0) > 1)
			                        OR (z.MDL_0 = 'BGR' and d.DLVDAT_0 <= CONVERT(date, getdate())))
                        ORDER BY z.SOHNUM_0
                        ]]></Sql>.Value
            Dim DeliveryConfirmations As DataTable = OpenDataSet(SQLStatement)
            Dim strOrdersSent As String = ""
            Dim strDebug As String = ""
            If DeliveryConfirmations.Rows.Count > 0 Then
                ' Loops through all the unsent records
                For i = 0 To DeliveryConfirmations.Rows.Count - 1
                    'Prevent multiple emails if 1+ shipments are status 1 for same SO
                    If InStr(strOrdersSent, DeliveryConfirmations.Rows(i).Item("SOHNUM_0")) > 0 Then
                    Else
                        ' Sets a variable with the order number
                        OrderNumber = DeliveryConfirmations.Rows(i).Item("SOHNUM_0")
                        CreateUser = DeliveryConfirmations.Rows(i).Item("CREUSR_0")
                        ' Sends the order confirmation
                        If SendDeliveryConfirmation(OrderNumber) Then
                            ' Sets the status to sent
                            ExecuteSQLQuery("UPDATE PILOT.ZDELCON SET STAT_0 = '2', SENDAT_0 = GETDATE() WHERE SOHNUM_0 = '" + OrderNumber + "' AND STAT_0 = '1'")
                            strOrdersSent += OrderNumber
                        Else
                            ExecuteSQLQuery("UPDATE PILOT.ZDELCON SET STAT_0 = '3', SENDAT_0 = GETDATE() WHERE SOHNUM_0 = '" + OrderNumber + "' AND STAT_0 = '1'")
                        End If
                    End If
                Next
                Console.WriteLine(strDebug)
            End If
            Console.WriteLine("DELCON Try Complete")
        Catch ex As Exception
            ' If the code encounters an error this sets the status to 3 and sends an email about the error
            ExecuteSQLQuery("UPDATE PILOT.ZDELCON SET STAT_0 = '3' WHERE SOHNUM_0 = '" + OrderNumber + "'")
            Console.WriteLine("ERROR - BIG LINES AT THE TOP")
            Console.WriteLine(ex.ToString())
            Console.WriteLine(SectionError)
            SendHTMLEmail("cdreyer@packbgr.com", "BGR <do-not-reply@packbgr.com>", "Delivery - ERROR - BIG LINES AT THE TOP; SOHNUM: " + OrderNumber, ex.ToString())
        End Try
    End Sub

    Private Function SendDeliveryConfirmation(ByVal OrderNumber As String) As Boolean

        Dim SQLStatement As String
        Dim Delivery As DataTable
        Dim CreateUserDT As DataTable
        Dim Footer As DataTable
        Dim PartialShip As DataTable
        Dim CreateUserEmail As String
        Dim CreateUserDept As Integer
        Dim FileReader As StreamReader
        Dim Body As String
        Dim RichTextBox1 As Object = Nothing
        Dim blnPartialShip As Boolean = False
        Dim strShipMethod As String = ""
        Dim intTabLevel As Integer = 14 ' Tracks indents of HTML code

        SendDeliveryConfirmation = False

        ' Pull Shipment Header Info
        SectionError = "DELCON - SQL Statement to pull the header information of the delivery confirmation"
        SQLStatement = <Sql><![CDATA[
                SELECT s.SOHNUM_0 As SalesOrderNum
	                ,s.SDHNUM_0 As ShipmentNum
	                ,o.CUSORDREF_0 As CustPONum
	                ,s.BPCORD_0 As BPNum
	                ,s.BPDNAM_0 As ShipToName1
	                ,s.BPDNAM_1 As ShipToName2
	                ,s.BPDADDLIG_0 As ShipToAdd1
	                ,s.BPDADDLIG_1 As ShipToAdd2
	                ,s.BPDADDLIG_2 As ShipToAdd3
	                ,s.BPDCTY_0 As ShipToCity
	                ,s.BPDSAT_0 As ShipToState
	                ,s.BPDPOSCOD_0 As ShipToZip
	                ,s.BPINAM_0 As BillToName1
	                ,s.BPINAM_1 As BillToName2
	                ,s.BPIADDLIG_0 As BillToAdd1
	                ,s.BPIADDLIG_1 As BillToAdd2
	                ,s.BPIADDLIG_2 As BillToAdd3
	                ,s.BPICTY_0 As BillToCity
	                ,s.BPISAT_0 As BillToState
	                ,s.BPIPOSCOD_0 As BillToZip
	                ,c.BPTNAM_0 As ShipMethod
	                ,s.MDL_0
	                ,o.STOFCY_0 As Facility
	                ,n.WEB_0 As Email
	                ,o.BPTNUM_0 
	                ,n.CNTOAEML1_0
	                ,n.CNTOAEML2_0
	                ,n.CNTOAEML3_0
	                ,n.CNTOAEML4_0
	                ,n.WEB_0
                FROM PILOT.SORDER o 
                LEFT JOIN PILOT.SDELIVERY s ON o.SOHNUM_0 = s.SOHNUM_0 
                LEFT JOIN PILOT.BPCARRIER c ON s.MDL_0 = c.BPTNUM_0 
                INNER JOIN PILOT.CONTACT n ON o.BPCORD_0 = n.BPANUM_0 and o.ZCONTACT_0 = n.CCNCRM_0
                LEFT JOIN PILOT.BPCARRIER b ON o.BPTNUM_0 = b.BPTNUM_0
                WHERE o.SOHNUM_0 = '***OrderNumber***'
	                And o.STOFCY_0 IN('BGR','IND','LVL','WVA', 'DET')
                ]]></Sql>.Value
        SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
        Delivery = OpenDataSet(SQLStatement)

        ' Pull email footer info
        SQLStatement = <Sql><![CDATA[
                SELECT b.USR_0 
                    ,a.REPNAME_0
                    ,a.REPPHONE_0 
                    ,a.REPEML_0 
                    ,b.REPNAME_0 as CSRNAME_0
                    ,b.REPPHONE_0 as CSRPHONE_0
                    ,b.REPEML_0 as CSREML_0 
                FROM PILOT.SORDER s 
                INNER JOIN PILOT.ZBGRREPS a ON s.REP_0 = a.USR_0 
                INNER JOIN PILOT.ZBGRREPS b ON s.ZCSR_0 = b.USR_0 
                WHERE s.SOHNUM_0 = '***OrderNumber***'
                ]]></Sql>.Value
        SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
        Footer = OpenDataSet(SQLStatement)

        ' Override web order creator with CSR on order
        If CreateUser = "PIP" Then
            CreateUser = Footer.Rows(0).Item("USR_0")
        End If

        ' Pull email delivery info
        SQLStatement = <Sql><![CDATA[
                SELECT REPEML_0
                    ,DEPARTMENT_0 
                    ,CSREML_0
                FROM PILOT.ZBGRREPS 
                WHERE USR_0 = '***CreateUser***'
                ]]></Sql>.Value
        SQLStatement = Replace(SQLStatement, "***CreateUser***", CreateUser)
        CreateUserDT = OpenDataSet(SQLStatement)
        CreateUserEmail = CreateUserDT.Rows(0).Item("REPEML_0")
        CreateUserDept = CreateUserDT.Rows(0).Item("DEPARTMENT_0")



        If Delivery.Rows.Count > 0 Then

            ' Opens the html template file and sets the entire file as a string "Body"
            SectionError = "DELCON - Opens the html template file and sets the entire file as a string Body"
            FileReader = New StreamReader("C:\Program Files (x86)\WebConfirmation\DeliveryConfirmation.html")
            Body = FileReader.ReadToEnd
            FileReader.Close()

            ' Set program flow variables
            strShipMethod = GetShipMethod(Delivery.Rows(0).Item("MDL_0"))
            SQLStatement = <Sql><![CDATA[
                    SELECT s.SOHNUM_0 As SalesOrderNum
	                    ,COUNT(s.SDHNUM_0) As ShipmentCount
                    FROM PILOT.SORDER o 
                    LEFT JOIN PILOT.SDELIVERY s ON o.SOHNUM_0 = s.SOHNUM_0 
                    WHERE o.SOHNUM_0 = '***OrderNumber***'
                    GROUP BY s.SOHNUM_0
                    ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
            PartialShip = OpenDataSet(SQLStatement)
            If PartialShip.Rows(0).Item("ShipmentCount") > 1 Then
                blnPartialShip = True
            Else
                blnPartialShip = False
            End If

            ' Inject Header information into the template with data from the SQL statement
            SectionError = "DELCON - Replaces the Header information in the template with data from the SQL statement"
            Body = Replace(Body, "***OrderNumber***", OrderNumber)
            Body = Replace(Body, "***PONumber***", Delivery.Rows(0).Item("CustPONum"))
            If strShipMethod = "CUST" Then
                If blnPartialShip Then
                    Body = Replace(Body, "***Title***", "Part of your order is ready for pick up")
                Else
                    Body = Replace(Body, "***Title***", "Your order is ready for pick up")
                End If
            ElseIf strShipMethod = "SALE" Then
                If blnPartialShip Then
                    Body = Replace(Body, "***Title***", "Part of your order has been picked up by your Sales Rep")
                Else
                    Body = Replace(Body, "***Title***", "Your order has been picked up by your Sales Rep")
                End If
            Else
                If blnPartialShip Then
                    Body = Replace(Body, "***Title***", "Part of your order has shipped")
                Else
                    Body = Replace(Body, "***Title***", "Your order has shipped")
                End If
            End If

            Body = Replace(Body, "***ShipMethod***", GetShipText(Delivery.Rows(0).Item("MDL_0"), Delivery.Rows(0).Item("BPTNUM_0")))

            If Footer.Rows.Count > 0 Then
                If Footer.Rows(0).Item("REPNAME_0") = "Inside sales" Then
                    Body = Replace(Body, "***REPNAME***", "Inside Sales Team")
                    Body = Replace(Body, "***REPEMAIL***", Footer.Rows(0).Item("REPEML_0"))
                    Body = Replace(Body, "***REPPHONE***", "513-759-8428")
                    Body = Replace(Body, "***CSRNAME***", Footer.Rows(0).Item("CSRNAME_0"))
                    Body = Replace(Body, "***CSREMAIL***", Footer.Rows(0).Item("CSREML_0"))
                    Body = Replace(Body, "***CSRPHONE***", Mid(Footer.Rows(0).Item("CSRPHONE_0"), 1, 3) _
                                           & "-" + Mid(Footer.Rows(0).Item("CSRPHONE_0"), 4, 3) _
                                           & "-" + Mid(Footer.Rows(0).Item("CSRPHONE_0"), 7, 4))
                Else
                    Body = Replace(Body, "***REPNAME***", Footer.Rows(0).Item("REPNAME_0"))
                    Body = Replace(Body, "***REPEMAIL***", Footer.Rows(0).Item("REPEML_0"))
                    Body = Replace(Body, "***REPPHONE***", Mid(Footer.Rows(0).Item("REPPHONE_0"), 1, 3) _
                                               & "-" + Mid(Footer.Rows(0).Item("REPPHONE_0"), 4, 3) _
                                               & "-" + Mid(Footer.Rows(0).Item("REPPHONE_0"), 7, 4))
                    Body = Replace(Body, "***CSRNAME***", Footer.Rows(0).Item("CSRNAME_0"))
                    Body = Replace(Body, "***CSREMAIL***", Footer.Rows(0).Item("CSREML_0"))
                    Body = Replace(Body, "***CSRPHONE***", Mid(Footer.Rows(0).Item("CSRPHONE_0"), 1, 3) _
                                           & "-" + Mid(Footer.Rows(0).Item("CSRPHONE_0"), 4, 3) _
                                           & "-" + Mid(Footer.Rows(0).Item("CSRPHONE_0"), 7, 4))
                End If
            Else
                SendDeliveryConfirmation = False
                Exit Function
            End If

            If strShipMethod = "Parcel" Then
                Dim strFoot As String = "<tr>" + vbCrLf
                strFoot += "<td align=""left"" style=""font-size: 13px; font-family: sans-serif;""><strong>UPS/FedEx Tracking</strong><br><br>" + vbCrLf
                strFoot += "The tracking # listed may not be accessible for up to 24 to 48 hours, even though your package has already shipped from our warehouse and is on its way to you.</td>" + vbCrLf
                strFoot += "</tr>" + vbCrLf
                Body = Replace(Body, "***Footer***", strFoot)
            Else
                Body = Replace(Body, "***Footer***", "")
            End If

            ' BillTo
            SectionError = "DELCON - BillTo"
            Dim BillTo As String = ""
            If Len(Delivery.Rows(0).Item("BillToName1")) > 2 Then BillTo += Delivery.Rows(0).Item("BillToName1")
            If Len(Delivery.Rows(0).Item("BPNum")) > 2 Then BillTo += " (Cust #: " + Delivery.Rows(0).Item("BPNum") + ")" + "<BR>" + vbCrLf
            If Len(Delivery.Rows(0).Item("BillToName2")) > 2 Then BillTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("BillToName2") + "<BR>" + vbCrLf
            If Len(Delivery.Rows(0).Item("BillToAdd1")) > 2 Then BillTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("BillToAdd1") + "<BR>" + vbCrLf
            If Len(Delivery.Rows(0).Item("BillToAdd2")) > 2 Then BillTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("BillToAdd2") + "<BR>" + vbCrLf
            If Len(Delivery.Rows(0).Item("BillToAdd3")) > 2 Then BillTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("BillToAdd3") + "<BR>" + vbCrLf
            BillTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("BillToCity") + ", " + Delivery.Rows(0).Item("BillToState") + " " + Delivery.Rows(0).Item("BillToZip")
            Body = Replace(Body, "***BillTo***", BillTo)

            ' ShipTo
            If strShipMethod <> "CUST" Then
                SectionError = "DELCON - ShipTo"
                Dim ShipTo As String = ""
                If Len(Delivery.Rows(0).Item("ShipToName1")) > 2 Then ShipTo += Delivery.Rows(0).Item("ShipToName1") + "<BR>" + vbCrLf
                If Len(Delivery.Rows(0).Item("ShipToName2")) > 2 Then ShipTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("ShipToName2") + "<BR>" + vbCrLf
                If Len(Delivery.Rows(0).Item("ShipToAdd1")) > 2 Then ShipTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("ShipToAdd1") + "<BR>" + vbCrLf
                If Len(Delivery.Rows(0).Item("ShipToAdd2")) > 2 Then ShipTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("ShipToAdd2") + "<BR>" + vbCrLf
                If Len(Delivery.Rows(0).Item("ShipToAdd3")) > 2 Then ShipTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("ShipToAdd3") + "<BR>" + vbCrLf
                ShipTo += StrDup(intTabLevel, vbTab) + Delivery.Rows(0).Item("ShipToCity") + ", " + Delivery.Rows(0).Item("ShipToState") + " " + Delivery.Rows(0).Item("ShipToZip")
                Body = Replace(Body, "***ShipTo***", ShipTo)
            Else
                Body = Replace(Body, "***ShipTo***", "BGR - Customer Pick Up<BR>6392 Gano Road<BR>West Chester, OH 45069")
                Body = Replace(Body, "<strong>Delivery Address</strong>", "<strong>Pick Up Address</strong>")
            End If

            ' Load items shipping today
            SectionError = "DELCON Today - SQL Statement to pull the product line information - shipping today"
            SQLStatement = <Sql><![CDATA[
                                SELECT DISTINCT 
	                                dd.SOHNUM_0
	                                ,dd.SDHNUM_0 As ShipmentNumber_Unshipped
	                                ,dd.ITMREF_0 As ItemRef
	                                ,i.ZWEBTITLE_0 As Title
	                                ,Case
		                                WHEN d.MDL_0 = 'BGR' or o.DLVPIO_0 = 2 and d.DLVDAT_0 < GETDATE() THEN GETDATE()
		                                WHEN d.MDL_0 = 'BGR' or o.DLVPIO_0 = 2 and d.DLVDAT_0 >= GETDATE() THEN d.DLVDAT_0
		                                WHEN d.MDL_0 <> 'BGR' THEN GETDATE()
		                                ELSE GETDATE()
	                                END As ShipDate
	                                ,d.SHIDAT_0 
	                                ,d.DLVDAT_0
	                                ,dd.QTY_0 As ShipQty
	                                ,q.QTY_0 As OrderQty
	                                ,dd.SAU_0 As UOM
                                    ,u.UOMDEC_0 AS NUMDEC
	                                ,o.DLVPIO_0 As SameDayIs2
                                FROM PILOT.SDELIVERYD dd
                                INNER JOIN PILOT.SDELIVERY d ON dd.SDHNUM_0 = d.SDHNUM_0
                                INNER JOIN PILOT.ZDELCON z ON dd.SOHNUM_0 = z.SOHNUM_0 and dd.SDHNUM_0 = z.SDHNUM_0
                                INNER JOIN PILOT.ITMMASTER i ON dd.ITMREF_0 = i.ITMREF_0
                                LEFT JOIN PILOT.SORDERQ q ON dd.SOHNUM_0 = q.SOHNUM_0 
							                                and dd.SOQSEQ_0 = q.SOQSEQ_0
							                                and dd.SOPLIN_0 = q.SOPLIN_0
                                LEFT JOIN PILOT.SORDER o ON dd.SOHNUM_0 = o.SOHNUM_0
                                LEFT JOIN PILOT.TABUNIT u ON dd.SAU_0 = u.UOM_0
                                WHERE z.STAT_0 = '1' 
                                AND dd.SOHNUM_0 = '***OrderNumber***'                                
                                AND Case
		                                WHEN d.MDL_0 = 'BGR' or o.DLVPIO_0 = 2 and d.DLVDAT_0 < GETDATE() THEN GETDATE()
		                                WHEN d.MDL_0 = 'BGR' or o.DLVPIO_0 = 2 and d.DLVDAT_0 >= GETDATE() THEN d.DLVDAT_0
		                                WHEN d.MDL_0 <> 'BGR' THEN GETDATE()
		                                ELSE GETDATE()
	                                END <= GETDATE()
            ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
            Dim DeliveryDetails As DataTable = OpenDataSet(SQLStatement)

            SectionError = "DELCON Today - Filling in the product lines"
            Dim Details As String = ""
            Dim NUMDEC As Integer = 0
            Dim LeadTimeDays As Integer = 0
            Dim blnDebugFlag As Boolean = True

            If DeliveryDetails.Rows.Count > 0 Then
                Details += "<table width=""575"" bgcolor=""#ffffff"" border=""0"" cellpadding=""10"" cellspacing=""0"" style=""border: 1px solid #ddd;margin: 0 auto;"">" + vbCrLf
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                intTabLevel = 12
                ' Adjust header depending on ship method
                Details += LoadTableHeader(strShipMethod, intTabLevel, "Today")
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                For i = 0 To DeliveryDetails.Rows.Count - 1
                    SectionError = "DELCON Today - Filling in the product lines 1"
                    NUMDEC = DeliveryDetails.Rows(i).Item("NUMDEC")
                    intTabLevel = 11
                    Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "<td style="" font-size:13px; font-family: sans-serif; border-bottom:  1px solid #ccc;"">" + vbCrLf
                    intTabLevel = 13
                    Details += StrDup(intTabLevel, vbTab) + "<strong>" + DeliveryDetails.Rows(i).Item("TITLE") + "</strong><br><br>" + vbCrLf
                    SectionError = "DELCON Today - Filling in the product lines 2"
                    Details += StrDup(intTabLevel, vbTab) + "Item #: " + DeliveryDetails.Rows(i).Item("ItemRef") + "  "
                    Details += "UOM: " + DeliveryDetails.Rows(i).Item("UOM").ToString() + "<br>"
                    Details += "Shipment #: " + DeliveryDetails.Rows(i).Item("ShipmentNumber_Unshipped").ToString() + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                    Details += StrDup(intTabLevel, vbTab) + "<td align=""center"" style=""font-size:13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">"
                    ' Tracking number for parcel, ship date for others
                    If strShipMethod = "Parcel" Then
                        Details += GetTrackingHyperlink(OrderNumber, DeliveryDetails.Rows(i).Item("ShipmentNumber_Unshipped")) + vbCrLf
                    ElseIf strShipMethod = "SALE" Then
                        Details += "Picked Up" + "</td>" + vbCrLf
                    ElseIf strShipMethod = "CUST" Then
                        Details += "Ready for<br>Pick Up" + "</td>" + vbCrLf
                    Else
                        Dim strBuffer As String = DatePart("m", DeliveryDetails.Rows(i).Item("ShipDate")).ToString() + "/" _
                                                + DatePart("d", DeliveryDetails.Rows(i).Item("ShipDate")).ToString() + "/" _
                                                + Mid(DatePart("yyyy", DeliveryDetails.Rows(i).Item("ShipDate")).ToString(), 3, 2)
                        Details += strBuffer + "</td>" + vbCrLf
                    End If
                    Details += StrDup(intTabLevel, vbTab) + "<td align=""center"" style=""font-size:13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">" + vbCrLf
                    intTabLevel = 13
                    ' Change Color when partial line ship
                    If Math.Round(DeliveryDetails.Rows(i).Item("ShipQty"), NUMDEC) <> Math.Round(DeliveryDetails.Rows(i).Item("OrderQty"), NUMDEC) Then
                        Details += StrDup(intTabLevel, vbTab) + "<p style=""background:#FFF2CE;color:#856311;padding:5px 0;"">" + vbCrLf
                    Else
                        Details += StrDup(intTabLevel, vbTab) + "<p style=""background:#cfeddc;color:#298250;padding:5px 0;"">"
                    End If
                    Details += Math.Round(DeliveryDetails.Rows(i).Item("ShipQty"), NUMDEC).ToString() + " of "
                    Details += Math.Round(DeliveryDetails.Rows(i).Item("OrderQty"), NUMDEC).ToString() + "</p>" + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                    intTabLevel = 11
                    Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                Next
                intTabLevel = 10
                Details += StrDup(intTabLevel, vbTab) + "</table>" + vbCrLf
                intTabLevel = 9
                Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                intTabLevel = 8
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
            Else
                'Shipment was deleted?
                SendDeliveryConfirmation = False
                Exit Function
            End If

            ' Load items awaiting shipment
            SectionError = "DELCON Future - SQL Statement to pull the product line information - awaiting shipment"
            SQLStatement = <Sql><![CDATA[
                    SELECT i.ITMREF_0 As ItemRef
	                    ,i.ZWEBTITLE_0 As TITLE
	                    ,CASE
		                    WHEN Del.Qty Is Null THEN q.QTY_0
		                    Else q.QTY_0 - Del.Qty
	                    END As BOQty
	                    ,q.QTY_0 As OrderQty
	                    ,i.SAU_0 As UOM
	                    ,EXTRCPDAT_0 + 1 As ShipDate
                        ,u.UOMDEC_0 As NUMDEC
                        ,q.SOHNUM_0
                        ,q.SOPLIN_0
                        ,q.SOQSEQ_0
                    FROM PILOT.SORDERQ q
                    LEFT JOIN PILOT.PORDERQ p ON q.SOHNUM_0 = p.SOHNUM_0
							                    and q.SOQSEQ_0 = p.SOQSEQ_0
							                    and q.SOPLIN_0 = p.SOPLIN_0
                    INNER JOIN PILOT.ITMMASTER i ON q.ITMREF_0 = i.ITMREF_0
                    LEFT JOIN PILOT.TABUNIT u ON i.SAU_0 = u.UOM_0
                    LEFT JOIN (
	                    SELECT dd.SOHNUM_0
	                    ,dd.SOQSEQ_0
	                    ,dd.SOPLIN_0
	                    ,d.BETFCY_0
	                    ,SUM(dd.QTY_0) As Qty
	                    FROM PILOT.SDELIVERY d
	                    LEFT JOIN PILOT.SDELIVERYD dd ON d.SOHNUM_0 = dd.SOHNUM_0
	                    WHERE dd.SOHNUM_0 = '***OrderNumber***'
	                    GROUP BY dd.SOHNUM_0, dd.SOQSEQ_0, dd.SOPLIN_0, d.BETFCY_0) As Del 
	                        ON q.SOHNUM_0 = Del.SOHNUM_0 
		                        and q.SOQSEQ_0 = Del.SOQSEQ_0 
		                        and q.SOPLIN_0 = Del.SOPLIN_0
                    WHERE q.SOHNUM_0 = '***OrderNumber***'
	                    and (Del.BETFCY_0 = 1 or Del.BETFCY_0 Is Null) 
	                    AND (q.QTY_0 - Del.Qty > 0 or Del.Qty is null)
            ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
            Dim PendingDelivery As DataTable = OpenDataSet(SQLStatement)
            If PendingDelivery.Rows.Count > 0 Then
                'Create ship status header
                intTabLevel = 8
                Details += StrDup(intTabLevel, vbTab) + "<td>" + vbCrLf
                intTabLevel = 9
                Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                intTabLevel = 10
                Details += vbCrLf + StrDup(intTabLevel, vbTab) + "<table width=""575"" border=""0"" cellspacing=""2"" cellpadding=""5"" style=""margin: 0 auto 20px;"">" + vbCrLf
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                intTabLevel = 12
                Details += StrDup(intTabLevel, vbTab) + "<td><span style=""font-size:16px; font-family: arial,sans-serif;""><strong>Awaiting Shipment</strong></span></td>" + vbCrLf
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                intTabLevel = 10
                Details += StrDup(intTabLevel, vbTab) + "</table>" + vbCrLf
                'Create table header
                Details += StrDup(intTabLevel, vbTab) + "<table width=""575"" bgcolor=""#ffffff"" border=""0"" cellpadding=""10"" cellspacing=""0"" style=""border: 1px solid #ddd;margin: 0 auto;"">" + vbCrLf
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                intTabLevel = 12
                Details += LoadTableHeader(strShipMethod, intTabLevel, "Future")
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                For i = 0 To PendingDelivery.Rows.Count - 1
                    'Create item lines
                    SectionError = "DELCON Future - Filling in the product lines 1"
                    NUMDEC = PendingDelivery.Rows(i).Item("NUMDEC")
                    intTabLevel = 11
                    Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "<td style=""font-size:13px; font-family: sans-serif; border-bottom:  1px solid #ccc;"">" + vbCrLf
                    intTabLevel = 13
                    Details += StrDup(intTabLevel, vbTab) + "<strong>" + PendingDelivery.Rows(i).Item("TITLE") + "</strong><br><br>" + vbCrLf
                    Details += StrDup(intTabLevel, vbTab) + "Item #: " + PendingDelivery.Rows(i).Item("ItemRef") + "  "
                    SectionError = "DELCON Future - Filling in the product lines 2"
                    Details += StrDup(intTabLevel, vbTab) + "UOM: " + PendingDelivery.Rows(i).Item("UOM").ToString() + "<br>" + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                    Details += StrDup(intTabLevel, vbTab) + "<td align=""center"" style=""font-size:13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">"
                    Details += GetReceiptDate(ItemRef:=PendingDelivery.Rows(i).Item("ItemRef"),
                                              SOHNUM:=PendingDelivery.Rows(i).Item("SOHNUM_0"),
                                              SOQSEQ:=PendingDelivery.Rows(i).Item("SOQSEQ_0"),
                                              SOPLIN:=PendingDelivery.Rows(i).Item("SOPLIN_0")) + "</td>" + vbCrLf
                    Details += StrDup(intTabLevel, vbTab) + "<td align=""center"" style=""font-size:13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">" + vbCrLf
                    intTabLevel = 13
                    Details += StrDup(intTabLevel, vbTab) + "<p style=""background:#FFF2CE;color:#856311;padding:5px 0;"">" + vbCrLf
                    Details += Math.Round(PendingDelivery.Rows(i).Item("BOQty"), NUMDEC).ToString() + " of "
                    Details += Math.Round(PendingDelivery.Rows(i).Item("OrderQty"), NUMDEC).ToString() + "</p>" + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                    intTabLevel = 11
                    Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                Next
                intTabLevel = 10
                Details += StrDup(intTabLevel, vbTab) + "</table>" + vbCrLf
                intTabLevel = 9
                Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                intTabLevel = 8
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
            End If

            'Load items already shipped
            SectionError = "DELCON Past - SQL Statement to pull the product line information - already shipped"
            SQLStatement = <Sql><![CDATA[
                        SELECT dd.SDHNUM_0 As ShipmentNumber_Shipped
	                        ,dd.ITMREF_0 As ItemRef
	                        ,i.ZWEBTITLE_0 As TITLE
	                        ,d.DLVDAT_0 As DeliveryDate
	                        ,dd.QTY_0 As ShipQty
	                        ,q.QTY_0 As OrderQty
	                        ,dd.SAU_0 As UOM
                            ,u.UOMDEC_0 AS NUMDEC
                        FROM PILOT.SDELIVERY d
                        LEFT JOIN PILOT.ZDELCON z ON d.SDHNUM_0 = z.SDHNUM_0
                        INNER JOIN PILOT.SDELIVERYD dd ON d.SDHNUM_0 = dd.SDHNUM_0
                        INNER JOIN PILOT.ITMMASTER i ON dd.ITMREF_0 = i.ITMREF_0
                        LEFT JOIN PILOT.SORDERQ q ON dd.SOHNUM_0 = q.SOHNUM_0 
							                        and dd.SOQSEQ_0 = q.SOQSEQ_0
							                        and dd.SOPLIN_0 = q.SOPLIN_0
                        LEFT JOIN PILOT.TABUNIT u ON dd.SAU_0 = u.UOM_0
                        WHERE d.SOHNUM_0 = '***OrderNumber***' and (z.STAT_0 = '2' or z.STAT_0 is Null)
                        ORDER BY d.DLVDAT_0 desc
                        ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***OrderNumber***", OrderNumber)
            Dim PastDeliveries As DataTable = OpenDataSet(SQLStatement)
            If PastDeliveries.Rows.Count > 0 Then
                'Create ship status header
                intTabLevel = 8
                Details += StrDup(intTabLevel, vbTab) + "<td>" + vbCrLf
                intTabLevel = 9
                Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                intTabLevel = 10
                Details += vbCrLf + StrDup(intTabLevel, vbTab) + "<table width=""575"" border=""0"" cellspacing=""2"" cellpadding=""5"" style=""margin: 0 auto 20px;"">" + vbCrLf
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                intTabLevel = 12
                Details += StrDup(intTabLevel, vbTab) + "<td><span style="" font-size:16px; font-family: arial,sans-serif;""><strong>Previously Delivered</strong></span></td>" + vbCrLf
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                intTabLevel = 10
                Details += StrDup(intTabLevel, vbTab) + "</table>" + vbCrLf
                'Create table header
                Details += StrDup(intTabLevel, vbTab) + "<table width=""575"" bgcolor=""#ffffff"" border=""0"" cellpadding=""10"" cellspacing=""0"" style=""border: 1px solid #ddd;margin: 0 auto;"">" + vbCrLf
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                intTabLevel = 12
                Details += LoadTableHeader(strShipMethod, intTabLevel, "Past")
                intTabLevel = 11
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                For i = 0 To PastDeliveries.Rows.Count - 1
                    'Create item lines
                    SectionError = "DELCON Past - Filling in the product lines 1"
                    NUMDEC = PastDeliveries.Rows(i).Item("NUMDEC")
                    intTabLevel = 11
                    Details += StrDup(intTabLevel, vbTab) + "<tr>" + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "<td style=""font-size:13px; font-family: sans-serif; border-bottom:  1px solid #ccc;"">" + vbCrLf
                    intTabLevel = 13
                    Details += StrDup(intTabLevel, vbTab) + "<strong>" + PastDeliveries.Rows(i).Item("TITLE") + "</strong><br><br>" + vbCrLf
                    SectionError = "DELCON Past - Filling in the product lines 2"
                    Details += StrDup(intTabLevel, vbTab) + "Item #: " + PastDeliveries.Rows(i).Item("ItemRef") + "  "
                    Details += "UOM: " + PastDeliveries.Rows(i).Item("UOM").ToString() + "<br>" + vbCrLf
                    Details += StrDup(intTabLevel, vbTab) + "Shipment #: " + PastDeliveries.Rows(i).Item("ShipmentNumber_Shipped") + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                    Details += StrDup(intTabLevel, vbTab) + "<td align=""center"" style=""font-size:13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">"
                    'Tracking number for parcel, ship date for others
                    If strShipMethod = "Parcel" Then
                        Details += GetTrackingHyperlink(OrderNumber, PastDeliveries.Rows(i).Item("ShipmentNumber_Shipped"))
                    Else
                        Details += PastDeliveries.Rows(i).Item("DeliveryDate") + "</td>" + vbCrLf
                    End If
                    Details += StrDup(intTabLevel, vbTab) + "<td align=""center"" style=""font-size:13px; font-family: sans-serif; border-bottom: 1px solid #ccc;"">" + vbCrLf
                    intTabLevel = 13
                    ' Change Color when partial line ship
                    If Math.Round(PastDeliveries.Rows(i).Item("ShipQty"), NUMDEC) <> Math.Round(PastDeliveries.Rows(i).Item("OrderQty"), NUMDEC) Then
                        Details += StrDup(intTabLevel, vbTab) + "<p style=""background:#FFF2CE;color:#856311;padding:5px 0;"">" + vbCrLf
                    Else
                        Details += StrDup(intTabLevel, vbTab) + "<p style=""background:#cfeddc;color:#298250;padding:5px 0;"">"
                    End If
                    Details += Math.Round(PastDeliveries.Rows(i).Item("ShipQty"), NUMDEC).ToString() + " of "
                    Details += Math.Round(PastDeliveries.Rows(i).Item("OrderQty"), NUMDEC).ToString() + "</p>" + vbCrLf
                    intTabLevel = 12
                    Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                    intTabLevel = 11
                    Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
                Next
                intTabLevel = 10
                Details += StrDup(intTabLevel, vbTab) + "</table>" + vbCrLf
                intTabLevel = 9
                Details += StrDup(intTabLevel, vbTab) + "</td>" + vbCrLf
                intTabLevel = 8
                Details += StrDup(intTabLevel, vbTab) + "</tr>" + vbCrLf
            End If

            ' Add Details to the HTML
            SectionError = "DELCON Past - Add Details to the HTML"
            Body = Replace(Body, "***LINES***", Details)

            ' Constructing the ToAddress of the email
            SectionError = "Constructing the ToAddress of the email for DELCON"
            Dim EmailTo As String = ""
            If Len(Delivery.Rows(0).Item("WEB_0")) > 1 Then
                EmailTo = EmailTo + Delivery.Rows(0).Item("WEB_0") + ";"
            End If
            If Len(Delivery.Rows(0).Item("CNTOAEML1_0")) > 1 And Not EmailTo.Contains(Delivery.Rows(0).Item("CNTOAEML1_0")) Then
                EmailTo = EmailTo + Delivery.Rows(0).Item("CNTOAEML1_0") + ";"
            End If
            If Len(Delivery.Rows(0).Item("CNTOAEML2_0")) > 1 And Not EmailTo.Contains(Delivery.Rows(0).Item("CNTOAEML2_0")) Then
                EmailTo = EmailTo + Delivery.Rows(0).Item("CNTOAEML2_0") + ";"
            End If
            If Len(Delivery.Rows(0).Item("CNTOAEML3_0")) > 1 And Not EmailTo.Contains(Delivery.Rows(0).Item("CNTOAEML3_0")) Then
                EmailTo = EmailTo + Delivery.Rows(0).Item("CNTOAEML3_0") + ";"
            End If
            If Len(Delivery.Rows(0).Item("CNTOAEML4_0")) > 1 And Not EmailTo.Contains(Delivery.Rows(0).Item("CNTOAEML4_0")) Then
                EmailTo = EmailTo + Delivery.Rows(0).Item("CNTOAEML4_0") + ";"
            End If

            Dim BCC As String = ""
            If Len(CreateUserEmail) > 2 Then
                BCC = CreateUserEmail
            End If
            If Footer.Rows.Count > 0 Then
                If Len(Footer.Rows(0).Item("CSREML_0")) > 2 Then
                    BCC += ";" + Footer.Rows(0).Item("CSREML_0")
                End If
            End If
            If Len(Footer.Rows(0).Item("REPEML_0")) > 2 And Not Footer.Rows(0).Item("REPEML_0") = "aj@packbgr.com" Then
                BCC += ";" + Footer.Rows(0).Item("REPEML_0")
            End If
            BCC += ";cdreyer@packbgr.com"

            'Me.RichTextBox1.Text = Body

            ' Sends the Email "BGR <do-not-reply@bgr.us>"
            SectionError = "DELCON - Sends the Email"
            If Len(EmailTo) > 1 Then
                EmailTo = EmailTo.Substring(0, Len(EmailTo) - 1)
                If CreateUserDept = 10 Then
                    SendHTMLEmail(EmailTo, CreateUserEmail, "Shipping Confirmation for Order Number " + OrderNumber, Body, "", BCC)
                    'SendHTMLEmail("cdreyer@packbgr.com", "cdreyer@packbgr.com", "Shipping Confirmation for Order Number " + OrderNumber, Body, "", "cdreyer@packbgr.com")
                    Console.WriteLine("Email sent")
                Else
                    If Footer.Rows.Count > 0 Then
                        If Footer.Rows(0).Item("CSREML_0") <> "" Then
                            SendHTMLEmail(EmailTo, Footer.Rows(0).Item("CSREML_0"), "Shipping Confirmation for Order Number " + OrderNumber, Body, "", BCC)
                            'SendHTMLEmail("cdreyer@packbgr.com", "cdreyer@packbgr.com", "Shipping Confirmation for Order Number " + OrderNumber, Body, "", "cdreyer@packbgr.com")
                            Console.WriteLine("Email sent")
                        Else
                            'SendHTMLEmail(EmailTo, "cdreyer@packbgr.com", "Shipping Confirmation for Order Number " + OrderNumber, Body, "", BCC)
                            'SendHTMLEmail("cdreyer@packbgr.com", "cdreyer@packbgr.com", "Shipping Confirmation for Order Number " + OrderNumber, Body, "", "cdreyer@packbgr.com")
                            Console.WriteLine("Email not sent, no CSR")
                        End If
                    Else
                        SendDeliveryConfirmation = False
                        Exit Function
                    End If
                End If
            End If
            If blnDebugFlag Then : Console.WriteLine(Body) : End If
            SendDeliveryConfirmation = True
            Exit Function
        End If
        SendDeliveryConfirmation = False
    End Function

    ' This function is used to make a datatable containing the SQL SELECT statement passed to it
    Public Function OpenDataSet(ByRef strSQL As String) As DataTable
        Dim dc As SqlConnection = New SqlConnection("Server=BGRSAGE\X3V6;Database=x3v6;UID=sa;PWD=tiger")
        Dim ds As New DataSet
        Dim cmd As New SqlCommand(strSQL, dc)
        cmd.CommandTimeout = 0
        Dim da As New SqlDataAdapter(cmd)
        da.Fill(ds, "1")
        OpenDataSet = ds.Tables("1")
        dc.Close()
    End Function

    ' This is used to run a SQL INSERT, UPDATE, or DELETE statement
    Public Sub ExecuteSQLQuery(ByVal strSQL As String)
        Dim dc As SqlConnection = New SqlConnection("Server=BGRSAGE\X3V6;Database=x3v6;UID=sa;PWD=tiger")
        Dim SQLcmd As New SqlCommand
        SQLcmd.Connection = dc
        SQLcmd.CommandText = strSQL
        SQLcmd.CommandTimeout = 0
        If Not dc.State = ConnectionState.Open Then dc.Open()
        SQLcmd.ExecuteNonQuery()
        dc.Close()
    End Sub

    ' This sends an HTML email
    Public Sub SendHTMLEmail(ByVal ToAddress As String, ByVal FromAddress As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CCAddress As String = "", Optional ByVal BCCAddress As String = "", Optional ByVal Attachments As String = "")
        Dim EmailMessage As New MailMessage
        Dim SimpleSMTP As New SmtpClient("10.200.11.30") '172.16.13.11 192.168.1.7
        SimpleSMTP.UseDefaultCredentials = False
        SimpleSMTP.DeliveryMethod = SmtpDeliveryMethod.Network
        Dim Attch As Mail.Attachment
        EmailMessage.From = New MailAddress(FromAddress)
        ' Pulling out ToAddresses
        Do While InStr(ToAddress, ";") <> 0
            EmailMessage.To.Add(ToAddress.Substring(0, InStr(ToAddress, ";") - 1))
            ToAddress = ToAddress.Substring(InStr(ToAddress, ";"), ToAddress.Length - InStr(ToAddress, ";"))
        Loop
        EmailMessage.To.Add(ToAddress)
        ' Pulling out CC Addresses
        If CCAddress <> "" Then
            Do While InStr(CCAddress, ";") <> 0
                EmailMessage.CC.Add(CCAddress.Substring(0, InStr(CCAddress, ";") - 1))
                CCAddress = CCAddress.Substring(InStr(CCAddress, ";"), CCAddress.Length - InStr(CCAddress, ";"))
            Loop
            EmailMessage.CC.Add(CCAddress)
        End If
        'Pulling out BCC Addresses
        If BCCAddress <> "" Then
            Do While InStr(BCCAddress, ";") <> 0
                EmailMessage.Bcc.Add(BCCAddress.Substring(0, InStr(BCCAddress, ";") - 1))
                BCCAddress = BCCAddress.Substring(InStr(BCCAddress, ";"), BCCAddress.Length - InStr(BCCAddress, ";"))
            Loop
            EmailMessage.Bcc.Add(BCCAddress)
        End If
        'Add Attachments
        If Attachments <> "" Then
            Do While InStr(Attachments, ",") <> 0
                Attch = New Mail.Attachment(Attachments.Substring(0, InStr(Attachments, ",") - 1))
                EmailMessage.Attachments.Add(Attch)
                Attachments = Attachments.Substring(InStr(Attachments, ","), Attachments.Length - InStr(Attachments, ","))
            Loop
            Attch = New Mail.Attachment(Attachments)
            EmailMessage.Attachments.Add(Attch)
        End If
        EmailMessage.Subject = (Subject)
        EmailMessage.Body = (Body)
        EmailMessage.IsBodyHtml = True
        SimpleSMTP.Port = 25
        'SimpleSMTP.EnableSsl = True

        SimpleSMTP.Credentials = New NetworkCredential("scan@packbgr.com", "7100Gano") '"bgr\ituser", "890iu890" bgr\tbailey
        SimpleSMTP.Send(EmailMessage)
    End Sub

    Private Function GetTrackingHyperlink(ByVal SOHNUM As String, ByVal SDHNUM As String) As String

        Dim strBuffer As String = ""
        SectionError = "DELCON GetHyperlink - RunQuery"
        Dim SQLStatement As String = <Sql><![CDATA[SELECT * FROM PILOT.XBSHPINT
                                                   WHERE LTRIM(RTRIM(SDHNUM_REF1)) = '***SDH***'
        ]]></Sql>.Value
        SQLStatement = Replace(SQLStatement, "***SDH***", SDHNUM)
        Dim Tracking As DataTable = OpenDataSet(SQLStatement)

        If Tracking.Rows.Count > 0 Then
            SectionError = "DELCON GetHyperlink - XBSHPINT"
            strBuffer = BuildHyperlink(Trim(Tracking.Rows(0).Item("SRVTYP")), Tracking.Rows(0).Item("TRKNUM"))
        Else
            SQLStatement = <Sql><![CDATA[SELECT * FROM PILOT.ZTRKNUMS
                                        WHERE ShipNum = '***SDH***'
                                    ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***SDH***", SDHNUM)
            Tracking = OpenDataSet(SQLStatement)

            If Tracking.Rows.Count > 0 Then
                SectionError = "DELCON GetHyperlink - View"
                strBuffer = BuildHyperlink(Trim(Tracking.Rows(0).Item("SRVTYP")), Tracking.Rows(0).Item("TRKNUM"))
            Else
                strBuffer = "Unavailable"
            End If
        End If

        GetTrackingHyperlink = strBuffer

    End Function

    Private Function BuildHyperlink(strType As String, strTrkNum As String) As String
        Dim strBuffer As String = ""
        Select Case strType
            Case "F3D", "FE2", "FEP", "FGH", "FGR", "FIE", "FIP", "FSO"
                strBuffer = "<a style=""color: #0000EE; text-decoration: none;"" href="
                strBuffer += "https://www.fedex.com/apps/fedextrack/?tracknumbers=" + strTrkNum + ">"
                strBuffer += strTrkNum + "</a>" + "</td>" + vbCrLf
            Case "U2AM", "U2DA", "U3DA", "UGRN", "UNAM", "UNDA", "UNDS", "UPS", "UST", "UWXP"
                strBuffer = "<a style=""color: #0000EE; text-decoration: none;"" href="
                strBuffer += "https://www.ups.com/track?loc=en_US&tracknum=" + strTrkNum + "&requester=WT/" + ">"
                strBuffer += strTrkNum + "</a>" + "</td>" + vbCrLf
            Case "Other"
                strBuffer = strTrkNum + "</td>" + vbCrLf
        End Select
        BuildHyperlink = strBuffer
    End Function

    Private Function GetShipMethod(ByVal strMDL As String) As String
        Select Case strMDL
            'BGR Truck
            Case "BGR", "MAIN"
                GetShipMethod = "BGR"
            'Parcel
            Case "F3D", "FE2", "FEP", "FFO", "FGH", "FGR", "FIE", "FIP", "FSO", "GRN", "U2AM", "U2DA", "U3DA", "UGRN", "UNAM", "UNDA", "UNDS", "UPS", "UST", "UWXP"
                GetShipMethod = "Parcel"
            'LTL
            Case "COUR", "FF2", "FFE", "FFP", "LTL", "PIN"
                GetShipMethod = "LTL"
            'Sales
            Case "SALE", "CUST"
                GetShipMethod = strMDL
            Case Else
                GetShipMethod = strMDL
        End Select
    End Function

    Private Function LoadTableHeader(ByVal strShipMethod As String, ByVal intTabLevel As Integer, ByVal strSection As String) As String
        Dim Details As String = ""
        If strSection = "Today" Then
            If strShipMethod = "Parcel" Then
                Details = StrDup(intTabLevel, vbTab) + "<td bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Description</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width = ""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Tracking #</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width=""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Shipped</td>" + vbCrLf
            ElseIf strShipMethod = "SALE" Or strShipMethod = "CUST" Then
                Details = StrDup(intTabLevel, vbTab) + "<td bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Description</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width = ""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Status</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width=""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Prepared</td>" + vbCrLf
            Else
                Details = StrDup(intTabLevel, vbTab) + "<td bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Description</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width = ""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Ship Date</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width=""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Shipped</td>" + vbCrLf
            End If
        ElseIf strSection = "Past" Then
            If strShipMethod = "Parcel" Then
                Details += StrDup(intTabLevel, vbTab) + "<td bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Description</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width = ""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Tracking #</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width=""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Shipped</td>" + vbCrLf
            ElseIf strShipMethod = "SALE" Or strShipMethod = "CUST" Then
                Details = StrDup(intTabLevel, vbTab) + "<td bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Description</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width = ""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Delivered</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width=""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Shipped</td>" + vbCrLf
            Else
                Details += StrDup(intTabLevel, vbTab) + "<td bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Description</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width = ""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Delivered</td>" + vbCrLf
                Details += StrDup(intTabLevel, vbTab) + "<td width=""100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Shipped</td>" + vbCrLf
            End If
        ElseIf strSection = "Future" Then
            Details += StrDup(intTabLevel, vbTab) + "<td bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Description</td>" + vbCrLf
            Details += StrDup(intTabLevel, vbTab) + "<td width="" 100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Estimated Ship</td>" + vbCrLf
            Details += StrDup(intTabLevel, vbTab) + "<td width="" 100"" align=""center"" bgcolor=""#666666"" style=""font-size: 13px; font-family: sans-serif; color: #fff;"">Left To Ship</td>" + vbCrLf
        End If
        LoadTableHeader = Details
    End Function

    Private Function GetShipText(ByVal strMDL As String, ByVal strBPT As String) As String
        Select Case strMDL
            Case "BGR", "MAIN"
                GetShipText = "Shipping via BGR Truck"
            Case "FF2", "FFE", "FFO", "FFP"
                GetShipText = "Shipping via FedEx Ground"
            Case "LTL"
                GetShipText = "Shipping via LTL Truck"
            Case "CUST"
                GetShipText = "Customer Pick Up"
            Case "SALE"
                GetShipText = "Sales Rep Delivery"
            Case "U2AM", "U2DA", "U3DA", "UGRN", "UNAM", "UNDA", "UNDS", "UPS", "UST", "UWXP"
                GetShipText = "Shipping via UPS Parcel"
            Case "F3D", "FE2", "FEP", "FGH", "FGR", "FIE", "FIP", "FSO"
                GetShipText = "Shipping via FedEx Parcel"
            Case "COUR"
                GetShipText = "Shipping via Courier"
            Case Else
                GetShipText = "Shipping via " + strMDL + " (" + strBPT + ")"
        End Select
    End Function

    Private Function GetReceiptDate(ByVal ItemRef As String, ByVal SOHNUM As String, ByVal SOQSEQ As String, ByVal SOPLIN As String) As String

        Dim dtmBuffer As Date

        ' Look for future date on SO
        Dim SQLStatement As String = <Sql><![CDATA[
                                SELECT ITMREF_0
	                                ,[DEMDLVDAT_0]
                                FROM PILOT.SORDERQ
                                WHERE ITMREF_0 = '***ItemRef***' 
                                    and SOHNUM_0 = '***SOH***' 
                                    and SOPLIN_0 = '***SOP***' 
                                    and SOQSEQ_0 = '***SOQ***' 
                                    and DEMDLVDAT_0 >= GETDATE()
                                ]]></Sql>.Value
        SQLStatement = Replace(SQLStatement, "***ItemRef***", ItemRef)
        SQLStatement = Replace(SQLStatement, "***SOH***", SOHNUM)
        SQLStatement = Replace(SQLStatement, "***SOP***", SOPLIN)
        SQLStatement = Replace(SQLStatement, "***SOQ***", SOQSEQ)
        Dim Receipt As DataTable = OpenDataSet(SQLStatement)

        If Receipt.Rows.Count > 0 Then
            dtmBuffer = Receipt.Rows(0).Item("DEMDLVDAT_0")
            GetReceiptDate = Mid(dtmBuffer.ToString(), 1, InStr(dtmBuffer.ToString(), " ") - 1)
            Exit Function
        Else
            ' Look for next available PO delivery
            SQLStatement = <Sql><![CDATA[
                                SELECT ITMREF_0
	                                ,EXTRCPDAT_0
	                                ,QTYPUU_0
	                                ,QTYUOM_0
                                FROM PILOT.PORDERQ
                                WHERE ITMREF_0 = '***ItemRef***' and EXTRCPDAT_0 >= GETDATE()
                                ORDER BY EXTRCPDAT_0
                                ]]></Sql>.Value
            SQLStatement = Replace(SQLStatement, "***ItemRef***", ItemRef)
            Receipt = OpenDataSet(SQLStatement)

            If Receipt.Rows.Count > 0 Then
                dtmBuffer = Receipt.Rows(0).Item("EXTRCPDAT_0").AddDays(1)
                GetReceiptDate = Mid(dtmBuffer.ToString(), 1, InStr(dtmBuffer.ToString(), " ") - 1)
                Exit Function
            Else
                GetReceiptDate = "Pending"
                Exit Function
            End If
        End If

    End Function

    Private Function GetShipDate(ByVal dtmShipDate As Date, ByVal strShipMethod As String, ByVal intSameDayIs2 As Integer) As Date
        If intSameDayIs2 = 2 Then
            GetShipDate = dtmShipDate
        Else
            Select Case dtmShipDate.DayOfWeek
                Case 0, 1, 2, 3, 4
                    GetShipDate = dtmShipDate.AddDays(1)
                Case 5
                    GetShipDate = dtmShipDate.AddDays(3)
                Case 6
                    GetShipDate = dtmShipDate.AddDays(2)
            End Select
        End If
        Return dtmShipDate
    End Function

End Class

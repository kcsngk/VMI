Module archvmi

    Sub Main()
        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12

        Dim RQ4_SERVICE_URL As String = "https://vmi1.iqmetrix.net/VMIService.asmx" ' PRODUCTION 
        'Dim RQ4_SERVICE_URL As String = "https://vmi7.iqmetrix.net/VMIService.asmx" ' TEST
        Dim VMI_ClientID As Guid = Nothing
        'Dim VMI_ClientID As New Guid("{aac62064-a98d-4c1e-95ac-6fc585c6bef8}") ' PRODUCTION ARCH

        Dim vmi As RQ4.VMIService = New RQ4.VMIService
        vmi.CookieContainer = New Net.CookieContainer
        vmi.Url = RQ4_SERVICE_URL

        Dim vid As RQ4.VendorIdentity = New RQ4.VendorIdentity
        vid.Username = "c2w1r3l355"
        vid.Password = "acc350r135"
        vid.VendorID = New Guid("F9B982C3-E7B1-4FD5-9C24-ABE752E058C7")

        Console.WriteLine("pending companies:")
        For Each p As RQ4.CompanyInformation In vmi.GetPendingCompanies(vid)
            Console.WriteLine(p.Name)
            Console.WriteLine(p.CompanyID.ToString.ToUpper)
            Console.WriteLine()

            'Console.Write("enable y/n: ")
            'If Console.ReadLine = "y" Then
            '    vid.Client = New RQ4.ClientAgent
            '    vid.Client.ClientID = p.CompanyID
            '    vmi.EnableCompany(vid, True)
            'End If
        Next
        Console.WriteLine("enter to continue")
        Console.ReadLine()


        Dim companies() As RQ4.CompanyInformation = vmi.GetCompanyList(vid)
        For x As Integer = 0 To companies.Length - 1
            Console.WriteLine(x + 1 & " => " & companies(x).Name & " - " & companies(x).CompanyID.ToString)
        Next
        Console.Write("which # : ")

        Try
            VMI_ClientID = companies(Console.ReadLine - 1).CompanyID
        Catch ex As Exception
            Console.WriteLine("failure")
            Console.ReadLine()
            End
        End Try

        If VMI_ClientID = Nothing Then
            Console.WriteLine("failure")
            Console.ReadLine()
            End
        End If

        Console.WriteLine()
        Console.WriteLine("1 => inventory")
        Console.WriteLine("2 => locations")
        Console.WriteLine("3 => sku list")
        Console.Write("which # (1): ")

        Dim whichthing As Integer = 1

        Select Case Console.ReadLine
            Case "1"
                whichthing = 1
            Case "2"
                whichthing = 2
            Case "3"
                whichthing = 3
            Case Else
                whichthing = 4
        End Select

        vid.Client = New RQ4.ClientAgent
        vid.Client.ClientID = VMI_ClientID

        Console.WriteLine()

        Dim outfilename As String = "output-" & Now.Ticks & ".csv"

        If whichthing = 1 Then
            Dim enddate As String = String.Format("{0:D2}/{1:D2}/{2:D4}", DateTime.Today.Month, DateTime.Today.Day, DateTime.Today.Year)
            Dim startdate As String = String.Format("{0:D2}/{1:D2}/{2:D4}", DateTime.Now.AddDays(-30).Month, DateTime.Now.AddDays(-30).Day, DateTime.Now.AddDays(-30).Year)

            Console.Write("sales FROM (" & startdate & "): ")
            Dim ndate As String = Console.ReadLine
            If ndate.IndexOf("/") > 0 Then
                startdate = ndate
            End If

            Console.Write("sales TO (" & enddate & "): ")
            Dim ndate2 As String = Console.ReadLine
            If ndate2.IndexOf("/") > 0 Then
                enddate = ndate2
            End If

            Console.WriteLine()
            Console.WriteLine("k wait")
            Console.WriteLine()

            Dim hinfo As RQ4.HierarchyInfo = vmi.GetHierarchyInfo(vid)

            Console.WriteLine("found " & hinfo.Channels.Count & " channels..")
            Console.WriteLine()

            Using outfile As New IO.StreamWriter(outfilename)
                outfile.WriteLine(quote("store") & "," & quote("rqsku") & "," & quote("vsku") & "," & quote("description") & "," & quote("sales") & "," & quote("voh") & "," & quote("min") & "," & quote("max") & "," & quote("cost") & "," & quote("dno") & "," & quote("category") & "," & quote("rqitemid"))

                For Each c As RQ4.Channel In hinfo.Channels
                    'If c.ChannelName = "West" Then
                    Console.Write("getting channel: " & c.ChannelName & " .. " & c.StoreCount & " stores .. wait .. ")

                    'vmi.GetProductSalesReport(vid, "8/1/2016", "10/31/2016", c.ChannelID, -1, -1, -1)


                    Dim pi() As RQ4.ProductAndStoreInformation = vmi.GetGeographicInventoryReport(vid, c.ChannelID, -1, -1, -1, startdate, enddate)
                    'Dim pi() As RQ4.ProductAndStoreInformation = vmi.GetGeographicInventoryReport(vid, Guid.Empty, -1, -1, 193, startdate, enddate)

                    Console.WriteLine("found " & pi.Length & " things")

                    For Each p As RQ4.ProductAndStoreInformation In pi
                        outfile.WriteLine(quote(p.StoreName) & "," & quote(p.ProductSKU) & "," & quote(p.VendorSKU) & "," & quote(p.ProductName) & "," & quote(p.GrossQuantitySold) & "," & quote(p.QuantityInStock) & "," & quote(p.MinimumLevel) & "," & quote(p.MaximumLevel) & "," & quote(p.ProductCost) & "," & quote(p.DoNotOrder) & "," & quote(p.CategoryPath) & "," & quote(p.ProductItemID))
                        'outfile.WriteLine(quote(p.StoreName) & "," & quote(p.ProductSKU) & "," & quote(p.VendorSKU) & "," & quote(p.ProductName) & "," & quote(p.GrossQuantitySold) & "," & quote(p.QuantityInStock + p.QuantityInTransfer + p.QuantityOnOrder + p.QuantityOnBackOrder) & "," & quote(p.MinimumLevel) & "," & quote(p.MaximumLevel) & "," & quote(p.ProductCost) & "," & quote(p.DoNotOrder) & "," & quote(p.CategoryPath) & "," & quote(p.ProductItemID))
                        'outfile.WriteLine(quote(p.StoreName) & "," & quote(p.ProductSKU) & "," & quote(p.VendorSKU) & "," & quote(p.ProductName) & "," & quote(p.GrossQuantitySold) & "," & quote(p.QuantityInStock) & "," & quote(p.MinimumLevel) & "," & quote(p.MaximumLevel) & "," & quote(p.ProductCost) & "," & quote(p.DoNotOrder) & "," & quote(p.CategoryPath) & "," & quote(p.ProductItemID))
                    Next
                    'End If
                Next
            End Using
        ElseIf whichthing = 2 Then
            Console.WriteLine("k wait")
            Console.WriteLine()

            Dim si() As RQ4.StoreInformation = vmi.GetStoreList(vid)

            Console.WriteLine("found " & si.Length & " locations..")

            Using outfile As New IO.StreamWriter(outfilename)
                outfile.WriteLine(quote("region") & "," & quote("district") & "," & quote("abbr") & "," & quote("name") & "," & quote("address") & "," & quote("city") & "," & quote("state") & "," & quote("zip") & "," & quote("phone") & "," & quote("country") & "," & quote("vendor#") & "," & quote("storeid") & "," & quote("billid") & "," & quote("shipid"))

                For Each s As RQ4.StoreInformation In si
                    outfile.WriteLine(quote(s.Region) & "," & quote(s.District) & "," & quote(s.Abbreviation) & "," & quote(s.Name) & "," & quote(s.Address) & "," & quote(s.City) & "," & quote(s.ProvinceState) & "," & quote(s.PostalZipCode) & "," & quote(s.PhoneNumber) & "," & quote(s.Country) & "," & quote(s.VendorAccountNumber) & "," & quote(s.StoreID) & "," & quote(s.BillToStoreID) & "," & quote(s.ShipToStoreID))
                Next
            End Using
        ElseIf whichthing = 3 Then
            'vid.Client.StoreID = 23

            'Dim a As RQ4.ReturnMerchandiseAuthorization() = vmi.GetRMAList(vid, False, False)
            'End

            Console.WriteLine("k wait")
            Console.WriteLine()

            Dim pi As RQ4.ProductInformation() = vmi.GetInventoryList(vid)

            Using fout As New IO.StreamWriter("skus.csv", False)
                fout.WriteLine("vendorsku,productsku,productitemid")
                For Each p As RQ4.ProductInformation In pi
                    If p.VendorSKU <> "" Then
                        fout.WriteLine(p.VendorSKU & "," & p.ProductSKU & "," & p.ProductItemID)
                    End If
                Next
                fout.Close()
            End Using

            Process.Start("skus.csv")


            End
        ElseIf whichthing = 4 Then
            vid.Client = New RQ4.ClientAgent
            vid.Client.ClientID = VMI_ClientID
            vid.Client.StoreID = 148

            Dim poi() As RQ4.PurchaseOrderInformation = vmi.GetPurchaseOrderList(vid)
            'vid.Client.StoreID = 117
            'Dim pis() As RQ4.ProductAndStoreInformation = vmi.GetGeographicInventoryReport(vid, Guid.Empty, -1, -1, 14, "07/01/2017", "07/06/2017")
            'Dim poi() As RQ4.PurchaseOrderInformation = vmi.GetPurchaseOrderList(vid)
            MsgBox("abc")
        End If

        Console.WriteLine()
        Console.WriteLine("k bye")

        Process.Start(outfilename)
    End Sub

    Function quote(what As String) As String
        Return """" & what.Replace(vbCr, "").Replace(vbLf, "") & """"
    End Function
End Module

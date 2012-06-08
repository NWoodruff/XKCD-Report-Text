Imports System.Net
Imports System.IO
Imports System.Object

Module Module1
    Structure GratTotal
        Public sGraticule As String
        Public sGratName As String
        Public iNum As Integer
        Public dtLast As Date
        Public bOverTake As Integer
    End Structure

    Structure GratGrandTotal
        Public sGraticule As String
        Public sGratName As String
        Public iGrandTot As Integer
        Public iRetroTot As Integer
        Public iMonthTot() As Integer
        Public sLast As String
        Public dtLast As Date
    End Structure

    Structure GratLookupStruct
        Public sGratName As String
        Public sGratLatLon As String
    End Structure

    Structure TotMax
        Public sGraticule As String
        Public iMonth As Integer
    End Structure

    Structure TotMaxNum
        Public iTotal As Integer
        Public aTotMax() As TotMax
    End Structure

    Structure GratYearTots
        Public iYear As Integer
        Public GratTot() As GratTotal
    End Structure
    Structure GratsMonYearTots
        Public MonYear As Integer
        Public GratTot() As GratTotal
    End Structure
    Structure GratGrandYearTots
        Public GratTots() As GratTotal
        Public GratYear() As GratYearTots
        Public GratMonYear() As GratsMonYearTots
    End Structure

    Public Enum Reports
        Coordinates_reached
        Coordinates_not_reached
        Retro_reached
        Retro_not_reached
    End Enum
    Sub sort(ByRef GGT() As GratGrandTotal, ByRef MyDate As Date)
        Dim MaxGrandTots As Integer = GGT.Length()
        For z As Integer = 0 To MaxGrandTots - 1
            Dim y As Integer

            If String.IsNullOrEmpty(GGT(z).sGratName) = False Then
                For x As Integer = z + 1 To MaxGrandTots - 1
                    If String.IsNullOrEmpty(GGT(x).sGratName) = False Then
                        If GGT(x).iGrandTot > GGT(z).iGrandTot Then
                            Dim TempGrat As GratGrandTotal = GGT(z)

                            If z + 1 = x Then

                                GGT(z) = GGT(x)
                                GGT(z).sLast = GGT(z).sLast.Substring(1) + "+"
                                GGT(x) = TempGrat
                                GGT(x).sLast = GGT(x).sLast.Substring(1) + "-"
                                GGT(x).dtLast = MyDate
                                GGT(z).dtLast = MyDate
                            Else

                                For y = x To z + 1 Step -1
                                    TempGrat = GGT(y - 1)
                                    GGT(y - 1) = GGT(y)
                                    GGT(y) = TempGrat
                                    GGT(y).sLast = GGT(y).sLast.Substring(1) + "-"
                                    GGT(y).dtLast = MyDate
                                Next
                                GGT(z).sLast = GGT(z).sLast.Substring(1) + "+"
                                GGT(z).dtLast = MyDate
                            End If
                        End If
                    Else
                        x = MaxGrandTots
                    End If
                Next
            Else
                z = MaxGrandTots

            End If
        Next

    End Sub
    Function GetWebPage(ByRef sURL As String) As String
        Dim myWriter As StreamWriter = Nothing
        Dim sWebPageData As String = ""

        Dim objRequestGrat As HttpWebRequest = CType(WebRequest.Create(sURL), HttpWebRequest)
        objRequestGrat.Method = "POST"

        Console.WriteLine(sURL)

        objRequestGrat.ContentLength = sWebPageData.Length()
        objRequestGrat.ContentType = "application/x-www-form-urlencoded"

        Try
            myWriter = New StreamWriter(objRequestGrat.GetRequestStream())
            myWriter.Write(sWebPageData)

        Catch e As Exception
            Console.Write(e.Message)
            Return ""
        Finally
            myWriter.Close()
            myWriter.Dispose()
            Try

                Dim objResponse As HttpWebResponse = CType(objRequestGrat.GetResponse(), HttpWebResponse)
                Dim sr As New StreamReader(objResponse.GetResponseStream())
                sWebPageData = sr.ReadToEnd()

                ' Close and clean up the StreamReader
                sr.Close()
                sr.Dispose()
                objResponse.Close()

            Catch ex As Exception

                Console.Write(ex.Message)


            End Try


        End Try


        Return sWebPageData

    End Function
    Sub Main()
        Dim myWriter As StreamWriter = Nothing
        Dim strPost As String = ""

        Dim sPostAll As String = ""
        Dim urlList() As String = {"http://wiki.xkcd.com/geohashing/Category:Coordinates_reached", _
                                   "http://wiki.xkcd.com/geohashing/Category:Coordinates_not_reached", _
                                   "http://wiki.xkcd.com/geohashing/Category:Retro_coordinates_reached", _
                                   "http://wiki.xkcd.com/geohashing/Category:Retro_coordinates_not_reached"}


        'Dim XKCDAll As String = "http://wiki.xkcd.com/geohashing/All_Graticules" ' No longer used.
        Dim XKCDGrats() As String = {"http://wiki.xkcd.com/geohashing/All_graticules/Africa", _
                                        "http://wiki.xkcd.com/geohashing/All_graticules/Antarctica", _
                                        "http://wiki.xkcd.com/geohashing/All_graticules/Australasia", _
                                        "http://wiki.xkcd.com/geohashing/All_graticules/Eurasia", _
                                        "http://wiki.xkcd.com/geohashing/All_graticules/North_America", _
                                        "http://wiki.xkcd.com/geohashing/All_graticules/Oceans", _
                                        "http://wiki.xkcd.com/geohashing/All_graticules/South_America"}
        Dim XKCDGratsData(XKCDGrats.Length()) As String
        Dim GratsYearReached As GratGrandYearTots = Nothing
        Dim GratsYearNotReached As GratGrandYearTots = Nothing

        Dim sAGratName(180, 360) As String

        Dim x As Integer
        Dim z As Integer
        Dim MyDate As Date
        Dim bDate As Boolean = False
        Dim iYear As Integer = 0
        Dim iMonth As Integer = 0
        Dim iTotalDate As Integer = 0

        Dim GratGrandTotals() As GratGrandTotal
        Dim GratGrandTotalsNotReached() As GratGrandTotal

        Dim GratLookupNum As Integer = 0

        ' Dim MaxGrandTots As Integer = 1000
        Dim MaxMonths As Integer = (Date.Today().Year() * 12 + Date.Today().Month() + 2) - ((2008 * 12) + 5)
        Dim tlong As Long = 0
        Dim dMaxDateTotals = Date.Today()

        ' to set the date for previous years totals
        'Date.TryParse("12/31/2008", dMaxDateTotals)
        'Date.TryParse("02/28/2009", dMaxDateTotals)
        'Date.TryParse("12/31/2009", dMaxDateTotals)
        'Date.TryParse("12/31/2010", dMaxDateTotals)


        GratGrandTotals = Nothing
        GratGrandTotalsNotReached = Nothing

        For XKCDGratNum As Integer = 0 To XKCDGrats.Length() - 1

            XKCDGratsData(XKCDGratNum) = GetWebPage(XKCDGrats(XKCDGratNum))

            Dim sTemp As String = XKCDGratsData(XKCDGratNum)
            Dim iTemp As Integer = sTemp.IndexOf("<!-- start content -->")
            Dim sGeoLookup As String = "<a href=""/geohashing/"

            If iTemp > -1 Then

                Dim sid As String = "mw-headline"

                iTemp = sTemp.IndexOf(sid, iTemp)
                iTemp = sTemp.IndexOf("</h3>", iTemp)

                Dim sHref As String = ""
                Dim iHRef As Integer = sTemp.IndexOf(sGeoLookup, iTemp)
                Dim iEnd As Integer = sTemp.IndexOf("printfooter")

                While iHRef < iEnd

                    Dim iTitleStart As Integer = sTemp.IndexOf("title=""", iHRef) + 7
                    Dim ititleEnd As Integer = sTemp.IndexOf("""", iTitleStart + 1)
                    Dim sTitle As String = sTemp.Substring(iTitleStart, ititleEnd - iTitleStart)

                    Dim sGratLatLon As String = sTemp.Substring(ititleEnd + 2, sTemp.IndexOf("(", ititleEnd + 2) - ititleEnd - 3)
                    If sGratLatLon.IndexOf("class=") > -1 Then
                        sGratLatLon = sGratLatLon.Substring(20)
                    End If
                    sGratLatLon = sGratLatLon.Trim(" ")

                    If sGratLatLon.Length() < 12 Then

                        Dim aLatLon = sGratLatLon.Split(", ")

                        If aLatLon.Length() > 1 Then

                            sAGratName(aLatLon(0) + 90, aLatLon(1) + 180) = sTitle
                        End If
                    End If


                    iHRef = sTemp.IndexOf(sGeoLookup, iHRef + sGeoLookup.Length())
                End While


            End If


        Next
        Dim GGT()() As GratGrandTotal = {GratGrandTotals, GratGrandTotalsNotReached}
        Dim GGYT() As GratGrandYearTots = {GratsYearReached, GratsYearNotReached}


        For urlNum As Integer = 0 To urlList.Length() - 1

            Dim url As String = urlList(urlNum)

            While String.IsNullOrEmpty(url) = False

                Dim result As String = ""
                Dim iStart As Integer = 0

                Dim iEnd As Integer = 0
                Dim sCopy As String = ""

                result = GetWebPage(url)

                iStart = result.IndexOf("<a href=""/geohashing/2", iStart)
                While iStart > 0
                    iStart = result.IndexOf(">", iStart) + 1
                    iEnd = result.IndexOf("</a>", iStart)

                    sCopy = result.Substring(iStart, iEnd - iStart)
                    Dim aVals
                    Dim sGrat As String = ""
                    Dim sGratName As String = ""

                    aVals = sCopy.Split(" ")
                    sGrat = aVals(1)
                    sGratName = aVals(1)

                    If aVals.length() > 2 Then


                        Dim iGratStart As Integer = 0
                        Dim iGratEnd As Integer = 0
                        sGrat = sGrat + " " + aVals(2)
                        sGratName = sGratName + ", " + aVals(2)
                        Dim bFound As Boolean = False

                        sGratName = sAGratName(aVals(1) + 90, aVals(2) + 180)

                        If sGratName Is Nothing Then

                            If bFound = False Then
                                sGratName = "Unknown"
                            End If
                        End If

                    End If

                    bDate = Date.TryParse(aVals(0), MyDate)

                    If (MyDate <= dMaxDateTotals) And (bDate = True) Then

                        iYear = MyDate.Year()
                        iMonth = MyDate.Month()

                        iTotalDate = ((iYear - 2008) * 12) + iMonth - 5

                        If (iTotalDate < MaxMonths) And (iTotalDate > -1) Then

                            ' Dim GGT()() As GratGrandTotal = {GratGrandTotals, GratGrandTotalsNotReached}
                            Dim iStat As Integer = 0

                            If urlNum > 0 Then
                                iStat = 1
                            End If

                            For iGGT As Integer = iStat To GGT.Length - 1

                                Dim bFound As Boolean = False

                                If GGT(iGGT) Is Nothing = False Then

                                    For x = 0 To GGT(iGGT).Length() - 1

                                        If GGT(iGGT)(x).sGraticule = sGrat Then

                                            Select Case urlNum
                                                Case 0, 1
                                                    GGT(iGGT)(x).iMonthTot(iTotalDate) = GGT(iGGT)(x).iMonthTot(iTotalDate) + 1
                                                Case 2, 3 ' retro url's
                                                    GGT(iGGT)(x).iRetroTot = GGT(iGGT)(x).iRetroTot + 1 ' retro's don't count toward monthy totals
                                            End Select

                                            GGT(iGGT)(x).iGrandTot = GGT(iGGT)(x).iGrandTot + 1
                                            x = GGT(iGGT).Length()
                                            bFound = True
                                        End If
                                    Next
                                    If bFound = True Then
                                        sort(GGT(iGGT), MyDate)
                                    End If
                                End If
                                If bFound = False Then
                                    Dim iGGTSize As Integer = 0
                                    If GGT(iGGT) Is Nothing = False Then

                                        iGGTSize = GGT(iGGT).Length()
                                    End If

                                    Array.Resize(GGT(iGGT), iGGTSize + 1)

                                    GGT(iGGT)(iGGTSize).sGraticule = sGrat
                                    GGT(iGGT)(iGGTSize).sGratName = sGratName
                                    GGT(iGGT)(iGGTSize).iGrandTot = 1
                                    GGT(iGGT)(iGGTSize).sLast = "*****"
                                    GGT(iGGT)(iGGTSize).dtLast = MyDate

                                    GGT(iGGT)(iGGTSize).iMonthTot = Array.CreateInstance(GetType(Integer), MaxMonths)
                                    For z = 0 To MaxMonths - 1
                                        GGT(iGGT)(iGGTSize).iMonthTot(z) = 0
                                    Next
                                    GGT(iGGT)(iGGTSize).iMonthTot(iTotalDate) = 1

                                End If

                            Next ' iGGT


                            Dim iGGYTNum As Integer = 0

                            Select Case urlNum
                                Case Reports.Coordinates_not_reached
                                    iGGYTNum = 1
                                Case Reports.Retro_reached, Reports.Retro_not_reached  ' retro url's
                                    iGGYTNum = GGYT.Length()
                            End Select


                            For iGGYT As Integer = iGGYTNum To GGYT.Length() - 1 ' skip over this for retro url's


                                If GGYT(iGGYT).GratTots Is Nothing Then
                                    Array.Resize(GGYT(iGGYT).GratTots, 1)
                                    GGYT(iGGYT).GratTots(0).sGraticule = sGrat
                                    GGYT(iGGYT).GratTots(0).sGratName = sGratName
                                    GGYT(iGGYT).GratTots(0).iNum = 1

                                Else
                                    Dim bNotFound As Boolean = True

                                    For x = 0 To GGYT(iGGYT).GratTots.Length - 1
                                        If (GGYT(iGGYT).GratTots(x).sGraticule = sGrat) Then
                                            GGYT(iGGYT).GratTots(x).iNum = GGYT(iGGYT).GratTots(x).iNum + 1
                                            bNotFound = False

                                            For y As Integer = x - 1 To 0 Step -1
                                                If (GGYT(iGGYT).GratTots(y + 1).iNum > GGYT(iGGYT).GratTots(y).iNum) Then
                                                    Dim GratTotsTemp As GratTotal
                                                    GratTotsTemp = GGYT(iGGYT).GratTots(y)
                                                    GGYT(iGGYT).GratTots(y) = GGYT(iGGYT).GratTots(y + 1)
                                                    GGYT(iGGYT).GratTots(y + 1) = GratTotsTemp

                                                Else
                                                    y = 0

                                                End If

                                            Next
                                            x = GGYT(iGGYT).GratTots.Length
                                        End If
                                    Next

                                    If bNotFound = True Then
                                        Array.Resize(GGYT(iGGYT).GratTots, x + 1)
                                        GGYT(iGGYT).GratTots(x).sGraticule = sGrat
                                        GGYT(iGGYT).GratTots(x).sGratName = sGratName
                                        GGYT(iGGYT).GratTots(x).iNum = 1

                                    End If
                                End If


                                If GGYT(iGGYT).GratYear Is Nothing Then
                                    Array.Resize(GGYT(iGGYT).GratYear, 1)

                                    GGYT(iGGYT).GratYear(0).iYear = iYear
                                    Array.Resize(GGYT(iGGYT).GratYear(0).GratTot, 1)
                                    GGYT(iGGYT).GratYear(0).GratTot(0).sGraticule = sGrat
                                    GGYT(iGGYT).GratYear(0).GratTot(0).sGratName = sGratName
                                    GGYT(iGGYT).GratYear(0).GratTot(0).iNum = 1


                                Else
                                    Dim bNotFound As Boolean = True
                                    For x = 0 To GGYT(iGGYT).GratYear.Length - 1
                                        If (GGYT(iGGYT).GratYear(x).iYear = iYear) Then

                                            Dim bNotFoundGrat As Boolean = True
                                            Dim y As Integer
                                            For y = 0 To GGYT(iGGYT).GratYear(x).GratTot.Length - 1

                                                If GGYT(iGGYT).GratYear(x).GratTot(y).sGraticule = sGrat Then
                                                    GGYT(iGGYT).GratYear(x).GratTot(y).iNum = GGYT(iGGYT).GratYear(x).GratTot(y).iNum + 1
                                                    bNotFoundGrat = False

                                                    For t As Integer = y - 1 To 0 Step -1
                                                        If (GGYT(iGGYT).GratYear(x).GratTot(t + 1).iNum > GGYT(iGGYT).GratYear(x).GratTot(t).iNum) Then
                                                            Dim GratTotsTemp As GratTotal
                                                            GratTotsTemp = GGYT(iGGYT).GratYear(x).GratTot(t)
                                                            GGYT(iGGYT).GratYear(x).GratTot(t) = GGYT(iGGYT).GratYear(x).GratTot(t + 1)
                                                            GGYT(iGGYT).GratYear(x).GratTot(t + 1) = GratTotsTemp

                                                        Else
                                                            t = 0
                                                        End If

                                                    Next
                                                    y = GGYT(iGGYT).GratYear(x).GratTot.Length
                                                End If
                                            Next

                                            If bNotFoundGrat = True Then
                                                Array.Resize(GGYT(iGGYT).GratYear(x).GratTot, y + 1)
                                                GGYT(iGGYT).GratYear(x).GratTot(y).sGraticule = sGrat
                                                GGYT(iGGYT).GratYear(x).GratTot(y).sGratName = sGratName
                                                GGYT(iGGYT).GratYear(x).GratTot(y).iNum = 1

                                            End If

                                            bNotFound = False
                                        End If
                                    Next
                                    If bNotFound = True Then
                                        Array.Resize(GGYT(iGGYT).GratYear, x + 1)

                                        GGYT(iGGYT).GratYear(x).iYear = iYear
                                        Array.Resize(GGYT(iGGYT).GratYear(x).GratTot, 1)
                                        GGYT(iGGYT).GratYear(x).GratTot(0).sGraticule = sGrat
                                        GGYT(iGGYT).GratYear(x).GratTot(0).sGratName = sGratName
                                        GGYT(iGGYT).GratYear(x).GratTot(0).iNum = 1

                                    End If
                                End If





                                If GGYT(iGGYT).GratMonYear Is Nothing Then
                                    Array.Resize(GGYT(iGGYT).GratMonYear, 1)

                                    GGYT(iGGYT).GratMonYear(0).MonYear = iTotalDate
                                    Array.Resize(GGYT(iGGYT).GratMonYear(0).GratTot, 1)
                                    GGYT(iGGYT).GratMonYear(0).GratTot(0).sGraticule = sGrat
                                    GGYT(iGGYT).GratMonYear(0).GratTot(0).sGratName = sGratName
                                    GGYT(iGGYT).GratMonYear(0).GratTot(0).iNum = 1
                                    GGYT(iGGYT).GratMonYear(0).GratTot(0).dtLast = MyDate
                                    GGYT(iGGYT).GratMonYear(0).GratTot(0).bOverTake = 0


                                Else
                                    Dim bNotFound As Boolean = True
                                    For x = 0 To GGYT(iGGYT).GratMonYear.Length - 1
                                        If (GGYT(iGGYT).GratMonYear(x).MonYear = iTotalDate) Then

                                            Dim bNotFoundGrat As Boolean = True
                                            Dim y As Integer
                                            For y = 0 To GGYT(iGGYT).GratMonYear(x).GratTot.Length - 1

                                                If GGYT(iGGYT).GratMonYear(x).GratTot(y).sGraticule = sGrat Then
                                                    GGYT(iGGYT).GratMonYear(x).GratTot(y).iNum = GGYT(iGGYT).GratMonYear(x).GratTot(y).iNum + 1
                                                    bNotFoundGrat = False

                                                    For t As Integer = y - 1 To 0 Step -1
                                                        If (GGYT(iGGYT).GratMonYear(x).GratTot(t + 1).iNum > GGYT(iGGYT).GratMonYear(x).GratTot(t).iNum) Then
                                                            Dim GratTotsTemp As GratTotal
                                                            GGYT(iGGYT).GratMonYear(x).GratTot(t + 1).dtLast = MyDate
                                                            GGYT(iGGYT).GratMonYear(x).GratTot(t + 1).bOverTake = 1
                                                            GratTotsTemp = GGYT(iGGYT).GratMonYear(x).GratTot(t)
                                                            GGYT(iGGYT).GratMonYear(x).GratTot(t) = GGYT(iGGYT).GratMonYear(x).GratTot(t + 1)
                                                            GGYT(iGGYT).GratMonYear(x).GratTot(t + 1) = GratTotsTemp
                                                            GGYT(iGGYT).GratMonYear(x).GratTot(t + 1).dtLast = MyDate
                                                            GGYT(iGGYT).GratMonYear(x).GratTot(t + 1).bOverTake = -1

                                                        Else
                                                            t = 0
                                                        End If

                                                    Next
                                                    y = GGYT(iGGYT).GratMonYear(x).GratTot.Length
                                                End If
                                            Next

                                            If bNotFoundGrat = True Then
                                                Array.Resize(GGYT(iGGYT).GratMonYear(x).GratTot, y + 1)
                                                GGYT(iGGYT).GratMonYear(x).GratTot(y).sGraticule = sGrat
                                                GGYT(iGGYT).GratMonYear(x).GratTot(y).sGratName = sGratName
                                                GGYT(iGGYT).GratMonYear(x).GratTot(y).iNum = 1
                                                GGYT(iGGYT).GratMonYear(x).GratTot(y).dtLast = MyDate
                                                GGYT(iGGYT).GratMonYear(x).GratTot(y).bOverTake = 0

                                            End If

                                            bNotFound = False
                                        End If
                                    Next
                                    If bNotFound = True Then
                                        Array.Resize(GGYT(iGGYT).GratMonYear, x + 1)

                                        GGYT(iGGYT).GratMonYear(x).MonYear = iTotalDate
                                        Array.Resize(GGYT(iGGYT).GratMonYear(x).GratTot, 1)
                                        GGYT(iGGYT).GratMonYear(x).GratTot(0).sGraticule = sGrat
                                        GGYT(iGGYT).GratMonYear(x).GratTot(0).sGratName = sGratName
                                        GGYT(iGGYT).GratMonYear(x).GratTot(0).iNum = 1
                                        GGYT(iGGYT).GratMonYear(x).GratTot(0).dtLast = MyDate
                                        GGYT(iGGYT).GratMonYear(x).GratTot(0).bOverTake = 0

                                    End If
                                End If



                            Next ' iGGYT


                        Else
                            If urlNum > 0 Then

                                If iTotalDate < 0 Then

                                    Dim bFound As Boolean = False

                                    If GGT(1) Is Nothing = False Then

                                        For x = 0 To GGT(1).Length() - 1

                                            If GGT(1)(x).sGraticule = sGrat Then

                                                Select Case urlNum
                                                    Case Reports.Coordinates_reached, Reports.Coordinates_not_reached
                                                        GGT(1)(x).iMonthTot(iTotalDate) = GGT(1)(x).iMonthTot(iTotalDate) + 1
                                                    Case Reports.Retro_reached, Reports.Retro_not_reached  ' retro url's
                                                        GGT(1)(x).iRetroTot = GGT(1)(x).iRetroTot + 1 ' retro's don't count toward monthy totals
                                                End Select

                                                GGT(1)(x).iGrandTot = GGT(1)(x).iGrandTot + 1
                                                x = GGT(1).Length()
                                                bFound = True
                                            End If
                                        Next
                                        If bFound = True Then
                                            sort(GGT(1), MyDate)
                                        End If
                                    End If
                                    If bFound = False Then
                                        Dim iGGTSize As Integer = 0
                                        If GGT(1) Is Nothing = False Then

                                            iGGTSize = GGT(1).Length()
                                        End If

                                        Array.Resize(GGT(1), iGGTSize + 1)

                                        GGT(1)(iGGTSize).sGraticule = sGrat
                                        GGT(1)(iGGTSize).sGratName = sGratName
                                        GGT(1)(iGGTSize).iGrandTot = 1
                                        GGT(1)(iGGTSize).sLast = "*****"
                                        GGT(1)(iGGTSize).dtLast = MyDate

                                        GGT(1)(iGGTSize).iMonthTot = Array.CreateInstance(GetType(Integer), MaxMonths)
                                        For z = 0 To MaxMonths - 1
                                            GGT(1)(iGGTSize).iMonthTot(z) = 0
                                        Next
                                        GGT(1)(iGGTSize).iMonthTot(iTotalDate) = 1

                                    End If





                                End If ' iTotalDate > -1


                            Else
                                bDate = Date.TryParse(aVals(0), MyDate)
                            End If


                        End If
                    End If


                    iStart = result.IndexOf("<a href=""/geohashing/2", iStart)
                End While ' start > 0 start date greater than May 2008

                Dim iNext As Integer = 0
                Dim iPrevious As Integer = 0

                iPrevious = result.IndexOf("previous 200", 0)
                If iPrevious > 0 Then
                    iNext = result.IndexOf("next 200", iPrevious)
                    sCopy = result.Substring(iPrevious, iNext - iPrevious)

                    Dim iHRefStart As Integer = 0
                    Dim iHRefEnd As Integer = 0

                    iHRefStart = sCopy.IndexOf("<a href=", 0)

                    If iHRefStart > 0 Then
                        url = "http://wiki.xkcd.com" + sCopy.Substring(iHRefStart + 9, sCopy.IndexOf("""", iHRefStart + 9) - (iHRefStart + 9))

                        url = url.Replace("&amp;", "&")

                    Else
                        url = ""
                    End If
                Else
                    url = ""
                End If
            End While
        Next ' urllist


        Dim iMaxMonths As Integer = (Date.Today().Year() * 12 + Date.Today().Month()) - ((2008 * 12) + 5)

        iMaxMonths = (dMaxDateTotals.Year() * 12 + dMaxDateTotals.Month()) - ((2008 * 12) + 5)

        Dim iStartDisplayYear As Integer = 2008
        iStartDisplayYear = dMaxDateTotals.Year()

        If dMaxDateTotals.Month() < 2 Then
            iStartDisplayYear = iStartDisplayYear - 1
        End If

        If iStartDisplayYear < 2008 Then
            iStartDisplayYear = 2008
        End If



        For iReport As Integer = 0 To 1

            Dim iStop As Integer = 0


            'Select Case (Date.Today().DayOfWeek)
            Select Case (dMaxDateTotals.DayOfWeek)

                Case DayOfWeek.Saturday
                    iStop = 390
                    iStop = GGT(iReport).Length() - 1 ' display all
                Case DayOfWeek.Sunday
                    iStop = 390
                    iStop = GGT(iReport).Length() - 1 ' display all
                Case Else
                    iStop = 99 ' Display top 100
            End Select

            If dMaxDateTotals < Date.Today() Then
                iStop = GGT(iReport).Length() - 1
            End If

            If iStop < GGT(iReport).Length() - 1 Then
                ' this is to make sure that you don't leave any out that have the same total
                If String.IsNullOrEmpty(GGT(iReport)(iStop).sGraticule) = False Then
                    While GGT(iReport)(iStop).iGrandTot = GGT(iReport)(iStop + 1).iGrandTot
                        iStop = iStop + 1
                    End While
                End If
            End If


            Dim irowspan As Integer = 0
            Dim bRowSpanPrint As Boolean = False
            Dim iRowSpan1 As Integer = -1
            Dim bRowSpanPrint1 As Boolean = False
            Dim dtNow As Date = Date.Now()

            Dim sFileName As String = "XKCD" + dtNow.Year().ToString() + dtNow.Month.ToString("d2") + _
                dtNow.Day.ToString("d2") + dtNow.Hour.ToString("d2") + dtNow.Minute.ToString("d2")

            Select Case iReport
                Case 0
                    sFileName = sFileName + "Reached.txt"
                Case 1
                    sFileName = sFileName + "Exp.txt"

            End Select

            Dim swText As New StreamWriter(sFileName, False)

            Select Case iReport

                Case 0
                    swText.Write("==Coordinates reached==" + vbCrLf + vbCrLf)
                Case 1
                    swText.Write("==Expeditions==" + vbCrLf + vbCrLf)
            End Select

            swText.Write("''as of " + Date.UtcNow.ToString() + " UTC''" + vbCrLf + vbCrLf)

            swText.Write("{| border=""2"" cellpadding=""4"" cellspacing=""0"" style=""margin-top:1em; margin-bottom:1em; background:#f9f9f9; border:1px #aaa solid; border-collapse:collapse;""" + vbCrLf)
            swText.Write("|-" + vbCrLf)
            swText.Write("|" + vbCrLf)
            swText.Write("! Graticule" + vbCrLf)
            swText.Write("! Meetup in" + vbCrLf)
            swText.Write("! Total" + vbCrLf)
            If iReport = 0 Then ' coordinates reached

                swText.Write("! Changes" + vbCrLf)
            End If

            irowspan = 0
            For z = 0 To iStop
                If String.IsNullOrEmpty(GGT(iReport)(z).sGraticule) = False Then
                    irowspan = irowspan + 1
                Else
                    z = GGT(iReport).Length()
                End If
            Next
            irowspan = irowspan + 1

            swText.Write("| rowspan=" + irowspan.ToString() + "|" + vbCrLf)
            If iReport = 1 Then ' coordinates not reached

                swText.Write("| Retro" + vbCrLf)
            End If

            For z = 2008 To iStartDisplayYear - 1
                swText.Write("! [[" + z.ToString() + "_Most_active_graticules|" + z.ToString() + "]]" + vbCrLf)
            Next

            Dim iBeginMonthTot As Integer = (iStartDisplayYear * 12) - ((2008 * 12) + 5)
            Dim iMonthOffset As Integer = 0
            If iStartDisplayYear = 2008 Then
                iMonthOffset = 4
            End If

            Dim iMaxMonthsPrint As Integer = iMaxMonths
            If iMaxMonthsPrint > ((dMaxDateTotals.Year() + 1) * 12) - ((2008 * 12) + 5) + 1 Then
                iMaxMonthsPrint = ((dMaxDateTotals.Year() + 1) * 12) - ((2008 * 12) + 5)
            End If

            For x = ((iBeginMonthTot + iMonthOffset) + 1) To iMaxMonthsPrint
                Dim iMonth1 As Integer = x - (iBeginMonthTot + iMonthOffset) + iMonthOffset
                If iMonth1 > 12 Then
                    iMonth1 = iMonth1 - 12
                End If
                swText.Write("! " + MonthName(iMonth1, True) + vbCrLf)
            Next


            irowspan = 0
            bRowSpanPrint = False
            iRowSpan1 = 0
            bRowSpanPrint1 = False

            For z = 0 To iStop
                Dim iRank As Integer = z + 1
                If String.IsNullOrEmpty(GGT(iReport)(z).sGraticule) = False Then
                    bRowSpanPrint = False

                    swText.Write("|-" + vbCrLf)

                    If irowspan < 1 Then
                        Dim bOverBounds As Boolean = False

                        If iRank + irowspan < GGT(iReport).Length() Then

                            While (GGT(iReport)(z).iGrandTot = GGT(iReport)(iRank + irowspan).iGrandTot) And (bOverBounds = False)
                                irowspan = irowspan + 1
                                bRowSpanPrint = True
                                If iRank + irowspan > GGT(iReport).Length() - 1 Then
                                    irowspan = GGT(iReport).Length() - 1 - iRank
                                    bOverBounds = True
                                End If
                            End While
                            If bOverBounds = True Then
                                irowspan = irowspan + 1
                            End If
                        End If
                        If bRowSpanPrint = True Then

                            swText.Write("| rowspan=" + (irowspan + 1).ToString() + " | " + iRank.ToString() + vbCrLf)
                        Else

                            swText.Write("| " + iRank.ToString() + vbCrLf)
                        End If
                    Else
                        irowspan = irowspan - 1
                    End If

                    swText.Write("| [[" + GGT(iReport)(z).sGratName + "]]" + vbCrLf)
                    swText.Write("|align=""center""| [[:Category:Meetup in " + GGT(iReport)(z).sGraticule + "|" + GGT(iReport)(z).sGraticule + "]]" + vbCrLf)






                    If iRowSpan1 < 1 Then
                        Dim bOverBounds As Boolean = False

                        If iRank + iRowSpan1 < GGT(iReport).Length() Then

                            While (GGT(iReport)(z).iGrandTot = GGT(iReport)(iRank + iRowSpan1).iGrandTot) And (bOverBounds = False)
                                iRowSpan1 = iRowSpan1 + 1
                                bRowSpanPrint = True
                                If iRank + iRowSpan1 > GGT(iReport).Length() - 1 Then
                                    iRowSpan1 = GGT(iReport).Length() - 1 - iRank
                                    bOverBounds = True
                                End If
                            End While
                            If bOverBounds = True Then
                                iRowSpan1 = iRowSpan1 + 1
                            End If
                        End If
                        If bRowSpanPrint = True Then

                            swText.Write("| rowspan=" + (iRowSpan1 + 1).ToString() + " | " + GGT(iReport)(z).iGrandTot.ToString() + vbCrLf)
                        Else

                            swText.Write("| " + GGT(iReport)(z).iGrandTot.ToString() + vbCrLf)
                        End If
                    Else
                        iRowSpan1 = iRowSpan1 - 1
                    End If


                    Select Case iReport
                        Case 0 'coordinates reached
                            Dim lDate As Long = 0
                            Dim lDays As Long = 0
                            Dim lWeeks As Long = 0
                            Dim lYears As Long = 0

                            lDate = Date.Today.Ticks() - GGT(iReport)(z).dtLast.Ticks()
                            lDate = dMaxDateTotals.Ticks() - GGT(iReport)(z).dtLast.Ticks()
                            lDays = lDate / TimeSpan.TicksPerDay

                            If lDays > 6 Then
                                lWeeks = Math.Truncate(lDays / 7.0)
                                lDays = lDays - (lWeeks * 7)

                                If lWeeks > 51 Then

                                    lYears = Math.Truncate(lWeeks / 52.0)
                                    lWeeks = lWeeks - (lYears * 52)
                                End If
                            End If

                            swText.Write("| ")
                            If lYears > 0 Then
                                swText.Write(lYears.ToString() + "y ")
                            End If
                            If lWeeks > 0 Then
                                swText.Write(lWeeks.ToString() + "w ")
                            End If
                            swText.Write(lDays.ToString() + "d")


                            swText.Write(GGT(iReport)(z).sLast.Substring(4) + vbCrLf)
                        Case 1 ' coordinates not reached
                            swText.Write("| " + GGT(iReport)(z).iRetroTot.ToString() + vbCrLf)

                    End Select

                    For iYearTotals As Integer = 2008 To iStartDisplayYear - 1
                        Dim iYearEndTots As Integer = 0

                        Dim iMonthBegin As Integer = (iYearTotals * 12) - ((2008 * 12) + 5)
                        Dim iMonthBeginOffset As Integer = 0
                        If iYearTotals = 2008 Then
                            iMonthBeginOffset = 4
                        End If

                        Dim iMonthEnd As Integer = iMaxMonths
                        If iMonthEnd > ((iYearTotals + 1) * 12) - ((2008 * 12) + 5) Then
                            iMonthEnd = ((iYearTotals + 1) * 12) - ((2008 * 12) + 5)
                        End If

                        For x = ((iMonthBegin + iMonthBeginOffset) + 1) To iMonthEnd
                            iYearEndTots = iYearEndTots + GGT(iReport)(z).iMonthTot(x)
                        Next
                        swText.Write("| " + iYearEndTots.ToString() + vbCrLf)
                    Next


                    For x = ((iBeginMonthTot + iMonthOffset) + 1) To iMaxMonthsPrint '8 To iMaxMonths
                        swText.Write("| " + GGT(iReport)(z).iMonthTot(x).ToString() + vbCrLf)
                    Next
                End If
            Next

            swText.Write("|}" + vbCrLf + vbCrLf)












            For iRecordsYear As Integer = iStartDisplayYear To dMaxDateTotals.Year() ' Date.Now().Year()

                swText.Write(":{| border=""2"" cellpadding=""4"" cellspacing=""0"" style=""margin-top:1em; margin-bottom:1em; background:#f9f9f9; border:1px #aaa solid; border-collapse:collapse;""" + vbCrLf)
                swText.Write("|+" + iRecordsYear.ToString() + " Monthly records, single graticule only" + vbCrLf)
                swText.Write("|-" + vbCrLf)
                swText.Write("!Month" + vbCrLf)
                swText.Write("!colspan=2|1st" + vbCrLf)
                swText.Write("!colspan=2|2nd" + vbCrLf)
                swText.Write("!colspan=2|3rd" + vbCrLf)

                iBeginMonthTot = (iRecordsYear * 12) - ((2008 * 12) + 5)
                iMonthOffset = 0
                If iRecordsYear = 2008 Then
                    iMonthOffset = 4
                End If
                iMaxMonthsPrint = iMaxMonths
                If iMaxMonthsPrint > ((iRecordsYear + 1) * 12) - ((2008 * 12) + 5) Then
                    iMaxMonthsPrint = ((iRecordsYear + 1) * 12) - ((2008 * 12) + 5)
                End If

                For x = ((iBeginMonthTot + iMonthOffset) + 1) To iMaxMonthsPrint
                    swText.Write("|-" + vbCrLf)
                    swText.Write("!" + MonthName(x - (iBeginMonthTot + iMonthOffset) + iMonthOffset) + vbCrLf)


                    Dim iTotMax(3) As Integer
                    Dim y As Integer
                    Dim bNotInserted As Boolean = False
                    For y = 0 To 3
                        iTotMax(y) = 0
                    Next

                    For iTotMonths As Integer = 0 To GGT(iReport).Length() - 1

                        bNotInserted = True
                        y = 0
                        While (bNotInserted) And (y < 3)

                            If (GGT(iReport)(iTotMonths).iMonthTot(x) > iTotMax(y)) Then
                                Dim itemp As Integer

                                For itemp = 2 To y + 1 Step -1
                                    iTotMax(itemp) = iTotMax(itemp - 1)
                                Next
                                iTotMax(y) = GGT(iReport)(iTotMonths).iMonthTot(x)
                                bNotInserted = False
                            End If

                            If (GGT(iReport)(iTotMonths).iMonthTot(x) = iTotMax(y)) Then
                                bNotInserted = False
                            End If

                            y = y + 1
                        End While
                    Next

                    For y = 0 To 2
                        swText.Write("|" + iTotMax(y).ToString() + vbCrLf)
                        If iTotMax(y) > 0 Then
                            swText.Write("|")
                            Dim bPrinted As Boolean = False
                            For iTotMonths As Integer = 0 To GGT(iReport).Length() - 1
                                If GGT(iReport)(iTotMonths).iMonthTot(x) = iTotMax(y) Then
                                    If bPrinted = True Then
                                        swText.Write("<br/>")
                                    Else
                                        bPrinted = True
                                    End If
                                    swText.Write("[[" + GGT(iReport)(iTotMonths).sGratName + "]]")
                                End If
                            Next
                            swText.Write(vbCrLf)
                        Else
                            swText.Write("|" + vbCrLf)
                        End If
                    Next
                Next

                swText.Write("|}" + vbCrLf + vbCrLf)

                swText.Write(vbCrLf + vbCrLf)

            Next





            swText.Write(":{| border=""2"" cellpadding=""4"" cellspacing=""0"" style=""margin-top:1em; margin-bottom:1em; background:#f9f9f9; border:1px #aaa solid; border-collapse:collapse;""" + vbCrLf)
            swText.Write("|+ Yearly Totals, single graticule only" + vbCrLf)

            For x = 0 To GGYT(iReport).GratYear.Length - 1
                swText.Write("!colspan=2|" + GGYT(iReport).GratYear(x).iYear.ToString() + vbCrLf)
            Next
            Dim iBMT As Integer = (dMaxDateTotals.Year * 12) - ((2008 * 12) + 5)
            Dim iMO As Integer = 0
            If dMaxDateTotals.Year = 2008 Then
                iMO = 4
            End If

            Dim iMMP As Integer = iMaxMonths
            If iMMP > ((dMaxDateTotals.Year() + 1) * 12) - ((2008 * 12) + 5) + 1 Then
                iMMP = ((dMaxDateTotals.Year() + 1) * 12) - ((2008 * 12) + 5)
            End If

            For x1 As Integer = ((iBMT + iMO) + 1) To iMMP
                Dim iMonth1 As Integer = x1 - (iBMT + iMO) + iMO
                If iMonth1 > 12 Then
                    iMonth1 = iMonth1 - 12
                End If
                swText.Write("!colspan=2|" + MonthName(iMonth1, True) + vbCrLf)
            Next


            Dim iLength1 As Integer = 0

            For x = 0 To GGYT(iReport).GratYear.Length - 1
                If GGYT(iReport).GratYear(x).GratTot.Length() > iLength1 Then
                    iLength1 = GGYT(iReport).GratYear(x).GratTot.Length - 1
                End If
            Next

            Select Case (dMaxDateTotals.DayOfWeek)

                Case DayOfWeek.Saturday
                Case DayOfWeek.Sunday
                Case Else
                    iLength1 = 100
            End Select

            For y As Integer = 0 To iLength1
                swText.Write("|-" + vbCrLf)
                For x = 0 To GGYT(iReport).GratYear.Length - 1
                    If y < GGYT(iReport).GratYear(x).GratTot.Length() Then

                        swText.Write("|" + GGYT(iReport).GratYear(x).GratTot(y).iNum.ToString() + vbCrLf)
                        swText.Write("|[[" + GGYT(iReport).GratYear(x).GratTot(y).sGratName + "]]<br/>")
                        swText.Write("[[:Category:Meetup in " + GGYT(iReport).GratYear(x).GratTot(y).sGraticule + "|" + GGYT(iReport).GratYear(x).GratTot(y).sGraticule + "]]" + vbCrLf)
                    Else
                        If y = GGYT(iReport).GratYear(x).GratTot.Length() Then
                            swText.Write("| colspan=2 rowspan=" + (iLength1 - y + 1).ToString() + " |" + vbCrLf)
                        End If
                    End If
                Next

                For x1 As Integer = ((iBMT + iMO) + 1) To iMMP
                    Dim iMonth1 As Integer = x1 - (iBMT + iMO) + iMO
                    If iMonth1 > 12 Then
                        iMonth1 = iMonth1 - 12
                    End If
                    If x1 < GGYT(iReport).GratMonYear.Length Then

                        If GGYT(iReport).GratMonYear(x1).GratTot Is Nothing = False Then

                            If y < GGYT(iReport).GratMonYear(x1).GratTot.Length Then
                                swText.Write("| " + GGYT(iReport).GratMonYear(x1).GratTot(y).iNum.ToString() + vbCrLf)
                                swText.Write("| [[" + GGYT(iReport).GratMonYear(x1).GratTot(y).sGratName + "]]<br/>")
                                swText.Write("[[:Category:Meetup in " + GGYT(iReport).GratMonYear(x1).GratTot(y).sGraticule + "|" + GGYT(iReport).GratMonYear(x1).GratTot(y).sGraticule + "]]" + vbCrLf)

                            Else
                                If y = GGYT(iReport).GratMonYear(x1).GratTot.Length Then
                                    swText.Write("| colspan=2 rowspan=" + (iLength1 - y + 1).ToString() + " |" + vbCrLf)
                                End If

                            End If

                        End If
                    End If
                Next

            Next
            swText.Write("|}" + vbCrLf + vbCrLf)







            Dim iMaxTotNum As Integer = 10 ' total number of overall results
            Dim iOverCount As Integer
            Dim bOverNotInserted As Boolean = False
            Dim bNoHeadingPrint As Boolean = False
            Dim iTotOverAllMax(iMaxTotNum) As Integer
            Dim OverAllReached(iMaxTotNum) As TotMaxNum

            For x = 0 To iMaxTotNum - 1
                OverAllReached(x).iTotal = 0
            Next



            For iTotalGrats As Integer = 0 To GGT(iReport).Length() - 1
                For iOverall As Integer = 0 To iMaxMonths
                    Dim iMonths As Integer = iMaxTotNum - 1

                    While (GGT(iReport)(iTotalGrats).iMonthTot(iOverall) > OverAllReached(iMonths).iTotal) And iMonths > 0
                        iMonths = iMonths - 1
                    End While
                    If (GGT(iReport)(iTotalGrats).iMonthTot(iOverall) > OverAllReached(iMonths).iTotal) Then

                        Dim iMonthsTemp As Integer = iMaxTotNum
                        While iMonthsTemp > iMonths
                            OverAllReached(iMonthsTemp) = OverAllReached(iMonthsTemp - 1)
                            iMonthsTemp = iMonthsTemp - 1
                        End While
                        If OverAllReached(iMonths).aTotMax Is Nothing = False Then
                            Array.Resize(OverAllReached(iMonths).aTotMax, 0)
                            OverAllReached(iMonths).aTotMax = Nothing
                        End If
                        If OverAllReached(iMaxTotNum).aTotMax Is Nothing = False Then
                            Array.Resize(OverAllReached(iMaxTotNum).aTotMax, 0)
                        End If

                        OverAllReached(iMonths).iTotal = GGT(iReport)(iTotalGrats).iMonthTot(iOverall)

                        If OverAllReached(iMonths).aTotMax Is Nothing Then
                            OverAllReached(iMonths).aTotMax = Array.CreateInstance(GetType(TotMax), 1)

                            OverAllReached(iMonths).aTotMax(0).iMonth = iOverall
                            OverAllReached(iMonths).aTotMax(0).sGraticule = GGT(iReport)(iTotalGrats).sGratName

                        Else


                            Dim iLength As Integer = OverAllReached(iMonths).aTotMax.Length()

                            Array.Resize(OverAllReached(iMonths).aTotMax, iLength + 1)

                            OverAllReached(iMonths).aTotMax(iLength).iMonth = iOverall
                            OverAllReached(iMonths).aTotMax(iLength).sGraticule = GGT(iReport)(iTotalGrats).sGratName

                        End If



                    Else
                        If GGT(iReport)(iTotalGrats).iMonthTot(iOverall) = OverAllReached(iMonths).iTotal Then

                            If OverAllReached(iMonths).aTotMax Is Nothing = False Then

                                Dim iLength As Integer = OverAllReached(iMonths).aTotMax.Length()

                                Array.Resize(OverAllReached(iMonths).aTotMax, iLength + 1)
                                OverAllReached(iMonths).aTotMax(iLength).iMonth = iOverall
                                OverAllReached(iMonths).aTotMax(iLength).sGraticule = GGT(iReport)(iTotalGrats).sGratName
                            Else

                                OverAllReached(iMonths).aTotMax = Array.CreateInstance(GetType(TotMax), 1)

                                OverAllReached(iMonths).aTotMax(0).iMonth = iOverall
                                OverAllReached(iMonths).aTotMax(0).sGraticule = GGT(iReport)(iTotalGrats).sGratName
                            End If

                        Else
                            If (GGT(iReport)(iTotalGrats).iMonthTot(iOverall) > OverAllReached(iMonths + 1).iTotal) Then

                                Dim iMonthsTemp As Integer = iMaxTotNum
                                While iMonthsTemp > iMonths + 1
                                    OverAllReached(iMonthsTemp) = OverAllReached(iMonthsTemp - 1)
                                    iMonthsTemp = iMonthsTemp - 1
                                End While
                                If OverAllReached(iMonths + 1).aTotMax Is Nothing = False Then
                                    Array.Resize(OverAllReached(iMonths + 1).aTotMax, 0)
                                    OverAllReached(iMonths + 1).aTotMax = Nothing
                                End If
                                If OverAllReached(iMaxTotNum).aTotMax Is Nothing = False Then
                                    Array.Resize(OverAllReached(iMaxTotNum).aTotMax, 0)
                                End If

                                OverAllReached(iMonths + 1).iTotal = GGT(iReport)(iTotalGrats).iMonthTot(iOverall)

                                If OverAllReached(iMonths + 1).aTotMax Is Nothing Then
                                    OverAllReached(iMonths + 1).aTotMax = Array.CreateInstance(GetType(TotMax), 1)

                                    OverAllReached(iMonths + 1).aTotMax(0).iMonth = iOverall
                                    OverAllReached(iMonths + 1).aTotMax(0).sGraticule = GGT(iReport)(iTotalGrats).sGratName

                                End If
                            End If
                        End If
                    End If

                Next
            Next








            swText.Write(":{| border=""2"" cellpadding=""4"" cellspacing=""0"" style=""margin-top:1em; margin-bottom:1em; background:#f9f9f9; border:1px #aaa solid; border-collapse:collapse;""" + vbCrLf)
            swText.Write("|+Overall record" + vbCrLf)
            swText.Write("|-" + vbCrLf)
            swText.Write("! Place" + vbCrLf)
            swText.Write("! Month" + vbCrLf)
            swText.Write("! Grat " + vbCrLf)
            swText.Write("! Total" + vbCrLf)

            For iOverCount = 0 To iMaxTotNum
                iTotOverAllMax(iOverCount) = 0
            Next

            bNoHeadingPrint = False
            For iOverCount = 0 To (iMaxTotNum - 1)
                If OverAllReached(iOverCount).aTotMax Is Nothing = True Then

                    iMaxTotNum = iOverCount
                End If

            Next

            For iOverCount = 0 To (iMaxTotNum - 1)
                Dim iPlace As Integer = iOverCount + 1
                swText.Write("|-" + vbCrLf)
                swText.Write("!" + (iOverCount + 1).ToString() + vbCrLf)


                If OverAllReached(iOverCount).aTotMax.Length > 1 Then

                    For x = 0 To OverAllReached(iOverCount).aTotMax.Length - 2
                        For y As Integer = x + 1 To OverAllReached(iOverCount).aTotMax.Length - 1
                            If OverAllReached(iOverCount).aTotMax(y).iMonth < OverAllReached(iOverCount).aTotMax(x).iMonth Then
                                Dim tTotMax As TotMax = Nothing

                                For z = y To x + 1 Step -1

                                    tTotMax = OverAllReached(iOverCount).aTotMax(z - 1)
                                    OverAllReached(iOverCount).aTotMax(z - 1) = OverAllReached(iOverCount).aTotMax(z)
                                    OverAllReached(iOverCount).aTotMax(z) = tTotMax

                                Next
                            End If
                        Next

                    Next
                End If

                bNoHeadingPrint = False
                For iOverall As Integer = 0 To OverAllReached(iOverCount).aTotMax.Length - 1
                    If bNoHeadingPrint = False Then
                        swText.Write("| ")
                        bNoHeadingPrint = True
                    Else
                        swText.Write("<br/>")
                    End If
                    Dim iYearOverall As Integer
                    Dim iMonthOverall As Integer
                    If OverAllReached(iOverCount).aTotMax(iOverall).iMonth < 8 Then
                        iYearOverall = 2008
                        iMonthOverall = OverAllReached(iOverCount).aTotMax(iOverall).iMonth + 5
                    Else
                        Dim iTemp As Integer = OverAllReached(iOverCount).aTotMax(iOverall).iMonth - 8

                        iYearOverall = Math.Truncate(iTemp / 12) + 2009

                        iMonthOverall = Math.Truncate(iTemp / 12)
                        iMonthOverall = (OverAllReached(iOverCount).aTotMax(iOverall).iMonth - 8) - (Math.Truncate(iTemp / 12) * 12) + 1
                    End If
                    swText.Write(MonthName(iMonthOverall))
                    swText.Write(" " + iYearOverall.ToString())
                Next
                swText.Write(vbCrLf)

                bNoHeadingPrint = False
                For iOverall As Integer = 0 To OverAllReached(iOverCount).aTotMax.Length - 1
                    If bNoHeadingPrint = False Then
                        swText.Write("| ")
                        bNoHeadingPrint = True
                    Else
                        swText.Write("<br/>")
                    End If
                    swText.Write("[[" + OverAllReached(iOverCount).aTotMax(iOverall).sGraticule + "]]")
                Next
                swText.Write(vbCrLf)

                swText.Write("|" + OverAllReached(iOverCount).iTotal.ToString() + vbCrLf)
            Next


            swText.Write("|}" + vbCrLf)






            swText.Flush()
            swText.Close()
            swText.Dispose()



        Next


    End Sub

End Module

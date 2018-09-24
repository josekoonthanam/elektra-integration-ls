Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports AdminBoard3._5




Public Class BookingAdapter
    Implements IMessageAdapter(Of CRMMessage.LSBooking)
    'Private Variable
    Private salesOfficeCode As String = String.Empty

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Load(ByVal id As Integer) As Production.Booking
        Return New Production.Booking(id)
    End Function

    Public Shared Function getcommunicationmanageridemailnumbersenddatebybookingid(ByVal bookingid As Integer, ByVal destinationcode As String, ByVal productcode As String) As DataView
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_GetCommunicationManagerIDEmailNumberSendDatebyBookingId
        Dim ds As New DataSet
        ' set parameters
        With cmd.Parameters
            .Booking_id = bookingid
            .DestinationCode = destinationcode
            .ProductCode = productcode
        End With
        ' execute query
        cm.Fill(ds, cmd, Constants.ElektraConnectionNamespace)
        Return ds.Tables(0).DefaultView
    End Function

    'Added for Auto Confirmation Changes
    Public Shared Sub DoAutoConfirmation(ByRef bkn As Production.Booking, ByRef coursebookingsForPoseidonConfirmationList As List(Of Production.CourseBooking),
                                         ByRef accArticlesForPoseidonConfirmationList As List(Of Production.BookingArticle),
                                         ByRef wl As BusinessObjects.WarningListBase)
        Try


            Dim modifiedCourseBookingsInDestination As New Hashtable
            Dim visaT4Destinations As New List(Of String) 'VISAT4
            Dim ok As New Production.AirportSearch

            Dim courseBookingData As New CourseBookingData

            'Sort Booking Articles By Start Week Code
            Dim sortedBookingArticleList As New List(Of Production.BookingArticle)

            For Each ba As Production.BookingArticle In bkn.BookingArticleList
                If Not String.IsNullOrEmpty(ba.DisplayTab.Trim()) Then
                    If ba.DisplayTab.Trim().ToLower() = ArticleLookup.DisplayTabs.Accommodation.Trim().ToLower() AndAlso
                    Not ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Cancelled Then
                        sortedBookingArticleList.Add(ba)
                    End If
                End If

                If Not String.IsNullOrEmpty(ba.DestinationCode.Trim()) Then
                    If Not modifiedCourseBookingsInDestination.ContainsKey(ba.DestinationCode.Trim()) Then
                        modifiedCourseBookingsInDestination.Add(ba.DestinationCode.Trim(), getModifiedBookingForAutoconfirmation("", bkn.Booking_id, "", "",
                                                                                                                               ba.DestinationCode.Trim(), "", "Mod"))
                    End If
                End If

                If String.Compare(ba.Article.Code.Trim(), "VISAT4", True) = 0 Then
                    If Not visaT4Destinations.Contains(ba.DestinationCode.Trim()) Then
                        visaT4Destinations.Add(ba.DestinationCode.Trim())
                    End If
                End If
            Next

            Dim tempba As Production.BookingArticle
            If sortedBookingArticleList.Count > 1 Then
                For outer As Integer = sortedBookingArticleList.Count - 1 To 0 Step -1
                    For inner As Integer = 0 To outer - 1
                        If sortedBookingArticleList(inner).StartWeekCode > sortedBookingArticleList(inner + 1).StartWeekCode Then
                            tempba = sortedBookingArticleList(inner)
                            sortedBookingArticleList(inner) = sortedBookingArticleList(inner + 1)
                            sortedBookingArticleList(inner + 1) = tempba
                        End If
                    Next
                Next
            End If

            'For each course, check if it can be auto confirmed
            ' Check to be done, check rules and  then check reservation
            ' If reservation not applicable, check RTI and auto confirm the course
            For Each cb As Production.CourseBooking In bkn.CourseBookingList
                'Rule 1: we are excluding Cancelled Courses --- They go through manual process
                If Not cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled Then
                    'AndAlso Not cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Confirmed Then
                    Dim accomodationArticleEligibleList As New List(Of Production.BookingArticle)
                    If cb.ISPIE = True And cb.CourseTypeCode.Trim().ToUpper() <> "MG" Then
                        AutoConfirmCourseBooking(cb, accomodationArticleEligibleList, coursebookingsForPoseidonConfirmationList,
                                                                                 accArticlesForPoseidonConfirmationList, bkn)
                    Else
                        'checking the flag for autoconfirmation rule
                        Dim courseHasCapacity As Boolean = True

                        For i As Integer = 0 To sortedBookingArticleList.Count - 1
                            If sortedBookingArticleList(i).DestinationCode = cb.CourseParent.DestinationCode Then
                                ' Check if article is applicable for this course
                                If sortedBookingArticleList(i).StartDate < cb.EndDate AndAlso
                                        sortedBookingArticleList(i).EndDate > cb.StartDate Then
                                    accomodationArticleEligibleList.Add(sortedBookingArticleList(i))
                                End If
                            End If
                        Next

                        For i As Integer = 0 To accomodationArticleEligibleList.Count - 1
                            'Rule 16: Check if accomodation article eligible for Auto confirmation
                            If isArticleEligibleForRTI(accomodationArticleEligibleList(i).ArticleCode) = 1 Then
                                'Rule 17: All group bookings which are not of course type Leader, need not check RTI
                                If Not bkn.IsReservation AndAlso (String.IsNullOrEmpty(bkn.GroupCode) Or
                                            (Not String.IsNullOrEmpty(bkn.GroupCode.Trim()) AndAlso cb.CourseParent.CourseTypeCode.Trim().ToUpper() = "L")) Then
                                    'Rule 18: Check RTI for capacity
                                    If accomodationArticleEligibleList(i).StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Active Then
                                        If Not IsCapacityAvailableInRTI(accomodationArticleEligibleList(i).DestinationCode.Trim(),
                                                                     accomodationArticleEligibleList(i).ArticleCode.Trim(),
                                                                     accomodationArticleEligibleList(i).StartDate, accomodationArticleEligibleList(i).EndDate,
                                                                     (New DateTime(accomodationArticleEligibleList(0).StartDate.Subtract(cb.Booking.Student.BirthDate).Ticks)).Year - 1,
                                                                     bkn.SalesBookingId, bkn.SalesOfficeCode, wl) Then
                                            courseHasCapacity = False
                                            Exit For
                                        End If
                                    End If
                                End If
                            Else
                                courseHasCapacity = False
                                Exit For
                            End If

                        Next
                        'for API calls
                        courseBookingData.CourseBooking = cb
                        courseBookingData.modifiedCourseBookingsInDestination = modifiedCourseBookingsInDestination
                        courseBookingData.visaT4Destinations = visaT4Destinations
                        If CourseTypeLookup.FindInOriginalList(cb.CourseTypeCode.Trim()).IsAutoConfirmed = 0 Then
                            courseBookingData.IsAutoConfCourse = False
                        Else
                            courseBookingData.IsAutoConfCourse = True
                        End If
                        courseBookingData.WhatChnaged = modifiedCourseBookingsInDestination(cb.DestinationCode.Trim())
                        courseBookingData.HasCapacityAvailable = courseHasCapacity

                        ' AutoConfirmationRulesCheck     
                        Dim autoConfirmRulecheckResult As Boolean
                        autoConfirmRulecheckResult = False
                        Dim manager = New AdminBoard3._5.ABManager

                        autoConfirmRulecheckResult = manager.AutoConfirmationRules(courseBookingData)
                        ' boolResult = AutoConfirmationRules(courseBookingData)

                        ' Check rules
                        'If IsAutoConfirmationApplicable(cb, modifiedCourseBookingsInDestination, visaT4Destinations) Then
                        If autoConfirmRulecheckResult Then
                            'Auto Confirm
                            AutoConfirmCourseBooking(cb, accomodationArticleEligibleList, coursebookingsForPoseidonConfirmationList,
                                                         accArticlesForPoseidonConfirmationList, bkn)


                        End If
                    End If

                End If
            Next

            Dim destinationsWhereAllCoursesAreNotConfirmed As New List(Of String)
            For Each cb As Production.CourseBooking In bkn.CourseBookingList
                If Not cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication AndAlso
                    Not cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Confirmed Then
                    If Not destinationsWhereAllCoursesAreNotConfirmed.Contains(cb.DestinationCode.Trim()) Then
                        destinationsWhereAllCoursesAreNotConfirmed.Add(cb.DestinationCode.Trim())
                    End If
                End If
            Next

            ' Confirm all transport and accomodation articles not confirmed, for all destinations where all courses are confirmed
            If Not bkn.StatusCode = BookingStatusLookup.BookingStatuses.PreApplication Then
                For Each ba As Production.BookingArticle In bkn.BookingArticleList
                    If ba.ISPIE = True Then
                        ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Confirmed
                        ba.StatusDate = Now
                        ba.AcceptedDate = Now
                        'ELEK-6127
                        If Not ba.IsAutoConfirmed Then
                            ba.IsAutoConfirmed = True
                        End If
                        If ba.DisplayTab = ArticleLookup.DisplayTabs.Accommodation Then
                            accArticlesForPoseidonConfirmationList.Add(ba)
                        End If
                    Else
                        If Not destinationsWhereAllCoursesAreNotConfirmed.Contains(ba.DestinationCode.Trim()) AndAlso
                                              ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Active AndAlso
                                              ((ba.DisplayTab = ArticleLookup.DisplayTabs.Accommodation AndAlso (Not ba.AccSupplierTypeCode Is Nothing AndAlso (ba.AccSupplierTypeCode <> "HF" AndAlso
                                              ba.AccSupplierTypeCode <> "HS" AndAlso ba.AccSupplierTypeCode <> "RE" AndAlso ba.AccSupplierTypeCode <> "HO" AndAlso ba.AccSupplierTypeCode <> "NA" AndAlso ba.AccSupplierTypeCode <> "OA"))) Or
                                              ba.DisplayTab = ArticleLookup.DisplayTabs.Transportation) Then
                            If Not ba.Article.ExtraNights Then
                                ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Confirmed
                                ba.StatusDate = Now
                                ba.AcceptedDate = Now
                                'ELEK-6127
                                If Not ba.IsAutoConfirmed Then
                                    ba.IsAutoConfirmed = True
                                End If
                                If ba.DisplayTab = ArticleLookup.DisplayTabs.Accommodation Then
                                    accArticlesForPoseidonConfirmationList.Add(ba) 'ELEK-6162
                                End If
                            End If
                        End If
                    End If


                Next
            End If

        Catch ex As Exception
            wl.Add(BusinessObjects.WarningSeverities.Info, "Auto Confirmation Exception Error:" + ex.Message, ex.Source)
        End Try

    End Sub

    'Added for Auto Confirmation Changes
    Public Shared Function IsAutoConfirmationApplicable(ByRef cb As Production.CourseBooking, ByRef modifiedCourseBookingsInDestination As Hashtable, ByRef visaT4Destinations As List(Of String)) As Boolean
        'Rule Checker, All rules are exclusion rules
        'ELEK-8515 stop auto-confirmations for AU and NZ schools

        If String.Compare(cb.CourseParent.DestinationCode.Trim(), "AU-BRS", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "AU-SYD", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "NZ-AUC", True) = 0 Then
            If String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "G", True) = 0 AndAlso cb.Weeks > 12 Then
                Return False
            End If
        End If

        'ELEK-8720 Update auto-confirmation rules in CAC
        If cb.CourseParent.DestinationCode.Trim().ToUpper().Equals("GB-CAC") Then
            Return False
        End If

        'ELEK-8596 stop auto-confirmation with Together With in US
        For Each bl As BookingArticle In cb.Booking.BookingArticleList
            If String.IsNullOrEmpty(bl.TogetherWith) Then
                Continue For
            Else
                Return False
            End If
        Next

        'ELEK-8320 stop auto-confirmations for Vancouver
        If String.Compare(cb.CourseParent.DestinationCode.Trim(), "CA-VAN", True) = 0 Then
            Return False
        End If
        'ELEK - 8310 - No auto-confirmation for bookings less than 8 days
        Dim Span As TimeSpan = cb.EndDate.Subtract(cb.StartDate)
        If (Span.Days < 8) Then
            Return False
        End If

        ' ELEK -8174 so if course code is MG (Mini group then do not confired for all school
        If cb.CourseParent.CourseTypeCode.Trim().ToUpper() = "MG" Then
            Return False
        End If
        ' If any rule is true, then auto confirmation is false

        'ELEK-9019
        If String.Compare(cb.CourseParent.DestinationCode.Trim(), "IE-DUB", True) = 0 Then
            Dim noOfWeeks As Int16
            For Each cl As CourseBooking In cb.Booking.CourseBookingList
                If cl.DestinationCode.Trim().ToUpper().Equals("IE-DUB") And cl.EndDate >= Date.Now Then
                    noOfWeeks = noOfWeeks + cl.Weeks
                End If
            Next
            If noOfWeeks >= 36 Then
                Return False
            End If
        End If

        'ELEK-9011
        For Each bl As BookingArticle In cb.Booking.BookingArticleList
            If bl.DisplayTab = ArticleLookup.DisplayTabs.Accommodation Then
                Dim age As Integer = 0
                age = (New DateTime(bl.StartDate.Subtract(cb.Booking.Student.BirthDate).Ticks)).Year - 1
                If age < 18 And (bl.ArticleCode.ToUpper().Trim() = "OA" Or bl.ArticleCode.ToUpper().Trim() = "OAH" Or bl.ArticleCode.ToUpper().Trim() = "OAY" Or bl.ArticleCode.ToUpper().Trim() = "OAYX" _
                    Or bl.ArticleCode.ToUpper().Trim() = "NA" Or bl.ArticleCode.ToUpper().Trim() = "NAABNB" Or bl.ArticleCode.ToUpper().Trim() = "NACE" Or bl.ArticleCode.ToUpper().Trim() = "NACL" _
                    Or bl.ArticleCode.ToUpper().Trim() = "NADBM" Or bl.ArticleCode.ToUpper().Trim() = "NAH" Or bl.ArticleCode.ToUpper().Trim() = "NAY" Or bl.ArticleCode.ToUpper().Trim() = "NAYX") Then
                    Return False
                End If
            End If
        Next

        'exclude Basic course for Auto Confirmation in Dublin ELEK-8043
        If String.Compare(cb.CourseParent.DestinationCode.Trim(), "IE-DUB", True) = 0 And cb.CourseParent.CourseTypeCode.Trim() = "B" Then
            Return False
        End If

        'Rule 3: All Language Learning Solutions bookings, Product: CLT, Programme code: EFC
        'ELEK-8681 Stop auto-confirmations for EFC courses
        'If cb.CourseParent.ProductCode.Trim() = ProductLookup.Products.CorporateLanguage Then
        '    If cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.EFCorporate AndAlso
        '        (String.Compare(cb.CourseParent.DestinationCode.Trim(), "US-BOE", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "GB-CAE", True) = 0) Then
        '        Return True
        '    End If
        'End If

        'Rule 26 implemented for ELEK-6737
        'Rule 26: Courses are not auto confirmed, if in Elektra’s Course Type Lookup Table column IsAutoConfirmed = 0
        If CourseTypeLookup.FindInOriginalList(cb.CourseTypeCode.Trim()).IsAutoConfirmed = 0 Then
            Return False
        End If

        ''Rule 2: All International Academy bookings, Product: International Academy, Programme codes: AL and IB
        'If String.Compare(cb.CourseParent.ProductCode.Trim(), "IA", True) = 0 Then
        '    If cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.ALevel Or String.Compare(cb.CourseParent.ProgramCode.Trim(), "IB", True) = 0 Then
        '        Return False
        '    End If
        'End If

        'ELEK-8739 IA Booking Confirmation Transfer should be stopped
        If String.Compare(cb.CourseParent.ProductCode.Trim(), "IA", True) = 0 Then
            Return False
        End If

        'Rule 3: All Language Learning Solutions bookings, Product: CLT, Programme code: EFC
        If cb.CourseParent.ProductCode.Trim() = ProductLookup.Products.CorporateLanguage Then
            If cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.EFCorporate Then
                Return False
            End If
        End If

        'Rule 4: All University Preparation bookings, Product: APP, Programme: UP and UPP
        If cb.CourseParent.ProductCode.Trim() = ProductLookup.Products.APP Then
            If cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.UniversityPreparationProfessionals Or cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.Brittin Then
                Return False
            End If
        End If

        'Rule 5: Bookings with Special Requirements – information entered in Poseidon
        If cb.Booking.IsSpecialRequirement Then
            Return False
        End If
        'Added disabled
        If cb.Booking.IsDisabled Then
            Return False
        End If
        'Never auto confirmed for Singalpore school
        If String.Compare(cb.CourseParent.DestinationCode.Trim(), "SG-SIN", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "SG-SIM", True) = 0 Then
            Return False
        End If
        'Rule 6: For Juniors- applicable in schools that accept Juniors, Product: LS, Programme: ILS and ILC, Course type JU or JI, destination code: BOU, CAC, CAM, BST, DUB, MLT, MSJ, NYC, SBA, SEA, HON, VAN, MAL, NIC, SIN, BRI, AUC
        If cb.CourseParent.ProductCode.Trim() = ProductLookup.Products.LanguageSchools Then
            If cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.InternationalLanguageSchool Or cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.InternationalLanguageClub Then
                If String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "JU", True) = 0 Or String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "JI", True) = 0 Then
                    If String.Compare(cb.CourseParent.DestinationCode.Trim(), "GB-BOU", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "GB-CAC", True) = 0 Or
                        String.Compare(cb.CourseParent.DestinationCode.Trim(), "GB-CAM", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "GB-BST", True) = 0 Or
                        String.Compare(cb.CourseParent.DestinationCode.Trim(), "IE-DUB", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "MT-MSJ", True) = 0 Or
                        String.Compare(cb.CourseParent.DestinationCode.Trim(), "US-NYC", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "US-SBA", True) = 0 Or
                        String.Compare(cb.CourseParent.DestinationCode.Trim(), "US-SEA", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "CA-VAN", True) = 0 Or
                        String.Compare(cb.CourseParent.DestinationCode.Trim(), "FR-NIC", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "SG-SIN", True) = 0 Or
                        String.Compare(cb.CourseParent.DestinationCode.Trim(), "GB-BRI", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "NZ-AUC", True) = 0 Or
                        String.Compare(cb.CourseParent.DestinationCode.Trim(), "US-HNL", True) = 0 Or String.Compare(cb.CourseParent.DestinationCode.Trim(), "ES-MAA", True) = 0 Then
                        Return False
                    End If
                End If
            End If
        End If

        If modifiedCourseBookingsInDestination.ContainsKey(cb.DestinationCode.Trim()) Then
            If IsSomething(modifiedCourseBookingsInDestination(cb.DestinationCode.Trim())) Then
                Dim dt As New DataTable()
                dt = modifiedCourseBookingsInDestination(cb.DestinationCode.Trim())
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        If Convert.ToInt32(dr("coursebooking_id").ToString()) = cb.CourseBooking_id Then
                            ' There is a modification

                            ' Check if accomodation has changed
                            'Rule 7: Once confirmed bookings where the accommodation article is changed but there is already an allocation for the initial request
                            ' or any of booking note or medical note or special diet note or allergic note or matching criteria changed and there is 
                            ' an allocation
                            If dr("whatchanged").ToString().Contains("Accomodation Details Changed") Or
                                dr("whatchanged").ToString().Contains("Booking Note Changed") Or
                                dr("whatchanged").ToString().Contains("Medical Note Changed") Or
                                dr("whatchanged").ToString().Contains("Special Diet Note Changed") Or
                                dr("whatchanged").ToString().Contains("Allergic Note Changed") Or
                                dr("whatchanged").ToString().Contains("Matching Criteria Changed") Then
                                If Not String.IsNullOrEmpty(dr("supplier_id").ToString()) Then
                                    Return False
                                End If
                            End If

                            ' Check if course has changed
                            'Rule 8: Once confirmed bookings where the date changes but there is academic allocation
                            If dr("whatchanged").ToString().Contains("Course Details Changed") Then
                                ' If academic allocation
                                If isClassPresent(cb.CourseBooking_id) = 1 Then
                                    Return False
                                End If
                            End If

                            'Rule 9: Check if transfer details have changed
                            If dr("whatchanged").ToString().Contains("Transfer Details Changed") Then
                                Return False
                            End If

                            'Rule 10: Check if article details have changed
                            If dr("whatchanged").ToString().Contains("Article Details Changed") Then
                                Return False
                            End If

                            'Rule 11: Check if gender has changed
                            If dr("whatchanged").ToString().Contains("Gender Changed") Then
                                Return False
                            End If

                            'Rule 27: Age has changed
                            If dr("whatchanged").ToString().Contains("Age Changed.") Then
                                Return False
                            End If

                            ' Check for VISAT4 and Sevis Rules for modified course bookings
                            'Rule 12: Bookings with Tier4 article and Sevis Articles
                            If visaT4Destinations.Contains(cb.DestinationCode.Trim()) Or String.Compare(cb.VisaTypeCode.Trim(), "I-20", True) = 0 Then
                                Return False
                            End If

                            'ELEK-6152
                            If cb.CourseParent.DestinationCode.ToUpper().Contains("US") Then
                                'Rule 19: US Schools for I-20 regulations --- All courses under the Academic Year (AY) and Academic Year Professionals (AYP) programme
                                If cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.AcademicYear Or
                                    cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.AcademicYearProfessionals Then
                                    Return False
                                End If

                                'Rules 20 - 23: US Schools for I-20 regulations --- All Exam courses (Course Code: E), All Volunteer courses (Course Code: IP)
                                ', All Intensive courses (Course Code: I)
                                If String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "E", True) = 0 Or String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "IP", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "I", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MBE", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MBE", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MBE", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MBE", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEDM", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEIA", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEFD", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MHT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MART", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MIT", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEHF", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MBE", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MBE", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MBE", True) = 0 Or
                                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MBE", True) = 0 Then
                                    Return False
                                End If

                                'Rule 24: US Schools for I-20 regulations --- General courses if the durations is 10 weeks or longer
                                If String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "G", True) = 0 Then
                                    If cb.Weeks >= 10 Then
                                        Return False
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            ' Check for VISAT4 and Sevis Rules for Active/New course bookings, check if destination has multiple courses or not
            'Rule 12: Bookings with Tier4 article and Sevis Articles
            If visaT4Destinations.Contains(cb.DestinationCode.Trim()) Or String.Compare(cb.VisaTypeCode.Trim(), "I-20", True) = 0 Then
                If hasMultipleCourses(cb.Booking.Booking_id, cb.DestinationCode.Trim()) = 1 Then
                    Return False
                End If
            End If

            'ELEK-6152
            If cb.CourseParent.DestinationCode.ToUpper().Contains("US") AndAlso Not String.Compare(cb.VisaTypeCode.Trim(), "I-20", True) = 0 Then
                'Rule 19: US Schools for I-20 regulations --- All courses under the Academic Year (AY) and Academic Year Professionals (AYP) programme
                If cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.AcademicYear Or
                    cb.CourseParent.ProgramCode.Trim() = ProgramLookup.Programs.AcademicYearProfessionals Then
                    Return False
                End If

                'Rules 20 - 23: US Schools for I-20 regulations --- All Exam courses (Course Code: E), All Volunteer courses (Course Code: IP)
                ', All Intensive courses (Course Code: I)
                If String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "E", True) = 0 Or String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "IP", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "I", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE01MBE", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE04MBE", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE06MBE", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "SE09MBE", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEDM", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEIA", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEFD", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MHT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MART", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MIT", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MEHF", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE01MBE", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE04MBE", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE06MBE", True) = 0 Or
                    String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "YE09MBE", True) = 0 Then
                    Return False
                End If

                'Rule 24: US Schools for I-20 regulations --- General courses if the durations is 10 weeks or longer
                If String.Compare(cb.CourseParent.CourseTypeCode.Trim(), "G", True) = 0 Then
                    If cb.Weeks >= 10 Then
                        Return False
                    End If
                End If
            End If

        End If

        Return True
    End Function

    'Added for Auto Confirmation Changes
    Public Shared Sub AutoConfirmCourseBooking(ByRef cb As Production.CourseBooking, ByRef baList As List(Of Production.BookingArticle),
                                               ByRef coursebookingsForPoseidonConfirmationList As List(Of Production.CourseBooking),
                                               ByRef accArticlesForPoseidonConfirmationList As List(Of Production.BookingArticle), ByRef bkn As Production.Booking)
        'Confirm Booking similar to what PoseidonUpdate does

        'update course booking
        If Not cb.IsAutoConfirmed Then
            cb.IsAutoConfirmed = True
        End If

        If cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Active Then
            cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Confirmed
            cb.AcceptedDate = Now
            CourseBookingChangeManager.AddChange(cb, CourseBookingChangeTypeLookup.CourseBookingChangeTypes.Confirmation, Now)
            cb.SetStatus(cb.StatusCode, CRMMessage.Constants.IntegrationUser)
        End If
        coursebookingsForPoseidonConfirmationList.Add(cb)

        'update accomodation article
        For Each ba As Production.BookingArticle In baList
            If ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Active Then
                If ba.ISPIE = True Then
                    ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Confirmed
                    ba.StatusDate = Now
                    ba.AcceptedDate = Now
                    'ELEK-6127
                    If Not ba.IsAutoConfirmed Then
                        ba.IsAutoConfirmed = True
                    End If
                Else
                    If Not ba.Article.ExtraNights Then
                        ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Confirmed
                        ba.StatusDate = Now
                        ba.AcceptedDate = Now
                        'ELEK-6127
                        If Not ba.IsAutoConfirmed Then
                            ba.IsAutoConfirmed = True
                        End If
                    End If
                End If
                If Not ba.Article.ExtraNights Then
                    ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Confirmed
                    ba.StatusDate = Now
                    ba.AcceptedDate = Now
                    'ELEK-6127
                    If Not ba.IsAutoConfirmed Then
                        ba.IsAutoConfirmed = True
                    End If
                End If
            End If
            accArticlesForPoseidonConfirmationList.Add(ba)  'ELEK-6162
        Next
    End Sub

    'Added for Auto Confirmation Changes
    Public Shared Function IsCapacityAvailableInRTI(ByVal DestinationCode As String, ByVal ArticleCode As String, ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal Age As Integer, ByVal SalesBookingId As String, ByVal SalesOfficeCode As String, ByRef wl As BusinessObjects.WarningListBase) As Boolean
        'Contact RTI to check availability
        Try
            Dim isAvailable As Boolean = False
            Dim rti As New RTIService.RTIAutoConfirmationServiceClient
            isAvailable = rti.IsWithinMargin(DestinationCode, ArticleCode, StartDate, EndDate, SalesBookingId, Age, SalesOfficeCode)
            If Not isAvailable Then
                wl.Add(BusinessObjects.WarningSeverities.Info, "RTI not available for this booking. DestinationCode:" + DestinationCode + ", ArticleCode:" +
                       ArticleCode + ", StartDate:" + StartDate.ToString() + ", EndDate:" + EndDate.ToString() + ", Age:" + Age.ToString() + ", SalesBookingId:" +
                       SalesBookingId + ", SalesOfficeCode:" + SalesOfficeCode)
            End If
            Return isAvailable
        Catch ex As Exception
            wl.Add(BusinessObjects.WarningSeverities.Info, "Auto Confirmation Exception Error, RTI Service Error: " + ex.Message, ex.Source)
            Return False
        End Try
    End Function

    'Added for Auto Confirmation Changes
    Public Shared Sub SendConfirmationToPoseidon(ByRef coursebookingsForPoseidonConfirmationList As List(Of Production.CourseBooking),
                                                 ByRef accArticlesForPoseidonConfirmationList As List(Of Production.BookingArticle),
                                                 ByVal booking_id As String, ByVal salesPersonEmailId As String,
                                                 ByRef wl As BusinessObjects.WarningListBase)
        'Send to Poseidon
        Try
            Dim courseDataContractList As New List(Of EF.Language.Elektra.Client.ServiceProxy.CAS.CASService.CourseDataContract)
            Dim aricleDataContractList As New List(Of EF.Language.Elektra.Client.ServiceProxy.CAS.CASService.BookingArticleDataContract)
            Dim ConfirmedAt As String = "Auto"

            For Each cb As Production.CourseBooking In coursebookingsForPoseidonConfirmationList
                Dim courseDataContract As New EF.Language.Elektra.Client.ServiceProxy.CAS.CASService.CourseDataContract()
                courseDataContract.PoseidonBookingId = cb.Booking.PoseidonBooking_Id
                courseDataContract.CourseStartDate = cb.StartDate
                courseDataContract.CourseEndDate = cb.EndDate
                courseDataContract.DestinationCode = cb.CourseParent.DestinationCode
                courseDataContract.IsAutoConfirmed = True
                courseDataContract.SalesOfficeCode = cb.Booking.SalesOfficeCode
                courseDataContract.SalesRegion = cb.Booking.SalesRegion
                If cb.Booking.StatusCode = BookingStatusLookup.BookingStatuses.PreApplication Then
                    courseDataContract.Status = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication
                Else
                    courseDataContract.Status = cb.StatusCode
                End If

                courseDataContractList.Add(courseDataContract)
            Next

            'ELEK-6162 --- send accomodationinfo also to poseidon
            If accArticlesForPoseidonConfirmationList.Count > 0 Then
                If Not accArticlesForPoseidonConfirmationList(0).Booking.StatusCode = BookingStatusLookup.BookingStatuses.PreApplication Then
                    For Each ba As Production.BookingArticle In accArticlesForPoseidonConfirmationList
                        Dim bknArticleDataContract As New EF.Language.Elektra.Client.ServiceProxy.CAS.CASService.BookingArticleDataContract
                        bknArticleDataContract.PoseidonBookingId = ba.Booking.PoseidonBooking_Id
                        bknArticleDataContract.StartDate = ba.StartDate
                        bknArticleDataContract.EndDate = ba.EndDate
                        bknArticleDataContract.DestinationCode = ba.DestinationCode
                        bknArticleDataContract.IsAutoConfirmed = True
                        bknArticleDataContract.SalesOfficeCode = ba.Booking.SalesOfficeCode
                        bknArticleDataContract.SalesRegion = ba.Booking.SalesRegion
                        bknArticleDataContract.Status = ba.StatusCode
                        bknArticleDataContract.ArticleCode = ba.ArticleCode
                        aricleDataContractList.Add(bknArticleDataContract)
                    Next
                End If
            End If

            Dim sndConfirmationToPoseidon As New EF.Language.Elektra.Client.Model.CAS.CASModelSource()
            Dim serviceStatus As Boolean = False
            serviceStatus = sndConfirmationToPoseidon.SaveAutoConfirmBookingToPoseidon(courseDataContractList, aricleDataContractList, booking_id, salesPersonEmailId,
                                                                       ConfirmedAt)

            wl.Add(BusinessObjects.WarningSeverities.Info, "Auto Confirmation send confirmation to poseidon --- Elektra Internal --- Cas Service status:" +
                   serviceStatus.ToString())

        Catch ex As Exception
            wl.Add(BusinessObjects.WarningSeverities.Info, "Auto Confirmation Exception at send confirmation to poseidon --- Elektra Internal Error:" +
                   ex.Message, ex.Source)
        End Try
    End Sub

    'Added for Auto Confirmation Changes
    'Function IsBookingNew --- to check if a booking is new or it is an update to existing booking
    Public Shared Function IsNewBooking(ByVal salesBookingId As String) As Boolean
        Dim bookingId As Integer

        bookingId = GetBookingIDBySalesBookingId(salesBookingId)

        If bookingId <> Null.IntegerNull Then
            Return False
        Else
            Return True
        End If
    End Function

    'Public Shared Function Load(ByVal salesBookingId As String, Optional ByVal createNew As Boolean = True) As Production.Booking
    '    Dim bkn As Production.Booking
    '    Dim bookingId As Integer

    '    bookingId = GetBookingIDBySalesBookingId(salesBookingId)
    '    If bookingId <> Null.IntegerNull Then
    '        bkn = New Production.Booking(bookingId)
    '    ElseIf createNew Then
    '        bkn = CreateNewBooking()
    '    Else
    '        bkn = Nothing
    '    End If
    '    Return bkn
    'End Function
    Public Shared Function Load(ByVal salesBookingId As String, ByVal customerId As Integer, ByVal salesRegion As String, Optional ByVal createNew As Boolean = True) As Production.Booking
        Dim bkn As Production.Booking
        Dim bookingId As Integer

        bookingId = GetBookingIDBySalesBookingId(salesBookingId)

        If bookingId <> Null.IntegerNull Then
            bkn = New Production.Booking(bookingId)
        ElseIf createNew Then
            bkn = CreateNewBooking(customerId, salesRegion)
        Else
            bkn = Nothing
        End If
        Return bkn
    End Function

    Public Shared Function GetBookingIDBySalesBookingId(ByVal salesBookingId As String) As Integer
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_BookingFindBySalesBookingId
        ' Set parameters
        With cmd.Parameters
            .SalesBookingId = salesBookingId
            'output parameter so set the value as null
            .Booking_id = Null.IntegerNull
        End With

        ' Execute query
        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)

        Return cmd.Parameters.Booking_id
    End Function

    'Auto Confirmation Changes
    Public Shared Function isArticleEligibleForRTI(ByVal articleCode As String) As Integer
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_IsAccomodationArticleEligibleForRTI
        ' Set parameters
        With cmd.Parameters
            .ArticleCode = articleCode
            'output parameter so set the value as null
            .IsEligible = Null.IntegerNull
        End With

        ' Execute query
        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)

        Return cmd.Parameters.IsEligible
    End Function

    'Auto Confirmation Changes
    Public Shared Function isClassPresent(ByVal CourseBooking_Id As Integer) As Integer
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_IsClassPresent
        ' Set parameters
        With cmd.Parameters
            .CourseBooking_Id = CourseBooking_Id
            'output parameter so set the value as null
            .IsPresent = Null.IntegerNull
        End With

        ' Execute query
        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)

        Return cmd.Parameters.IsPresent
    End Function

    'Auto Confirmation Changes
    Public Shared Function hasMultipleCourses(ByVal Booking_Id As Integer, ByVal DestinationCode As String) As Integer
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_HasMultipleCourses
        ' Set parameters
        With cmd.Parameters
            .Booking_id = Booking_Id
            .DestinationCode = DestinationCode
            'output parameter so set the value as null
            .HasCourses = Null.IntegerNull
        End With

        ' Execute query
        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)

        Return cmd.Parameters.HasCourses
    End Function

    'Auto Confirmation Changes
    Public Shared Function getModifiedBookingForAutoconfirmation(ByVal GroupStatus As String, ByVal Bookingid As Integer, ByVal GroupCode As String,
                                                                     ByVal PROGRAMCODE As String, ByVal Destination As String, ByVal COURSEWEEK As String,
                                                                     ByVal STATUS As String) As DataTable
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_GetModifiedBookingForAutoconfirmation
        Dim dt As New DataTable
        ' Set parameters
        With cmd.Parameters
            .GroupStatus = GroupStatus
            .Bookingid = Bookingid
            .GroupCode = GroupCode
            .PROGRAMCODE = PROGRAMCODE
            .Destination = Destination
            .STATUS = STATUS
            .COURSEWEEK = COURSEWEEK
        End With

        ' Execute query
        cm.Fill(dt, cmd, Constants.ElektraConnectionNamespace)

        Return dt
    End Function

    Public Shared Function GetStudentIdBySalesCustomerIdAndSalesRegion(ByVal customerId As Integer, ByVal salesregion As String) As Integer
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_GetStudentIdBySalesCustomerIdAndSalesRegion
        ' Set parameters
        With cmd.Parameters
            .SalesCustomerId = customerId
            .SalesRegion = salesregion
            'output parameter so set the value as null
            .Student_id = Null.IntegerNull
        End With

        ' Execute query
        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)

        Return cmd.Parameters.Student_id
    End Function
    'Public Shared Function CreateNewBooking() As Production.Booking
    '    Dim bkn As Production.Booking
    '    Dim stu As New Production.Student
    '    bkn = New Production.Booking(stu)
    '    Return bkn
    'End Function

    Public Shared Function CreateNewBooking(ByVal customerId As Integer, ByVal salesRegion As String) As Production.Booking
        Dim studentId As Integer
        Dim student As Student
        Dim booking As Production.Booking
        studentId = GetStudentIdBySalesCustomerIdAndSalesRegion(customerId, salesRegion)
        If studentId > 0 Then
            student = New Production.Student(studentId)
        Else
            student = New Production.Student()
        End If
        booking = New Production.Booking(student)
        Return booking
    End Function

    Public Function GetFightItin_id(ByVal salesFlightItinId As String)
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New asp_FlightItinFindBySalesFlightItinId

        ' Set parameters
        With cmd.Parameters
            .SalesFlightItinId = salesFlightItinId
            .FlightItin_id = Null.IntegerNull
        End With

        ' Execute query
        cm.ExecuteNonQuery(cmd, ProdLookups.LookupConstants.MainDbNamespace)

        Return cmd.Parameters.FlightItin_id
    End Function
    ''' <summary>
    '''  Updates the value from the msg-derived Individual object to the 
    '''  mapping Individual object loaded from the production system.
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="sourceSysCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update(ByVal msg As CRMMessage.LSBooking, ByVal sourceSysCode As String) _
            As ResultStepInfo Implements IMessageAdapter(Of CRMMessage.LSBooking).Update

        Dim stepInfo As ResultStepInfo
        stepInfo = Me.UpdateLSBooking(msg, sourceSysCode)
        Return stepInfo
    End Function
    Private Function SalesOfficeCodeStartsWith(ByVal countryCode As String) As Boolean
        If Me.salesOfficeCode.StartsWith(countryCode.Trim) Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function IsRebooking(ByVal msg As CRMMessage.LSBooking) As Boolean
        For Each item As BookingLine In msg.BookingLineItems
            If item.StatusCode = "RBK" Then
                Return True
            End If
            Return False
        Next
    End Function

    Private Function HandleUpdateInfo(ByVal wl As BusinessObjects.WarningListBase _
                                        , ByVal isStaleMsg As Boolean) As ResultStepInfo
        Dim iEnum As IEnumerator
        Dim stepInfo As ResultStepInfo

        stepInfo = New ResultStepInfo

        If isStaleMsg Then
            stepInfo.Status = IntegrationResult.Stale

        ElseIf Not wl.HasError Then
            stepInfo.Status = IntegrationResult.Success

        Else
            stepInfo.Status = IntegrationResult.Failure
        End If

        If wl.Count > 0 Then
            iEnum = wl.GetEnumerator

            While iEnum.MoveNext
                stepInfo.Errors.Add(iEnum.Current.ToString)
            End While

        End If

        Return stepInfo

    End Function

    Private Function UpdateLSBooking(ByVal msg As CRMMessage.LSBooking, ByVal sourceSysCode As String) _
       As ResultStepInfo

        Dim poseidonElektraBooking As New PoseidonElektraBookingData()
        Dim adminboard As New ABManager()
        Dim bookingMatching As New BookingMatching
        Dim bkn As Production.Booking
        Dim bookingHistory As Production.Booking
        Dim stepInfo As ResultStepInfo
        Dim stepInfo2 As ResultStepInfo
        Dim wl As BusinessObjects.WarningListBase
        Dim isStaleMsg As Boolean
        Dim stuAdapter As New StudentAdapter
        Dim bknLineAdapter As BookingLineAdapter
        Dim deletedList As New ArrayList

        Dim isSpecialReqt As Boolean = False
        Dim isDisabled As Boolean = False
        Dim PoseidonGroupCode As String = ""
        wl = New BusinessObjects.WarningListBase
        isStaleMsg = False

        Dim GenderCode As String = String.Empty
        'Auto Confirmation Changes
        Dim acWl As BusinessObjects.WarningListBase
        acWl = New BusinessObjects.WarningListBase
        Dim countWS1T1ArticleElektra As Int16 = 0
        Dim countWS1T1ArticlePoseidon As Int16 = 0
        Dim WS1T1Autoconfirmaiton As Boolean = False
        Dim ElektraTogetherWith As String = String.Empty
        Dim PosTogetherWith As String = String.Empty
        Dim IsPosVisaExits As Boolean = False
        Dim IsElekVisaExits As Boolean = False


        Try
            ' check for BOF destination and if yes convert it to BOI for AY and UP programs           
            If msg.BookingLineItems.Count > 0 Then
                For Each bknLine As CRMMessage.BookingLine In msg.BookingLineItems
                    If (((bknLine.ProgramCode.Trim() = "AY") Or (bknLine.ProgramCode.Trim() = "UP") Or (bknLine.ProgramCode.Trim() = "AYP") Or (bknLine.ProgramCode.Trim() = "UPP") Or (bknLine.ProgramCode.Trim() = "ILSP") Or (bknLine.ProgramCode.Trim() = "MLYP")) And (bknLine.DestinationCode.Contains("BOF"))) Then
                        bknLine.DestinationCode = "US-BOI"
                    End If
                Next
            End If

            If msg.BookingLineChangeItems.Count > 0 Then
                For Each bknLineChange As CRMMessage.BookingLineChange In msg.BookingLineChangeItems
                    If (((msg.ProgramCode.Trim() = "AY") Or (msg.ProgramCode.Trim() = "UP") Or (msg.ProgramCode.Trim() = "AYP") Or (msg.ProgramCode.Trim() = "UPP") Or (msg.ProgramCode.Trim() = "ILSP") Or (msg.ProgramCode.Trim() = "MLYP")) And (bknLineChange.DestinationCode.Contains("BOF"))) Then
                        bknLineChange.DestinationCode = "US-BOI"
                    End If
                Next
            End If

            If msg.ArticleBookingLineItems.Count > 0 Then
                For Each bknLineArticle As CRMMessage.BookingLine In msg.ArticleBookingLineItems
                    If (((bknLineArticle.ProgramCode.Trim() = "AY") Or (bknLineArticle.ProgramCode.Trim() = "UP") Or (bknLineArticle.ProgramCode.Trim() = "AYP") Or (bknLineArticle.ProgramCode.Trim() = "UPP") Or (bknLineArticle.ProgramCode.Trim() = "ILSP") Or (bknLineArticle.ProgramCode.Trim() = "MLYP")) And (bknLineArticle.DestinationCode.Contains("BOF"))) Then
                        bknLineArticle.DestinationCode = "US-BOI"
                    End If
                Next
            End If

            If msg.BookingLineItems.Count > 0 Then
                If msg.PoseidonGroup_Id > 0 Then
                    For Each bknLine As CRMMessage.BookingLine In msg.BookingLineItems
                        bknLine.DestinationCode = msg.GroupDestinationCode
                    Next
                End If
            End If

            If msg.BookingLineChangeItems.Count > 0 Then
                If msg.PoseidonGroup_Id > 0 Then
                    For Each bknLineChange As CRMMessage.BookingLineChange In msg.BookingLineChangeItems
                        bknLineChange.DestinationCode = msg.GroupDestinationCode
                    Next
                End If
            End If

            If msg.ArticleBookingLineItems.Count > 0 Then
                If msg.PoseidonGroup_Id > 0 Then
                    For Each bknLineArticle As CRMMessage.BookingLine In msg.ArticleBookingLineItems
                        bknLineArticle.DestinationCode = msg.GroupDestinationCode
                    Next
                End If
            End If

            If msg.BookingLineItems.Count > 0 Then
                For Each bknLine As CRMMessage.BookingLine In msg.BookingLineItems
                    If bknLine.CourseNumber.Trim() = "SG" AndAlso (bknLine.DestinationCode <> "SG-SIN" Or bknLine.DestinationCode <> "SG-SIM") Then
                        deletedList.Add(bknLine)
                    ElseIf bknLine.CourseNumber.Trim() = "SG" AndAlso (bknLine.DestinationCode = "SG-SIN" Or bknLine.DestinationCode = "SG-SIM") AndAlso bknLine.StartDate.Year <> 2014 Then
                        deletedList.Add(bknLine)
                    End If
                Next
                For Each bknLine As CRMMessage.BookingLine In deletedList
                    msg.BookingLineItems.Remove(bknLine)
                Next
                deletedList = New ArrayList
            End If

            If msg.BookingLineItems.Count > 0 Then
                For Each bknLine As CRMMessage.BookingLine In msg.BookingLineItems
                    If Not DoesDestinationExistInElektra(bknLine.DestinationCode) Then
                        deletedList.Add(bknLine)
                    End If
                Next
                For Each bknLine As CRMMessage.BookingLine In deletedList
                    msg.BookingLineItems.Remove(bknLine)
                Next
                deletedList = New ArrayList
            End If

            If msg.ArticleBookingLineItems.Count > 0 Then
                For Each bknLine As CRMMessage.BookingLine In msg.ArticleBookingLineItems
                    If Not DoesDestinationExistInElektra(bknLine.DestinationCode) Then
                        deletedList.Add(bknLine)
                    End If
                Next
                For Each bknLine As CRMMessage.BookingLine In deletedList
                    msg.ArticleBookingLineItems.Remove(bknLine)
                Next
                deletedList = New ArrayList
            End If

            If msg.BookingLineChangeItems.Count > 0 Then
                For Each bknLineChange As CRMMessage.BookingLineChange In msg.BookingLineChangeItems
                    If Not DoesDestinationExistInElektra(bknLineChange.DestinationCode) Then
                        deletedList.Add(bknLineChange)
                    End If
                Next
                For Each bknLineChange As CRMMessage.BookingLineChange In deletedList
                    msg.BookingLineChangeItems.Remove(bknLineChange)
                Next
                deletedList = New ArrayList
            End If

            If msg.BookingLineItems.Count > 0 Then
                For Each bknLine As CRMMessage.BookingLine In msg.BookingLineItems
                    If msg.GroupProgramCode = "LT" And msg.GroupProgramCode <> "" Then
                        deletedList.Add(bknLine)
                    End If
                Next
                For Each bknLine As CRMMessage.BookingLine In deletedList
                    msg.BookingLineItems.Remove(bknLine)
                Next
                deletedList = New ArrayList
            End If

            bkn = Load(msg.BookingNum, msg.Customer_id, msg.InstanceName)
            bookingHistory = Load(msg.BookingNum, msg.Customer_id, msg.InstanceName, False)

            If IsSomething(bkn) Then
                isSpecialReqt = bkn.IsSpecialRequirement
                isDisabled = bkn.IsDisabled
                PoseidonGroupCode = bkn.PoseidonGroupCode
                GenderCode = bkn.Student.GenderCode
                IsPosVisaExits = msg.NeedsVisa
                For Each cb As CourseBooking In bkn.CourseBookingList
                    If Not String.IsNullOrEmpty(cb.VisaTypeCode) And cb.VisaTypeCode.Trim().ToUpper() = "I-20" Then
                        IsElekVisaExits = True
                        Exit For
                    End If
                Next


                If IsSomething(bkn.CourseBookingList) And bkn.CourseBookingList.Count > 0 Then
                    For Each cb As CourseBooking In bkn.CourseBookingList
                        For Each ba As BookingArticle In bkn.BookingArticleList
                            If Not String.IsNullOrEmpty(ba.TogetherWith) AndAlso Not ba.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Cancelled Then
                                ElektraTogetherWith = ba.TogetherWith.Trim().ToUpper()
                                Exit For
                            End If
                        Next
                    Next
                End If

                If IsSomething(msg) Then
                    If IsSomething(msg.BookingLineItems) AndAlso msg.BookingLineItems.Count > 0 Then
                        For Each msgbknLine As CRMMessage.BookingLine In msg.BookingLineItems
                            If msgbknLine.DestinationCode.Trim().ToUpper().StartsWith("US-") Then
                                For Each msgarticle As BookingLine In msg.ArticleBookingLineItems
                                    If Not msgarticle.TogetherWith Is Nothing Then
                                        If msgarticle.TogetherWith.Count > 0 AndAlso Not msgarticle.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Cancelled AndAlso msgarticle.DestinationCode.Trim().ToUpper().StartsWith("US-") Then
                                            PosTogetherWith = msgarticle.TogetherWith(0).ToString().Trim().ToUpper()
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If

                If (Not IsPosVisaExits) And (IsElekVisaExits) Then
                    For Each bl As CourseBooking In bkn.CourseBookingList
                        If bl.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Confirmed AndAlso bl.StatusCode <> CourseBookingStatusLookup.CourseBookingStatuses.Cancelled Then
                            bl.StatusCode = "AC"
                            bl.IsAutoConfirmed = False
                            bl.StatusDate = Now
                            bl.AcceptedDate = Nothing
                            bl.ConfirmedDate = New Date(1800, 1, 1)
                        End If
                    Next
                End If


                If ((Not PosTogetherWith.Equals(String.Empty) Or Not ElektraTogetherWith.Equals(String.Empty)) AndAlso Not PosTogetherWith.Equals(ElektraTogetherWith)) Then
                    For Each bl As CourseBooking In bkn.CourseBookingList
                        If bl.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Confirmed Then
                            bl.StatusCode = "AC"
                            bl.IsAutoConfirmed = False
                            bl.StatusDate = Now
                            bl.AcceptedDate = Nothing
                            bl.ConfirmedDate = New Date(1800, 1, 1)
                        End If
                    Next
                    For Each bl As BookingArticle In bkn.BookingArticleList
                        If bl.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Confirmed Then
                            bl.StatusCode = "AC"
                            bl.IsAutoConfirmed = False
                            bl.StatusDate = Now
                            bl.AcceptedDate = Nothing
                        End If
                    Next
                End If

                If IsSomething(bkn.BookingArticleList) AndAlso bkn.BookingArticleList.Count > 0 Then
                    For Each balist As BookingArticle In bkn.BookingArticleList
                        If IsSomething(balist.ArticleCode) And balist.ArticleCode.Trim().ToUpper().StartsWith("WS1T1") Then
                            countWS1T1ArticleElektra = countWS1T1ArticleElektra + 1
                        End If
                    Next
                End If
                If IsSomething(msg) AndAlso IsSomething(msg.ArticleBookingLineItems) AndAlso msg.ArticleBookingLineItems.Count > 0 Then
                    For Each bli As BookingLine In msg.ArticleBookingLineItems
                        If IsSomething(bli.ArticleCode) AndAlso bli.ArticleCode.Trim().ToUpper().StartsWith("WS1T1") Then
                            countWS1T1ArticlePoseidon = countWS1T1ArticlePoseidon + 1
                        End If
                    Next
                End If

                If (countWS1T1ArticlePoseidon > countWS1T1ArticleElektra) Then
                    WS1T1Autoconfirmaiton = True
                End If

            End If

            'ELEK -9753
            Dim PosbookingItems As New List(Of BookingLine)
            If IsSomething(msg.BookingLineItems) And msg.BookingLineItems.Count > 1 Then
                For Each cb As BookingLine In msg.BookingLineItems
                    If cb.StatusCode <> BookingArticleStatusLookup.BookingArticleStatuses.Cancelled AndAlso cb.DestinationCode.Trim().ToUpper().Equals("KR-SEL") Then
                        PosbookingItems.Add(cb)
                    End If
                Next
            End If

            'ELEK -9753
            If IsSomething(PosbookingItems) And PosbookingItems.Count > 1 Then
                Dim _previousItmList As New List(Of BookingLine)
                Dim _bookingForActive As New List(Of BookingLine)

                PosbookingItems.Sort(Function(a, b) a.StartDate.CompareTo(b.StartDate))
                For Each BookedItem As BookingLine In PosbookingItems
                    If _previousItmList.Count = 0 Then
                        _previousItmList.Add(BookedItem)
                    Else
                        Dim item As BookingLine = _previousItmList(0)
                        Dim _weeksGapInDays = Math.Abs((BookedItem.StartDate - item.EndDate).Days)
                        If _weeksGapInDays <= 28 Then
                            If Not _bookingForActive.Contains(item) Then
                                _bookingForActive.Add(item)
                            End If
                            If Not _bookingForActive.Contains(BookedItem) Then
                                _bookingForActive.Add(BookedItem)
                            End If
                        Else
                            If Not _bookingForActive.Contains(item) Then
                                _bookingForActive.Add(item)
                            End If
                            If _bookingForActive.Count > 0 Then
                                AutoConfirmWithTotalWeeksCalculate(_bookingForActive, bkn.CourseBookingList)
                            End If
                        End If

                        _previousItmList = New List(Of BookingLine)
                        _previousItmList.Add(BookedItem)
                        Dim LastItem As Int16 = PosbookingItems.Count - 1
                        If BookedItem.Equals(PosbookingItems(LastItem)) Then
                            If Not _bookingForActive.Contains(BookedItem) Then
                                _bookingForActive.Add(BookedItem)
                            End If
                        End If
                    End If
                Next
                If _bookingForActive.Count > 0 Then
                    AutoConfirmWithTotalWeeksCalculate(_bookingForActive, bkn.CourseBookingList)
                End If
            End If


            'History Tracking for Admin Board - Poseidon(Modified), Elektra(Existing)
            Try
                poseidonElektraBooking.bookingId = bkn.Booking_id
                If (Not IsNothing(bookingHistory)) Then
                    bookingMatching = adminboard.GetElektraBookingMatching(bookingHistory.Booking_id)
                    poseidonElektraBooking.elektraData = bookingHistory
                End If
                poseidonElektraBooking.poseidonData = msg
                poseidonElektraBooking.elektraMatchingNotes = bookingMatching
            Catch ex As Exception
                wl.Add(BusinessObjects.WarningSeverities.Info, "Adding data - Track Booking Changes Exception: " + ex.Message + " Stack Trace: " + Convert.ToString(ex.StackTrace), ex.Source)
            End Try

            UpdateLSBookingProperties(msg, bkn)
            UpdateFlights(msg, bkn, wl)

            bknLineAdapter = New BookingLineAdapter
            If WS1T1Autoconfirmaiton = True Then
                For Each bl As CourseBooking In bkn.CourseBookingList
                    bl.StatusCode = "AC"
                    bl.IsAutoConfirmed = False
                    bl.StatusDate = Now
                    bl.AcceptedDate = Nothing
                    bl.ConfirmedDate = New Date(1800, 1, 1)
                Next
            End If
            If ((Not isSpecialReqt AndAlso bkn.IsSpecialRequirement) OrElse (Not isDisabled AndAlso bkn.IsDisabled) OrElse (msg.Customer.GenderCode <> bkn.Student.GenderCode) OrElse (msg.Customer.DateOfBirth <> bkn.Student.BirthDate)) Then
                For Each bl As CourseBooking In bkn.CourseBookingList
                    bl.StatusCode = "AC"
                    bl.IsAutoConfirmed = False
                    bl.StatusDate = Now
                    bl.AcceptedDate = Nothing
                    bl.ConfirmedDate = New Date(1800, 1, 1)
                Next
                For Each bl As BookingArticle In bkn.BookingArticleList
                    bl.StatusCode = "AC"
                    bl.IsAutoConfirmed = False
                    bl.StatusDate = Now
                    bl.AcceptedDate = Nothing
                Next
            End If

            'ELEK - ELEK-7664
            Dim dtGetFirstCoruseAndDOBForBooking As DataTable = BookingAdapter.GetFirstCourseAndDOBForBooking(bkn.SalesBookingId)
            stepInfo = bknLineAdapter.UpdateFromMsg(bkn, msg)

            msg.Customer.InstanceName = msg.InstanceName

            stuAdapter.Booking = bkn
            stepInfo2 = stuAdapter.Update(msg.Customer, sourceSysCode)

            '******* Dont remove - If removed history tracking for Admin Board doesnt work *******
            Console.WriteLine("Booking Id: {0}", bookingMatching.BookingId)
            If (Not IsNothing(bookingHistory)) Then
                Console.WriteLine("Existing Age: {0}", bookingHistory.Student.CurrentAge)
                Console.WriteLine("Existing Course count: {0}", bookingHistory.CourseBookingList.Count)
            End If
            Console.WriteLine("Modified Course count: {0}", bkn.CourseBookingList.Count)
            If (Not IsNothing(bookingHistory)) Then
                Console.WriteLine("Existing Article count: {0}", bookingHistory.BookingArticleList.Count)
            End If
            Console.WriteLine("Modified Article count: {0}", bkn.BookingArticleList.Count)
            Console.WriteLine("")
            '******* Dont remove *******

            If stepInfo.Status = IntegrationResult.Success AndAlso stepInfo2.Status = IntegrationResult.Success Then
                bkn.Save(CRMMessage.Constants.IntegrationUser, wl)
            End If

            If stepInfo.Status = IntegrationResult.Success Then

                'AutoConfirmation changes, start
                Try
                    bkn = New Production.Booking(bkn.Booking_id)
                    Dim coursebookingsForPoseidonConfirmationList As New List(Of Production.CourseBooking)
                    Dim bookingInfo As New StudentMatching()
                    Dim accArticlesForPoseidonConfirmationList As New List(Of Production.BookingArticle)    'ELEK-6162
                    DoAutoConfirmation(bkn, coursebookingsForPoseidonConfirmationList, accArticlesForPoseidonConfirmationList, acWl)
                    'Elek-7960 StudentMatchingnotes
                    bookingInfo.IsVegetarian = msg.IsVegetarian
                    bookingInfo.IsWheelChairUser = msg.IsWheelChaired
                    bookingInfo.CarriesMedicine = Not String.IsNullOrEmpty(msg.CarriesMedicationNote) 'msg.CarriesMedication
                    bookingInfo.HasInsectsAllergy = msg.HasInsectsAllergy
                    bookingInfo.HasNutsAllergy = msg.HasNutAllergic
                    bookingInfo.HasOtherAllergy = msg.HasOtherAllergies
                    bookingInfo.HasOtherDietaryNeed = msg.HasSpecialDietNotes
                    bookingInfo.HasOtherDisability = msg.HasOtherDisabilities
                    bookingInfo.HasOtherMedicalNeed = msg.HasOtherMedicalNeeds
                    bookingInfo.HasPenicillinAllergy = msg.NeedsPencillin
                    bookingInfo.HasPetAllergy = msg.HasPetsAllergy
                    bookingInfo.HasPollenAllergy = msg.HasPollenAllergy
                    bookingInfo.IsAsthmatic = msg.IsAsthmatic
                    bookingInfo.IsBedwetter = msg.IsBedwetter
                    bookingInfo.IsDeaf = msg.IsDeaf
                    bookingInfo.IsDiabetic = msg.IsDiabetic
                    bookingInfo.IsEpileptic = msg.IsEpileptic
                    bookingInfo.IsGlutenFree = msg.IsGlutenFree
                    bookingInfo.IsHalal = msg.IsHalal
                    bookingInfo.IsKosher = msg.IsKosher
                    If Not String.IsNullOrEmpty(msg.IsMilkFree) AndAlso msg.IsMilkFree Then
                        bookingInfo.IsLactoseFree = msg.IsMilkFree
                    Else
                        bookingInfo.IsLactoseFree = msg.IsLactoseFree
                    End If
                    bookingInfo.IsVegan = msg.IsVegan
                    bookingInfo.IsVisuallyImpaired = msg.IsVisuallyImpaired
                    bookingInfo.Medication = msg.CarriesMedicationNote
                    bookingInfo.OtherDisabilityNote = msg.OtherDisabilityNote
                    bookingInfo.PreferFamilyWithChildren = msg.PrefersFamilyWithChildren
                    bookingInfo.PreferPets = msg.PreferswithPets
                    bookingInfo.PreferPiano = msg.PrefersPiano
                    bookingInfo.PreferSmokingInHouse = msg.PrefersSmokinginHouse
                    bookingInfo.HasDustAllergy = msg.HasDustAllergy
                    bookingInfo.IsSpecialDiet = msg.IsSpecialDiet
                    bookingInfo.IsAllergic = msg.IsAllergic
                    bookingInfo.UserId = CRMMessage.Constants.IntegrationUser.Id
                    bookingInfo.IsSmoker = msg.IsSmoker
                    bookingInfo.DietaryNote = msg.OtherSpecialDietNotes
                    bkn.StudentMatchingConfirmed.SaveIntoBookingInfoTable(bookingInfo)
                    If coursebookingsForPoseidonConfirmationList.Count > 0 Or accArticlesForPoseidonConfirmationList.Count > 0 Then
                        If Not bkn.StatusCode = BookingStatusLookup.BookingStatuses.PreApplication Then
                            bkn.StudentMatchingConfirmed.UpdateFromBooking(True)    'ELEK-6127                                              
                            bkn.Save(CRMMessage.Constants.IntegrationUser, wl)
                        End If
                        SendConfirmationToPoseidon(coursebookingsForPoseidonConfirmationList, accArticlesForPoseidonConfirmationList, bkn.Booking_id,
                                                   bkn.SalesPersonEmailId, acWl)
                    End If
                Catch ex As Exception
                    acWl.Add(BusinessObjects.WarningSeverities.Error, ex.Message, ex.Source)
                End Try
                'AutoConfirmation changes, end

                If msg.ArticleBookingLineItems.Count > 0 Then
                    For Each baItems As CRMMessage.BookingLine In msg.ArticleBookingLineItems
                        If baItems.Description.Contains("EX-") Then
                            Dim dv As New DataView
                            Dim communicationManager_id As Integer
                            Dim EmailNumber As String
                            Dim SendDate As DateTime
                            Dim productCode As String
                            EmailNumber = String.Empty

                            productCode = ProgramLookup.FindInOriginalListByProductAndProgramCode(baItems.ProgramCode, msg.ProductCode).ProductCode

                            dv = getcommunicationmanageridemailnumbersenddatebybookingid(bkn.Booking_id, baItems.DestinationCode, productCode)
                            If IsSomething(dv) Then
                                For Each drv As DataRowView In dv
                                    communicationManager_id = Convert.ToInt32(drv("CommunicationManager_id"))
                                    EmailNumber = Convert.ToInt32(drv("EmailNumber"))
                                    SendDate = Convert.ToDateTime(drv("SendDate"))
                                    If baItems.StartDate > SendDate Then
                                        Dim statusDataContract As New EF.Language.Elektra.Client.ServiceProxy.CommunicationManager.CourseExtension.StatusDataContract()
                                        Dim statusDataContractList As New List(Of EF.Language.Elektra.Client.ServiceProxy.CommunicationManager.CourseExtension.StatusDataContract)
                                        Dim dem As New EF.Language.Elektra.Client.Model.Communications.CommunicationsModelSource()
                                        statusDataContract.BookingId = bkn.Booking_id
                                        statusDataContract.CommunicationManagerId = communicationManager_id
                                        statusDataContract.EmailStatus = 7
                                        statusDataContract.Comments = String.Empty
                                        statusDataContract.EmailNumber = EmailNumber
                                        statusDataContract.Id = "D54FAB26-4054-49C5-BFE2-8CD7F03EB09E"
                                        statusDataContract.EmailId = "Integration@PoseidonII.com"
                                        statusDataContract.UpdatedBy = 1
                                        statusDataContractList.Add(statusDataContract)
                                        dem.UpdateStatus(statusDataContractList)
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            End If
            'ELEK-7664
            ExcludeJUJICheckList(bkn, msg, dtGetFirstCoruseAndDOBForBooking)
            'Start insert destination info if generic doc exists
            GenericDocumentInsert(bkn.Booking_id)
            'End insert estination info if generic doc exists
            If (Not String.IsNullOrEmpty(bkn.PoseidonGroupCode) Or (PoseidonGroupCode <> bkn.PoseidonGroupCode)) Then
                Dim isGroupRemoved As Boolean = False
                Dim isLeaderBooking As Boolean = False
                If IsSomething(bkn) And IsSomething(bkn.CourseBookingList) And bkn.CourseBookingList.Count > 0 Then
                    For Each cb As CourseBooking In bkn.CourseBookingList
                        If (cb.IsCourseLeader) Then
                            isLeaderBooking = True
                            Exit For
                        End If
                    Next
                End If

                If (String.IsNullOrEmpty(bkn.PoseidonGroupCode)) Then
                    isGroupRemoved = True
                End If
                PreGroupBookingCancelled(bkn.PoseidonGroupCode, bkn.SalesBookingId, bkn.StatusCode, isGroupRemoved, PoseidonGroupCode, isLeaderBooking)
            End If

            If Not String.IsNullOrEmpty(GenderCode) And Not String.IsNullOrEmpty(bkn.Student.GenderCode) AndAlso GenderCode.Trim().ToUpper() <> bkn.Student.GenderCode.Trim().ToUpper() Then
                UpdateAccomArticleIsCaxWhenGenderChanged(bkn.Booking_id, GenderCode, bkn.Student.GenderCode)
            End If

            If IsSomething(bkn) And IsSomething(bkn.CourseBookingList) And bkn.CourseBookingList.Count > 0 Then
                For Each cb As CourseBooking In bkn.CourseBookingList
                    UpdateLSTransporationChanges(cb.CourseBooking_id, cb.Booking_id)
                Next
            End If
        Catch ex As Exception
            wl.Add(BusinessObjects.WarningSeverities.Error, ex.Message, ex.Source)
            ExceptionManager.Publish(ex)
        Finally
            'Auto Confirmation Changes
            wl.AddRange(acWl)

            'Dim sb As New System.Text.StringBuilder()
            'sb.AppendLine("DateTime:" + Now.ToString())
            'For Each w As BusinessObjects.WarningBase In wl
            '    sb.AppendLine(w.ToString())
            'Next

            'If (wl.Count > 0) Then
            '    Using outfile As New System.IO.StreamWriter("D:\error" + msg.BookingNum + ".txt")
            '        outfile.Write(sb.ToString())
            '    End Using
            'End If

            stepInfo = Me.HandleUpdateInfo(wl, isStaleMsg)
        End Try
        Try
            'History tracking - Admin Board - Elektra(Modified)
            poseidonElektraBooking.modifiedElektraData = bkn

            adminboard.TrackBookingChanges(poseidonElektraBooking)
        Catch ex As Exception
            wl.Add(BusinessObjects.WarningSeverities.Info, "Track Booking Changes Exception: " + ex.Message + " Stack Trace: " + Convert.ToString(ex.StackTrace), ex.Source)
        End Try

        Return stepInfo
    End Function



    Private Sub AutoConfirmWithTotalWeeksCalculate(ByRef bookingForActive As List(Of BookingLine), ByRef elekBknList As CourseBookingList)

        Dim totalWeeks As Int16 = 0
        For Each item As BookingLine In bookingForActive
            totalWeeks = totalWeeks + item.Quantity
        Next
        If totalWeeks > 24 Then
            For Each item As BookingLine In bookingForActive
                For Each ebi As CourseBooking In elekBknList
                    If item.StartDate.Equals(ebi.StartDate) AndAlso item.EndDate.Equals(ebi.EndDate) AndAlso item.DestinationCode.Trim().ToUpper().Equals(ebi.DestinationCode.Trim().ToUpper()) AndAlso ebi.DestinationCode.Trim().ToUpper().Equals("KR-SEL") Then
                        If ebi.StatusCode.Equals(CourseBookingStatusLookup.CourseBookingStatuses.Confirmed) Then
                            ebi.StatusCode = "AC"
                            ebi.IsAutoConfirmed = False
                            ebi.StatusDate = Now
                            ebi.AcceptedDate = Nothing
                            ebi.ConfirmedDate = New Date(1800, 1, 1)
                        End If
                        Exit For
                    End If
                Next
            Next
        End If
        bookingForActive = New List(Of BookingLine)

    End Sub

    Public Shared Function UpdateAccomArticleIsCaxWhenGenderChanged(ByVal bookingId As Integer, ByVal oldGenderCode As String, newGenderCode As String) As DataTable

        Dim cm As New DataH.ConnectionManager
        Dim dt As New DataTable
        Dim cmd As New Sprocs.asp_UpdateAccomArticleIsCaxWhenGenderChanged
        With cmd.Parameters
            .BookingId = bookingId
            .oldGenderCode = oldGenderCode
            .newGenderCode = newGenderCode
        End With

        cm.Fill(dt, cmd, Constants.ElektraConnectionNamespace)
        Return dt
    End Function
    Public Sub UpdateLSTransporationChanges(ByVal CoursebookingId As Integer, ByVal Booking_id As Integer)

        Dim cm As New DataH.ConnectionManager
        Dim isExist As Boolean = False
        Dim cmd As New Sprocs.UpdateLSTransporationChanges
        With cmd.Parameters
            .Booking_Id = Booking_id
            .CoursebookingId = CoursebookingId
        End With

        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)
    End Sub
    Public Sub ExcludeJUJICheckList(ByVal bkn As Production.Booking, ByVal msg As CRMMessage.LSBooking, ByVal dtGetFirstCoruseAndDOBForBooking As DataTable)
        'Excluding JCC

        Dim CourseStartDate As Nullable(Of DateTime) = Nothing
        Dim FirstCourseType As String = ""
        Dim StudentAge As Integer = 0
        Dim StudentModifedAge As Integer = 0
        Dim FirstCourseStatusCode As String = ""
        If dtGetFirstCoruseAndDOBForBooking.Rows.Count > 0 Then
            If bkn.CourseBookingList.Count > 0 Then
                For Each cblist As CRMMessage.BookingLine In msg.BookingLineItems
                    If CourseStartDate Is Nothing Then
                        CourseStartDate = cblist.StartDate
                        FirstCourseType = cblist.CourseNumber
                        FirstCourseStatusCode = cblist.StatusCode
                    End If
                    If CourseStartDate > cblist.StartDate Then
                        CourseStartDate = cblist.StartDate
                        FirstCourseType = cblist.CourseNumber
                        FirstCourseStatusCode = cblist.StatusCode
                    End If
                Next
                'If Old Age is Below 18 and Modified age is > 18 then call RemoveRequiredCheckItemForBooking
                'if Old age is >18 and Modified age is < 18 then call RemoveRequiredCheckItemForBooking other no need to call
                'If Old first coruse type code is JI or JU and modified course type code is Not JI or JU then call RemoveRequiredCheckItemForBooking
                'If Old firsts course type is NOT JI OR JU and modified coruse is ji or ju then no need to call as it will be inserting first time in checklist table
                StudentAge = (New DateTime(CDate(dtGetFirstCoruseAndDOBForBooking.Rows(0)("CourseStartDate")).Subtract(CDate(dtGetFirstCoruseAndDOBForBooking.Rows(0)("BirthDate"))).Ticks)).Year - 1
                StudentModifedAge = (New DateTime(CDate(CourseStartDate).Subtract(bkn.Student.BirthDate).Ticks)).Year - 1
                If (StudentAge <= 18 AndAlso StudentModifedAge > 18) Or (StudentAge >= 18 AndAlso StudentModifedAge < 18) Then
                    BookingAdapter.RemoveRequiredCheckItemForBooking(bkn.SalesBookingId)
                ElseIf (dtGetFirstCoruseAndDOBForBooking.Rows(0)("CourseTypeCode").ToString().Trim() = "JU" AndAlso FirstCourseType <> "JU") Or
                    (dtGetFirstCoruseAndDOBForBooking.Rows(0)("CourseTypeCode").ToString().Trim() = "JI" AndAlso FirstCourseType <> "JI") Or
                    ((FirstCourseType = "JU" Or FirstCourseType = "JI") AndAlso FirstCourseStatusCode = "CA") Then
                    BookingAdapter.RemoveRequiredCheckItemForBooking(bkn.SalesBookingId)
                End If
            End If
        End If
    End Sub
    Public Shared Function GetFirstCourseAndDOBForBooking(ByVal salesBookingId As String) As DataTable

        Dim cm As New DataH.ConnectionManager
        Dim dt As New DataTable
        Dim cmd As New Sprocs.asp_GetFirstCourseAndDOBForBooking
        With cmd.Parameters
            .SalesBookingId = salesBookingId
        End With

        cm.Fill(dt, cmd, Constants.ElektraConnectionNamespace)
        Return dt
    End Function
    Public Shared Sub RemoveRequiredCheckItemForBooking(ByVal salesBookingId As String)

        Dim cm As New DataH.ConnectionManager
        Dim dt As New DataTable
        Dim cmd As New Sprocs.asp_RemoveRequiredCheckItemForBooking
        With cmd.Parameters
            .SalesBookingId = salesBookingId
        End With
        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)
    End Sub
    Public Shared Sub PreGroupBookingCancelled(ByVal PosGroupCode As String, ByVal SalesBookingId As String, ByVal BookingStatusCode As String, ByVal IsGroupRemoved As Boolean, ByVal OldPosGroupCode As String, ByVal IsLeaderBooking As Boolean)

        Dim cm As New DataH.ConnectionManager
        Dim ds As New DataSet
        Dim cmd As New Sprocs.asp_PreGroupBookingCancelled

        With cmd.Parameters
            .PosGroupCode = PosGroupCode
            .SalesBookingId = SalesBookingId
            .BookingStatusCode = BookingStatusCode
            .IsGroupRemoved = IsGroupRemoved
            .OldPosGroupCode = OldPosGroupCode
            .IsLeaderBooking = IsLeaderBooking
        End With

        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)
    End Sub
    Private Function DoesDestinationExistInElektra(ByVal destinationCode As String) As Boolean
        Dim exist As Boolean = True

        Try
            ProdLookups.DestinationLookup.FindInOriginalList(destinationCode)
        Catch ex As Exception
            exist = False
        End Try

        Return exist
    End Function

    'ELEK-6923 -LSupdatracking changes
    Public Shared Sub LTupdateTrackerInsert(ByVal salesBookingId As String, ByVal requiredItemId As Integer)

        Dim cm As New DataH.ConnectionManager
        Dim ds As New DataSet
        Dim cmd As New Sprocs.asp_LSUpdateTrackerInsert

        With cmd.Parameters
            .salesBookingId = salesBookingId
            .RequiredItem_id = requiredItemId
        End With

        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)
    End Sub

    'ELEK-6923 -LSupdatracking changes
    Public Shared Sub GenericDocumentInsert(ByVal BookingId As Integer)

        Dim cm As New DataH.ConnectionManager
        Dim ds As New DataSet
        Dim cmd As New Sprocs.asp_GenericDocumentInsert

        With cmd.Parameters
            .BookingId = BookingId
        End With

        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)
    End Sub

    Private Sub UpdateFlights(ByVal msg As CRMMessage.LSBooking, ByVal bkn As Production.Booking, ByVal wl As BusinessObjects.WarningListBase)
        Dim courseId As Integer = 0
        Dim FlightItin_id As Integer
        Dim FlgtItin As Production.FlightItin
        Dim Flgtitin_lnk As Production.BookingFlightItin_lnk
        Dim Flgt As Production.Flight
        Dim touchedList As New ArrayList
        Dim salesFlightItin_Id As String
        Dim _poseidonFlightItinList As New ArrayList

        Dim _elektraArrayList As New ArrayList
        Dim _elektrabookingFlightItinList As Production.BookingFlightItin_lnkList

        If Not msg.FlightItinItems Is Nothing Then
            For Each FlghtItin As CRMMessage.FlightItin In msg.FlightItinItems
                If Trim(FlghtItin.FlightItin_id) <> 0 Then
                    salesFlightItin_Id = bkn.SalesOfficeCode.Trim + FlghtItin.FlightItin_id.ToString.Trim
                    FlightItin_id = GetFightItin_id(salesFlightItin_Id)

                    If FlightItin_id = Null.IntegerNull Then
                        FlgtItin = New Production.FlightItin
                    Else
                        FlgtItin = New Production.FlightItin(FlightItin_id)
                    End If

                    _poseidonFlightItinList.Add(FlgtItin.FlightItin_id)

                    With FlghtItin
                        FlgtItin.Name = Trim(.Name)
                        FlgtItin.RecordLocator = Trim(.RecordLocator)
                        If (FlghtItin.IsPnr) = False Then
                            FlgtItin.IsPnr = False
                        ElseIf (FlghtItin.IsPnr) = True Then
                            FlgtItin.IsPnr = True
                        End If
                        FlgtItin.SalesSystemCode = Trim(.SalesSystemCode)
                        FlgtItin.TypeCode = Trim(.TypeCode)
                        FlgtItin.SalesFlightItinId = salesFlightItin_Id
                    End With


                    For Each flt As CRMMessage.Flight In FlghtItin.Flight
                        'Flgt variable has to be made null
                        Flgt = Nothing
                        Flgt = FlgtItin.FlightList.Find(flt.DepGateCode, flt.ArrGateCode, flt.FlightNum, flt.StatusCode, flt.ArrDateTime, flt.DepDateTime)
                        If Flgt Is Nothing Then
                            Flgt = FlgtItin.FlightList.Add()
                        End If
                        UpdateFlightProperties(msg, flt, Flgt)
                        touchedList.Add(Flgt)
                    Next

                    ' Remove flights
                    For Each Flgt In FlgtItin.FlightList.ToArray
                        If Not touchedList.Contains(Flgt) Then
                            'fi.FlightList.Remove(f)
                            'Do not delete the flight - rather cancel it
                            Flgt.StatusCode = FlightStatusLookup.FlightStatuses.Cancelled
                        End If
                    Next

                    If Not IsNull(FlgtItin) Then
                        FlgtItin.Save(CRMMessage.Constants.IntegrationUser, wl)
                    End If

                    Flgtitin_lnk = FlgtItin.BookingLnkList.Find(bkn.Booking_id, FlgtItin.FlightItin_id)

                    For Each bookingflightitin As CRMMessage.BookingFlightItin_lnk In msg.BookingFlightItin_lnkItems
                        If bookingflightitin.FlightItin_id = FlghtItin.FlightItin_id Then
                            If bookingflightitin.IsEnabled And Flgtitin_lnk Is Nothing Then
                                bkn.FlightItinLnkList.Add(FlgtItin)
                            ElseIf Not bookingflightitin.IsEnabled And Not Flgtitin_lnk Is Nothing Then
                                If bkn.FlightItinLnkList.Contains(FlgtItin) Then
                                    ' ELEK-3994 Manually entered flights gets overwritten by integration 
                                    If Flgtitin_lnk.InsertBy_id = 1837 Or Flgtitin_lnk.InsertBy_id = 11 Then
                                        bkn.FlightItinLnkList.Remove(FlgtItin)
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    Next

                End If
            Next

            _elektrabookingFlightItinList = bkn.FlightItinLnkList
            _elektraArrayList.AddRange(_elektrabookingFlightItinList)
            If IsSomething(_elektrabookingFlightItinList) AndAlso _elektrabookingFlightItinList.Count > 0 Then
                For Each _elektrabookingFlightItin As Production.BookingFlightItin_lnk In _elektraArrayList
                    If Not _poseidonFlightItinList.Contains(_elektrabookingFlightItin.FlightItin_id) Then
                        ' ELEK-3994 Manually entered flights gets overwritten by integration 
                        If _elektrabookingFlightItin.InsertBy_id = 1837 Or _elektrabookingFlightItin.InsertBy_id = 11 Then
                            bkn.FlightItinLnkList.Remove(_elektrabookingFlightItin.FlightItin_id)
                        End If
                    End If
                Next
            End If
        End If

        'If Not msg.BookingFlightItin_lnkItems Is Nothing AndAlso msg.BookingFlightItin_lnkItems.Count > 0 Then
        '    For Each BookingFlightItin As CRMMessage.BookingFlightItin_lnk In msg.BookingFlightItin_lnkItems
        '        If BookingFlightItin.IsEnabled = False Then
        '            FlightItin_id = GetFightItin_id(BookingFlightItin.FlightItin_id)
        '            If FlightItin_id <> Null.IntegerNull Then
        '                FlgtItin = New Production.FlightItin(FlightItin_id)
        '                If IsSomething(FlgtItin) Then
        '                    Flgtitin_lnk = FlgtItin.BookingLnkList.Find(bkn.Booking_id, FlgtItin.FlightItin_id)
        '                    If IsSomething(Flgtitin_lnk) Then
        '                        FlgtItin.BookingLnkList.Remove(Flgtitin_lnk)
        '                        FlgtItin.Save(CRMMessage.Constants.IntegrationUser, wl)
        '                        If IsSomething(bkn.FlightItinLnkList.Find(bkn.Booking_id, Flgtitin_lnk.FlightItin_id)) Then
        '                            bkn.FlightItinLnkList.Remove(Flgtitin_lnk)
        '                        End If
        '                    End If 
        '                End If
        '            End If
        '        End If
        '    Next
        'End If
    End Sub
    ' Code changes to rectify Null Error whenever BookingNote, MedicalNote etc are Nothing
    Private Sub UpdateLSBookingProperties(ByVal msg As CRMMessage.LSBooking, ByVal bkn As Production.Booking)

        If bkn.BookingNote Is Nothing Then bkn.BookingNote = String.Empty
        If bkn.MedicalNote Is Nothing Then bkn.MedicalNote = String.Empty
        If bkn.SpecialDietNote Is Nothing Then bkn.SpecialDietNote = String.Empty
        If bkn.AllergicNote Is Nothing Then bkn.AllergicNote = String.Empty
        If bkn.SubStatusCode Is Nothing Then bkn.SubStatusCode = String.Empty
        If msg.BookingNote Is Nothing Then msg.BookingNote = String.Empty
        If msg.MedicalNote Is Nothing Then msg.MedicalNote = String.Empty
        If msg.SpecialDietNote Is Nothing Then msg.SpecialDietNote = String.Empty
        If msg.AllergicNote Is Nothing Then msg.AllergicNote = String.Empty
        If msg.SubStatusCode Is Nothing Then msg.SubStatusCode = String.Empty
        If msg.Customer.GenderCode Is Nothing Then msg.Customer.GenderCode = String.Empty
        If msg.Customer.DateOfBirth Is Nothing Then msg.Customer.DateOfBirth = DateNull


        With msg
            If msg.BookingNote.Trim() <> bkn.BookingNote.Trim() OrElse msg.MedicalNote.Trim() <> bkn.MedicalNote.Trim() _
                OrElse msg.SpecialDietNote.Trim() <> bkn.SpecialDietNote.Trim() OrElse msg.AllergicNote.Trim() <> bkn.AllergicNote.Trim() _
                OrElse msg.IsSmoker <> bkn.IsSmoker OrElse msg.IsVegetarian <> bkn.IsVegetarian OrElse msg.IsSpecialDiet <> bkn.IsSpecialDiet _
                OrElse (msg.Customer.DateOfBirth > bkn.Student.BirthDate Or msg.Customer.DateOfBirth < bkn.Student.BirthDate) OrElse msg.IsAllergic <> bkn.IsAllergic OrElse msg.SubStatusCode <> bkn.SubStatusCode OrElse
                msg.Customer.GenderCode <> bkn.Student.GenderCode OrElse msg.NeedsVisa <> IIf(IsSomething(bkn.CourseBookingList.CourseBookingVisaTypeCode), (IIf(bkn.CourseBookingList.CourseBookingVisaTypeCode = "", False, True)), False) Then 'OrElse msg.Customer.Then 
                bkn.StudentMatchingConfirmed.ExternalChangeCofirmed = False
            End If
            bkn.SalesBookingId = Trim(.BookingNum)
            bkn.PoseidonBookingKey = .Booking_id
            bkn.SalesOfficeCode = .SalesOfficeCode

            If msg.GroupProgramCode = "LT" And msg.GroupProgramCode <> "" Then
                bkn.StatusCode = BookingStatusLookup.BookingStatuses.Cancelled
            Else
                bkn.StatusCode = .StatusCode
            End If

            If Not IsSomething(msg.BookingNote) Then
                msg.BookingNote = " "
            End If
            bkn.BookingNote = .BookingNote

            'Togetherwith for Groups
            If msg.ProgramCode = "LT" Then
                For Each b As BookingLine In msg.BookingLineItems
                    If Not b.TogetherWith Is Nothing Then
                        If b.TogetherWith.Count > 0 Then
                            bkn.BookingNote = bkn.BookingNote + " Togetherwith: " + b.TogetherWith(0).ToString() + IIf(Not String.IsNullOrEmpty(msg.IsOnlyLanguage) AndAlso msg.IsOnlyLanguage, " Only Language", "")
                        End If
                    End If
                Next
            End If
            If Not IsSomething(msg.MedicalNote) Then
                msg.MedicalNote = " "
            End If
            bkn.MedicalNote = .MedicalNote
            If Not IsSomething(msg.SpecialDietNote) Then
                msg.SpecialDietNote = " "
            End If
            bkn.SpecialDietNote = .SpecialDietNote
            bkn.IsSmoker = .IsSmoker
            bkn.IsVegetarian = .IsVegetarian
            bkn.IsSpecialDiet = .IsSpecialDiet
            'Auto COnfirmation changes, start
            bkn.IsReservation = .IsReservation
            bkn.IsDisabled = .IsDisabled
            bkn.IsMedicalOther = .IsMedicalOther
            bkn.SalesPersonEmailId = .SalesPersonEmailId
            bkn.IsSpecialRequirement = .IsSpecialRequirement
            'Auto Confirmation changes, end
            bkn.IsAllergic = .IsAllergic
            bkn.SalesPersonName = .SalesPersonName
            bkn.SubStatusCode = .SubStatusCode
            bkn.AgentName = .AgentName
            bkn.GroupName = .LSGroupName
            bkn.BookingDate = .BookingDate
            If Not IsSomething(msg.AllergicNote) Then
                msg.AllergicNote = " "
            End If
            bkn.AllergicNote = .AllergicNote
            If String.IsNullOrEmpty(bkn.Company) Then
                bkn.Company = .AgentName
            End If
            If IsSomething(.AgentName) Then
                bkn.AgentName = .AgentName
            End If
            ' ELEK-4031 Add Agent ID to Edit student screen
            If IsSomething(.AgentID) AndAlso .AgentID > 0 Then
                bkn.AgentId = Left(.BookingNum.Trim(), 2) + .AgentID.ToString()
            Else
                bkn.AgentId = ""
            End If
            bkn.Student.Company = .AgentName
            bkn.Student.Profession = .Profession
            If .SpecialDietNote <> "" Then
                bkn.IsSpecialDiet = True
            End If
            bkn.PoseidonVer = Production.Constants.PosedionVersion.PoseionII
            bkn.SalesRegion = .InstanceName  ' Code added for Tissino Sales Office By Gaurav Naithani

            'ELEK-5860
            If IsSomething(.SingaporeCourseFee) Then
                bkn.CourseFee = CType(.SingaporeCourseFee, Decimal)
            End If
            If IsSomething(.SingaporeCourseMaterial) Then
                bkn.CourseMaterial = CType(.SingaporeCourseMaterial, Decimal)
            End If
            If IsSomething(.SingaporeFpsInsurance) Then
                bkn.FpsInsurance = CType(.SingaporeFpsInsurance, Decimal)
            End If
            If IsSomething(.SingaporeRegistrationFee) Then
                bkn.RegistrationFee = CType(.SingaporeRegistrationFee, Decimal)
            End If


            If .PoseidonGroup_Id = 0 Then
                bkn.PoseidonGroup_Id = -1
                bkn.PoseidonGroupCode = String.Empty
                bkn.PoseidonGroupStatus = String.Empty
                bkn.SoldProductCode = String.Empty
                bkn.SoldProgramCode = String.Empty
            Else
                bkn.PoseidonGroup_Id = .PoseidonGroup_Id ' Added for group functionality by gaurav naithani
                bkn.PoseidonGroupCode = .GroupCode
                bkn.PoseidonGroupStatus = .GroupStatus
                bkn.SoldProductCode = .ProductCode
                bkn.SoldProgramCode = .ProgramCode
            End If
            bkn.PoseidonBooking_Id = .Booking_id
            bkn.PoseidonCAXReason = .SubStatusReason
            If Not String.IsNullOrEmpty(.EstimatedTravelDate) Then
                bkn.EstimatedTravelDate = CType(.EstimatedTravelDate, Date)
            End If

            bkn.SalesCurrencyCode = .CurrencyCode 'ELEK-6127
            bkn.Interest_ID = .InterestId
            bkn.HasDiscount = .HasDiscount
            bkn.MainSalesResponsible = .MainSalesResponsible
            bkn.HasBeenRebooked = .HasBeenRebooked
            bkn.NoOfRebookings = .NoOfRebookings
            bkn.LastExPaxProd = .LastExPaxProduct
            bkn.NoOFPrevBKN = .NoOfPrevBKNs
            bkn.AttnPreDepMtn = .AttnPreDepMtn
            bkn.HasEFbookedFlight = .HasEFBookedFlight

        End With
    End Sub

    Private Sub UpdateFlightProperties(ByVal msg As CRMMessage.LSBooking, ByVal fltMsg As CRMMessage.Flight, ByVal Flgt As Production.Flight)
        With fltMsg
            If msg.GroupProgramCode = "LT" And msg.GroupProductCode <> "" Then
                Flgt.StatusCode = FlightStatusLookup.FlightStatuses.Cancelled
            Else
                Flgt.StatusCode = Trim(.StatusCode)
            End If

            Flgt.ArrDateTime = Trim(.ArrDateTime)
            Flgt.ArrDestinationCode = Trim(.ArrDestinationCode)
            Flgt.ArrGateCode = Trim(.ArrGateCode)
            Flgt.ArrTerminal = Trim(.ArrTerminal)
            Flgt.Carrier = Trim(.Carrier)
            Flgt.DepDateTime = Trim(.DepDateTime)
            Flgt.DepDestinationCode = Trim(.DepDestinationCode)
            Flgt.DepGateCode = Trim(.DepGateCode)
            Flgt.DepTerminal = Trim(.DepTerminal)
            Flgt.FlightNum = Trim(.FlightNum)
            Flgt.FlightDirectionCode = Trim(.FlightDirectionCode)
        End With

    End Sub

End Class


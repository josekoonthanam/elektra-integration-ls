Public Class BookingLineAdapter

    Private _findBookingLine_id As Integer
    Private Shared _soldProgram As String = ""
    Private Const SOLDPROGRAM_LT As String = "LT"
    Private Const SOLDPROGRAM_ILC As String = "ILC" 'ELEK-4806

    Friend Function UpdateFromMsg(ByVal bkn As Production.Booking, ByVal msg As CRMMessage.LSBooking) As ResultStepInfo
        Dim stepInfo As ResultStepInfo
        Dim wl As New BusinessObjects.WarningListBase
        Dim isStaleMsg As Boolean = False
        Try
            Dim poseidonCourseBookings As ArrayList
            Dim cb As CourseBooking
            Dim c As Production.Course
            Dim cl As New CourseList
            Dim cbw As CourseBookingWeeks
            Dim touchedCourseBookings As New ArrayList
            Dim cbk As CRMMessage.BookingLine
            Dim LSVisaTypeCode As String
            Dim productMappingList As IntegrationProductMappingList
            Dim productMapping As IntegrationProductMapping
            Dim ProductDestination As ProdLookups.ProductDestination
            Dim productDestinationList As ProdLookups.ProductDestinationLookup
            'ELEK-4537
            Dim ProductCodeExist As Boolean = False
            Dim ILSPProgramme As Boolean = False
            Dim IsEligibleForLW As Boolean = False
            Dim IsEligibleForJCC As Boolean = False 'JCC - Juniour cord of conduct
            Dim PosProductCode As String
            Dim PosProgramCode As String
            Dim PoseidonProgramList As New Dictionary(Of Integer, String)

            For Each bknLine As CRMMessage.BookingLine In msg.BookingLineItems
                PoseidonProgramList.Add(bknLine.BookingLine_id, bknLine.ProgramCode)
            Next

            PosProductCode = msg.ProductCode
            PosProgramCode = msg.ProgramCode
            _soldProgram = msg.ProgramCode  ' PII-17061
            'ELEK-4537
            ILSPProgramme = isILSPProgramme(msg)
            'ELEK-6923 -LSupdatracking changes
            IsEligibleForLW = isLWEligible(msg)

            IsEligibleForJCC = IsJCCEligible(msg)

            productMappingList = New IntegrationProductMappingList
            ' code changes for groups
            If msg.BookingLineItems.Count > 0 Then
                For Each bknLine As CRMMessage.BookingLine In msg.BookingLineItems
                    If (bknLine.ProgramCode.Trim() = "EFC") And (bknLine.DestinationCode.Contains("GB-LRS") OrElse bknLine.DestinationCode.Contains("US-CHI") OrElse bknLine.DestinationCode.Contains("MT-MSJ") OrElse bknLine.DestinationCode.Contains("SG-SIN") OrElse bknLine.DestinationCode.Contains("US-SFD") OrElse bknLine.DestinationCode.Contains("SG-SIM")) Then
                        bknLine.ProductCode = "LSP"
                        bknLine.ProgramCode = "ILSP"
                        bkn.SoldProductCode = "CLT"
                        bkn.SoldProgramCode = "EFC"
                        msg.ProductCode = "LSP"
                        msg.ProgramCode = "ILSP"
                    End If
                Next
            End If

            If msg.PoseidonGroup_Id > 0 Then
                If msg.GroupProgramCode = "ILS" Or msg.GroupProgramCode = "ILC" Then
                    productMapping = productMappingList.Find(msg.GroupProductCode, msg.GroupProgramCode)
                    msg.ProductCode = msg.GroupProductCode
                    'ELEK-4537
                    If ILSPProgramme And msg.GroupProgramCode = "ILS" Then
                        msg = UpdatePrograms(msg)
                    Else
                        msg.ProgramCode = msg.GroupProgramCode
                    End If
                ElseIf (msg.GroupProgramCode = "LT") And IsSomething(bkn) Then
                    If bkn.CourseBookingList.Count > 0 Then
                        productMapping = productMappingList.Find(bkn.CourseBookingList.Item(0).CourseParent.ProductCode, bkn.CourseBookingList.Item(0).CourseParent.ProgramCode)
                        msg.ProductCode = bkn.CourseBookingList.Item(0).CourseParent.ProductCode
                        msg.ProgramCode = bkn.CourseBookingList.Item(0).CourseParent.ProgramCode
                    Else
                        productMapping = productMappingList.Find(msg.ProductCode, msg.ProgramCode)
                    End If
                End If
            Else
                'ELEK-4537
                If ILSPProgramme And msg.ProgramCode = "ILS" Then
                    msg = UpdatePrograms(msg)
                End If
                productMapping = productMappingList.Find(msg.ProductCode, msg.ProgramCode)
            End If

            For Each cbk In msg.BookingLineItems
                If msg.GroupProgramCode = "LT" And msg.GroupProgramCode <> "" Then
                    cbk.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                End If
            Next


            For Each cbk In msg.BookingLineItems
                If Not String.IsNullOrEmpty(PosProductCode) And (PosProductCode.Trim().ToUpper() = "LS" Or PosProductCode.Trim().ToUpper() = "LSP") AndAlso cbk.ProgramCode.Trim().ToUpper() = "EXC" Then
                    cbk.ProgramCode = "EFC"
                    cbk.ProductCode = "CLT"
                    msg.ProductCode = "CLT"
                End If
            Next
            productDestinationList = ProdLookups.ProductDestinationLookup.CreateInstance()

            ' Interpret course bookings
            poseidonCourseBookings = CreateCourseBookingWeeks(msg, wl)

            If msg.NeedsVisa Then
                LSVisaTypeCode = "I-20"
            Else
                LSVisaTypeCode = String.Empty
            End If
            ' Go through and update
            For Each cb In bkn.CourseBookingList.ToArray
                If cb.Status.IsActive Then
                    Dim schoolClassCache As New SchoolClassCacheList
                    For Each cbw In poseidonCourseBookings.ToArray
                        If cbw.DestinationCode.Trim = "EC-QUI" Then
                            If cbw.ProgramCode.Trim = "CC" Then
                                cbw.DestinationCode = "EC-QUIE"
                            End If
                        End If

                        ' checked with IntegrationProductmapping with ElektraProduct
                        For Each ProductDestination In productDestinationList
                            If IsSomething(ProductDestination) _
                            AndAlso ProductDestination.DestinationCode = cbw.DestinationCode _
                            AndAlso Date.Compare(ProductDestination.ValidFromDate, Today) <= 0 _
                            AndAlso Date.Compare(ProductDestination.ValidToDate, Today) >= 0 AndAlso ProductDestination.ProductCode = msg.ProductCode Then
                                If ProductDestination.ProductCode = msg.ProductCode Then
                                    ProductCodeExist = True
                                    Exit For
                                End If
                            End If
                        Next

                        If IsSomething(productMapping) AndAlso ProductCodeExist = False Then
                            msg.ProductCode = productMapping.DestinationProductCode
                            msg.ProgramCode = productMapping.DestinationProgramCode
                        End If
                        ' Code updated by Pagalavan - PII-6624
                        c = CourseManager.GetCourse( _
                                            cbw.DestinationCode _
                                            , cbw.ProgramCode _
                                            , msg.ProductCode _
                                            , cbw.CourseTypeCode.Trim().ToString() _
                                            , cbw.StartDate _
                                            , cbw.EndDate _
                                            , msg.SalesOfficeCode _
                                            , cl _
                                            , CRMMessage.Constants.IntegrationUser _
                                            )

                        ' Only consider correct destination, program and course type
                        If cb.DestinationCode = cbw.DestinationCode Then
                            If cb.CourseParent.ProgramCode = c.ProgramCode _
                                AndAlso cb.CourseParent.CourseTypeCode = c.CourseTypeCode Then
                                If cb.StartWeek.Week_id <= cbw.EndWeek.Week_id _
                                        And cb.EndWeek.Week_id >= cbw.StartWeek.Week_id Then

                                    ' Keep track of which ones were touched
                                    ' This is used later to know what to delete
                                    touchedCourseBookings.Add(cb)
                                    'Update - Sangamithra
                                    cb.ExamCode = cbw.ExamCode
                                    If cbw.YearsOfStudy > 0 Then
                                        cb.YearsOfStudy = cbw.YearsOfStudy
                                    End If

                                    cb.TotalPrice = cbw.TotalPrice  'ELEK-6127 --- Total Price update for course
                                    cb.ISPRW = cbw.ISPRW
                                    cb.ISPIE = cbw.ISPIE
                                    cb.IsCourseLeader = cbw.IsCourseLeader
                                    cb.PoseidonProduct = PosProductCode
                                    If PoseidonProgramList.ContainsKey(cbw.BookingLine_ID) Then
                                        cb.PoseidonProgram = PoseidonProgramList.Item(cbw.BookingLine_ID)
                                    Else
                                        cb.PoseidonProgram = PosProgramCode
                                    End If


                                    If cb.Status.IsActive AndAlso cb.StatusCode <> cbw.StatusCode AndAlso cb.StatusCode <> CourseBookingStatusLookup.CourseBookingStatuses.Confirmed Then
                                        cb.StatusCode = IIf(cbw.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active)
                                        cb.SetStatus(IIf(cbw.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active), CRMMessage.Constants.IntegrationUser)
                                    End If
                                    cb.StatusDate = msg.BookingDate
                                    If cb.StatusCode <> "CF" Then       ' code changes for ELEK-3174,3365 by gaurav naithani
                                        cb.AcceptedDate = msg.BookingDate
                                    End If
                                    ' Business rule: treat pre and post departure differently:
                                    '  - If not arrived, always move current item
                                    '  - If arrived, add rows for extensions, but shorten if changed
                                    If Not HasStarted(bkn, cb.DestinationCode) Then
                                        ' Not arrived yet
                                        ' Find correct course
                                        c = CourseManager.GetCourse( _
                                            cbw.DestinationCode _
                                            , cbw.ProgramCode _
                                            , msg.ProductCode _
                                            , cbw.CourseTypeCode.Trim().ToString() _
                                            , cbw.StartDate _
                                            , cbw.EndDate _
                                            , msg.SalesOfficeCode _
                                            , cl _
                                            , CRMMessage.Constants.IntegrationUser _
                                            )
                                        With cb
                                            .Course_id = c.Course_id
                                            .StartWeekCode = cbw.StartWeek.Code
                                            .EndWeekCode = cbw.EndWeek.Code
                                            .StartDate = cbw.StartDate
                                            .EndDate = cbw.EndDate
                                            .Weeks = cbw.Weeks.Count
                                            .VisaTypeCode = LSVisaTypeCode
                                            .StatusDate = msg.BookingDate
                                            If cb.StatusCode <> "CF" Then
                                                .AcceptedDate = msg.BookingDate
                                            End If
                                            If .HasDateChanges Then
                                                .StatusCode = IIf(cbw.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active)
                                                .SetStatus(IIf(cbw.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active), CRMMessage.Constants.IntegrationUser)
                                            End If
                                        End With
                                        ' Remove entry
                                        poseidonCourseBookings.Remove(cbw)
                                        Exit For
                                    Else
                                        ' Handle start week changes
                                        If cbw.StartWeek.Week_id < cb.StartWeek.Week_id Then
                                            ' Create copy of current poseidon booking guy, but remove all weeks after start of the course
                                            Dim startCbw As New CourseBookingWeeks(cbw)
                                            startCbw.Weeks.Remove(cb.StartWeekCode, startCbw.EndWeek.Code)
                                            startCbw.EndWeek = startCbw.Weeks.LastWeek
                                            startCbw.EndDate = startCbw.EndWeek.Friday

                                            ' Add to list so it gets created in the next step
                                            poseidonCourseBookings.Add(startCbw)

                                        ElseIf cbw.StartWeek.Week_id > cb.StartWeek.Week_id Then
                                            Dim WeeksToRemove As New WeekSpan(cb.StartWeek, cbw.StartWeek.AddWeeks(-1))
                                            Dim HasData As Boolean
                                            HasData = False
                                            If cb.HasClasses(cb.StartWeekCode, WeeksToRemove.LastWeek.Code) Then
                                                HasData = True
                                            End If
                                            cb.StartDate = cbw.StartDate
                                            cb.StartWeekCode = (New Week(cb.StartDate)).Code
                                            cb.EndDate = cbw.EndDate
                                            cb.EndWeekCode = (New Week(cb.EndDate)).Code
                                            cb.StatusDate = msg.BookingDate
                                            If cb.StatusCode <> "CF" Then
                                                cb.AcceptedDate = msg.BookingDate
                                            End If

                                        End If

                                        ' Handle end week changes
                                        If cbw.EndWeek.Week_id < cb.EndWeek.Week_id Then
                                            ' Arrived, and the course is shortened
                                            ' Never update start week, since the student has arrived: cb.StartWeekCode = cbw.StartWeek.Code
                                            cb.EndWeekCode = cbw.EndWeek.Code
                                            cb.StartDate = cbw.StartDate
                                            cb.EndDate = cbw.EndDate

                                            cb.Weeks = cbw.Weeks.Count

                                            cb.VisaTypeCode = LSVisaTypeCode
                                            cb.StatusCode = IIf(cbw.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active)
                                            cb.StatusDate = msg.BookingDate
                                            If cb.StatusCode <> "CF" Then
                                                cb.AcceptedDate = msg.BookingDate
                                            End If

                                            ' Remove entry
                                            poseidonCourseBookings.Remove(cbw)
                                            Exit For

                                        ElseIf cbw.EndWeek.Week_id > cb.EndWeek.Week_id Then
                                            ' Student has arrived, and the course is extended
                                            ' Cut out only relevant weeks from cbw to add item
                                            cbw.Weeks.Remove(cbw.StartWeek.Code, cb.EndWeekCode)
                                            cbw.StartWeek = cbw.Weeks.FirstWeek
                                            cbw.StartDate = cbw.StartWeek.Monday
                                            'ElseIf (Not cbw.Description Is Nothing) AndAlso _
                                            '        (cbw.Description.StartsWith("TR-") OrElse _
                                            '        cbw.Description.StartsWith("TE-") OrElse _
                                            '        cbw.Description.StartsWith("EX-")) Then
                                            '    cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                                            '    poseidonCourseBookings.Remove(cbw)
                                        Else
                                            poseidonCourseBookings.Remove(cbw)
                                        End If
                                    End If  '  HasStarted(bkn, cb.DestinationCode) 
                                End If  ' cb.StartWeek.Week_id <= cbw.EndWeek.Week_id ...
                            End If  '  cb.CourseParent.ProgramCode = cbw.ProgramCode
                        End If  '  cb.DestinationCode = cbw.DestinationCode
                    Next ' cbw
                    schoolClassCache.Dispose()
                End If
                ' Clear flag for user acceptance
                If cb.Status.IsActive And cb.HasDateChanges Then
                    cb.StatusCode = IIf(cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active)
                    cb.SetStatus(IIf(cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active), CRMMessage.Constants.IntegrationUser)
                End If
            Next ' cb
            ' Create new course bookings
            For Each cbw In poseidonCourseBookings.ToArray
                'If cbw.EndWeek.Code >= Week.Current.AddWeeks(-36).Code Then
                cb = bkn.CourseBookingList.Add(bkn)

                If cbw.DestinationCode.Trim = "EC-QUI" Then 'Added for defect Id 26069 by Thanigai
                    If cbw.ProgramCode.Trim = "CC" Then
                        cbw.DestinationCode = "EC-QUIE"
                    End If
                End If

                ' Find correct course
                c = CourseManager.GetCourse( _
                    cbw.DestinationCode _
                    , cbw.ProgramCode _
                    , msg.ProductCode _
                    , cbw.CourseTypeCode.Trim().ToString() _
                    , cbw.StartDate _
                    , cbw.EndDate _
                    , msg.SalesOfficeCode _
                    , cl _
                    , CRMMessage.Constants.IntegrationUser _
                    )
                ' Set all properties
                cb.Course_id = c.Course_id
                cb.StartWeekCode = cbw.StartWeek.Code
                cb.EndWeekCode = cbw.EndWeek.Code
                cb.StartDate = cbw.StartDate
                cb.EndDate = cbw.EndDate
                cb.Weeks = cbw.Weeks.Count
                cb.VisaTypeCode = LSVisaTypeCode
                cb.PoseidonTermReason = cbw.PoseidonTermReason
                cb.TotalPrice = cbw.TotalPrice 'ELEK-6127, add total price of course
                cb.ISPRW = cbw.ISPRW
                cb.ISPIE = cbw.ISPIE
                cb.IsCourseLeader = cbw.IsCourseLeader
                cb.PoseidonProduct = PosProductCode
                If PoseidonProgramList.ContainsKey(cbw.BookingLine_ID) Then
                    cb.PoseidonProgram = PoseidonProgramList.Item(cbw.BookingLine_ID)
                Else
                    cb.PoseidonProgram = PosProgramCode
                End If

                If msg.ProductCode.Trim() = "IA" Then
                    cb.IsNoShow = True
                End If

                'Added By - Pagalavan Check whether It is Tranfer or Termination
                'If (Not cbw.Description Is Nothing) AndAlso _
                '                            (cbw.Description.StartsWith("TR-") OrElse _
                '                            cbw.Description.StartsWith("TE-") OrElse _
                '                            cbw.Description.StartsWith("EX-")) Then
                '    cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                'Else

                cb.StatusCode = IIf(cbw.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.PreApplication, CourseBookingStatusLookup.CourseBookingStatuses.Active)

                'End If
                cb.ExamCode = cbw.ExamCode
                If cbw.YearsOfStudy > 0 Then
                    cb.YearsOfStudy = cbw.YearsOfStudy
                End If

                cb.StatusDate = msg.BookingDate
                If cb.StatusCode <> "CF" Then
                    cb.AcceptedDate = msg.BookingDate
                End If
                touchedCourseBookings.Add(cb)
                poseidonCourseBookings.Remove(cbw)
                'End If
            Next
            ' Cancel unused course bookings
            For Each cb In bkn.CourseBookingList
                If Not touchedCourseBookings.Contains(cb) Then
                    'Only cax the course booking if there are no classes.
                    cb.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                End If
            Next

            ' Find weeks to move
            Dim moveList As New ArrayList
            For Each cb In bkn.CourseBookingList
                For Each cbweek As CourseBookingWeek In cb.CourseBookingWeekList
                    If Not touchedCourseBookings.Contains(cb) _
                            OrElse Not (cbweek.WeekCode >= cb.StartWeekCode AndAlso cbweek.WeekCode <= cb.EndWeekCode) Then
                        moveList.Add(cbweek)
                    End If
                Next
            Next
            ' Find place for weeks
            Dim schoolClassCache2 As New SchoolClassCacheList
            For Each cbweek As CourseBookingWeek In moveList.ToArray
                For Each cb In touchedCourseBookings
                    Dim targetWeek As CourseBookingWeek
                    Dim targetClassMember As SchoolClassMember

                    If cb.Status.IsActive _
                            AndAlso cbweek.WeekCode >= cb.StartWeekCode _
                            AndAlso cbweek.WeekCode <= cb.EndWeekCode Then

                        ' Find or create target week
                        targetWeek = cb.CourseBookingWeekList.Find(cb.CourseBooking_id, cbweek.WeekCode)
                        If targetWeek Is Nothing Then
                            targetWeek = cb.CourseBookingWeekList.Add(cb.CourseBooking_id, cbweek.WeekCode)
                        End If

                        ' Add info and/or classes
                        If Not targetWeek.LevelValue <> String.Empty Then
                            targetWeek.LevelValue = cbweek.LevelValue
                        End If

                        For Each scm As SchoolClassMember In cbweek.SchoolClassMemberList.ToArray
                            Dim shouldCopy As Boolean = False

                            ' Check if the class should be copied
                            shouldCopy = scm.HasUserData

                            'made sure only if its same destination copy classes
                            If String.Compare(cbweek.CourseBooking.CourseParent.DestinationCode.Trim(), targetWeek.CourseBooking.CourseParent.DestinationCode.Trim(), True) = 0 Then

                                shouldCopy = Not IsInList( _
                                        schoolClassCache2.GetSchoolClass(scm.Class_id).ClassType.ClassCategoryCode _
                                        , ClassCategoryLookup.ClassCategories.ArrivalClasses _
                                        , ClassCategoryLookup.ClassCategories.DepartureClasses _
                                        )

                            End If
                            'ELEK-5633 --- Two Placement test show up in the Classes Tab, end

                            If shouldCopy Then
                                ' Get or create class member
                                targetClassMember = targetWeek.SchoolClassMemberList.Find(scm.Class_id)
                                If targetClassMember Is Nothing Then
                                    targetClassMember = targetWeek.SchoolClassMemberList.Add(scm.Class_id)
                                End If

                                ' update info
                                targetClassMember.Attended = MaxOf(targetClassMember.Attended, scm.Attended)
                                targetClassMember.Lessons = MaxOf(targetClassMember.Lessons, scm.Lessons)
                                targetClassMember.GradeCode = MaxOf(targetClassMember.GradeCode, scm.GradeCode)
                                targetClassMember.ResultCode = MaxOf(targetClassMember.ResultCode, scm.ResultCode)
                                targetClassMember.Score = MaxOf(targetClassMember.Score, scm.Score)
                                targetClassMember.TestTypeCode = MaxOf(targetClassMember.TestTypeCode, scm.TestTypeCode)
                            End If
                            cbweek.SchoolClassMemberList.Remove(scm)
                        Next
                        ' Remove the week
                        'cbweek.CourseBookingWeekList.Remove(cbweek.CourseBooking_id, cbweek.WeekCode)
                        'moveList.Remove(cbweek)
                        Exit For
                    End If
                Next

                ' Find out if any relevant classes are left
                Dim hasClasses As Boolean = False
                For Each scm As SchoolClassMember In cbweek.SchoolClassMemberList.ToArray
                    ' Check if the class should be copied
                    hasClasses = hasClasses Or scm.HasUserData
                    If Not hasClasses Then
                        hasClasses = Not IsInList( _
                                schoolClassCache2.GetSchoolClass(scm.Class_id).ClassType.ClassCategoryCode _
                                , ClassCategoryLookup.ClassCategories.ArrivalClasses _
                                , ClassCategoryLookup.ClassCategories.DepartureClasses _
                                )
                    End If
                Next
                If cbweek.CourseBooking.StatusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled Then
                    cbweek.SchoolClassMemberList.RemoveAll()
                End If
                ' code commented by Gaurav Naithani so that class should not be deleted even if booking is CAX
                ' Remove empty guys
                'If Not hasClasses Then
                'If moveList.Contains(cbweek) Then
                'cbweek.CourseBookingWeekList.Remove(cbweek)
                'cbweek.CourseBookingWeekList.Remove(cbweek.CourseBooking_id, cbweek.WeekCode)
                'moveList.Remove(cbweek)
                ' End If
                'End If
            Next ' cbweek in moveList
            schoolClassCache2.Dispose()
            If moveList.Count > 0 Then
                wl.Add( _
                    BusinessObjects.WarningSeverities.Warning _
                    , "Could not move {0} classes from booking {1}." _
                    , moveList.Count _
                    , bkn.SalesBookingId _
                    )
            End If

            For Each cb In bkn.CourseBookingList
                cb.Weeks = New WeekSpan(cb.StartWeekCode, cb.EndWeekCode).Count
                If _soldProgram <> SOLDPROGRAM_LT Then
                    If _soldProgram <> Programs.EFCorporate Then
                        cb.StartDate = cb.StartWeek.Monday
                        cb.EndDate = cb.EndWeek.Friday
                    End If
                End If

            Next
            UpdateBookingArticles(bkn, msg)
            UpdateBookingLineChanges(bkn, msg)
            'ELEK-6923 -LSupdatracking changes for JCC
            If IsEligibleForJCC Then
                BookingAdapter.LTupdateTrackerInsert(bkn.SalesBookingId, 2)
            ElseIf IsEligibleForLW Then
                BookingAdapter.LTupdateTrackerInsert(bkn.SalesBookingId, 1)
            End If
        Catch ex As Exception
            wl.Add(BusinessObjects.WarningSeverities.Error, ex.Message, ex.Source)
            Throw ex
            'wl.Add(ex)
        Finally
            stepInfo = Me.HandleUpdateInfo(wl, isStaleMsg)
        End Try
        Return stepInfo
    End Function
    'ELEK-4537
    Private Function isILSPProgramme(ByVal msg As CRMMessage.LSBooking) As Boolean
        Dim cbk As CRMMessage.BookingLine
        For Each cbk In msg.BookingLineItems

            If (ProdLookups.CalculateAge(msg.Customer.DateOfBirth, cbk.StartDate) >= 25) And (cbk.ProgramCode = Programs.InternationalLanguageSchools) Then
                Return True

            End If
        Next
        Return False
    End Function
    ' Start ELEK-6923 -LSupdatracking changes
    Private Function isLWEligible(ByVal msg As CRMMessage.LSBooking) As Boolean
        Dim cbk As CRMMessage.BookingLine
        For Each cbk In msg.BookingLineItems
            If cbk.StartDate > DateTime.Now Then
                If IsDate(cbk.StartDate) Then
                    If (ProdLookups.CalculateAge(msg.Customer.DateOfBirth, cbk.StartDate) < 18) Then
                        Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function
    ' Start ELEK-6923 -LSupdatracking changes for JCC
    Private Function IsJCCEligible(ByVal msg As CRMMessage.LSBooking) As Boolean
        Dim cbk As CRMMessage.BookingLine
        For Each cbk In msg.BookingLineItems
            If Not String.IsNullOrEmpty(cbk.CourseNumber) Then
                If cbk.StartDate >= DateTime.Now AndAlso (cbk.CourseNumber.Trim().ToUpper() = "JU" Or cbk.CourseNumber.Trim().ToUpper() = "JI") Then
                    If IsDate(cbk.StartDate) Then
                        Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function
    'ELEK-4537
    Private Function UpdatePrograms(ByVal bookingMsgs As CRMMessage.LSBooking) As CRMMessage.LSBooking
        Dim bl As CRMMessage.BookingLine
        bookingMsgs.ProgramCode = Programs.InternationalLanguageSchools25Plus
        For Each bl In bookingMsgs.BookingLineItems
            If (ProdLookups.CalculateAge(bookingMsgs.Customer.DateOfBirth, bl.StartDate) >= 25) And (bl.ProgramCode = Programs.InternationalLanguageSchools) Then
                bl.ProgramCode = Programs.InternationalLanguageSchools25Plus
            End If

        Next

        For Each aBlI As CRMMessage.BookingLine In bookingMsgs.ArticleBookingLineItems
            If IsDate(aBlI.StartDate) Then
                If (ProdLookups.CalculateAge(bookingMsgs.Customer.DateOfBirth, aBlI.StartDate) >= 25) And (aBlI.ProgramCode = Programs.InternationalLanguageSchools) Then
                    aBlI.ProgramCode = Programs.InternationalLanguageSchools25Plus
                End If
            End If
        Next
        Return bookingMsgs
    End Function

    Private Sub UpdateBookingArticles(ByVal bkn As Production.Booking, ByVal msg As CRMMessage.LSBooking)
        Dim bar As CRMMessage.BookingLine
        Dim a As Article
        Dim ba As Production.BookingArticle
        Dim list As New ArrayList(bkn.BookingArticleList)
        Dim hasSignificantChanges As Boolean
        Dim art As CRMMessage.BookingLine
        'ELEK-4609
        Dim resArtCnt As New Hashtable
        Dim resBkArtCnt As New Hashtable
        Dim artCnt As Integer = 0
        'Dim timeSpan As TimeSpan
        'Dim quantity As Integer
        '_soldProgram = msg.ProgramCode
        ' code changes for groups


        For Each art In msg.ArticleBookingLineItems
            If msg.GroupProgramCode = "LT" And msg.GroupProgramCode <> "" Then
                art.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Cancelled
            End If
            'ELEK-4609
            If art.ArticleCode.StartsWith("RES") Then
                If resArtCnt.ContainsKey(art.ArticleCode) Then
                    resArtCnt(art.ArticleCode) = CObj(CInt(resArtCnt(art.ArticleCode)) + 1)
                Else
                    resArtCnt.Add(art.ArticleCode, 1)
                End If
            End If
        Next

        For Each articles As BookingArticle In list
            If articles.ArticleCode.StartsWith("RES") AndAlso (articles.StatusCode = "AC" Or articles.StatusCode = "CF") Then
                If resBkArtCnt.ContainsKey(articles.ArticleCode.Trim) Then
                    resBkArtCnt(articles.ArticleCode) = CObj(CInt(resBkArtCnt(articles.ArticleCode)) + 1)
                Else
                    resBkArtCnt.Add(articles.ArticleCode, 1)
                End If
            End If
        Next
        msg.ArticleBookingLineItems.Sort(Function(x, y) x.BookingLine_id.CompareTo(y.BookingLine_id))

        For Each bar In msg.ArticleBookingLineItems
            hasSignificantChanges = False
            ' code changes by gaurav naithani so that no duplicate rows should be created in booking article table
            If (bar.ArticleCode = "EX") Or (bar.ArticleCode = "TE") Then
                bar.ArticleCode = bar.Description
            End If

            Try
                If bar.ArticleCode = SystemArticleTypeCodes.Extension OrElse _
                    bar.ArticleCode = SystemArticleTypeCodes.Termination Then
                    a = ArticleLookup.FindInOriginalList(bar.Description)
                Else
                    a = ArticleLookup.FindInOriginalList(bar.ArticleCode)
                End If

            Catch ex As Exception
                a = Nothing
            End Try
            If IsSomething(a) Then
                If a.Code <> a.ParentCode Then
                    a = ArticleLookup.FindInOriginalList(a.ParentCode)
                End If
                If a.AllowImport Then
                    If bar.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.Active OrElse bar.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.PreApplication Then
                        ' Try to find a match
                        If (bar.ArticleCode = "EX") Or (bar.ArticleCode = "TE") Then
                            ba = FindBookingArticle(bar, list)
                        Else
                            ba = FindBookingArticleNew(bar, list)
                            'ELEK-4609
                            If IsSomething(ba) AndAlso bar.ArticleCode.StartsWith("RES") Then
                                If IsSomething(resBkArtCnt(bar.ArticleCode)) Then
                                    If IsSomething(resArtCnt(bar.ArticleCode) AndAlso CInt(resArtCnt(bar.ArticleCode)) > CInt(resBkArtCnt(bar.ArticleCode))) Then
                                        'If bar.StartDate <> ba.StartDate AndAlso bar.EndDate <> ba.EndDate Then
                                        '    ba = Nothing
                                        'End If
                                        resArtCnt(bar.ArticleCode) = CInt(resArtCnt(bar.ArticleCode)) - 1
                                        resBkArtCnt(bar.ArticleCode) = CInt(resBkArtCnt(bar.ArticleCode)) - 1
                                    End If
                                End If
                            End If

                        End If
                        If ba Is Nothing Then
                            ' Create new
                            ba = bkn.BookingArticleList.Add()
                        Else
                            ' Remove from list - remaining items will be cancelled
                            list.Remove(ba)
                            If a.AccSupplierTypeCode = "RE" AndAlso bar.StartDate >= ba.StartDate AndAlso bar.EndDate <= ba.EndDate Then
                                'modify ResidenceAllocatedStudents
                                Dim raTable As DataTable = bkn.GetAllocatedAllocationForResManager(ServiceBookingStatusLookup.ServiceBookingStatuses.Allocated).Tables(0)
                                If IsSomething(raTable) AndAlso raTable.Rows.Count > 0 Then
                                    For Each raRow As DataRow In raTable.Rows
                                        If IsSomething(raRow("BookingArticle_id")) Then
                                            If ba.BookingArticle_id = CInt(raRow("BookingArticle_id")) Then
                                                ResidenceAllocation.ResidenceAllocatedStudentsUpdate(ba.BookingArticle_id, bar.StartDate, bar.EndDate, CBool(raRow("IsDeleted")))
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If

                        ' Set values
                        With ba
                            If .ArticleCode <> a.Code Then
                                hasSignificantChanges = True
                            End If
                            .ArticleCode = a.Code

                            If bar.DestinationCode <> .DestinationCode Then
                                hasSignificantChanges = True
                            End If

                            If Not bar.DestinationCode = "" Then
                                .DestinationCode = bar.DestinationCode
                            Else
                                .DestinationCode = Nothing
                            End If

                            If .StartDate <> bar.StartDate Then
                                hasSignificantChanges = True
                            End If

                            If bar.StartDate = Null.DateNull Then
                                .StartDate = Null.DateNull
                                .StartWeekCode = Nothing
                            Else
                                .StartDate = bar.StartDate
                                .StartWeekCode = New Week(bar.StartDate).Code
                            End If

                            If .EndDate <> bar.EndDate Then
                                hasSignificantChanges = True
                            End If

                            If bar.EndDate = Null.DateNull Then
                                If bar.StartDate <> Null.DateNull Then
                                    .EndDate = bar.StartDate
                                Else
                                    .EndDate = Null.DateNull
                                End If
                                .EndWeekCode = Nothing
                            Else
                                .EndDate = bar.EndDate
                                .EndWeekCode = New Week(bar.EndDate).Code
                            End If

                            If a.RoomTypeCode > "" Then
                                .RoomTypeCode = a.RoomTypeCode
                            Else
                                .RoomTypeCode = Nothing
                            End If
                            Select Case bar.UnitCode.Trim()
                                Case "UNIT"
                                    .UnitCode = "Unit"
                                Case "LESSON"
                                    .UnitCode = "Lesson"
                                Case "DAY"
                                    .UnitCode = "Day"
                                Case "WEEK"
                                    .UnitCode = "Week"
                            End Select

                            'TimeSpan = bar.EndDate.Subtract(bar.StartDate)
                            'quantity = TimeSpan.Days / 7

                            If msg.PoseidonGroup_Id > 0 And bar.StartDate <> Null.DateNull And bar.EndDate <> Null.DateNull Then
                                If Not String.IsNullOrEmpty(bar.ArticleCode) AndAlso _soldProgram = "LT" AndAlso bar.ArticleCode.ToUpper().StartsWith("GAV") Then
                                    bar.Quantity = bar.Quantity
                                Else
                                    bar.Quantity = NumberOfWeeks(bar.StartDate, bar.EndDate)
                                End If
                            End If

                            '.UnitCode = bar.UnitCode
                            If .Units <> bar.Quantity OrElse (ba.StatusCode <> bar.StatusCode AndAlso ba.StatusCode <> BookingArticleStatusLookup.BookingArticleStatuses.Confirmed) Then
                                hasSignificantChanges = True
                            End If

                            .ISPIE = bar.IsPIERequest

                            .Units = bar.Quantity

                            .TotalPrice = bar.Totalprice    'ELEK-6127, total price for articles



                            'ELEK-6275, Together With, start
                            If Not bar.TogetherWith Is Nothing Then 'condition added to make sure that even if togetherwith is nothing, then there is no exception for ELEK-6448
                                If bar.TogetherWith.Count > 0 Then
                                    .TogetherWith = bar.TogetherWith(0).ToString()
                                Else
                                    .TogetherWith = ""
                                End If
                            End If
                            'ELEK-6275, Together With, end

                            If hasSignificantChanges OrElse ba.StatusCode = BookingStatusLookup.BookingStatuses.Cancelled Then

                                ba.SetStatus(IIf(bar.StatusCode = BookingArticleStatusLookup.BookingArticleStatuses.PreApplication, BookingArticleStatusLookup.BookingArticleStatuses.PreApplication, BookingArticleStatusLookup.BookingArticleStatuses.Active), CRMMessage.Constants.IntegrationUser)


                            End If

                            .StatusDate = Now()
                            .AcceptedDate = msg.BookingDate
                        End With
                    End If

                End If
            End If
        Next
        ' Cancel remaining items
        For Each ba In list
            ba.SetStatus(BookingArticleStatusLookup.BookingArticleStatuses.Cancelled, CRMMessage.Constants.IntegrationUser)
            'Set Ispie to false for Cancelled Booking Articles.
            If (ba.ISPIE) Then
                ba.ISPIE = False
            End If
            If (ba.IsUpgrade) Then
                ResidenceAllocation.CancelUpgradeRquest(ba.BookingArticle_id)
            End If
        Next
    End Sub


    Public Shared Function NumberOfWeeks(ByVal dateFrom As DateTime, ByVal dateTo As DateTime) As Integer
        Dim Span As TimeSpan = dateTo.Subtract(dateFrom)

        If Span.Days <= 7 Then
            If dateFrom.DayOfWeek > dateTo.DayOfWeek Then
                Return 2
            End If

            Return 1
        End If

        Dim Days As Integer = Span.Days - 7 + CInt(dateFrom.DayOfWeek)
        Dim WeekCount As Integer = 1
        Dim DayCount As Integer = 0

        WeekCount = 1
        While DayCount < Days
            DayCount += 7
            WeekCount += 1
        End While

        Return WeekCount
    End Function

    Private Sub UpdateBookingLineChanges(ByVal bkn As Production.Booking, ByVal msg As CRMMessage.LSBooking)
        Dim touchedlist As New ArrayList
        Dim done As Boolean

        UpdateCbChangeRemoveCourseBookingData(bkn)
        ' Get change history from poseidon
        If msg.BookingLineChangeItems.Count > 0 Then
            Try
                UpdateCbChanges2(bkn, msg, touchedlist)
                done = True

            Catch ex As BaseApplicationException
                Throw ex
            End Try
        End If
    End Sub

    Private Shared Sub UpdateCbChanges2(ByVal bkn As Production.Booking, ByVal msg As CRMMessage.LSBooking, ByVal touchedList As ArrayList)
        Dim sw As Week
        Dim ew As Week
        Dim changeStartWeek As Week
        Dim changeEndWeek As Week
        Dim courseBookingStartWeek As Week
        Dim courseBookingEndWeek As Week
        Dim salesCourseBookingId As Integer

        Dim cbc As CourseBookingChange
        Dim courseTypeCode As String
        Dim changeType As String
        Dim weeks As Integer
        Dim statusCode As String

        For Each cbr As CRMMessage.BookingLineChange In msg.BookingLineChangeItems

            If cbr.StartDate.Year <> 1800 AndAlso cbr.EndDate.Year <> 1800 Then

                sw = New Week(CourseBookingWeeks.GetStartDate(msg.ProgramCode, cbr.StartDate, _soldProgram))
                ew = New Week(CourseBookingWeeks.GetEndDate(msg.ProgramCode, cbr.EndDate, _soldProgram))
                weeks = cbr.Quantity
                ' code changes for groups
                If msg.GroupProgramCode = "LT" And msg.GroupProgramCode <> "" Then
                    statusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                Else
                    statusCode = CourseBookingStatusLookup.CourseBookingStatuses.Active
                End If
                courseTypeCode = cbr.CourseTypeCode
                changeType = Nothing

                ' Get weeks
                If cbr.BookingLine_id <> salesCourseBookingId Then
                    changeStartWeek = sw
                    changeEndWeek = ew
                    courseBookingStartWeek = sw
                    courseBookingEndWeek = ew
                    salesCourseBookingId = cbr.BookingLine_id

                ElseIf cbr.Quantity > 0 Then
                    If sw.Week_id < courseBookingStartWeek.Week_id Then
                        changeStartWeek = sw
                        changeEndWeek = courseBookingStartWeek.AddWeeks(-1)
                        courseBookingStartWeek = changeStartWeek

                    ElseIf ew.Week_id > courseBookingEndWeek.Week_id Then
                        changeEndWeek = ew
                        changeStartWeek = courseBookingEndWeek.AddWeeks(1)
                        courseBookingEndWeek = changeEndWeek
                    Else
                        changeStartWeek = sw
                        changeEndWeek = ew
                        'Throw New InvalidOperationException("Can't figure out change history.")

                    End If

                ElseIf cbr.Quantity < 0 Then
                    If sw.Week_id = courseBookingStartWeek.Week_id And ew.Week_id = courseBookingEndWeek.Week_id Then
                        changeStartWeek = sw
                        changeEndWeek = ew
                        statusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                        changeType = ProdLookups.CourseBookingChangeTypeLookup.CourseBookingChangeTypes.Cancellation
                    ElseIf ew.Week_id = courseBookingEndWeek.Week_id Then
                        changeStartWeek = sw
                        changeEndWeek = courseBookingEndWeek
                        courseBookingEndWeek = sw.AddWeeks(-1)
                    ElseIf sw.Week_id = courseBookingStartWeek.Week_id Then
                        changeStartWeek = courseBookingStartWeek
                        changeEndWeek = ew
                        courseBookingEndWeek = ew.AddWeeks(1)
                    Else
                        changeStartWeek = sw
                        changeEndWeek = ew
                        'Throw New InvalidOperationException("Can't figure out change history.")
                    End If

                End If

                If cbr.CourseTypeCode.StartsWith(ArticleLookup.ArticleCategoryCodes.Extension) Then
                    ' Flip stuff for cancelled extensions
                    If weeks > 0 Then
                        changeType = ProdLookups.CourseBookingChangeTypeLookup.CourseBookingChangeTypes.Extension
                        statusCode = CourseBookingStatusLookup.CourseBookingStatuses.Active
                    Else
                        changeType = ProdLookups.CourseBookingChangeTypeLookup.CourseBookingChangeTypes.Termination
                        statusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                    End If
                    courseTypeCode = courseTypeCode.Substring(3).Trim()

                ElseIf cbr.CourseTypeCode.StartsWith(ArticleLookup.ArticleCategoryCodes.Termination) Then
                    ' Flip stuff for terminations
                    ' weeks = (-weeks)
                    If weeks < 0 Then
                        changeType = ProdLookups.CourseBookingChangeTypeLookup.CourseBookingChangeTypes.Termination
                        statusCode = CourseBookingStatusLookup.CourseBookingStatuses.Cancelled
                    Else
                        changeType = ProdLookups.CourseBookingChangeTypeLookup.CourseBookingChangeTypes.Extension
                        statusCode = CourseBookingStatusLookup.CourseBookingStatuses.Active
                    End If
                    courseTypeCode = courseTypeCode.Substring(3).Trim()

                ElseIf cbr.Quantity < 0 Then
                    If changeType Is Nothing Then
                        changeType = ProdLookups.CourseBookingChangeTypeLookup.CourseBookingChangeTypes.BookingReduction
                    End If
                Else
                    changeType = ProdLookups.CourseBookingChangeTypeLookup.CourseBookingChangeTypes.BookingAddition
                End If

                ' -- Add change record
                cbc = FindCourseBookingChange( _
                    bkn.ChangeList _
                    , msg.ProgramCode, courseTypeCode, cbr.DestinationCode, changeStartWeek.Code, changeEndWeek.Code, weeks, cbr.ChangeDate.Date _
                    , changeType _
                    , statusCode _
                    , touchedList _
                    )

                If cbc Is Nothing Then
                    cbc = bkn.ChangeList.AddNew( _
                    msg.ProgramCode, msg.ProductCode, courseTypeCode, cbr.DestinationCode, changeStartWeek, changeEndWeek, weeks, cbr.ChangeDate.Date _
                    , changeType _
                    , statusCode _
                    )
                End If
                touchedList.Add(cbc)
            End If
        Next
    End Sub

    Private Shared Function FindCourseBookingChange( _
                ByVal list As CourseBookingChangeListBase _
                , ByVal programCode As String _
                , ByVal courseTypeCode As String _
                , ByVal destinationCode As String _
                , ByVal startWeekCode As String _
                , ByVal endWeekCode As String _
                , ByVal weeks As Integer _
                , ByVal changeDate As Date _
                , ByVal type As String _
                , ByVal statusCode As String _
                , ByVal touchedList As ArrayList _
                ) As CourseBookingChange

        For Each c As CourseBookingChange In list
            If c.ProgramCode = programCode.Trim _
                    AndAlso c.CourseTypeCode = courseTypeCode.Trim _
                    AndAlso c.DestinationCode = destinationCode.Trim _
                    AndAlso c.StartWeekCode = startWeekCode.Trim _
                    AndAlso c.EndWeekCode = endWeekCode.Trim _
                    AndAlso c.Weeks = weeks _
                    AndAlso c.ChangeDate = changeDate _
                    AndAlso c.ChangeTypeCode = type.Trim _
                    AndAlso c.StatusCode = statusCode.Trim _
                    Then
                ' Only return the guy once
                If touchedList Is Nothing OrElse Not touchedList.Contains(c) Then
                    Return c
                End If
            End If
        Next
        Return Nothing
    End Function

    Private Function LoadCourse(ByVal DestinationCode As String, ByVal ProgramCode As String, ByVal CourseNumber As String, ByVal SalesOfficeCode As String) As Production.Course
        Dim CourseId As Integer
        CourseId = GetCourseID(DestinationCode, ProgramCode, CourseNumber, SalesOfficeCode)
        If IsSomething(CourseId) Then
            Return New Production.Course(CourseId)
        End If
        Return Nothing
    End Function

    Public Function GetCourseID(ByVal DestinationCode As String, ByVal ProgramCode As String, ByVal BookingNum As String, ByVal SalesOfficeCode As String) As Integer
        Dim cm As New DataH.ConnectionManager
        Dim cmd As New Sprocs.asp_CourseFindByCourseTypeCode

        ' Set parameters
        With cmd.Parameters
            .DestinationCode = DestinationCode
            .ProgramCode = ProgramCode
            .CourseNumber = BookingNum
            .salesOfficeCode = SalesOfficeCode
            'output parameter so set the value as null
            .Course_id = Null.IntegerNull
        End With

        ' Execute query
        cm.ExecuteNonQuery(cmd, Constants.ElektraConnectionNamespace)
        Return cmd.Parameters.Course_id
    End Function

    Public Function GetActiveBookingLine(ByVal msg As CRMMessage.LSBooking) As CRMMessage.BookingLine
        For Each item As CRMMessage.BookingLine In msg.BookingLineItems
            If item.StatusCode = "AC" Then
                Return item
            End If
        Next
        Return Nothing
    End Function

    Private Function CreateCourseBookingWeeks(ByVal msg As CRMMessage.LSBooking, ByVal wl As BusinessObjects.WarningListBase _
          ) As ArrayList
        Dim cbw As CourseBookingWeeks
        Dim cbwList As New ArrayList
        Dim sysArticles As New List(Of CRMMessage.BookingLine)
        Dim extItems As List(Of CRMMessage.BookingLine)
        Dim terItems As List(Of CRMMessage.BookingLine)


        For Each bknline As CRMMessage.BookingLine In msg.BookingLineItems
            If bknline.StatusCode = "AC" OrElse bknline.StatusCode = "PRE" Then

                'Load System Articles - Transfer, Termination and Extension
                'TODO - Need to Change Based on the New TR Logic
                For Each articles As CRMMessage.BookingLine In msg.ArticleBookingLineItems
                    If articles.IsSytemArticle Then
                        If (Not articles.ArticleCode = SystemArticleTypeCodes.Transfer AndAlso _
                                articles.CourseNumber = bknline.CourseNumber AndAlso _
                            articles.DestinationCode = bknline.DestinationCode) OrElse _
                            (articles.ArticleCode = SystemArticleTypeCodes.Transfer AndAlso _
                               articles.Description.Trim = "TR-" + bknline.CourseNumber AndAlso _
                                articles.DestinationCode = bknline.DestinationCode) Then
                            sysArticles.Add(articles)
                        End If
                    End If
                Next

                'Filter Ex- for Create New Course Booking Weeks
                If IsSomething(sysArticles) AndAlso sysArticles.Count > 0 Then
                    'sysArticles.Sort(AddressOf SortArticle)
                    extItems = Me.FilterBySysTypeCode(SystemArticleTypeCodes.Extension, sysArticles, bknline)

                    ' If the EX is Not Present then Update the Current EndDate to EndDate
                    If IsSomething(extItems) AndAlso extItems.Count > 0 Then
                        'extItems.Sort()
                        '*****Remove the Termination Weeks from Course and Extensions
                        '1) Load All the Terminations order by EndDate
                        '2) Take a Termination, Load Related Extensions
                        '3) Correct the Quantity
                        '   a) Qty (of TE) = Quantity( of TE ) - Qty ( of EX )
                        '


                        terItems = Me.FilterTEAndTR(sysArticles, bknline)

                        If IsSomething(terItems) AndAlso terItems.Count > 0 Then
                            'terItems.Sort()
                            Dim extsByTE As New List(Of CRMMessage.BookingLine)

                            For Each ter As CRMMessage.BookingLine In terItems
                                'extsByTE = LoadExtsBeforeBknLine(ter, extItems)
                                'ELEK-4917
                                extsByTE = LoadExtsBeforeBknLineByIndex(sysArticles.IndexOf(ter), ter.BookingLine_id, extItems, sysArticles, bknline)
                                If IsSomething(extsByTE) AndAlso extsByTE.Count > 0 Then

                                    For Each ext As CRMMessage.BookingLine In extsByTE
                                        If ter.Quantity = ext.Quantity Then
                                            sysArticles.Remove(ext)
                                            ter.Quantity = 0
                                            Exit For
                                        ElseIf ter.Quantity > ext.Quantity Then
                                            sysArticles.Remove(ext)
                                            ter.Quantity -= ext.Quantity
                                        ElseIf ter.Quantity < ext.Quantity Then
                                            Me._findBookingLine_id = ext.BookingLine_id
                                            sysArticles.Find(AddressOf FindByBookingLineId).Quantity -= ter.Quantity
                                            ter.Quantity = 0
                                            Exit For
                                        End If
                                    Next
                                    'Remove the Quantity From Parent Course
                                    If ter.Quantity > 0 Then
                                        bknline.Quantity -= ter.Quantity
                                    End If
                                ElseIf ter.Quantity > 0 Then
                                    bknline.Quantity -= ter.Quantity
                                End If
                            Next

                        End If

                        'Create EX- Weeks
                        extItems = Me.FilterBySysTypeCode(SystemArticleTypeCodes.Extension, sysArticles, bknline)

                        If IsSomething(extItems) AndAlso extItems.Count > 0 Then
                            For Each ext As CRMMessage.BookingLine In extItems
                                'ELEK-4955
                                'Its an ILC Booking and Enddate is calculated in a different way for ILC booking so added this condition
                                If _soldProgram = SOLDPROGRAM_ILC Then
                                    If terItems.Count = 0 Then
                                        ext.EndDate = CalculateEndDate(ext.StartDate, ext.Quantity, bknline.ProgramCode, ext.EndDate, msg.PoseidonGroup_Id)
                                    ElseIf terItems.Count = 1 Then
                                        'ELEK-5011
                                        ext.EndDate = CalculateEndDate(ext.StartDate, ext.Quantity, bknline.ProgramCode, bknline.CurrentEndDate, msg.PoseidonGroup_Id)
                                    Else
                                        ext.EndDate = CalculateEndDate(ext.StartDate, ext.Quantity, bknline.ProgramCode, ext.EndDate, msg.PoseidonGroup_Id)
                                    End If
                                Else
                                    ext.EndDate = CalculateEndDate(ext.StartDate, ext.Quantity, bknline.ProgramCode, bknline.EndDate, msg.PoseidonGroup_Id)
                                End If
                                cbw = New CourseBookingWeeks(ext, _soldProgram)
                                If (cbw.WeekCount > 0) Then
                                    cbwList.Add(cbw)
                                End If
                            Next
                        End If

                        extItems.Clear()
                        terItems.Clear()
                        'Update BookingLine
                        bknline.EndDate = CalculateEndDate(bknline.StartDate, bknline.Quantity, bknline.ProgramCode, bknline.EndDate, msg.PoseidonGroup_Id)
                    Else
                        bknline.EndDate = CalculateEndDate(bknline.StartDate, bknline.CurrentQuantity, bknline.ProgramCode, bknline.EndDate, msg.PoseidonGroup_Id)
                    End If
                Else
                    bknline.EndDate = CalculateEndDate(bknline.StartDate, bknline.CurrentQuantity, bknline.ProgramCode, bknline.EndDate, msg.PoseidonGroup_Id)
                End If

                'reset SysArticles
                sysArticles.Clear()
            End If
            'Need to check this code with Gulrez
            'This is not required as we handled this in UpdateProgram method
            'If msg.ProgramCode = "ILSP" And bknline.ProgramCode = "ILS" Then
            '    bknline.ProgramCode = "ILSP"
            'End If
            cbw = New CourseBookingWeeks(bknline, _soldProgram)
            If (cbw.WeekCount > 0) Then
                cbwList.Add(cbw)
            End If
        Next
        Return cbwList
    End Function
    'ELEK-4917
    Private Function LoadExtsBeforeBknLineByIndex(ByVal bknLineIndex As Integer, ByVal terBookingLineId As Integer, ByVal extItems As List(Of CRMMessage.BookingLine), ByVal sysArticlesBag As List(Of CRMMessage.BookingLine), ByVal bknLine As BookingLine)
        Dim extList As New List(Of CRMMessage.BookingLine)
        For Each ext As CRMMessage.BookingLine In extItems
            If bknLineIndex < sysArticlesBag.IndexOf(ext) AndAlso ext.BookingLine_id < terBookingLineId AndAlso ext.ParentBookingLine_id = bknLine.BookingLine_id Then
                extList.Add(ext)
            End If
        Next
        Return extList
    End Function

    Private Function LoadExtsBeforeBknLine(ByVal bknLine As CRMMessage.BookingLine, ByVal extItems As List(Of CRMMessage.BookingLine))
        Dim extList As New List(Of CRMMessage.BookingLine)
        For Each ext As CRMMessage.BookingLine In extItems
            If ext.BookingLine_id < bknLine.BookingLine_id Then
                extList.Add(ext)
            End If
        Next
        Return extList
    End Function

    Public Function FindByBookingLineId(ByVal bkl As CRMMessage.BookingLine) As Boolean
        Dim bknLine As CRMMessage.BookingLine
        bknLine = DirectCast(bkl, BookingLine)
        If Me._findBookingLine_id = bknLine.BookingLine_id Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CalculateEndDate(ByVal startDate As Date, ByVal weeks As Integer, ByVal programCode As String, ByVal courseEndDate As Date, ByVal poseidonGroup_Id As Integer) As Date

        Dim endDate As Date
        If (_soldProgram = SOLDPROGRAM_LT Or _soldProgram = SOLDPROGRAM_ILC) And poseidonGroup_Id > 0 Then 'ELEK-4241 ,'ELEK-4806
            Select Case courseEndDate.DayOfWeek
                Case DayOfWeek.Monday
                    endDate = courseEndDate.AddDays(-3)
                Case DayOfWeek.Sunday
                    endDate = courseEndDate.AddDays(-2)
                Case Else
                    endDate = courseEndDate.AddDays(-1)
            End Select
        Else
            If programCode = Programs.AcademicYear OrElse _
                           programCode = Programs.PreparationWeeks OrElse _
                           programCode = Programs.InternationalBaccalaureate OrElse _
                           programCode = Programs.IBPrep OrElse _
                           programCode = Programs.InternationalBaccalaureatePreparation OrElse _
                           programCode = Programs.InternationalGerneralCertificate OrElse _
                           programCode = Programs.AcademicYearProfessionals OrElse _
                           programCode = Programs.ALevels Then
                If startDate.DayOfWeek > 1 Then
                    weeks -= 1
                End If
            End If
            If programCode = Programs.EFCorporate Then
                endDate = courseEndDate
            Else
                endDate = startDate.AddDays(weeks * 7)
            End If

        End If
        Return endDate
    End Function

    Private Function FilterBySysTypeCode(ByVal sysTypeCode As String, ByVal sysArticles As List(Of BookingLine), ByVal bknLine As BookingLine) As List(Of BookingLine)
        Dim filteredSysArticles As New List(Of BookingLine)
        For Each sysArticle As CRMMessage.BookingLine In sysArticles
            If sysArticle.ArticleCode = sysTypeCode AndAlso sysArticle.ParentBookingLine_id = bknLine.BookingLine_id Then
                filteredSysArticles.Add(sysArticle)
            End If
        Next
        Return filteredSysArticles
    End Function

    Private Function FilterTEAndTR(ByVal sysArticles As List(Of BookingLine), ByVal bknLine As BookingLine) As List(Of BookingLine)
        Dim filteredSysArticles As New List(Of BookingLine)
        For Each sysArticle As CRMMessage.BookingLine In sysArticles
            If (sysArticle.ArticleCode = SystemArticleTypeCodes.Termination OrElse sysArticle.ArticleCode = SystemArticleTypeCodes.Transfer) AndAlso sysArticle.ParentBookingLine_id = bknLine.BookingLine_id Then
                filteredSysArticles.Add(sysArticle)
            End If
        Next
        Return filteredSysArticles
    End Function
    'Private Function CorrectionsOnBookingLine(ByVal msg As CRMMessage.LSBooking, ByVal bknLine As CRMMessage.BookingLine) As CRMMessage.BookingLine
    '    Dim quantity As Integer
    '    Dim sysArticles As New List(Of CRMMessage.BookingLine)
    '    Dim extItems As List(Of BookingLine)

    '    quantity = bknLine.CurrentQuantity
    '    For Each articles As CRMMessage.BookingLine In msg.ArticleBookingLineItems
    '        If articles.IsSytemArticle AndAlso articles.CourseNumber = bknLine.CourseNumber AndAlso _
    '                articles.DestinationCode = bknLine.DestinationCode Then
    '            sysArticles.Add(articles)
    '        End If
    '    Next

    '    If IsSomething(sysArticles) AndAlso sysArticles.Count > 0 Then
    '        'Create the Extension Weeks
    '        extItems = Me.FilterBySysTypeCode(SystemArticleTypeCodes.Extension, sysArticles)
    '        If IsSomething(extItems) AndAlso extItems.Count > 0 Then
    '            For Each extItem As CRMMessage.BookingLine In extItems
    '                quantity -= extItem.Quantity
    '            Next
    '        End If
    '    End If
    '    'update the bknLine info with corrected Quantity
    '    bknLine.CurrentQuantity = quantity
    '    bknLine.EndDate = Me.CalculateEndDate(bknLine.StartDate, quantity)

    '    Return bknLine
    'End Function

    ' Version 1
    'Private Function CreateCourseBookingWeeks(ByVal msg As CRMMessage.LSBooking, ByVal wl As BusinessObjects.WarningListBase _
    '       ) As ArrayList
    '    Dim cbw As CourseBookingWeeks
    '    Dim extWeeks As CourseBookingWeeks
    '    Dim currentCrsWeeksList As New List(Of CourseBookingWeeks)
    '    Dim cbwList As New ArrayList
    '    Dim systemArticles As New List(Of BookingLine)
    '    Dim extItems As List(Of BookingLine)
    '    Dim terItems As List(Of BookingLine)
    '    Dim traItems As List(Of BookingLine)

    '    For Each bknline As CRMMessage.BookingLine In msg.BookingLineItems
    '        If bknline.StatusCode = "AC" Then
    '            cbw = New CourseBookingWeeks(bknline)
    '            If (cbw.WeekCount > 0) Then
    '                currentCrsWeeksList.Add(cbw)
    '            End If
    '            For Each articles As CRMMessage.BookingLine In msg.ArticleBookingLineItems
    '                If articles.IsSytemArticle AndAlso articles.CourseNumber = bknline.CourseNumber AndAlso articles.DestinationCode = bknline.DestinationCode Then
    '                    systemArticles.Add(articles)
    '                End If
    '            Next

    '            If IsSomething(systemArticles) AndAlso systemArticles.Count > 0 Then
    '                'Create the Extension Weeks
    '                extItems = Me.FilterBySysTypeCode(SystemArticleTypes.Extension, systemArticles)
    '                If IsSomething(extItems) AndAlso extItems.Count > 0 Then
    '                    For Each extItem As CRMMessage.BookingLine In extItems
    '                        extWeeks = New CourseBookingWeeks(extItem)
    '                        If extWeeks.WeekCount > 0 Then
    '                            currentCrsWeeksList.Add(cbw)
    '                        End If
    '                    Next
    '                End If

    '                'Remove the Termination Weeks
    '                terItems = Me.FilterBySysTypeCode(SystemArticleTypes.Termination, systemArticles)
    '                If IsSomething(terItems) AndAlso terItems.Count > 0 Then
    '                    For Each terItem As CRMMessage.BookingLine In terItems
    '                        UpdateCurrentCourseWeeks(terItem, currentCrsWeeksList)
    '                    Next
    '                End If

    '                'Remove the Transfer Weeks
    '                traItems = Me.FilterBySysTypeCode(SystemArticleTypes.Transfer, systemArticles)
    '                If IsSomething(traItems) AndAlso traItems.Count > 0 Then
    '                    For Each terItem As CRMMessage.BookingLine In terItems
    '                        UpdateCurrentCourseWeeks(terItem, currentCrsWeeksList)
    '                    Next
    '                End If
    '            End If
    '            'Need to Do Corrections Based on Parent Quantity.



    '            If (cbw.WeekCount > 0) Then
    '                cbwList.Add(cbw)
    '            End If
    '            systemArticles = Nothing
    '            currentCrsWeeksList = Nothing
    '        End If
    '    Next
    '    Return cbwList
    'End Function

    'Private Sub UpdateCurrentCourseWeeks(ByVal removeWeeks As CRMMessage.BookingLine, ByVal currCrsWeeks As List(Of CourseBookingWeeks))
    '    For Each currCrsWeek As CourseBookingWeeks In currCrsWeeks
    '        'Remove complete weeks
    '        If currCrsWeek.StartDate = removeWeeks.StartDate AndAlso currCrsWeek.EndDate = removeWeeks.EndDate Then
    '            currCrsWeek.StartDate = CourseBookingWeeks.GetStartDate(currCrsWeek.ProgramCode, removeWeeks.StartDate)
    '        End If
    '        'Move the Start Date
    '        If currCrsWeek.StartDate <= removeWeeks.StartDate AndAlso currCrsWeek.StartDate < removeWeeks.EndDate Then
    '            currCrsWeek.StartDate = CourseBookingWeeks.GetStartDate(currCrsWeek.ProgramCode, removeWeeks.EndDate)
    '        End If
    '        'Move the End Date
    '        If currCrsWeek.EndDate > removeWeeks.StartDate AndAlso currCrsWeek.EndDate <= removeWeeks.EndDate Then
    '            currCrsWeek.EndDate = CourseBookingWeeks.GetEndDate(currCrsWeek.ProgramCode, removeWeeks.StartDate)
    '        End If

    '        currCrsWeek.StartWeek = New Week(currCrsWeek.StartDate)
    '        currCrsWeek.EndWeek = New Week(currCrsWeek.EndDate)
    '        currCrsWeek.Weeks = New WeekSpan(currCrsWeek.StartWeek, currCrsWeek.EndWeek)
    '    Next
    'End Sub



    'Private Function FilterTEAndTRItems(ByVal sysArticles As List(Of BookingLine)) As List(Of BookingLine)
    '    Dim filteredSysArticles As New List(Of BookingLine)
    '    For Each sysArticle As CRMMessage.BookingLine In sysArticles
    '        If sysArticle.ArticleCode.Trim = SystemArticleTypeCodes.Transfer OrElse _
    '            sysArticle.ArticleCode.Trim = SystemArticleTypeCodes.Termination Then
    '            filteredSysArticles.Add(sysArticle)
    '        End If
    '    Next
    '    Return filteredSysArticles
    'End Function

    Private Shared Function HasStarted(ByVal bkn As Production.Booking, ByVal destinationCode As String) As Boolean
        Dim cb As CourseBooking
        For Each cb In bkn.CourseBookingList
            If cb.Status.IsActive _
                    And cb.DestinationCode = destinationCode _
                    And cb.StartDate <= Now Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Function FindBookingArticle(ByVal bar As CRMMessage.BookingLine, ByVal bookingArticles As ICollection) As Production.BookingArticle
        Dim ba As Production.BookingArticle
        For Each ba In bookingArticles
            If IsBookingArticleMatch(bar, ba) Then
                Return ba
            End If
        Next
        Return Nothing
    End Function

    Private Function FindBookingArticleNew(ByVal bar As CRMMessage.BookingLine, ByVal bookingArticles As ICollection) As Production.BookingArticle
        Dim ba As Production.BookingArticle
        'ELEK-4609
        Dim temp As Production.BookingArticle = Nothing
        For Each ba In bookingArticles
            If IsBookingArticleMatchNew(bar, ba) Then
                'ELEK-4609
                If bar.ArticleCode.StartsWith("RES") Then
                    If ba.StartDate = bar.StartDate AndAlso bar.EndDate = ba.EndDate Then
                        Return ba
                    End If
                    temp = ba
                Else
                    Return ba
                End If
            End If
        Next
        'ELEK-4609
        If IsSomething(temp) Then
            Return temp
        End If
        Return Nothing
    End Function
    Private Function IsBookingArticleMatch(ByVal arl As CRMMessage.BookingLine, ByVal ba As Production.BookingArticle) As Boolean
        Dim uc As String = ""
        Select Case arl.UnitCode.Trim()
            Case "UNIT"
                uc = "Unit"
            Case "LESSON"
                uc = "Lesson"
            Case "DAY"
                uc = "Day"
            Case "WEEK"
                uc = "Week"
        End Select

        Select Case True
            Case arl.DestinationCode = "" And ba.DestinationCode > "" : Return False
            Case Not arl.DestinationCode = "" AndAlso arl.DestinationCode <> ba.DestinationCode : Return False
            Case arl.ArticleCode <> ba.ArticleCode : Return False
            Case arl.StatusCode <> ba.StatusCode : Return False
            Case ToSafeDate(arl.StartDate) <> ba.StartDate : Return False
            Case ToSafeDate(arl.EndDate) <> ba.EndDate : Return False
                'Case arl.Quantity <> ba.Units : Return False
            Case uc <> ba.UnitCode : Return False
            Case Else : Return True
        End Select
    End Function
    Private Function IsBookingArticleMatchNew(ByVal arl As CRMMessage.BookingLine, ByVal ba As Production.BookingArticle) As Boolean
        Dim uc As String = ""
        Select Case arl.UnitCode.Trim()
            Case "UNIT"
                uc = "Unit"
            Case "LESSON"
                uc = "Lesson"
            Case "DAY"
                uc = "Day"
            Case "WEEK"
                uc = "Week"
        End Select

        Select Case True
            Case arl.DestinationCode = "" And ba.DestinationCode > "" : Return False
            Case Not arl.DestinationCode = "" AndAlso arl.DestinationCode <> ba.DestinationCode : Return False
            Case arl.ArticleCode <> ba.ArticleCode : Return False
            Case ba.StatusCode = "AC" OrElse ba.StatusCode = "CF" : Return True
            Case ToSafeDate(arl.StartDate) <> ba.StartDate : Return False
            Case ToSafeDate(arl.EndDate) <> ba.EndDate : Return False
                'Case arl.Quantity <> ba.Units : Return False
            Case uc <> ba.UnitCode : Return False
            Case Else : Return True
        End Select
    End Function

    Private Function ToSafeDate(ByVal value As Object) As Date
        Return CDate(Null.ReplaceNull(value, Null.DateNull))
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

    Private Sub UpdateCbChangeRemoveCourseBookingData(ByVal bkn As Production.Booking)
        ' Remove course booking change data for course bookings
        For Each cb As CourseBooking In bkn.CourseBookingList
            For Each cbc As CourseBookingChange In cb.ChangeList.ToArray
                If cbc.ChangeType.Adjustment <> 0 Then
                    cb.ChangeList.Remove(cbc)
                End If
            Next
        Next
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Function SortArticle(ByVal extList1 As CRMMessage.BookingLine, ByVal extList2 As CRMMessage.BookingLine)
        Return extList1.StartDate.CompareTo(extList2.StartDate)
    End Function
End Class
' Code Befor Ex Changes
'Public Class CourseBookingWeeks
'    Public DestinationCode As String
'    Public ReadOnly ProgramCode As String
'    Public ReadOnly CourseTypeCode As String
'    Public Weeks As WeekSpan
'    Public StartWeek As Week
'    Public EndWeek As Week
'    Public StartDate As Date
'    Public EndDate As Date
'    Public Description As String
'    'Public SalesDate As Date
'    Public ExamCode As String
'    Public YearsOfStudy As String
'    'Public ReadOnly SalesOfficeCode As String
'    Sub New(ByVal bknLine As CRMMessage.BookingLine)
'        Me.StartDate = GetStartDate(bknLine.ProgramCode, bknLine.StartDate)
'        Me.EndDate = GetEndDate(bknLine.ProgramCode, bknLine.EndDate)
'        Me.DestinationCode = bknLine.DestinationCode
'        Me.ProgramCode = bknLine.ProgramCode
'        Me.CourseTypeCode = bknLine.CourseNumber
'        Me.StartWeek = New Week(Me.StartDate)
'        Me.EndWeek = New Week(Me.EndDate)
'        Me.Weeks = New WeekSpan(Me.StartWeek, Me.EndWeek)
'        Me.Description = bknLine.Description
'        Me.ExamCode = bknLine.ExamCode
'        Me.YearsOfStudy = bknLine.YearsOfStudy
'    End Sub

Public Class CourseBookingWeeks
    Public DestinationCode As String
    Public ReadOnly ProgramCode As String
    Public ReadOnly CourseTypeCode As String
    Public Weeks As WeekSpan
    Public StartWeek As Week
    Public EndWeek As Week
    Public StartDate As Date
    Public EndDate As Date
    Public Description As String
    Public ExamCode As String
    Public YearsOfStudy As String
    Public StatusCode As String
    Public PoseidonTermReason As String
    Public TotalPrice As Decimal    'ELEK-6127
    Public ISPRW As Boolean
    Public ISPIE As Boolean
    Public IsCourseLeader As Boolean
    Public BookingLine_ID As Integer


    Sub New(ByVal bknLine As CRMMessage.BookingLine, ByVal soldProgramCode As String)
        Me.DestinationCode = bknLine.DestinationCode
        Me.ProgramCode = bknLine.ProgramCode
        Me.CourseTypeCode = bknLine.CourseNumber
        Me.Description = bknLine.Description
        Me.ExamCode = bknLine.ExamCode
        Me.YearsOfStudy = bknLine.YearsOfStudy
        Me.StartDate = GetStartDate(bknLine.ProgramCode, bknLine.StartDate, soldProgramCode)
        Me.EndDate = GetEndDate(bknLine.ProgramCode, bknLine.EndDate, soldProgramCode)
        Me.StartWeek = New Week(Me.StartDate)
        Me.EndWeek = New Week(Me.EndDate)
        Me.Weeks = New WeekSpan(Me.StartWeek, Me.EndWeek)
        Me.StatusCode = bknLine.StatusCode
        Me.PoseidonTermReason = bknLine.TerminationReason
        Me.TotalPrice = bknLine.Totalprice  'ELEK-6127
        Me.ISPRW = bknLine.IsPRWCourse
        Me.ISPIE = bknLine.IsPIERequest
        Me.IsCourseLeader = bknLine.IsCourseLeader
        Me.BookingLine_ID = bknLine.BookingLine_id
    End Sub

    Sub New(ByVal cbw As CourseBookingWeeks)
        Me.StartDate = cbw.StartDate
        Me.EndDate = cbw.EndDate
        Me.DestinationCode = cbw.DestinationCode
        Me.ProgramCode = cbw.ProgramCode
        Me.CourseTypeCode = cbw.CourseTypeCode
        Me.StartWeek = New Week(Me.StartDate)
        Me.EndWeek = New Week(Me.EndDate)
        Me.Weeks = New WeekSpan(Me.StartWeek, Me.EndWeek)
        'Me.SalesOfficeCode = cbw.SalesOfficeCode
        Me.ExamCode = cbw.ExamCode
        Me.YearsOfStudy = cbw.YearsOfStudy
        'Me.SalesDate = cbw.SalesDate
        Me.StatusCode = cbw.StatusCode
        Me.PoseidonTermReason = cbw.PoseidonTermReason
        Me.TotalPrice = cbw.TotalPrice  'ELEK-6127
        Me.ISPRW = cbw.ISPRW
        Me.ISPIE = cbw.ISPIE
        Me.IsCourseLeader = cbw.IsCourseLeader
        Me.BookingLine_ID = cbw.BookingLine_ID
    End Sub

    Public ReadOnly Property StartWeekCode() As String
        Get
            Return Me.StartWeek.Code
        End Get
    End Property

    Public ReadOnly Property EndWeekCode() As String
        Get
            Return Me.EndWeek.Code
        End Get
    End Property

    Public ReadOnly Property WeekCount() As Integer
        Get
            Return (Me.EndWeek.Week_id - Me.StartWeek.Week_id + 1)
        End Get
    End Property
    Public Shared Function GetStartDate(ByVal programCode As String, ByVal startDate As Date, ByVal soldProgram As String) As Date
        If (soldProgram = "LT" Or soldProgram = "ILC") Then
            Select Case startDate.DayOfWeek  'ELEK-4241, 'ELEK-4806
                Case DayOfWeek.Friday
                    Return startDate.AddDays(3)
                Case DayOfWeek.Saturday
                    Return startDate.AddDays(2)
                Case Else
                    Return startDate.AddDays(1)
            End Select
        Else
            If programCode = Programs.EFCorporate Then
                Return startDate
            End If
            If startDate.DayOfWeek = DayOfWeek.Sunday Then
                Return startDate.AddDays(1)
            ElseIf startDate.DayOfWeek = DayOfWeek.Monday Then
                Return startDate
            Else
                If programCode = Programs.AcademicYear OrElse _
                    programCode = Programs.AcademicYearProfessionals OrElse _
                    programCode = Programs.PreparationWeeks OrElse _
                    programCode = Programs.InternationalBaccalaureate OrElse _
                    programCode = Programs.IBPrep OrElse _
                    programCode = Programs.InternationalBaccalaureatePreparation OrElse _
                    programCode = Programs.ALevels OrElse _
                    programCode = Programs.InternationalGerneralCertificate Then 'ELEK-4970

                    Return startDate.AddDays(7 + DayOfWeek.Monday - startDate.DayOfWeek)
                Else
                    Return startDate
                End If
            End If
        End If
    End Function
    Public Shared Function GetEndDate(ByVal programCode As String, ByVal startDate As Date, ByVal soldProgram As String) As Date
        If soldProgram = "LT" Then
            Return startDate
        Else
            If programCode = Programs.EFCorporate Then
                Return startDate
            End If
            If startDate.DayOfWeek = DayOfWeek.Sunday Then
                Return startDate.AddDays(-2)
            ElseIf startDate.DayOfWeek = DayOfWeek.Monday Then
                Return startDate.AddDays(-3)
            Else
                Return startDate.AddDays(DayOfWeek.Friday - startDate.DayOfWeek)
            End If
        End If

    End Function

End Class



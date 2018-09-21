Public Class StudentAdapter
    Implements IMessageAdapter(Of CRMMessage.Customer)

    Private _booking As Production.Booking
    '''' <summary>
    '''' 
    '''' </summary>
    '''' <param name="bookingNum"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Load(ByVal bookingNum As String) As Production.Student
    '    Dim ba As New BookingAdapter
    '    Return ba.Load(bookingNum).Student
    'End Function
    Public Property Booking() As Production.Booking
        Get
            Return Me._booking
        End Get
        Set(ByVal value As Production.Booking)
            _booking = value
        End Set
    End Property

    ''' <summary>
    '''  Updates the value from the msg-derived Individual object to the 
    '''  mapping Individual object loaded from the production system.
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="sourceSysCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update(ByVal msg As CRMMessage.Customer, ByVal sourceSysCode As String) _
            As ResultStepInfo Implements IMessageAdapter(Of CRMMessage.Customer).Update

        Dim stepInfo As ResultStepInfo

        'If (msg.Booking_id = "") Then 'did those bastards send us a contact again?
        ''CODE CHANGED: Change this
        'stepInfo = New ResultStepInfo

        'stepInfo.Status = IntegrationResult.Unwanted
        'stepInfo.Warnings.Add("The supplied booking does not have a booking number.")

        'Else
        stepInfo = Me.UpdateStudent(msg, sourceSysCode)
        'End If

        Return stepInfo
    End Function


    Private Function UpdateStudent(ByVal msg As CRMMessage.Customer, ByVal sourceSysCode As String) _
        As ResultStepInfo

        Dim stu As Production.Student
        Dim stepInfo As ResultStepInfo
        Dim wl As BusinessObjects.WarningListBase
        Dim isStaleMsg As Boolean

        wl = New BusinessObjects.WarningListBase
        isStaleMsg = False

        stu = Me.Booking.Student
        Try
            With msg
                'stu.NickName = .NickName
                stu.FirstName = .FirstNameEn
                stu.LastName = .LastNameEn
                stu.MiddleName = .MiddleNameEn                   ' code added for elek-3405
                stu.SalesCustomerId = .Customer_id               ' code changes to save customer_id in student table for my LT.
                stu.HomeRegionCode = .RegionStateCode
                stu.PostalCode = .PostalCode
                stu.HomeCountryCode = .CountryCode
                stu.HomePhone = .MainPhone
                stu.CellPhone = .MobilePhone
                stu.CitizenshipCode = .NationalityCode
                stu.CountryOfBirthCode = .BirthCountryCode
                stu.CityOfBirth = .BirthCity
                stu.NativeLanguageCode = .LanguageCode
                stu.Email = .Email1
                If stu.OtherEmail = String.Empty Then
                    stu.OtherEmail = .Email1
                End If
                stu.BirthDate = .DateOfBirth
                stu.GenderCode = .GenderCode
                stu.HomeAddress1 = .HomeAddress1
                stu.HomeAddress2 = .HomeAddress2
                stu.HomeAddress3 = .HomeAddress3
                stu.StateOrProvince = .StateOrProvince
                stu.IntegrationTimestamp = Now()
                stu.SalesSystemCode = "Poseidon"
                stu.HomeCity = .City
                stu.PassportNumber = .PassportNum
                stu.DoNotEmail = .PoseidonDoNotEmail
                stu.SalesRegion = .InstanceName
                stu.AetnaNumber = .AetnaNumber
                stu.IsVIP = .IsVipCustomer 'Changes for VIP students
                stu.ISProtectedIdentity = .IsProtectedIdentity 'changes for ISProtectedIdentity ELEK-8762

            End With

        Catch ex As Exception
            'wl.Add(BusinessObjects.WarningSeverities.Error, ex.Message, ex.Source)
            'ExceptionManager.Publish(ex)
            Throw (ex)
        Finally
            stepInfo = Me.HandleUpdateInfo(wl, isStaleMsg)
        End Try

        Return stepInfo

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
End Class

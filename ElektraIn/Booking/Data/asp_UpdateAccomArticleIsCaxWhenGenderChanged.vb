Imports System.Data.SqlClient
Namespace Sprocs
    Friend Class asp_UpdateAccomArticleIsCaxWhenGenderChanged
        Inherits DataH.EFSqlCommand
        Dim _params As New ParametersClass(Me.SqlCmd)

        Public Sub New()
            MyBase.New()
            Dim p As SqlParameter
            Me.SqlCmd.CommandText = "asp_UpdateAccomArticleIsCaxWhenGenderChanged"
            Me.SqlCmd.CommandType = System.Data.CommandType.StoredProcedure

            With Me.SqlCmd.Parameters
                p = .Add(New SqlParameter("@RETURN_VALUE", SqlDbType.Int, 0)) '
                p.Direction = ParameterDirection.ReturnValue  'special direction
                p = .Add(New SqlParameter("@BookingId", SqlDbType.Int, 100)) '
                p = .Add(New SqlParameter("@oldGenderCode", SqlDbType.Char, 1)) '
                p = .Add(New SqlParameter("@NewGenderCode", SqlDbType.Char, 1)) '

            End With

        End Sub

        Friend ReadOnly Property Parameters() As ParametersClass
            Get
                Return _params
            End Get
        End Property

        Friend Class ParametersClass
            Dim _sqlCmd As New SqlCommand

            Friend Sub New(ByVal cmd As SqlCommand)
                _sqlCmd = cmd
            End Sub

            Property RETURN_VALUE() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@RETURN_VALUE").Value, Integer)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@RETURN_VALUE").Value = Value
                End Set
            End Property

            Property BookingId() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@BookingId").Value, String)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@BookingId").Value = Value
                End Set
            End Property
            Property oldGenderCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@OldGenderCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@OldGenderCode").Value = Value
                End Set
            End Property
            Property newGenderCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@NewGenderCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@NewGenderCode").Value = Value
                End Set
            End Property
        End Class

    End Class

    Friend Class UpdateLSTransporationChanges
        Inherits DataH.EFSqlCommand
        Dim _params As New ParametersClass(Me.SqlCmd)

        Public Sub New()
            MyBase.New()
            Dim p As SqlParameter
            Me.SqlCmd.CommandText = "LSTransportation_UpdateLSTransporationChanges"
            Me.SqlCmd.CommandType = System.Data.CommandType.StoredProcedure

            With Me.SqlCmd.Parameters
                p = .Add(New SqlParameter("@RETURN_VALUE", SqlDbType.Int, 0)) '
                p.Direction = ParameterDirection.ReturnValue  'special direction
                p = .Add(New SqlParameter("@Booking_Id", SqlDbType.Int, 100)) '
                p = .Add(New SqlParameter("@CourseBooking_Id", SqlDbType.Int, 100)) '

            End With

        End Sub

        Friend ReadOnly Property Parameters() As ParametersClass
            Get
                Return _params
            End Get
        End Property

        Friend Class ParametersClass
            Dim _sqlCmd As New SqlCommand

            Friend Sub New(ByVal cmd As SqlCommand)
                _sqlCmd = cmd
            End Sub

            Property RETURN_VALUE() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@RETURN_VALUE").Value, Integer)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@RETURN_VALUE").Value = Value
                End Set
            End Property
            Property Booking_Id() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@Booking_Id").Value, Integer)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@Booking_Id").Value = Value
                End Set
            End Property
            Property CoursebookingId() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@CourseBooking_Id").Value, Integer)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@CourseBooking_Id").Value = Value
                End Set
            End Property


        End Class

    End Class
End Namespace
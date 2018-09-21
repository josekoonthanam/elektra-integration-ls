Imports System.Data.SqlClient

Namespace Sprocs
    Friend Class asp_GetModifiedBookingForAutoconfirmation
        Inherits DataH.EFSqlCommand
        'Builds command for sproc asp_GetModifiedBookingForAutoconfirmation
        'Added for Auto Confirmation Changes
        Dim _params As New ParametersClass(Me.SqlCmd)

        Public Sub New()
            MyBase.New()
            Dim p As SqlParameter
            Me.SqlCmd.CommandText = "asp_GetModifiedBookingForAutoconfirmation"
            Me.SqlCmd.CommandType = System.Data.CommandType.StoredProcedure

            With Me.SqlCmd.Parameters
                p = .Add(New SqlParameter("@RETURN_VALUE", SqlDbType.Int, 0))  '
                p.Direction = ParameterDirection.ReturnValue  'special direction
                p = .Add(New SqlParameter("@GroupStatus", SqlDbType.Char, 10))
                p = .Add(New SqlParameter("@Bookingid", SqlDbType.Int))
                p = .Add(New SqlParameter("@GroupCode", SqlDbType.VarChar, 50))
                p = .Add(New SqlParameter("@PROGRAMCODE", SqlDbType.VarChar, 1000))
                p = .Add(New SqlParameter("@Destination", SqlDbType.VarChar, 20))
                p = .Add(New SqlParameter("@COURSEWEEK", SqlDbType.VarChar, 10))
                p = .Add(New SqlParameter("@STATUS", SqlDbType.VarChar, 50))
            End With

        End Sub

        Friend ReadOnly Property Parameters() As ParametersClass
            Get
                Return _params
            End Get
        End Property

        Friend Class ParametersClass
            Dim _sqlCmd As New SqlCommand()

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

            Property GroupStatus() As String
                Get
                    Return CType(_sqlCmd.Parameters("@GroupStatus").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@GroupStatus").Value = Value
                End Set
            End Property

            Property Bookingid() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@Bookingid").Value, Integer)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@Bookingid").Value = Value
                End Set
            End Property

        Property GroupCode() As String
            Get
                Return CType(_sqlCmd.Parameters("@GroupCode").Value, String)
            End Get

            Set(ByVal Value As String)
                _sqlCmd.Parameters("@GroupCode").Value = Value
            End Set
        End Property

        Property PROGRAMCODE() As String
            Get
                Return CType(_sqlCmd.Parameters("@PROGRAMCODE").Value, String)
            End Get

            Set(ByVal Value As String)
                _sqlCmd.Parameters("@PROGRAMCODE").Value = Value
            End Set
        End Property

        Property Destination() As String
            Get
                Return CType(_sqlCmd.Parameters("@DESTINATION").Value, String)
            End Get

            Set(ByVal Value As String)
                _sqlCmd.Parameters("@DESTINATION").Value = Value
            End Set
        End Property

        Property COURSEWEEK() As String
            Get
                Return CType(_sqlCmd.Parameters("@COURSEWEEK").Value, String)
            End Get

            Set(ByVal Value As String)
                _sqlCmd.Parameters("@COURSEWEEK").Value = Value
            End Set
        End Property

        Property STATUS() As String
            Get
                Return CType(_sqlCmd.Parameters("@STATUS").Value, String)
            End Get

            Set(ByVal Value As String)
                _sqlCmd.Parameters("@STATUS").Value = Value
            End Set
            End Property
        End Class

    End Class

End Namespace


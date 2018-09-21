Imports System.Data.SqlClient
Namespace Sprocs
    Friend Class asp_RemoveRequiredCheckItemForBooking
        Inherits DataH.EFSqlCommand
        Dim _params As New ParametersClass(Me.SqlCmd)

        Public Sub New()
            MyBase.New()
            Dim p As SqlParameter
            Me.SqlCmd.CommandText = "RemoveRequiredCheckItemForBooking"
            Me.SqlCmd.CommandType = System.Data.CommandType.StoredProcedure

            With Me.SqlCmd.Parameters
                p = .Add(New SqlParameter("@RETURN_VALUE", SqlDbType.Int, 0)) '
                p.Direction = ParameterDirection.ReturnValue  'special direction
                p = .Add(New SqlParameter("@SalesBookingId", SqlDbType.VarChar, 100)) '

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

            Property SalesBookingId() As String
                Get
                    Return CType(_sqlCmd.Parameters("@SalesBookingId").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@SalesBookingId").Value = Value
                End Set
            End Property
        End Class

    End Class
End Namespace
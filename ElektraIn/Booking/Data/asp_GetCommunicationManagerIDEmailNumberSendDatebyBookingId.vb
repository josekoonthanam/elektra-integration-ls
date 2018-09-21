﻿Imports System.Data.SqlClient

Namespace Sprocs
    Friend Class asp_GetCommunicationManagerIDEmailNumberSendDatebyBookingId
        Inherits DataH.EFSqlCommand
        'Builds command for sproc asp_BookingFindBySalesBookingId
        'Generated by the code spewer on 2/1/2007 3:45:00 PM
        Dim _params As New ParametersClass(Me.SqlCmd)

        Public Sub New()
            MyBase.New()
            Dim p As SqlParameter
            Me.SqlCmd.CommandText = "asp_GetCommunicationManagerIDEmailNumberSendDatebyBookingId"
            Me.SqlCmd.CommandType = System.Data.CommandType.StoredProcedure

            With Me.SqlCmd.Parameters
                p = .Add(New SqlParameter("@RETURN_VALUE", SqlDbType.Int, 0))  '
                p.Direction = ParameterDirection.ReturnValue  'special direction
                p = .Add(New SqlParameter("@Booking_id", SqlDbType.Int, 0))
                p = .Add(New SqlParameter("@DestinationCode", SqlDbType.VarChar, 10))
                p = .Add(New SqlParameter("@ProductCode", SqlDbType.VarChar, 10))  '
                'p.Direction = ParameterDirection.InputOutput  'special direction
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

            Property Booking_id() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@Booking_id").Value, Integer)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@Booking_id").Value = Value
                End Set
            End Property

            Property DestinationCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@DestinationCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@DestinationCode").Value = Value
                End Set
            End Property

            Property ProductCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@ProductCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@ProductCode").Value = Value
                End Set
            End Property
        End Class
    End Class
End Namespace


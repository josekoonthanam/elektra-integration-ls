﻿Imports System.Data.SqlClient
Namespace Sprocs
    Friend Class asp_PreGroupBookingCancelled
        Inherits DataH.EFSqlCommand
        'Builds command for sproc asp_RemoveInpBookings
        'Generated by the code spewer on 5/16/2007 5:40:57 PM
        Dim _params As New ParametersClass(Me.SqlCmd)

        Public Sub New()
            MyBase.New()
            Dim p As SqlParameter
            Me.SqlCmd.CommandText = "asp_PreGroupBookingCancelled"
            Me.SqlCmd.CommandType = System.Data.CommandType.StoredProcedure

            With Me.SqlCmd.Parameters
                p = .Add(New SqlParameter("@RETURN_VALUE", SqlDbType.Int, 0)) '
                p.Direction = ParameterDirection.ReturnValue  'special direction
                p = .Add(New SqlParameter("@SalesBookingId", SqlDbType.VarChar, 8000)) '
                p = .Add(New SqlParameter("@PosGroupCode", SqlDbType.VarChar, 8000))
                p = .Add(New SqlParameter("@BookingStatusCode", SqlDbType.VarChar, 8000)) '
                p = .Add(New SqlParameter("@IsGroupRemoved", SqlDbType.Bit, 0)) '
                p = .Add(New SqlParameter("@OldPosGroupCode", SqlDbType.VarChar, 8000))
                p = .Add(New SqlParameter("@IsLeaderBooking", SqlDbType.Bit, 0))
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

            Property IsLeaderBooking() As Boolean
                Get
                    Return CType(_sqlCmd.Parameters("@IsLeaderBooking").Value, String)
                End Get

                Set(ByVal Value As Boolean)
                    _sqlCmd.Parameters("@IsLeaderBooking").Value = Value
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
            Property PosGroupCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@PosGroupCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@PosGroupCode").Value = Value
                End Set
            End Property
            Property BookingStatusCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@BookingStatusCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@BookingStatusCode").Value = Value
                End Set
            End Property
            Property IsGroupRemoved() As Boolean
                Get
                    Return CType(_sqlCmd.Parameters("@IsGroupRemoved").Value, Boolean)
                End Get

                Set(ByVal Value As Boolean)
                    _sqlCmd.Parameters("@IsGroupRemoved").Value = Value
                End Set
            End Property
            Property OldPosGroupCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@OldPosGroupCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@OldPosGroupCode").Value = Value
                End Set
            End Property
        End Class
    End Class
End Namespace

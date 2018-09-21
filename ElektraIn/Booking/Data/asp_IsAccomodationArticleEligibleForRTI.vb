Imports System.Data.SqlClient

Namespace Sprocs
    Friend Class asp_IsAccomodationArticleEligibleForRTI
        Inherits DataH.EFSqlCommand
        'Builds command for sproc asp_IsAccomodationArticleEligibleForRTI
        'Added for Auto Confirmation Changes
        Dim _params As New ParametersClass(Me.SqlCmd)

        Public Sub New()
            MyBase.New()
            Dim p As SqlParameter
            Me.SqlCmd.CommandText = "asp_IsAccomodationArticleEligibleForRTI"
            Me.SqlCmd.CommandType = System.Data.CommandType.StoredProcedure

            With Me.SqlCmd.Parameters
                p = .Add(New SqlParameter("@RETURN_VALUE", SqlDbType.Int, 0))  '
                p.Direction = ParameterDirection.ReturnValue  'special direction
                p = .Add(New SqlParameter("@ArticleCode", SqlDbType.VarChar, 20))  '
                p = .Add(New SqlParameter("@IsEligible", SqlDbType.Int, 0))  '
                p.Direction = ParameterDirection.InputOutput  'special direction
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

            Property ArticleCode() As String
                Get
                    Return CType(_sqlCmd.Parameters("@ArticleCode").Value, String)
                End Get

                Set(ByVal Value As String)
                    _sqlCmd.Parameters("@ArticleCode").Value = Value
                End Set
            End Property

            Property IsEligible() As Integer
                Get
                    Return CType(_sqlCmd.Parameters("@IsEligible").Value, Integer)
                End Get

                Set(ByVal Value As Integer)
                    _sqlCmd.Parameters("@IsEligible").Value = Value
                End Set
            End Property

        End Class

    End Class
End Namespace


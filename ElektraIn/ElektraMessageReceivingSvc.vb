Public Class ElektraMessageReceivingSvc
    Inherits MarshalByRefObject
    Implements IntegrationMessageReceiver
    Implements IntegrationMessageReceiver2

    Private Const msg_Not_Supported As String = "Message type '{0}' is not supported."
    Private Const GC_Every_X_Messages As Integer = 5

    Private Shared _callsSinceGC As Integer = 0


    Public Function ProcessMessage(ByVal efMsg As EFIntegrationMessage _
            ) As ResultInfo Implements IntegrationMessageReceiver.ProcessMessage

        Dim rInfo As ResultInfo = Nothing

        Select Case efMsg.Header.MessageType
            'Case CrmMessaging.MessageTypes.SalesTour

            'TODO: Call SalesTour Message Handler

            Case CRMMessage.MessageTypes.Leader
                Dim msgHandler As New BookingMessageHandler
                Dim bknMessage As CRMMessage.LSBookingMessage

                For Each bknMessage In efMsg.Body                    
                    rInfo = msgHandler.ProcessMessage(bknMessage, efMsg.Header.SourceSystem)
                Next

            Case CRMMessage.MessageTypes.Booking
                Dim msgHandler As New BookingMessageHandler
                Dim bknMessage As CRMMessage.LSBookingMessage

                bknMessage = DirectCast(efMsg.Body, CRMMessage.LSBookingMessage)
                rInfo = msgHandler.ProcessMessage(bknMessage, efMsg.Header.SourceSystem)

            Case Else
                Dim stepResult As New ResultStepInfo

                rInfo = New ResultInfo

                stepResult.Errors.Add(String.Format(msg_Not_Supported, efMsg.Header.MessageType))
                stepResult.Status = IntegrationResult.UnknownType
                rInfo.StepResults.Add(stepResult)

        End Select

        _callsSinceGC += 1

        If (_callsSinceGC >= GC_Every_X_Messages) Then
            _callsSinceGC = 0

            Debug.WriteLine(String.Format("forcing GC {0:T}.", Date.Now))
            GCH.ForceGarbageCollection()
        End If

        Return rInfo
    End Function

    Public Function ProcessMessage(ByVal message As String) As String _
            Implements IntegrationMessageReceiver2.ProcessMessage

        Dim rec As ElektraMessageReceivingSvc
        Dim resultStr As String = ""
        Dim msgObj As EFIntegrationMessage
        Dim rsInfo As New ResultStepInfo
        Dim rInfo As New ResultInfo

        Try
            rec = New ElektraMessageReceivingSvc

            msgObj = SerializationH.ObjectFormatter(Of EFIntegrationMessage).Deserialize(message)

            rInfo = rec.ProcessMessage(msgObj)

            resultStr = SerializationH.ObjectFormatter(Of ResultInfo).Serialize(rInfo)

        Catch ex As Exception
            rsInfo.Status = IntegrationResult.Failure
            rsInfo.Errors.Add(ex.Message)
            rInfo.StepResults.Add(rsInfo)
        End Try

        Return resultStr
    End Function

End Class

Imports IntCommon = EFTours.Framework.IntegrationFramework.Common

Public Class BookingMessageHandler
    Implements IMessageHandler(Of CRMMessage.LSBookingMessage)

    Public Function ProcessMessage(ByVal message As CRMMessage.LSBookingMessage, ByVal systemCode As String _
                        ) As IntCommon.ResultInfo Implements IMessageHandler(Of CRMMessage.LSBookingMessage).ProcessMessage
        Dim rInfo As IntCommon.ResultInfo
        Dim stepInfo As IntCommon.ResultStepInfo
        Dim bknAdapter As BookingAdapter


        rInfo = New IntCommon.ResultInfo

        bknAdapter = New BookingAdapter

        stepInfo = bknAdapter.Update(message.Booking, systemCode)
        rInfo.StepResults.Add(stepInfo)

        Return rInfo
    End Function
End Class


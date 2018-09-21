Public Class CourseBookingData
    Public Property CourseBooking As Production.CourseBooking
    Public Property modifiedCourseBookingsInDestination As New Hashtable
    Public Property visaT4Destinations As New List(Of String)
    Public Property IsAutoConfCourse As Boolean
    Public Property WhatChnaged As New DataTable
    Public Property HasCapacityAvailable As Boolean

End Class


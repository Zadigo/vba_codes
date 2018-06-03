'
' Use this sub to create appointments in Outlook programatically
' John PM (2017)
'
'
Sub CreateAppointment()
    Dim olAppt As AppointmentItem
    Set olAppt = Application.CreateItem(olAppointmentItem)
    
    With olAppt
        .Subject = "My Subject"
        .Body = "This is the body"
        .RequiredAttendees = "something@gmail.com"
        .Location = "Lille"
        .ReminderMinutesBeforeStart = "30"
        .Start = #11/19/2017 2:00:00 AM#
        .End = #11/19/2017 4:00:00 AM#
        '.BillingInformation = "something"
        .Categories = "Business"
        .Display
    End With
End Sub
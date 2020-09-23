Attribute VB_Name = "mMain"

Public cLog     As clsLogging

Public Sub Main()
frmMain.Show
Set_Interceptor
End Sub

Public Sub Set_Interceptor()

    Set cLog = New clsLogging
    With cLog
        '** just a little of my sick humor here, I'm Canadian, eh..
        '//app log title
        .AppTitle = "www.twulaisietwuvowt.com - Fine Software"
        '//app log footer
        .AppFtr = "Thank You for choosing !LamoWare!"
        '//app log message
        .AppMsg = "Teach your dog to read with !LamoWare! - Ruff! Ruff!"
        '//app log copyright
        .AppCpr = "All Rights Reserved© TLTV - Entertain the Brutes!"
        '//output wmi computer info
        .DataCmp = True
        '//output wmi os info
        .DataOS = True
        '//output wmi software info
        .DataSft = True
        '//output wmi process info
        .DataPrc = True
        '//output service data - turned off because may
        '//raise firewall alert
        .DataSrv = False
        '//mail error log To:
        .MailTo = "bitbucket@lamoware.com"
        '//mail error log Subject:
        .MailSubject = "Custom Error Report"
        '//dump error context to log
        .EDump = True
        '//raise user error message on gpf
        .EResume = False
        '//set properties first - then start the logging engine
        '//enable error handler
        .EHandler = True
        '//start logging now
        .OnStart = True
        '//start the logging engine
        .Log_Start
        
        '//application logging subheader
        .Log_SubHeader "Application Events"
        '//example of application event logging
        '//can be used for any event you want to log
        .Log_Event "Main Finished Processing", _
        "Example test completed - ok", _
        "log status - check", _
        CStr(Now)
        
        '//web log style example
        .WebImage = App.Path & "\title.gif"
        .CssStyle = True
        .CssBackColor = "ffffff"
        .CssForeColor = "555555"
        .CssFontSize = 9
        .CssFont = "8pt arial"
        .CssHeadColor = "333333"
        .CssHeadSize = 14
        .CssHeader = True
    End With
    
End Sub

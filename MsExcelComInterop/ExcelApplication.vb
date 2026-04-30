''' <summary>
''' A wrapper for an Excel application instance created with Interaction.CreateObject without requiring MS Office type libraries.
''' </summary>
Public Class ExcelApplication
    Inherits ComRootObject(Of Object)

    ''' <summary>
    ''' Creates a new Microsoft Excel application instance.
    ''' </summary>
    Public Sub New()
#Disable Warning CA1416
        MyBase.New(CreateObject("Excel.Application"), Sub(instance) If instance IsNot Nothing Then instance.InvokeMethod("Quit"))
#Enable Warning CA1416
        Me.Workbooks = New ExcelWorkbooksCollection(Me)
    End Sub

    ''' <summary>
    ''' Gets the workbook collection for this Excel application.
    ''' </summary>
    Public ReadOnly Property Workbooks As ExcelWorkbooksCollection

    ''' <summary>
    ''' Gets or sets whether the Excel application is controlled by the user.
    ''' </summary>
    Public Property UserControl As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("UserControl")
        End Get
        Set(value As Boolean)
            InvokePropertySet("UserControl", value)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets whether Excel displays alerts.
    ''' </summary>
    Public Property DisplayAlerts As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("DisplayAlerts")
        End Get
        Set(value As Boolean)
            InvokePropertySet("DisplayAlerts", value)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets whether the Excel application window is visible.
    ''' </summary>
    Public Property Visible As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("Visible")
        End Get
        Set(value As Boolean)
            InvokePropertySet("Visible", value)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets whether Excel accepts user interaction.
    ''' </summary>
    Public Property Interactive As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("Interactive")
        End Get
        Set(value As Boolean)
            InvokePropertySet("Interactive", value)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets whether Excel updates the screen while operations run.
    ''' </summary>
    Public Property ScreenUpdating As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("ScreenUpdating")
        End Get
        Set(value As Boolean)
            InvokePropertySet("ScreenUpdating", value)
        End Set
    End Property

#Disable Warning CA1712 ' Enumerationswerten nicht den Typnamen voranstellen
    ''' <summary>
    ''' https://learn.microsoft.com/de-de/office/vba/api/office.msoautomationsecurity.
    ''' </summary>
    Public Enum MsoAutomationSecurity As Integer
        ''' <summary>
        ''' Uses the security setting specified in the Security dialog box.
        ''' </summary>
        msoAutomationSecurityByUI = 2
        ''' <summary>
        ''' Disables all macros in all files opened programmatically without showing any security alerts.
        ''' </summary>
        msoAutomationSecurityForceDisable = 3
        ''' <summary>
        ''' Enables all macros. This is the default value when the application is started.
        ''' </summary>
        msoAutomationSecurityLow = 1
    End Enum
#Enable Warning CA1712 ' Enumerationswerten nicht den Typnamen voranstellen

    ''' <summary>
    ''' Configure security level for macros/VBA.
    ''' </summary>
    ''' <returns></returns>
    Public Property AutomationSecurity As MsoAutomationSecurity
        Get
            Return CType(InvokePropertyGet(Of Integer)("AutomationSecurity"), MsoAutomationSecurity)
        End Get
        Set(value As MsoAutomationSecurity)
            InvokePropertySet("AutomationSecurity", CType(value, Integer))
        End Set
    End Property

    ''' <summary>
    ''' Gets a built-in Excel dialog.
    ''' </summary>
    ''' <param name="type">Built-in dialog type.</param>
    ''' <returns>A wrapper for the requested Excel dialog.</returns>
    Public Function Dialogs(type As Enumerations.XlBuiltInDialog) As ExcelDialog
        Return New ExcelDialog(Me, InvokePropertyGet(Of Object)("Dialogs", CType(type, Integer)))
    End Function

    ''' <summary>
    ''' Runs a VBA method.
    ''' </summary>
    ''' <param name="vbaMethodNameInclWorkbookName">VBA method name, optionally including the workbook name.</param>
    ''' <returns>The value returned by Excel.</returns>
    Public Function Run(vbaMethodNameInclWorkbookName As String) As Object
        Return InvokeFunction(Of Object)("Run", New Object() {vbaMethodNameInclWorkbookName})
    End Function

    ''' <summary>
    ''' Runs a VBA method in a workbook.
    ''' </summary>
    ''' <param name="workbookName">Workbook name containing the VBA method.</param>
    ''' <param name="vbaMethod">VBA method name to run.</param>
    ''' <returns>The value returned by Excel.</returns>
    Public Function Run(workbookName As String, vbaMethod As String) As Object
        Return InvokeFunction(Of Object)("Run", New Object() {"'" & workbookName & "'!" & vbaMethod})
    End Function

    Private AdditionalDisposeChildrenList As New List(Of ComObjectBase)

    ''' <inheritdoc/>
    Protected Overrides Sub OnDisposeChildren()
        If Me.Workbooks IsNot Nothing Then Me.Workbooks.Dispose()
    End Sub

    ''' <inheritdoc/>
    Protected Overrides Sub onBeforeClosing()
        UserControl = True
    End Sub

End Class

Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl5 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl6 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Gallery1 = Me.Factory.CreateRibbonGallery
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "Eki Tools"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Gallery1)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Gallery1
        '
        Me.Gallery1.ColumnCount = 3
        Me.Gallery1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Gallery1.Image = Global.Ekimetrics_Powerpoint_AddIn.My.Resources.Resources.EkiLogo1
        Me.Gallery1.ItemImageSize = New System.Drawing.Size(50, 50)
        RibbonDropDownItemImpl1.Image = Global.Ekimetrics_Powerpoint_AddIn.My.Resources.Resources.Artificial_Intelligence11
        RibbonDropDownItemImpl1.Label = "Artificial Intelligence1.emf"
        RibbonDropDownItemImpl2.Image = Global.Ekimetrics_Powerpoint_AddIn.My.Resources.Resources.Calendar_Checked11
        RibbonDropDownItemImpl2.Label = "Calendar Checked1.emf"
        RibbonDropDownItemImpl3.Image = Global.Ekimetrics_Powerpoint_AddIn.My.Resources.Resources.Cash
        RibbonDropDownItemImpl3.Label = "Cash.emf"
        RibbonDropDownItemImpl4.Image = Global.Ekimetrics_Powerpoint_AddIn.My.Resources.Resources.Coins
        RibbonDropDownItemImpl4.Label = "Coins2.emf"
        RibbonDropDownItemImpl5.Image = Global.Ekimetrics_Powerpoint_AddIn.My.Resources.Resources.Forecast1
        RibbonDropDownItemImpl5.Label = "Forescast1.emf"
        RibbonDropDownItemImpl6.Image = Global.Ekimetrics_Powerpoint_AddIn.My.Resources.Resources.Cycle1
        RibbonDropDownItemImpl6.Label = "Cycle1.emf"
        Me.Gallery1.Items.Add(RibbonDropDownItemImpl1)
        Me.Gallery1.Items.Add(RibbonDropDownItemImpl2)
        Me.Gallery1.Items.Add(RibbonDropDownItemImpl3)
        Me.Gallery1.Items.Add(RibbonDropDownItemImpl4)
        Me.Gallery1.Items.Add(RibbonDropDownItemImpl5)
        Me.Gallery1.Items.Add(RibbonDropDownItemImpl6)
        Me.Gallery1.Label = "Gallery1"
        Me.Gallery1.Name = "Gallery1"
        Me.Gallery1.ShowImage = True
        Me.Gallery1.ShowItemLabel = False
        Me.Gallery1.ShowItemSelection = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Public WithEvents Gallery1 As Microsoft.Office.Tools.Ribbon.RibbonGallery
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class

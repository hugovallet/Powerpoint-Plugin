Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core.MsoTriState

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Gallery1_Click(sender As Object, e As RibbonControlEventArgs) _
        Handles Gallery1.Click

        'Get application path
        Dim path As String
        path = "C:\Users\hvallet\Desktop\Ekimetrics Powerpoint AddIn\Ekimetrics Powerpoint AddIn\Resources\"

        'Get current slide index
        Dim pptLayout As CustomLayout
        Dim active_presentation As Presentation
        active_presentation = Globals.ThisAddIn.Application.ActivePresentation
        pptLayout = active_presentation.Slides(1).CustomLayout

        Dim active_slide_idx As Integer
        active_slide_idx = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideNumber

        ' Display the id of the object clicked number.
        MsgBox("Ribbon clicked: " & e.Control.Id, vbInformation)

        ' Display the id of the item clicked.
        Dim item As RibbonDropDownItem
        item = Gallery1.SelectedItem()
        MsgBox("Item clicked: " & item.Label, vbInformation)

        ' Get the current active slide
        Dim active_slide As Slide
        active_slide = active_presentation.Slides(active_slide_idx)
        MsgBox(path & item.Label, vbInformation)
        active_slide.Shapes.AddPicture(path & item.Label, msoTrue, msoTrue, 50, 50, 50, 50)




    End Sub



End Class

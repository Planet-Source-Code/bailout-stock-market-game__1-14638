Attribute VB_Name = "ModPrint"
Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
Const vbHiMetric As Integer = 8
Dim PicRatio As Double
Dim PrnWidth As Double
Dim PrnHeight As Double
Dim PrnRatio As Double
Dim PrnPicWidth As Double
Dim PrnPicHeight As Double

    ' Determine if picture should be printed in landscape or portrait
    ' and set the orientation
    If Pic.Height >= Pic.Width Then
        Prn.Orientation = vbPRORPortrait   ' Taller than wide
    Else
        Prn.Orientation = vbPRORLandscape  ' Wider than tall
    End If

    ' Calculate device independent Width to Height ratio for picture
        PicRatio = Pic.Width / Pic.Height

    ' Calculate the dimentions of the printable area in HiMetric
        PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
        PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    ' Calculate device independent Width to Height ratio for printer
        PrnRatio = PrnWidth / PrnHeight

    ' Scale the output to the printable area
    If PicRatio >= PrnRatio Then
    ' Scale picture to fit full width of printable area
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
        Prn.ScaleMode)
    Else
    ' Scale picture to fit full height of printable area
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
        Prn.ScaleMode)
    End If

    ' Print the picture using the PaintPicture method
        Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub


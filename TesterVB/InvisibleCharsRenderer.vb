Namespace TesterVB
    Public Class InvisibleCharsRenderer
        Inherits Style

        Private ReadOnly pen As Pen

        Public Sub New(pen As Pen)
            Me.pen = pen
        End Sub

        Public Overrides Sub Draw(gr As Graphics, position As Point, range As FastColoredTextBoxNS.Range)
            Dim tb As FastColoredTextBox = range.tb
            Using brush As Brush = New SolidBrush(Me.pen.Color)
                For Each place As Place In range
                    Dim c As Char = tb(place).c
                    If c <> " " Then
                        GoTo IL_BC
                    End If
                    Dim point As Point = tb.PlaceToPoint(place)
                    point.Offset(tb.CharWidth / 2, tb.CharHeight / 2)
                    gr.DrawLine(Me.pen, point.X, point.Y, point.X + 1, point.Y)
                    If tb(place.iLine).Count - 1 = place.iChar Then
                        GoTo IL_BC
                    End If
                    Continue For
IL_BC:
                    If tb(place.iLine).Count - 1 = place.iChar Then
                        point = tb.PlaceToPoint(place)
                        point.Offset(tb.CharWidth, 0)
                        gr.DrawString("¶", tb.Font, brush, point)
                    End If
                Next
            End Using
        End Sub
    End Class
End Namespace

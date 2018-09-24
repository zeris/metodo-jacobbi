Public Class Form1

    Dim a1n(3), a2n(3), a3n(3), b(3), n, ec As Single
    Dim f1, f2, f3 As Single
    Dim x(50), y(50), z(50), erx(50), ery(50), erz(50) As Single
    Dim i, redon As Integer
    Dim cadena As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        a11.Text = ""
        a12.Text = ""
        a13.Text = ""
        a21.Text = ""
        a22.Text = ""
        a23.Text = ""
        a31.Text = ""
        a32.Text = ""
        a33.Text = ""
        b1.Text = ""
        b2.Text = ""
        b3.Text = ""
        tcs.Text = ""
        Respuesta.Text = ""
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        End
    End Sub

    Private Sub btnCalcular_Click(sender As Object, e As EventArgs) Handles btnCalcular.Click
        salida.Rows.Clear()
        n = tcs.Text
        redon = n + 2
        ec = 0.5 * 10 ^ (-n)
        If Math.Abs(Convert.ToInt32(a13.Text) + Convert.ToInt32(a12.Text)) < Math.Abs(Convert.ToInt32(a11.Text)) Then
            a1n(0) = a11.Text
            a1n(1) = a12.Text
            a1n(2) = a13.Text
            b(0) = b1.Text
            If Math.Abs(Convert.ToInt32(a21.Text) + Convert.ToInt32(a23.Text)) < Math.Abs(Convert.ToInt32(a22.Text)) Then
                a2n(0) = a21.Text
                a2n(1) = a22.Text
                a2n(2) = a23.Text
                b(1) = b2.Text
                a3n(0) = a31.Text
                a3n(1) = a32.Text
                b(2) = b3.Text
                a3n(2) = a33.Text
            Else
                a2n(0) = a31.Text
                a2n(1) = a32.Text
                a2n(2) = a33.Text
                b(1) = b3.Text
                a3n(0) = a21.Text
                a3n(1) = a22.Text
                a3n(2) = a23.Text
                b(2) = b2.Text
            End If
        ElseIf Math.Abs(Convert.ToInt32(a22.Text) + Convert.ToInt32(a23.Text)) < Math.Abs(Convert.ToInt32(a21.Text)) Then
            n = tcs.Text
            a1n(0) = a21.Text
            a1n(1) = a22.Text
            a1n(2) = a23.Text
            b(0) = b2.Text
            redon = n + 2
            ec = 0.5 * 10 ^ (-n)
            If Math.Abs(Convert.ToInt32(a31.Text) + Convert.ToInt32(a33.Text)) < Math.Abs(Convert.ToInt32(a32.Text)) Then
                a2n(0) = a31.Text
                a2n(1) = a32.Text
                a2n(2) = a33.Text
                b(1) = b3.Text
                a3n(0) = a11.Text
                a3n(1) = a12.Text
                a3n(2) = a13.Text
                b(2) = b1.Text
            Else
                a2n(0) = a11.Text
                a2n(1) = a12.Text
                a2n(2) = a13.Text
                b(1) = b1.Text
                a3n(0) = a31.Text
                a3n(1) = a32.Text
                a3n(2) = a33.Text
                b(2) = b3.Text
            End If
        Else
            a1n(0) = a31.Text
            a1n(1) = a32.Text
            a1n(2) = a33.Text
            b(0) = b3.Text
            If Math.Abs(Convert.ToInt32(a11.Text) + Convert.ToInt32(a13.Text)) < Math.Abs(Convert.ToInt32(a12.Text)) Then
                a2n(0) = a11.Text
                a2n(1) = a12.Text
                a2n(2) = a13.Text
                b(1) = b1.Text
                a3n(0) = a21.Text
                a3n(1) = a22.Text
                a3n(2) = a23.Text
                b(2) = b2.Text
            Else
                a2n(0) = a21.Text
                a2n(1) = a22.Text
                a2n(2) = a23.Text
                b(1) = b2.Text
                a3n(0) = a11.Text
                a3n(1) = a12.Text
                a3n(2) = a13.Text
                b(2) = b1.Text
            End If
        End If
        i = 0
        erx(0) = 1
        ery(0) = 1
        erz(0) = 1
        x(0) = 0
        y(0) = 0
        z(0) = 0
        salida.Rows.Add(i, Math.Round(x(i), redon), Math.Round(y(i), redon), Math.Round(z(i), redon), "-", "-", "-")
        Do While (erx(i) > ec)
            x(i + 1) = (b(0) - (a1n(1) * y(i)) - (a1n(2) * z(i))) / a1n(0)
            y(i + 1) = (b(1) - (a2n(0) * x(i)) - (a2n(2) * z(i))) / a2n(1)
            z(i + 1) = (b(2) - (a3n(0) * x(i)) - (a3n(1) * y(i))) / a3n(2)
            erx(i + 1) = Math.Abs((x(i + 1) - x(i)) / x(i + 1))
            ery(i + 1) = Math.Abs((y(i + 1) - y(i)) / y(i + 1))
            erz(i + 1) = Math.Abs((z(i + 1) - z(i)) / z(i + 1))
            salida.Rows.Add(i + 1, Math.Round(x(i + 1), redon), Math.Round(y(i + 1), redon), Math.Round(z(i + 1), redon), Math.Round(erx(i + 1), redon), Math.Round(ery(i + 1), redon), Math.Round(erz(i + 1), redon))
            i = i + 1
        Loop
        Respuesta.Text = "Los valores aproximados de las raices son: X = " + Math.Round(x(i), redon).ToString() + " Y = " + Math.Round(y(i), redon).ToString() + " Z = " + Math.Round(z(i), redon).ToString()

    End Sub
End Class

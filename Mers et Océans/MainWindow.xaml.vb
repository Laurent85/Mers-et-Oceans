Imports System.IO

Class MainWindow

    Public Temp As Integer
    Public Score As Integer
    Public Coups1 As Integer
    Public P As Integer = 0
    Public Triche As Integer = 0

    Sub active_drag(ByVal cib As Label)
        AddHandler cib.DragEnter, AddressOf lbl_Dragenter
        AddHandler cib.Drop, AddressOf lbl_Drop
        AddHandler cib.MouseDown, AddressOf lbl_MouseDown
    End Sub

    Sub interdit_drag(ByVal cib As Label)
        RemoveHandler cib.DragEnter, AddressOf lbl_Dragenter
        RemoveHandler cib.Drop, AddressOf lbl_Drop
        RemoveHandler cib.MouseDown, AddressOf lbl_MouseDown
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        For i As Int32 = 0 To Numéros.Children.Count - 1
            Dim numero As Label = CType(Numéros.Children.Item(i), Label)
            active_drag(numero)
        Next

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            active_drag(cible)
            AddHandler cible.MouseRightButtonDown, AddressOf lbl_clear
            cible.AllowDrop = True
            cible.Foreground = Brushes.White
        Next

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim numéro As Label = CType(Numéros.Children.Item(i), Label)
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            source.Foreground = Brushes.Black
            AddHandler numéro.MouseEnter, AddressOf Souris
            AddHandler numéro.MouseLeave, AddressOf plus_souris
        Next

        Bilan_des_placements()
        Points.Foreground = Brushes.SaddleBrown
        Points.HorizontalContentAlignment = HorizontalAlignment.Center
        Points.VerticalContentAlignment = VerticalAlignment.Top
        Points.Content = score
        Coups.Content = "Coups : " & coups1
        Carte.IsChecked = True
        zoom()

    End Sub

    Private Sub lbl_MouseDown(sender As Object, e As MouseButtonEventArgs)
        'Activation du "glisser"
        Dim lbl As Label = CType(sender, Label)
        DragDrop.DoDragDrop(lbl, lbl.Content, DragDropEffects.Copy)
    End Sub

    Private Sub lbl_Dragenter(sender As Object, e As DragEventArgs)
        'Activation de la copie de l'élément glissé
        If e.Data.GetDataPresent(DataFormats.Text) Then
            e.Effects = DragDropEffects.Copy
        Else
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub lbl_Drop(sender As Object, e As DragEventArgs)

        For j = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(j), Label)
            Dim cercle As Ellipse = CType(Cercles.Children.Item(j), Ellipse)
            Dim couleurGris As New SolidColorBrush()
            couleurGris.Color = Color.FromArgb(255, 137, 140, 143)
            Dim couleurBleu As New SolidColorBrush()
            couleurBleu.Color = Color.FromArgb(255, 0, 116, 255)
            'Cherche si existe déjà
            'Si cible = source et source différent de "?" alors..
            If cible.Content = e.Data.GetData(DataFormats.Text) And e.Data.GetData(DataFormats.Text) <> "?" Then
                cible.Content = "?"
                cible.ToolTip = vbNullString
                cible.Foreground = Brushes.White
                CType(sender, Label).Content = e.Data.GetData(DataFormats.Text)
                CType(sender, Label).Foreground = Brushes.White
                cercle.Fill = couleurGris
            Else
                CType(sender, Label).Content = e.Data.GetData(DataFormats.Text)
                CType(sender, Label).Foreground = Brushes.White
            End If
            If cible.Content = "?" Then
                cercle.Fill = couleurBleu
            End If
            If cible.Content <> "?" And Not Equals(cercle.Fill, Brushes.LawnGreen) And Not Equals(cercle.Fill, Brushes.Red) Then
                cercle.Fill = couleurGris
            End If
        Next j

        'Numéros gris et infobulle pour les départements placés
        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim numero As Label = CType(Numéros.Children.Item(i), Label)
            Dim couleurGris As New SolidColorBrush()
            couleurGris.Color = Color.FromArgb(255, 137, 140, 143)
            For j As Int32 = 0 To Cibles.Children.Count - 1
                Dim cible As Label = CType(Cibles.Children.Item(j), Label)
                If cible.Content = numero.Content Then
                    numero.Foreground = Brushes.LightGray
                    cible.ToolTip = CType(Sources.Children.Item(i), Label).Content
                    Exit For
                Else
                    numero.Foreground = Brushes.Red
                End If
                If cible.Content = "?" Then
                    cible.ToolTip = vbNullString
                    cible.Foreground = Brushes.White
                End If
            Next j
        Next i

        p = 0
        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If cible.Content <> "?" And (Equals(cible.Foreground, Brushes.White)) Then
                p = p + 1
                En_cours.Foreground = Brushes.Gray
                En_cours.Content = "Eléments non vérifiés : " & p.ToString
            End If
        Next

        Dim m As Integer = 60
        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If cible.Content <> "?" And (Equals(cible.Foreground, Brushes.White)) Then
                m = m - 1
                'En_cours.Foreground = Brushes.Gray
                Absent.Content = "Eléments non placés : " & m.ToString
            End If
        Next

        Test_réponse()

    End Sub

    Sub Zoom()

        If Europe.IsChecked = True Then
            ZoomSlider.Value = 4
            MyScrollViewer.ScrollToHorizontalOffset(1400)
            MyScrollViewer.ScrollToVerticalOffset(200)

        ElseIf Afrique.IsChecked = True Then
            ZoomSlider.Value = 2
            MyScrollViewer.ScrollToHorizontalOffset(500)
            MyScrollViewer.ScrollToVerticalOffset(300)

        ElseIf Carte.IsChecked = True Then
            ZoomSlider.Value = 1
            MyScrollViewer.ScrollToHorizontalOffset(0)
            MyScrollViewer.ScrollToVerticalOffset(0)

        ElseIf Amerique_Nord.IsChecked = True Then
            ZoomSlider.Value = 2
            MyScrollViewer.ScrollToHorizontalOffset(0)
            MyScrollViewer.ScrollToVerticalOffset(0)

        ElseIf Océanie.IsChecked = True Then
            ZoomSlider.Value = 2
            MyScrollViewer.ScrollToHorizontalOffset(1000)
            MyScrollViewer.ScrollToVerticalOffset(370)

        ElseIf Asie.IsChecked = True Then
            ZoomSlider.Value = 2
            MyScrollViewer.ScrollToHorizontalOffset(1000)
            MyScrollViewer.ScrollToVerticalOffset(0)
        End If

    End Sub

    Private Sub Europe_Checked(sender As Object, e As RoutedEventArgs) Handles Europe.Checked
        zoom()
    End Sub

    Private Sub Afrique_Checked(sender As Object, e As RoutedEventArgs) Handles Afrique.Checked
        zoom()
    End Sub

    Private Sub Carte_Checked(sender As Object, e As RoutedEventArgs) Handles Carte.Checked
        zoom()
    End Sub

    Private Sub Amerique_Nord_Checked(sender As Object, e As RoutedEventArgs) Handles Amerique_Nord.Checked
        zoom()
    End Sub

    Private Sub Océanie_checked(sender As Object, e As RoutedEventArgs) Handles Océanie.Checked
        zoom()
    End Sub

    Private Sub Asie_Checked(sender As Object, e As RoutedEventArgs) Handles Asie.Checked
        zoom()
    End Sub

    Private Sub lbl_clear(sender As Object, e As MouseButtonEventArgs) Handles Océan_Pacifique1.MouseRightButtonDown, Océan_Atlantique.MouseRightButtonDown

        'effacement par clic-droit
        CType(sender, Label).AllowDrop = True
        Dim couleurBleu As New SolidColorBrush()
        couleurBleu.Color = Color.FromArgb(255, 0, 116, 255)

        Dim b As Integer = 0
        Dim m As Integer = -1
        Dim n As Integer = 0
        Dim j As String
        'Dim total As Integer
        p = 0
        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If cible.Content <> "?" And (Equals(cible.Foreground, Brushes.White)) Then
                p = p + 1
                If p = 1 Then
                    p = 0
                End If
                En_cours.Foreground = Brushes.Gray
                En_cours.Content = "Eléments non vérifiés : " & p.ToString
            End If
        Next

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim cercle As Ellipse = CType(Cercles.Children.Item(i), Ellipse)
            If i < 9 Then
                j = "0" & i + 1.ToString
            Else
                j = i + 1.ToString
            End If

            If cible.Content <> "?" And cible.Content <> j.ToString Then
                cible.Foreground = Brushes.White
                cercle.Fill = Brushes.Red
                m = m + 1
                Mal.Foreground = Brushes.Red
                Mal.Content = "Eléments mal placés : " & m.ToString
            End If
        Next i

        For i As Int32 = 0 To Numéros.Children.Count - 1
            Dim source As Label = CType(Numéros.Children.Item(i), Label)
            If CType(sender, Label).Content = source.Content Then
                source.Foreground = Brushes.Red
                AddHandler source.MouseDown, AddressOf lbl_MouseDown 'activation du glissé
            End If
        Next i

        CType(sender, Label).Content = "?"
        CType(sender, Label).Foreground = Brushes.White
        CType(sender, Label).ToolTip = vbNullString

        For i As Int32 = 0 To Numéros.Children.Count - 1
            Dim cercle As Ellipse = CType(Cercles.Children.Item(i), Ellipse)
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            If cible.Content = "?" Then
                cercle.Fill = couleurBleu
            End If
        Next i

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim cercle As Ellipse = CType(Cercles.Children.Item(i), Ellipse)
            If i < 9 Then
                j = "0" & i + 1.ToString
            Else
                j = i + 1.ToString
            End If

            If cible.Content = j.ToString Then
                cercle.Fill = Brushes.LawnGreen
                cible.Foreground = Brushes.Black
                b = b + 1
                Bien.Foreground = Brushes.Green
                Bien.Content = "Eléments bien placés : " & b.ToString
            End If
            If cible.Content = "?" Then
                n = n + 1
                Absent.Foreground = Brushes.Blue
                Absent.Content = "Eléments non placés : " & n.ToString
            End If
        Next i

    End Sub

    Private Sub Vérifier_Click(sender As Object, e As RoutedEventArgs) Handles vérifier.Click

        Dim b As Integer = 0
        Dim m As Integer = 0
        Dim n As Integer = 0
        Dim j As String
        Dim total As Integer

        Bilan_des_placements()

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim cercle As Ellipse = CType(Cercles.Children.Item(i), Ellipse)
            If i < 9 Then
                j = "0" & i + 1.ToString
            Else
                j = i + 1.ToString
            End If

            If cible.Content = j.ToString Then
                cercle.Fill = Brushes.LawnGreen
                cible.Foreground = Brushes.Black
                b = b + 1
                Bien.Foreground = Brushes.Green
                Bien.Content = "Eléments bien placés : " & b.ToString
            End If
            If cible.Content = "?" Then
                Dim couleurBleu As New SolidColorBrush()
                couleurBleu.Color = Color.FromArgb(255, 0, 116, 255)
                cible.Foreground = Brushes.White
                cercle.Fill = couleurBleu
                n = n + 1
                Absent.Foreground = Brushes.Blue
                Absent.Content = "Eléments non placés : " & n.ToString
            End If
            If cible.Content <> "?" And cible.Content <> j.ToString Then
                cible.Foreground = Brushes.White
                cercle.Fill = Brushes.Red
                m = m + 1
                Mal.Foreground = Brushes.Red
                Mal.Content = "Eléments mal placés : " & m.ToString
            End If
        Next i

        Test_réponse()

        If Temp <> b Then
            total = (b - Temp) * 2
        Else
            total = 0
        End If
        Score = Score + Val(total) - m
        Temp = b
        Coups1 = Coups1 + 1
        Coups.Content = "Coups : " & Coups1

        Points.Foreground = Brushes.SaddleBrown
        Points.HorizontalContentAlignment = HorizontalAlignment.Center
        Points.VerticalContentAlignment = VerticalAlignment.Top
        Points.Content = Score
        'My.Computer.Audio.Play(My.Resources.Bruit, AudioPlayMode.Background)

        Test_réponse()

        P = 0
        En_cours.Content = "Eléments non vérifiés : " & P.ToString
        'Cacher_Checked(sender, e)
        'Voir_Checked(sender, e)

        If b = 60 And Triche = 0 Then
            Dim result As MessageBoxResult = MessageBox.Show("Bravo " & copie.Content & " !" & Chr(10) & "Tu as placé les 60 Eléments en " & Coups1 & " coups." & Chr(10) & "Tu as totalisé " & Score & " points." & Chr(10) & "Veux-tu refaire une partie ?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.Yes Then
                fichier_Scores()
                Effacer_Click(sender, e)
            Else
                fichier_Scores()
                Close()
            End If
        End If
    End Sub

    Sub Bilan_des_placements()
        Bien.Foreground = Brushes.Green
        Mal.Foreground = Brushes.Red
        Absent.Foreground = Brushes.Blue
        Bien.Content = "Eléments bien placés : 0"
        Mal.Content = "Eléments mal placés : 0"
        Absent.Content = "Eléments non placés : 0"
        En_cours.Content = "Eléments non vérifiés : 0"
    End Sub

    Sub fichier_Scores()
        If triche = 0 Then
            Try
                Dim fichierScores As StreamWriter = New StreamWriter("\\laurent\d\Mers_Oceans.txt", True)
                fichierScores.WriteLine(Now & Chr(9) & copie.Content & Chr(9) & "Points : " & Points.Content & Chr(9) & Coups.Content & Chr(9) & Bien.Content & Chr(9) & Mal.Content & Chr(9) & Absent.Content)
                fichierScores.Close()

            Catch ex As Exception

            End Try

            Try
                Dim fichierScores1 As StreamWriter = New StreamWriter("c:\Mers_Oceans.txt", True)
                fichierScores1.WriteLine(Now & Chr(9) & copie.Content & Chr(9) & "Points : " & Points.Content & Chr(9) & Coups.Content & Chr(9) & Bien.Content & Chr(9) & Mal.Content & Chr(9) & Absent.Content)
                fichierScores1.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Solutions_Click(sender As Object, e As RoutedEventArgs) Handles Solutions.Click

        Dim b As Integer = 0
        Dim m As Integer = 0
        Dim n As Integer = 0
        Dim j As String
        triche = 1

        Bilan_des_placements()

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim cercle As Ellipse = CType(Cercles.Children.Item(i), Ellipse)
            If i < 9 Then
                j = "0" & i + 1.ToString
                cible.Content = j.ToString
                cible.ToolTip = CType(Sources.Children.Item(i), Label).Content
            Else
                j = i + 1.ToString
                cible.Content = j.ToString
                cible.ToolTip = CType(Sources.Children.Item(i), Label).Content
            End If

            If cible.Content = j.ToString Then
                cercle.Fill = Brushes.LawnGreen
                cible.Foreground = Brushes.Black
                b = b + 1
                Bien.Foreground = Brushes.Green
                Bien.Content = "Eléments bien placés : " & b.ToString
            End If
            If cible.Content = "?" Then
                cible.Foreground = Brushes.Blue
                n = n + 1
                Absent.Foreground = Brushes.Blue
                Absent.Content = "Eléments non placés : " & n.ToString
            End If
            If cible.Content <> "?" And cible.Content <> j.ToString Then
                cible.Foreground = Brushes.Red
                m = m + 1
                Mal.Foreground = Brushes.Red
                Mal.Content = "Eléments mal placés : " & m.ToString
            End If
        Next i

        For i As Integer = 0 To Cibles.Children.Count - 1
            Dim cible As Label = Cibles.Children.Item(i)
            If cible IsNot Nothing Then
                cible.AllowDrop = False
            End If
        Next

        fichier_Scores()
        Test_réponse()

    End Sub

    Private Sub Effacer_Click(sender As Object, e As RoutedEventArgs) Handles Effacer.Click

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            Dim numéro As Label = CType(Numéros.Children.Item(i), Label)
            Dim cercle As Ellipse = CType(Cercles.Children.Item(i), Ellipse)
            Dim couleurBleu As New SolidColorBrush()
            couleurBleu.Color = Color.FromArgb(255, 0, 116, 255)
            cible.Content = "?"
            cible.AllowDrop = True
            cible.Foreground = Brushes.White
            cible.ToolTip = vbNullString
            source.Foreground = Brushes.Black
            numéro.Foreground = Brushes.Red
            cercle.Fill = couleurBleu
            AddHandler source.MouseDown, AddressOf lbl_MouseDown
            AddHandler cible.MouseRightButtonDown, AddressOf lbl_clear
            active_drag(numéro)
            active_drag(cible)
            Bilan_des_placements()
        Next i

        fichier_Scores()
        'Window_Loaded(sender, e)

        temp = 0
        score = 0
        coups1 = 0
        Coups.Content = "Coups : " & coups1
        Points.Content = score

    End Sub

    Sub Test_réponse()
        'Si bonne réponse, on met la source en vert et on bloque le déplacement du département
        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim numéro As Label = CType(Numéros.Children.Item(i), Label)
            Dim cible As Label = CType(Cibles.Children.Item(i), Label)
            Dim source As Label = CType(Sources.Children.Item(i), Label)

            If (Equals(cible.Foreground, Brushes.Black)) Then
                numéro.Foreground = Brushes.Green
                source.Foreground = Brushes.Green
                RemoveHandler numéro.MouseDown, AddressOf lbl_MouseDown
                interdit_drag(cible)
                RemoveHandler cible.MouseRightButtonDown, AddressOf lbl_clear
                cible.AllowDrop = False
                cible.ToolTip = CType(Sources.Children.Item(i), Label).Content
            End If
        Next i
    End Sub

    Private Sub Window1_Closing(ByVal sender As Object, ByVal e As ComponentModel.CancelEventArgs) Handles Me.Closing
        Dim result As MessageBoxResult = MessageBox.Show("Voulez-vous quitter le jeu ?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Question)

        If result = MessageBoxResult.Yes Then
            fichier_Scores()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub Souris(sender As Object, e As MouseEventArgs)

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            Dim numéro As Label = CType(Numéros.Children.Item(i), Label)
            If numéro.IsMouseOver = True And numéro.Foreground.ToString = "#FFFF0000" Then
                source.Foreground = Brushes.Red
                source.FontSize = 12
            End If
            If numéro.IsMouseOver = True And (Equals(numéro.Foreground, Brushes.Red)) Then
                source.Foreground = Brushes.Red
                source.FontSize = 12
            End If
        Next

    End Sub

    Private Sub plus_souris(sender As Object, e As MouseEventArgs)

        For i As Int32 = 0 To Cibles.Children.Count - 1
            Dim source As Label = CType(Sources.Children.Item(i), Label)
            Dim numéro As Label = CType(Numéros.Children.Item(i), Label)
            If numéro.IsMouseOver = False And (Equals(numéro.Foreground, Brushes.Red)) Then
                source.Foreground = Brushes.Black
                source.FontSize = 12
            End If
        Next

    End Sub

End Class

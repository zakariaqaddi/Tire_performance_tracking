Sub CommandButton1_Click()

    Dim Feuille As Worksheet
    Dim sh As Worksheet
    Nomdelasociete = UserForm1.ComboBox1.Value
    Datee = UserForm1.TextBox2.Value
    Imm = UserForm1.TextBox3.Value
    path = "C:\Users\r.tarmi\Desktop\Projet Siége\Data"
    
    
    'Vérifier si le fichier excel de la société déja existe.
    
    If Dir(path & "\" & Nomdelasociete & ".xlsx") = "" Then
     
        'Sinon on doit vérifier si le nom est bien écrit(oui), le fichier n'éxiste pas.
         MsgBox "Nom de la société n'existe pas, Merci de le vérifier"
         
    Else
     
        'Ouvrir le fichier Excel de la société
         classeurcible = Nomdelasociete & ".xlsx"
         Workbooks.Open Filename:="C:\Users\r.tarmi\Desktop\Projet Siége\Data" & "\" & classeurcible
         Workbooks(classeurcible).Activate
         
         If Datee <> "" Or Imm <> "" Then
            cherche = Imm & Datee
            For Each Feuille In Worksheets
              If InStr(1, Feuille.Name, cherche) <= 0 Then
              'Ton traitement
              On Error Resume Next
              Worksheets(Feuille.Name).Visible = False
              End If
              Next

            
            UserForm1.Hide
            Workbooks(classeurcible).Activate
              
         Else
         
            For Each sh In Worksheets
            sh.Visible = True
            Next
            
             Workbooks(classeurcible).Activate
             UserForm1.Hide
        
        End If
     
     
     End If
     

End Sub


Private Sub TextBox1_Change()

End Sub

Private Sub ListBox1_Click()

End Sub

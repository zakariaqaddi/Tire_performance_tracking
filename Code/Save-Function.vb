     
   Sub Enregistrer()
     
        Dim Projet As String, NomFichier As String, nomComplet As String
        Dim path As String, DocName As String
        Dim Feuille As Worksheet
        Dim chbox As CheckBox
        Projet = "Projet.xlsm"
        path = "C:\Users\r.tarmi\Desktop\Projet Siége\Data"
        cherche = ActiveSheet.Range("AW8") & "__" & ActiveSheet.Range("J6")
        
    If ActiveSheet.Range("J8") <> "" Then    'Tester si la cellule Nom de la société n'est pas vide

             DocName = ActiveSheet.Range("J8").Value   'Nom de la société
             classeurcible = path & "\" & DocName & ".xlsx"   'Emplacement database
           
             If Dir(path & "\" & DocName & ".xlsx") = "" Then     'Tester si le fichier N'existe pas

                       Workbooks.Add                         'Créer un classeur
                       ActiveWorkbook.SaveAs classeurcible   'Enregister et le nommer par le nom de la société
                       Workbooks(Projet).Activate            'Activer le classeur projet
                       ActiveSheet.Copy after:=Workbooks(DocName & ".xlsx").Sheets(Workbooks(DocName & ".xlsx").Worksheets.Count)  'Copier la fiche de suivi dans le classeur creer
                       Workbooks(DocName & ".xlsx").Activate 'Activer le classeur cible
                       ActiveSheet.Buttons.Delete  'Delete all buttons
                       'Call deleteboxes 'Delete checkboxes
                       ActiveSheet.Name = ActiveSheet.Range("AW8").Value & "__" & ActiveSheet.Range("J6").Value 'Renommer la feuille par IMM & Date
                      
                       On Error Resume Next
                       Workbooks(DocName & ".xlsx").Sheets("Sheet1").Delete 'Effacer sheet 1 crée par défaut
                       Workbooks(DocName & ".xlsx").Close savechanges:=True  'Fermer et sauvegarder le classeur créé
                       Workbooks(Projet).Activate   'Afficher la fiche de suivi
                       Call Remplir_le_tableau 'Remplir le tableau d'historique
                       Call Remplir_le_tableau_de_suivi 'Remplir le tableau de suivi
                       MsgBox "Enregistré avec succes" 'Message save done
             
             Else                          'Le classeur existe déja
             
                
                Projet = "Projet.xlsm" 'le nom du Classeur du projet
                classeurcible = ActiveSheet.Range("J8") & ".xlsx" 'le nom du Classeur Cible
                Workbooks.Open Filename:="C:\Users\r.tarmi\Desktop\Projet Siége\Data" & "\" & classeurcible  'Ouvrir le Classeur Cible
                Workbooks(classeurcible).Activate 'Activer le classeurcible
                
                'Faire un test si la feuille existe déja

                K = 0
                For Each Feuille In Worksheets
                If Feuille.Name = cherche Then
                K = K + 1
                End If
                Next
                
                If K = 0 Then     'La feuille n'éxiste pas
                    
                    Workbooks(Projet).Activate 'Activer le Classeur Projet
                    ActiveSheet.Copy after:=Workbooks(classeurcible).Sheets(Workbooks(classeurcible).Sheets.Count) 'Copier la Feuille Active et coller dans le Classeur Cible avec un nom = position du sheet
                    Workbooks(classeurcible).ActiveSheet.Buttons.Delete 'Delete all buttons
                    'Call deleteboxes 'Delete checkboxes
                    Workbooks(classeurcible).ActiveSheet.Name = ActiveSheet.Range("AW8").Value & "__" & ActiveSheet.Range("J6").Value 'Renommer la feuille
                    Workbooks(classeurcible).Close savechanges:=True 'Fermer et sauvgarder
                    Workbooks(Projet).Activate 'Activer le fichier de suivi
                    Call Remplir_le_tableau   'Remplir les tableaux de données
                    Call Remplir_le_tableau_de_suivi 'Remplir le tableau de suivi
                    MsgBox "Enregistré avec succes" 'Message save done
                Else
                    MsgBox " Déja enregistré"
                    Workbooks(classeurcible).Close savechanges:=True
                End If
               
             End If
    Else
          MsgBox "Vérifier les données d'entrer"
    End If
        
   End Sub
      
Sub ClearSheet()
On Error Resume Next
    ActiveSheet.UsedRange.Value = vbNullString
On Error GoTo 0
End Sub

Sub Buttonchercher()
UserForm1.Show
End Sub

Sub Remplir_le_tableau()

   Dim Lastrow As Long
   Lastrow = WorksheetFunction.CountA(Sheets("Historique").Range("A:A"))

                 Sheets("Historique").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 Sheets("Historique").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 Sheets("Historique").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 Sheets("Historique").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD6").Value
                 Sheets("Historique").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 Sheets("Historique").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 Sheets("Historique").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 Sheets("Historique").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 Sheets("Historique").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 Sheets("Historique").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BO6").Value
                 Sheets("Historique").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BO8").Value
                 Sheets("Historique").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BO10").Value
End Sub


Sub printer()
ActiveWorkbook.ActiveSheet.PrintOut Copies:=1
End Sub


Sub SavePdf()
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\Users\r.tarmi\Desktop\Projet Siége\" & ActiveSheet.Name & " " & ActiveSheet.Range("J4") & " " & ActiveSheet.Range("AJ4") & ".pdf"
End Sub


Sub deleteboxes()
Dim chbox As CheckBox
For Each chbox In Workbooks(classeurcible).Sheets(ActiveSheet).CheckBoxes
chbox.Delete
Next
End Sub


Sub Remplir_le_tableau_de_suivi()

                       Call CheckBox1
                       Call CheckBox2
                       Call CheckBox3
                       Call CheckBox4
                       Call CheckBox5
                       Call CheckBox6
                       Call CheckBox7
                       Call CheckBox8
                       Call CheckBox9
                       Call CheckBox10
                       Call CheckBox11
                       Call CheckBox12
                       Call CheckBox13
                       Call CheckBox14
                       Call CheckBox15
                       Call CheckBox16
                       Call CheckBox17
                       Call CheckBox18
                       Call CheckBox19
                       Call CheckBox20
                       Call CheckBox21
                       Call CheckBox22
                       Call CheckBox23
                       Call CheckBox24
                       Call CheckBox25
                       Call CheckBox26
                       Call CheckBox27
                       Call CheckBox28
                       Call CheckBox29
                       Call CheckBox30
                       Call CheckBox31
                       Call CheckBox32
                       Call CheckBox33
                       Call CheckBox34
                       Call CheckBox35
                       Call CheckBox36
End Sub

'----------------------------------------Aprés montage Checkboxes----------------------------------------------
Sub basedata()
                 'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
End Sub

Sub CheckBox1()
If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox1.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))
                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BD18").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BD17").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BG17").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BD16").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BD15").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "AVG"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BG19").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BH19").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BI19").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BJ19").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BD19").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
End If
End Sub

Sub CheckBox2()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox2.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BL18").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BL17").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BO17").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BL16").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BL15").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "AVD"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BO19").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BP19").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BQ19").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BR19").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BL19").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
End If
End Sub

Sub CheckBox3()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox3.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AV25").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AV24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AY24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AV23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AV22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARGE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AY26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AZ26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BA26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BB26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AV26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
End If
End Sub

Sub CheckBox4()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox4.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BD28").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BD24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BG24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BD23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BD22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARGI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BG26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BH26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BI26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BJ26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BD26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
End If
End Sub

Sub CheckBox5()

If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox5.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BL25").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BL24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BO24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BL23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BL22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARDI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BO26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BP26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BQ26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BR26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BL26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
End If
End Sub

Sub CheckBox6()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox6.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BT25").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BT24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BW24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BT23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BT22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("AJ6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARDE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BW26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BX26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BY26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BZ26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BT26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub

Sub CheckBox7()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox7.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AV32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AV31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AY31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AV30").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AV29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1GE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AY3").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AZ33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BA33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BB33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AV33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub

Sub CheckBox8()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox8.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BD32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BD31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BG31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BD30").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BD29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1GI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BG33").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BH33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BI33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BJ33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BD33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub

Sub CheckBox9()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox9.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BL32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BL31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BO31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BL30").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BL29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1DI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BO33").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BP33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BQ33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BR33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BL33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox10()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox10.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BT32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BT31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BW31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BT30").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BT29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1DE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BW33").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BX33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BY33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BZ33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BT33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox11()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox11.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AV39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AV38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AY38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AV37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AV36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2GE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AY40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AZ40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BA40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BB40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AV40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub

Sub CheckBox12()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox12.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BD39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BD38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BG38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BD37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BD36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2GI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BG40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BH40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BI40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BJ40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BD40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox13()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox13.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BL39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BL38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BO38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BL37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BL36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2DI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BO40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BP40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BQ40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BR40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BL40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox14()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox14.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))
                 
                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BT39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BT38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BW38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BT37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BT36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2DE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BW40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BX40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BY40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BZ40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BT40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox15()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox15.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AV46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AV45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AY45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AV44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AV43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3GE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AY47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AZ47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BA47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BB47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AV47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox16()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox16.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BD46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BD45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BG45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BD44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BD43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3GI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BG47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BH47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BI47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BJ47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BD47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox17()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox17.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BL46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BL45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BO45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BL44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BL43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3DI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BO47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BP47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BQ47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BR47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BL47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub
Sub CheckBox18()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox18.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("BT46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("BT45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("BW45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("BT44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("BT43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3DE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("BW47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("BX47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("BY47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("BZ47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("BT47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "New"
   End If
End Sub

'-----------------------------------------------Avant montage checkboxes------------------------------------------------------------

Sub CheckBox19()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox19.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                  'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("W18").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("W17").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("Z17").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("W16").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("W15").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "AVG"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("Z19").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AA19").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AB19").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AC19").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("W19").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   End If
End Sub


Sub CheckBox20()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox20.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                  'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AE18").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AE17").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AH17").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AE16").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AE15").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "AVD"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AH19").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AI19").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AJ19").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AK19").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AE19").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"

   End If
End Sub

Sub CheckBox21()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox21.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("O25").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("O24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("R24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("O23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("O22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARGE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("R26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("S26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("T26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("U26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("O26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   
   End If
End Sub

Sub CheckBox22()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox22.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("W25").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("W24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("Z24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("W23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("W22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARGI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("Z26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AA26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AB26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AC26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("W26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   
   End If
End Sub

Sub CheckBox23()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox23.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AE25").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AE24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AH24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AE23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AE22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARDI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AH26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AI26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AJ26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AK26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AE26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   
   End If
End Sub

Sub CheckBox24()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox24.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AM25").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AM24").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AP24").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AM23").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AM22").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ARDE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AP26").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AQ26").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AR26").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AS26").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AM26").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   
   End If
End Sub

Sub CheckBox25()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox25.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("O32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("O31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("R31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("O30").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("O29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1GE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("R33").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("S33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("T33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("U33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("O33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"

   End If
End Sub

Sub CheckBox26()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox26.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                  'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("W32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("W31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("Z31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("W20").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("W29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1GI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("Z33").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AA33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AB33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AC33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("W33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"

   End If
End Sub

Sub CheckBox27()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox27.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AE32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AE31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AH31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AE30").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AE29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1DI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AH33").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AI33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AJ33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AK33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AE33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"

   End If
End Sub

Sub CheckBox28()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox28.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                  'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AM32").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AM31").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AP31").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AM30").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AM29").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES1DE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AP33").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AQ33").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AR33").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AS33").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AM33").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   
   End If
End Sub

Sub CheckBox29()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox29.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("O39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("O38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("R38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("O37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("O36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2GE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("R40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("S40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("T40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("U40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("O40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   
   End If
End Sub

Sub CheckBox30()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox30.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("W39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("W38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("Z38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("W37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("W36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2GI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("Z40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AA40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AB40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AC40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("W40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   End If
End Sub

Sub CheckBox31()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox31.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AE39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AE38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AH38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AE37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AE36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2DI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AH40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AI40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AJ40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AK40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AE40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   End If
End Sub

Sub CheckBox32()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox32.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))
 
                 
                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AM39").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AM38").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AP38").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AM37").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AM36").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES2DE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AP40").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AQ40").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AR40").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AS40").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AM40").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   End If
End Sub

Sub CheckBox33()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox33.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                
                  'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("O46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("O45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("R45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("O44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("O43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3GE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("R47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("S47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("T47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("U47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("O47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"

   End If
End Sub

Sub CheckBox34()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox34.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                 
                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("W46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("W45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("Z45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("W44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("W43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3GI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("Z47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AA47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AB47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AC47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("W47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   End If
End Sub

Sub CheckBox35()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox35.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                
                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AE46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AE45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AH45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AE44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AE43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3DI"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AH47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AI47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AJ47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AK47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AE47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   End If
End Sub

Sub CheckBox36()


If Workbooks("Projet.xlsm").Sheets("Fiche de suivi").CheckBox36.Value = True Then
Lastrow = WorksheetFunction.CountA(Sheets("Tableau de suivi").Range("A:A"))

                 
                   'Date
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 2).Value = Sheets("Fiche de suivi").Range("J6")
                 'Nom de la société
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 3).Value = Sheets("Fiche de suivi").Range("J8").Value
                 'ville
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 4).Value = Sheets("Fiche de suivi").Range("J10").Value
                 'Opérateur
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 5).Value = Sheets("Fiche de suivi").Range("AD8").Value
                 'Configuration
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 6).Value = Sheets("Fiche de suivi").Range("AD10").Value
                 'ON-OFF
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 7).Value = Sheets("Fiche de suivi").Range("AW6").Value
                 'Imm
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 8).Value = Sheets("Fiche de suivi").Range("AW8").Value
                 'Type de visite
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 23).Value = Sheets("Fiche de suivi").Range("BO10").Value
                 'Dimension
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 9).Value = Sheets("Fiche de suivi").Range("AM46").Value
                 'Marque
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 10).Value = Sheets("Fiche de suivi").Range("AM45").Value
                 'Profil
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 11).Value = Sheets("Fiche de suivi").Range("AP45").Value
                 'N série
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 12).Value = Sheets("Fiche de suivi").Range("AM44").Value
                 'DOT
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 13).Value = Sheets("Fiche de suivi").Range("AM43").Value
                 'Date de montage
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 14).Value = Sheets("Fiche de suivi").Range("J6")
                 'Position
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 15).Value = "ES3DE"
                 'Profondeur 1
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 16).Value = Sheets("Fiche de suivi").Range("AP47").Value
                 'Profondeur 2
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 17).Value = Sheets("Fiche de suivi").Range("AQ47").Value
                 'Profondeur 3
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 18).Value = Sheets("Fiche de suivi").Range("AR47").Value
                 'Profondeur 4
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 19).Value = Sheets("Fiche de suivi").Range("AS47").Value
                 'Km
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 20).Value = Sheets("Fiche de suivi").Range("AW10").Value
                 'Pression
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 21).Value = Sheets("Fiche de suivi").Range("AM47").Value
                 'Status
                 Sheets("Tableau de suivi").Cells(Lastrow + 1, 22).Value = "Old"
   End If
End Sub


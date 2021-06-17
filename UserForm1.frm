VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "EXCEL DICE"
   ClientHeight    =   3200
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4410
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dice As Integer
Dim umcnko As String
Dim uscore, mscore, cscore As Long
Dim score As Long
Dim unkocombo, unchicombo, chinkocombo, mankocombo, omankocombo, chinchincombo, ochinchincombo As Integer

Private Sub UserForm_Initialize()
    
    Label11.Caption = 0
    Label17.Caption = 5
    Label13.Caption = 5
    Label16.Caption = 3
    
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    
    CommandButton1.Enabled = True
    CommandButton2.Enabled = False
    
    uscore = 0
    mscore = 0
    cscore = 0
    score = 0
    
    unkocombo = 0
    unchicombo = 0
    chinkocombo = 0
    mankocombo = 0
    omankoconbo = 0
    chinchincombo = 0
    ochinchincombo = 0
    
End Sub


Private Sub CommandButton1_Click()
    
    Dim i As Integer
    
    If CommandButton1.Caption = "Roll" Then
    
        Label16.Caption = Label16.Caption - 1
        
        Label1.Caption = ""
        Label2.Caption = ""
        Label3.Caption = ""
        Label4.Caption = ""
        Label5.Caption = ""
        Label6.Caption = ""
        Label7.Caption = ""
        Label8.Caption = ""
        Label9.Caption = ""
        Label10.Caption = ""
    
        For i = 1 To Label17.Caption
            Randomize
            dice = Int(6 * Rnd + 1)
            Select Case dice
                Case 1: umcnko = "う"
                Case 2: umcnko = "ま"
                Case 3: umcnko = "ち"
                Case 4: umcnko = "ん"
                Case 5: umcnko = "こ"
                Case 6: umcnko = "お"
            End Select
            Dice_Result (i)
        Next i
        
        If Label13.Caption > 0 Then
            CommandButton2.Enabled = True
        ElseIf Label13.Caption = 0 Then
            CommandButton2.Enabled = False
        Else
            MsgBox "エラーが発生しました。終了します。", vbOKOnly + vbCritical, "エラー"
            End
        End If
        CommandButton1.Caption = "Result"
    Else
        Result
        CommandButton1.Caption = "Roll"
    End If
End Sub

Private Sub CommandButton2_Click()

    Dim j As Integer
    
    Label13.Caption = Label13.Caption - 1
    
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    
    For j = 1 To Label17.Caption
        Randomize
        dice = Int(6 * Rnd + 1)
        Select Case dice
            Case 1: umcnko = "う"
            Case 2: umcnko = "ま"
            Case 3: umcnko = "ち"
            Case 4: umcnko = "ん"
            Case 5: umcnko = "こ"
            Case 6: umcnko = "お"
        End Select
        Dice_Result (j)
    Next j
    
    If Label13.Caption < 1 Then
        CommandButton2.Enabled = False
    End If
    
End Sub

Private Sub CommandButton3_Click()

    End
    
End Sub

Sub Dice_Result(ByVal n As Integer)
    
    If n = Label1.Tag Then
        Label1.Caption = umcnko
    ElseIf n = Label2.Tag Then
        Label2.Caption = umcnko
    ElseIf n = Label3.Tag Then
        Label3.Caption = umcnko
    ElseIf n = Label4.Tag Then
        Label4.Caption = umcnko
    ElseIf n = Label5.Tag Then
        Label5.Caption = umcnko
    ElseIf n = Label6.Tag Then
        Label6.Caption = umcnko
    ElseIf n = Label7.Tag Then
        Label7.Caption = umcnko
    ElseIf n = Label8.Tag Then
        Label8.Caption = umcnko
    ElseIf n = Label9.Tag Then
        Label9.Caption = umcnko
    ElseIf n = Label10.Tag Then
        Label10.Caption = umcnko
    Else
        MsgBox "エラーが発生しました。終了します。", vbOKOnly + vbCritical, "エラー"
        End
    End If
    
End Sub

Sub Result()
    Dim u, m, c, n, k, o As Integer
    Dim uz, mz, cz As Long
    Dim l As Integer
    Dim p, q As Integer
    Dim yakucount As Integer
    Dim ochinchin As Boolean
    u = 0
    m = 0
    c = 0
    n = 0
    k = 0
    o = 0
    uz = 0
    mz = 0
    cz = 0
    yakucount = 0
    ochinchin = False
    
    Label17.Caption = 5
    
    For l = 1 To 10
        If l = Label1.Tag Then
            Select Case Label1.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label2.Tag Then
            Select Case Label2.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label3.Tag Then
            Select Case Label3.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label4.Tag Then
            Select Case Label4.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label5.Tag Then
            Select Case Label5.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label6.Tag Then
            Select Case Label6.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label7.Tag Then
            Select Case Label7.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label8.Tag Then
            Select Case Label8.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label9.Tag Then
            Select Case Label9.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        ElseIf l = Label10.Tag Then
            Select Case Label10.Caption
                Case "う": u = u + 1
                Case "ま": m = m + 1
                Case "ち": c = c + 1
                Case "ん": n = n + 1
                Case "こ": k = k + 1
                Case "お": o = o + 1
            End Select
        Else
            MsgBox "エラーが発生しました。終了します。", vbOKOnly + vbCritical, "エラー"
            End
        End If
    Next l
    
    If u > 0 And n > 0 And k > 0 Then
        If unkocombo = 0 Then
            MsgBox "UNKO"
            uscore = uscore + 1000
            yakucount = yakucount + 1
            unkocombo = 1
        Else
            MsgBox ("UNKO" & vbCrLf & "combo " & unkocombo + 1)
            If unkocombo > 2 Then
                uscore = uscore + 8000
            ElseIf unkocombo = 2 Then
                uscore = uscore + 4000
            ElseIf unkocombo = 1 Then
                uscore = uscore + 2000
            End If
            yakucount = yakucount + 1
            unkocombo = unkocombo + 1
        End If
    Else
        unkocombo = 0
    End If
    
    If u > 0 And n > 0 And c > 0 Then
        If unchicombo = 0 Then
            MsgBox "UNCHI"
            uscore = uscore + 1000
            yakucount = yakucount + 1
            unchicombo = 1
        Else
            MsgBox ("UNCHI" & vbCrLf & "combo " & unchicombo + 1)
            If unchicombo > 2 Then
                uscore = uscore + 8000
            ElseIf unchicombo = 2 Then
                uscore = uscore + 4000
            ElseIf unchicombo = 1 Then
                uscore = uscore + 2000
            End If
            yakucount = yakucount + 1
            unchicombo = unchicombo + 1
        End If
    Else
        unchicombo = 0
    End If
    
    If c > 0 And n > 0 And k > 0 Then
        If chinkocombo = 0 Then
            MsgBox "CHINKO"
            cscore = cscore + 1000
            yakucount = yakucount + 1
            chinkocombo = 1
        Else
            MsgBox ("CHINKO" & vbCrLf & "combo " & chinkocombo + 1)
            If chinkocombo > 2 Then
                cscore = cscore + 8000
            ElseIf chinkocombo = 2 Then
                cscore = cscore + 4000
            ElseIf chinkocombo = 1 Then
                cscore = cscore + 2000
            End If
            yakucount = yakucount + 1
            chinkocombo = chinkocombo + 1
        End If
    Else
        chinkocombo = 0
    End If
    
    If o > 0 And m > 0 And n > 0 And k > 0 Then
        If omankocombo = 0 Then
            MsgBox "OMANKO"
            mscore = mscore + 5000
            yakucount = yakucount + 1
            omankocombo = 1
        Else
            MsgBox ("OMANKO" & vbCrLf & "combo " & omankocombo + 1)
            If omankocombo > 2 Then
                mscore = mscore + 40000
            ElseIf omankocombo = 2 Then
                mscore = mscore + 20000
            ElseIf omankocombo = 1 Then
                mscore = mscore + 10000
            End If
            yakucount = yakucount + 1
            omankocombo = omankocombo + 1
        End If
        mankocombo = 0
    ElseIf m > 0 And n > 0 And k > 0 Then
        If mankocombo = 0 Then
            MsgBox "MANKO"
            mscore = mscore + 1000
            yakucount = yakucount + 1
            mankocombo = 1
        Else
            MsgBox ("MANKO" & vbCrLf & "combo " & mankocombo + 1)
            If mankocombo > 2 Then
                mscore = mscore + 8000
            ElseIf mankocombo = 2 Then
                mscore = mscore + 4000
            ElseIf mankocombo = 1 Then
                mscore = mscore + 2000
            End If
            yakucount = yakucount + 1
            mankocombo = mankocombo + 1
        End If
        omankocombo = 0
    Else
        mankocombo = 0
        omankocombo = 0
    End If
    
    If o > 0 And c > 1 And n > 1 Then
        If ochinchincombo = 0 Then
            MsgBox "OCHINCHIN"
            cscore = cscore + 10000
            yakucount = yakucount + 1
            ochinchin = True
            ochinchincombo = 1
        Else
            MsgBox ("OCHINCHIN" & vbCrLf & "combo " & ochinchincombo + 1)
            If ochinchincombo > 2 Then
                cscore = cscore + 80000
            ElseIf ochinchincombo = 2 Then
                cscore = cscore + 40000
            ElseIf ochinchincombo = 1 Then
                cscore = cscore + 20000
            End If
            yakucount = yakucount + 1
            ochinchin = True
            ochinchincombo = ochinchincombo + 1
        End If
        chinchincombo = 0
    ElseIf c > 1 And n > 1 Then
        If chinchincombo = 0 Then
            MsgBox "CHINCHIN"
            cscore = cscore + 3000
            yakucount = yakucount + 1
            chinchincombo = 1
        Else
            MsgBox ("CHINCHIN" & vbCrLf & "combo " & chinchincombo + 1)
            If chinchincombo > 2 Then
                cscore = cscore + 24000
            ElseIf chinchincombo = 2 Then
                cscore = cscore + 12000
            ElseIf chinchincombo = 1 Then
                cscore = cscore + 6000
            End If
            yakucount = yakucount + 1
            chinchincombo = chinchincombo + 1
        End If
        ochinchincombo = 0
    Else
        chinchincombo = 0
        ochinchincombo = 0
    End If
    
    uz = u * 500 + n * 50 + k * 100 + o * 300
    mz = m * 500 + n * 50 + k * 100 + o * 300
    cz = c * 500 + n * 50 + k * 100 + o * 300
    
    uscore = uscore + uz
    mscore = mscore + mz
    cscore = cscore + cz
    
    If u > 6 Then
        uscore = uscore * 6
    ElseIf u = 6 Then
        uscore = uscore * 5
    ElseIf u = 5 Then
        uscore = uscore * 4
    ElseIf u = 4 Then
        uscore = uscore * 3
    ElseIf u = 3 Then
        uscore = uscore * 2
    End If
    
    If m > 6 Then
        mscore = mscore * 6
    ElseIf m = 6 Then
        mscore = mscore * 5
    ElseIf m = 5 Then
        mscore = mscore * 4
    ElseIf m = 4 Then
        mscore = mscore * 3
    ElseIf m = 3 Then
        mscore = mscore * 2
    End If
    
    If c > 6 Then
        cscore = cscore * 6
    ElseIf c = 6 Then
        cscore = cscore * 5
    ElseIf c = 5 Then
        cscore = cscore * 4
    ElseIf c = 4 Then
        cscore = cscore * 3
    ElseIf c = 3 Then
        cscore = cscore * 2
    End If
    
    score = uscore + mscore + cscore
    
    If n > 6 Then
        uscore = uscore * -7
        mscore = mscore * -7
        cscore = cscore * -7
    ElseIf n = 6 Then
        uscore = uscore * -6
        mscore = mscore * -6
        cscore = cscore * -6
    ElseIf n = 5 Then
        uscore = uscore * -5
        mscore = mscore * -5
        cscore = cscore * -5
    ElseIf n = 4 Then
        uscore = uscore * -4
        mscore = mscore * -4
        cscore = cscore * -4
    ElseIf n = 3 Then
        uscore = uscore * -3
        mscore = mscore * -3
        cscore = cscore * -3
    End If
    
    score = uscore + mscore + cscore
    
    If k > 6 Then
        uscore = uscore * 5.5
        mscore = mscore * 5.5
        cscore = cscore * 5.5
    ElseIf k = 6 Then
        uscore = uscore * 4.5
        mscore = mscore * 4.5
        cscore = cscore * 4.5
    ElseIf k = 5 Then
        uscore = uscore * 3.5
        mscore = mscore * 3.5
        cscore = cscore * 3.5
    ElseIf k = 4 Then
        uscore = uscore * 2.5
        mscore = mscore * 2.5
        cscore = cscore * 2.5
    ElseIf k = 3 Then
        uscore = uscore * 1.5
        mscore = mscore * 1.5
        cscore = cscore * 1.5
    End If
    
    score = uscore + mscore + cscore
    
    If o > 6 Then
        If score < 0 Then
            uscore = uscore * -5.5
            mscore = mscore * -5.5
            cscore = cscore * -5.5
        Else
            uscore = uscore * 5.5
            mscore = mscore * 5.5
            cscore = cscore * 5.5
        End If
    ElseIf o = 6 Then
        If score < 0 Then
            uscore = uscore * -4.5
            mscore = mscore * -4.5
            cscore = cscore * -4.5
        Else
            uscore = uscore * 4.5
            mscore = mscore * 4.5
            cscore = cscore * 4.5
        End If
    ElseIf o = 5 Then
        If score < 0 Then
            uscore = uscore * -3.5
            mscore = mscore * -3.5
            cscore = cscore * -3.5
        Else
            uscore = uscore * 3.5
            mscore = mscore * 3.5
            cscore = cscore * 3.5
        End If
    ElseIf o = 4 Then
        If score < 0 Then
            uscore = uscore * -2.5
            mscore = mscore * -2.5
            cscore = cscore * -2.5
        Else
            uscore = uscore * 2.5
            mscore = mscore * 2.5
            cscore = cscore * 2.5
        End If
    ElseIf o = 3 Then
        If score < 0 Then
            uscore = uscore * -1.5
            mscore = mscore * -1.5
            cscore = cscore * -1.5
        Else
            uscore = uscore * 1.5
            mscore = mscore * 1.5
            cscore = cscore * 1.5
        End If
    End If
    
    score = uscore + mscore + cscore
    
    Label11.Caption = score
    
    If yakucount > 0 Then
        Label16.Caption = Label16.Caption + 1
        For p = 1 To yakucount
            Label13.Caption = Label13.Caption + 1
        Next p
    End If
    
    If ochinchin = True Then
        Label17.Caption = 10
    ElseIf yakucount > 1 Then
        For q = 1 To yakucount - 1
            Label17.Caption = Label17.Caption + 1
        Next q
    End If
    
    ochinchin = False
    
    If Label16.Caption < 1 Then
        CommandButton1.Enabled = False
    Else
        CommandButton1.Enabled = True
    End If
    
    CommandButton2.Enabled = False
    
End Sub

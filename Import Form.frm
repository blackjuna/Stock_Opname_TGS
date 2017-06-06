VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import From Excel"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13515
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_import 
      Caption         =   "IMPORT"
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton cmd_cvt 
      Caption         =   "CONVERT"
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton cmd_del 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   12000
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd_sf 
      Caption         =   "SEARCH FILE"
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txt_fileadd 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4200
      Width           =   8655
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   9960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvw_so 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "File Address"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_cvt_Click()
    OpenExcel
End Sub

Private Function ClearAll()
    lvw_so.ListItems.Clear
    txt_fileadd.Text = ""
    txt_fileadd.SetFocus
End Function

Private Function ImportTable()
    Dim I As Integer
    Dim C1 As String
    Dim C2 As String
    Dim C3, C4, C5, C6 As String
    
    'On Error Resume Next
    If lvw_so.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang disimpan", vbExclamation, "Informasi Import"
        Exit Function
    End If
    
    With lvw_so
        For I = 1 To .ListItems.Count
            C1 = .ListItems(I).SubItems(1) 'tag no
            C2 = .ListItems(I).SubItems(2) 'part code
            C3 = .ListItems(I).SubItems(3) 'part name
            C4 = .ListItems(I).SubItems(4) 'category
            C5 = .ListItems(I).SubItems(5) 'location
            C6 = .ListItems(I).SubItems(6) 'grup
            C7 = .ListItems(I).SubItems(7) 'tag code
            C8 = .ListItems(I).SubItems(8) 'satuan
            C9 = .ListItems(I).SubItems(9) 'ukuran
            
                        
            sql = "insert into tag_stock_opname_tgs (tag_no,location,part_no,part_name," & _
                "category,grup,tag_code,status,satuan,qty,qty_admin,qty_selisih) " & _
                "values('" & C1 & "','" & C5 & "'," & _
                "'" & C2 & "','" & C3 & "','" & C4 & "'," & _
                "'" & C6 & "','" & C7 & "','','" & C8 & "','" & Val(C9) & "',0," & _
                "0)"
            conn.Execute (sql)
        Next
    End With
    
    MsgBox "Anda berhasil impor data", vbInformation + vbOKOnly, "Informasi import"
    ClearAll
    
End Function

Private Sub cmd_del_Click()
    ClearAll
End Sub

Private Sub cmd_import_Click()
        ImportTable
        Unload Me
'    'On Err GoTo err_rpt
'    msg = MsgBox("Apakah anda ingin impor data?", vbInformation + vbOKCancel, "Informasi")
'    If msg = vbOK Then
'        'db
'        ImportTable
'        Unload Me
'    Else
'        txt_fileadd.SetFocus
'    End If
''err_rpt:
''    MsgBox Err.Description, vbExclamation, Err.Number
End Sub

Private Sub cmd_sf_Click()
    cd.Filter = "Microsoft Excel (*.xlsx)|*xls;*.xlsx"
    cd.InitDir = App.Path
    cd.ShowOpen

    txt_fileadd.Text = cd.FileName
End Sub

Private Function OpenExcel()
'    strnamadatabase = "stock_opname_db"
'    strnamaserver = "proliant\sqlexpress"
'    strnamapemakai = "sa"
'    strpassword = "admin123"
'    strnamapemakaicf = "cfuser"
'    intconnectiontimeout = 60
'    strprovider = "SQLOLEDB.1"
'
'    strconstr = "Provider=" & strprovider & "; Data Source=" & "192.168.0.108, 1433" & "; Network Library=DBMSSOCN; Initial Catalog=" & strnamadatabase & ";User ID=" & strnamapemakai & ";Password=" & strpassword

    Dim strAlamat As String
    Dim strKonek As String
    Dim struser As String
    Dim Lis As ListItem
    Dim no As String
    
'    On Error GoTo 1
    strAlamat = txt_fileadd.Text
    strKonek = "provider=Microsoft.ACE.OLEDB.12.0;"
    strKonek = strKonek & "data source=" & strAlamat & ";"
    strKonek = strKonek & "extended properties= 'Excel 12.0;HDR=yes;imex=1';"
    If conn_excel.State = 1 Then conn_excel.Close
    conn_excel.CursorLocation = adUseClient
    conn_excel.Open strKonek
    
    Dim strKueri As String
    Set rs_import = New ADODB.Recordset
    
    strKueri = "select * from [Sheet1$]"
    rs_import.Open strKueri, conn_excel, adOpenDynamic, adLockOptimistic
    
    If rs_import.RecordCount > 0 Then
        rs_import.MoveFirst
        lvw_so.ListItems.Clear
        no = 0
        While Not rs_import.EOF
            Set Lis = lvw_so.ListItems.Add(, , no + 1)
                Lis.SubItems(1) = IIf(IsNull(rs_import.Fields(0).Value), "", rs_import.Fields(0).Value)
                Lis.SubItems(2) = IIf(IsNull(rs_import.Fields(1).Value), "", rs_import.Fields(1).Value)
                Lis.SubItems(3) = IIf(IsNull(rs_import.Fields(2).Value), "", rs_import.Fields(2).Value)
                Lis.SubItems(4) = IIf(IsNull(rs_import.Fields(3).Value), "", rs_import.Fields(3).Value)
                Lis.SubItems(5) = IIf(IsNull(rs_import.Fields(4).Value), "", rs_import.Fields(4).Value)
                Lis.SubItems(6) = IIf(IsNull(rs_import.Fields(5).Value), "", rs_import.Fields(5).Value)
                Lis.SubItems(7) = IIf(IsNull(rs_import.Fields(6).Value), "", rs_import.Fields(6).Value)
                Lis.SubItems(8) = IIf(IsNull(rs_import.Fields(7).Value), "", rs_import.Fields(7).Value)
                Lis.SubItems(9) = IIf(IsNull(rs_import.Fields(8).Value), "", rs_import.Fields(8).Value)
                no = no + 1
            rs_import.MoveNext
        Wend
        Tb = no
        'Lis.Selected
        Lis.EnsureVisible
    End If
    If conn_excel.State = 1 Then conn_excel.Close
    Exit Function
'1:
'    MsgBox Err.Description, vbExclamation, Err.Number
'    Exit Function
End Function

Private Function TABEL()
    With lvw_so
    .View = lvwReport
    .Gridlines = True
    .FullRowSelect = True
    .HotTracking = True
        With .ColumnHeaders
        .Add , , "NO", 600
        .Add , , "TAG NO", 1200
        .Add , , "PART CODE", 1200
        .Add , , "PART NAME", 3900
        .Add , , "CATEGORY", 1200
        .Add , , "LOCATION", 1200
        .Add , , "GRUP", 1200
        .Add , , "TAG CODE", 1200
        .Add , , "SATUAN", 900
        .Add , , "QTY", 900
        End With
    End With
End Function

Private Sub Form_Load()
    TABEL
End Sub


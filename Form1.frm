VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEIDH 
      Caption         =   "¡–œ»« ≈’”« ≈…ƒŸÕ Ã≈ –À«—« œÕœÃ¡”…¡"
      Height          =   480
      Left            =   0
      TabIndex        =   6
      Top             =   7800
      Width           =   3855
   End
   Begin VB.CommandButton cmdYPOLOGISMOSSECSV 
      Caption         =   "’–œÀœ√…”ÃŸÕ ”E CSV"
      Height          =   480
      Left            =   4320
      TabIndex        =   5
      Top             =   3840
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6480
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox text1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Text            =   " "
      Top             =   312
      Width           =   5580
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   5688
      TabIndex        =   2
      Top             =   312
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1800
      Top             =   6480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "’–œÀœ√…”Ãœ” ”’ÕœÀŸÕ ”≈ EXCEL"
      Height          =   480
      Left            =   7440
      TabIndex        =   1
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "¡–œ»« ≈’”«  ¡‘¡√—¡÷«” ”≈ ¬¡”«"
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label lbl¡Ò˜ÂﬂÔExcel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "¡Ò˜ÂﬂÔ Excel Ôı ÂÒÈ›˜ÂÈ ÙÔÌ ÙÈÏÔÍ·Ù‹ÎÔ„Ô"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4332
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gdb As New ADODB.Connection




Private Sub cmdCommand1_Click()
       'Set xlsheet = xlwbook.Sheets.Item(1)

gdb.Execute "DELETE FROM MEGGTIM"

    

390     Open text1.Text For Input As #1

        Dim ko   As String

        Dim mNew As Long, mUpd As Long

400     mNew = 0
410     mUpd = 0


        Dim AA, DUM

        Dim mRow As Long

        ' data1.Recordset.MoveFirst


        Dim ELEM(1 To 60)

       ' On Error GoTo error_name
       Dim CC As String
       

470     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

480         Line Input #1, AA

            ' DUM = to437(AA)
490         If EOF(1) Then

                Exit Do

            End If

               CC = Left(Split(AA, ";")(7), 60)
               gdb.Execute "insert into MEGGTIM (BARC,ONO,POSO) VALUES ('" + Split(AA, ";")(2) + "','" + CC + "',1)"
               




740         DoEvents

760         mRow = mRow + 1    'data1.Recordset.MoveNext
        Loop

        'xl.Quit
        'Set xlwbook = Nothing
        'Set xl = Nothing

770     Close #1

End Sub

Private Sub cmdCommand2_Click()
    '--------- excel  kartella --------------------------------
        '<EhHeader>
      '  On Error GoTo Command11_Click_Err

        '</EhHeader>

        Dim Excel    As Excel.Application

        Dim workbook As Excel.workbook

        Dim myXL     As Excel.Worksheet

100     Set Excel = New Excel.Application
        ' Excel.Visible = True
110     Set workbook = Excel.Workbooks.Add

        On Error Resume Next

120     workbook.Activate

130     Set myXL = workbook.ActiveSheet
        
        
        
Dim R As New ADODB.Recordset
R.Open "SELECT SUM(POSO) AS SS,BARC,ONO  From [MEGGTIM]  GROUP BY BARC , ONO", gdb, adOpenDynamic, adLockOptimistic

'Open "c:\mercvb\synola.csv" For Output As #3
Dim N As Integer
N = 0
Do While Not R.EOF
   N = N + 1
     myXL.Cells(N, 1) = R!barc
     myXL.Cells(N, 2) = R!ono
     myXL.Cells(N, 3) = R!ss
     
    ' Print #3, "'" + R!barc + ";" + R!ono + ";" + Str(R!ss)
     
     
     
   R.MoveNext
  


Loop


        
R.Close


 ' Close #3
  
        
        

540     DoEvents

550     myXL.SaveAs "C:\MERCVB\EKTYP.XLS"

 Excel.Visible = True

        Dim ANS3 As Long

560     ANS3 = MsgBox(" ÎÂﬂÌ˘ ÙÔ EXCEL", vbYesNo)

570     If ANS3 = vbYes Then
580         Call workbook.Close(False)
590         Excel.Quit
600         Set Excel = Nothing
        End If

        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
End Sub

Private Sub cmdSUMES_Click()

Dim R As New ADODB.Recordset
R.Open "SELECT SUM(POSO) AS SS,BARC,ONO  From [MEGGTIM]  GROUP BY BARC , ONO", gdb, adOpenDynamic, adLockOptimistic







End Sub

Private Sub cmdYPOLOGISMOSSECSV_Click()

'--------- excel  kartella --------------------------------
        '<EhHeader>
      '  On Error GoTo Command11_Click_Err

        '</EhHeader>

        Dim Excel    As Excel.Application

        Dim workbook As Excel.workbook

        Dim myXL     As Excel.Worksheet


        On Error Resume Next


        
Dim R As New ADODB.Recordset
R.Open "SELECT SUM(POSO) AS SS,BARC,ONO  From [MEGGTIM]  GROUP BY BARC , ONO", gdb, adOpenDynamic, adLockOptimistic

Open "c:\mercvb\synola.csv" For Output As #3
Dim N As Integer
N = 0
Do While Not R.EOF
   N = N + 1
    ' myXL.Cells(N, 1) = R!barc
    ' myXL.Cells(N, 2) = R!ono
    ' myXL.Cells(N, 3) = R!ss
     
     Print #3, "'" + R!barc + ";" + R!ono + ";" + Str(R!ss)
     
     
     
   R.MoveNext
  


Loop


        
R.Close


  Close #3
  
        
    'Shell "EXCEL c:\mercvb\synola.csv", vbMaximizedFocus
    
Shell "EXPLORER.EXE c:\mercvb\synola.csv", vbMaximizedFocus

End Sub

Private Sub cmdEIDH_Click()


gdb.Execute "DELETE FROM MEIDH"

    

390     Open text1.Text For Input As #1

        Dim ko   As String

        Dim mNew As Long, mUpd As Long

400     mNew = 0
410     mUpd = 0


        Dim AA, DUM

        Dim mRow As Long

        ' data1.Recordset.MoveFirst


        Dim ELEM(1 To 60)

       ' On Error GoTo error_name
       Dim CC As String
       

470     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

480         Line Input #1, AA

            ' DUM = to437(AA)
490         If EOF(1) Then

                Exit Do

            End If

               CC = Left(Split(AA, ";")(1), 60)
               gdb.Execute "insert into MEIDH (BARC,ONO) VALUES ('" + Split(AA, ";")(0) + "','" + CC + "')"
               




740         DoEvents

760         mRow = mRow + 1    'data1.Recordset.MoveNext
        Loop

        'xl.Quit
        'Set xlwbook = Nothing
        'Set xl = Nothing

770     Close #1










End Sub

Private Sub Command2_Click()


100     If Len(Trim(text1.Text)) = 0 Then
110         cd1.ShowOpen
120         text1.Text = cd1.FileName
        Else

130         If Len(Dir(LTrim(text1.Text), vbNormal)) < 2 Then
140             MsgBox "‰ÂÌ ı‹Ò˜ÂÈ ÙÔ ·Ò˜ÂﬂÔ " + text1.Text

                Exit Sub

            End If
        End If
 
      

150     Me.MousePointer = vbHourglass
     
       

160     Me.MousePointer = vbNormal






End Sub

Private Sub Form_Load()

gdb.Open "DSN=MERCSQL;uid=sa;pwd=p@ssw0rd;"
Dim sql As String

sql = "IF  NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MEGGTIM]') AND type in (N'U'))"
sql = sql + "  CREATE TABLE [dbo].[MEGGTIM](ID INT NOT NULL IDENTITY(1,1),BARC VARCHAR(20),ONO VARCHAR(60),POSO NUMERIC(10,2) ) "
gdb.Execute sql


sql = "IF  NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MEIDH]') AND type in (N'U'))"
sql = sql + "  CREATE TABLE [dbo].[MEIDH](ID INT NOT NULL IDENTITY(1,1),BARC VARCHAR(20),ONO VARCHAR(60),POSO NUMERIC(10,2) ) "
gdb.Execute sql


End Sub

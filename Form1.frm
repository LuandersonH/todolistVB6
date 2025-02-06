VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstabToDoList 
      Height          =   9045
      Left            =   -75
      TabIndex        =   0
      Top             =   0
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   15954
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   5292
      BackColor       =   8421440
      TabCaption(0)   =   "ToDoList"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTodolist"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "listTasks"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tboxInsertTask"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnInsertTask"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnDeleteTask"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnFinishedTask"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnClearAll"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.CommandButton btnClearAll 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "LIMPAR TUDO"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11085
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6795
         Width           =   1725
      End
      Begin VB.CommandButton btnFinishedTask 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "CONCLUÍDA"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   11100
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4185
         Width           =   1725
      End
      Begin VB.CommandButton btnDeleteTask 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "EXCLUIR"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   11085
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5505
         Width           =   1725
      End
      Begin VB.CommandButton btnInsertTask 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "ADICIONAR"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   11055
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1875
         Width           =   1725
      End
      Begin VB.TextBox tboxInsertTask 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   2955
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1875
         Width           =   8080
      End
      Begin VB.ListBox listTasks 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         ItemData        =   "Form1.frx":0038
         Left            =   2955
         List            =   "Form1.frx":003A
         TabIndex        =   1
         Top             =   4245
         Width           =   8080
      End
      Begin VB.Label lblTodolist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTA DE TAREFAS"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   810
         Left            =   4380
         TabIndex        =   7
         Top             =   780
         Width           =   6810
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim handleTaskValue As String


Private Sub tabToDoList_DblClick()

Private Sub btnClearAll_Click()
resposta = MsgBox("Isso apagará TODAS AS TAREFAS e não poderá ser desfeito, deseja continuar?", vbOKCancel, "Apagar todas as tarefas")
If resposta = vbOK Then
listTasks.Clear
End If
End Sub

Private Sub btnInsertTask_Click()
If tboxInsertTask.Text <> "" Then
handleTaskValue = tboxInsertTask.Text
listTasks.AddItem handleTaskValue
tboxInsertTask.Text = ""
Else
MsgBox "Crie uma tarefa antes de adicionar a lista!", vbExclamation, "Aviso"
End If
End Sub

Private Sub btnDeleteTask_Click()
If listTasks.ListIndex <> -1 Then
listTasks.RemoveItem listTasks.ListIndex
Else
MsgBox "Selecione uma tarefa a ser removida!", vbExclamation, "Aviso"
End If
End Sub

Private Sub btnFinishedTask_Click()
If listTasks.ListIndex <> -1 Then
listTasks.List(listTasks.ListIndex) = "[CHECK!] " & listTasks.List(listTasks.ListIndex)
Else
MsgBox "Selecione uma tarefa a ser concluída!", vbExclamation, "Aviso"
End If
End Sub


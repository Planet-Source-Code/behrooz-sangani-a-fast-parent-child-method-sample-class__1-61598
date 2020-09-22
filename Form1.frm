VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parent Child Method Sample"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView tv1 
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5106
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton cmdGetRoot 
      Caption         =   "Get Root Item(s)"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Parent Field:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblParent 
      Caption         =   "..."
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   4080
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Parent Child Method Sample
'  Sample usage of ParentChild class
'=========================================================================================
'  Created By: Behrooz Sangani
'  Published Date: 7/9/2005
'  E-Mail: sangani@gmail.com
'  Legal Copyright: Behrooz Sangani © 7/9/2005
'  Free for personal use but do not remove the copyright notice.
'  For comercial use please contact the author.
'=========================================================================================

Dim objParentChild As ParentChild

Private Sub cmdGetRoot_Click()

    Dim RS As Recordset 'temporary object

    tv1.Nodes.Clear     'clear the nodes first

    With objParentChild
        'set the db path first
        .DatabasePath = App.Path & "\db.mdb"
        If .Connect Then    'if no error in connection

            .pkField = "ID"     'PrimaryKey Field name in table
            .piField = "Boss"   'ParentID Field name in table

            If .OpenRecordset("select * from Employee") Then    'If no error opening the recordset
                'get root items. (ParentID = 0)
                Set RS = .ChildFields(0)
                For i = 1 To RS.RecordCount
                    'add items to the tree. (node key = "I" & [primary key])
                    tv1.Nodes.Add , , "I" & RS("ID"), RS("Name") & " " & RS("Family") & "  [" & RS("Job") & "]"
                    If Not RS.EOF Then RS.MoveNext

                Next
            Else
                MsgBox .LastError
            End If
        Else
            MsgBox .LastError
        End If
    End With

    Set RS = Nothing

End Sub

Private Sub Form_Load()

    'load a new object
    Set objParentChild = New ParentChild

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'destroy objects on unload
    Set objParentChild = Nothing

End Sub

'load child nodes when clicked on each node
Private Sub tv1_Click()

    Dim RS As Recordset     'temporary object

    If tv1.SelectedItem.Children = 0 Then   'if child nodes are not yet loaded
        With objParentChild
            'get child nodes for the selected node using its key
            Set RS = .ChildFields(Replace(tv1.SelectedItem.Key, "I", ""))
            For i = 1 To RS.RecordCount
                'add items in tree under the right parent
                tv1.Nodes.Add tv1.SelectedItem.Key, tvwChild, "I" & RS("ID"), RS("Name") & " " & RS("Family") & "  [" & RS("Job") & "]"
                If Not RS.EOF Then RS.MoveNext

            Next
        End With
    End If

    If tv1.SelectedItem.Parent Is Nothing Then
        'well it's a root element
        lblParent.Caption = "No top level field"
    Else
        'show the parent item
        Set RS = objParentChild.ParentField(Replace(tv1.SelectedItem.Key, "I", ""))
        lblParent.Caption = RS("Name") & " " & RS("Family")
    End If

    Set RS = Nothing

End Sub

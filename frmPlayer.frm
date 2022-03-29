VERSION 5.00
Begin VB.Form frmPlayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "玩家信息"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "微软雅黑 Light"
      Size            =   10.5
      Charset         =   134
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   12030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdWorldId 
      Caption         =   "世界ID"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   9
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1140
      TabIndex        =   5
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdCopyID 
      Caption         =   "复制ID"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   9
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   855
   End
   Begin VB.Line Line3 
      Tag             =   "H"
      X1              =   8640
      X2              =   8640
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Line Line2 
      Tag             =   "H"
      X1              =   5400
      X2              =   5400
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Line Line1 
      Tag             =   "H"
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   6135
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   6135
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   6135
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   -360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PlayerID As String, WorldID As String
Attribute WorldID.VB_VarUserMemId = 1073938432


Private Sub cmdCopyID_Click()
    Clipboard.SetText PrettyID(PlayerID)
    MsgBox PrettyID(PlayerID) & " 复制成功！"
End Sub
Private Sub cmdWorldID_Click()
    Clipboard.SetText WorldID
    MsgBox WorldID & " 复制成功！"
End Sub

Private Sub Form_Load()
    Me.Caption = "隐藏の无敌星 " & App.Major & "." & App.Minor & "." & App.Revision & " - 玩家信息"
End Sub
Public Sub PlayerInfo(Player As Variant)
'显示玩家信息
    Dim stdpic As New stdPicEx2, SingleBadge As Variant
    Label1 = Player("name")
    LblAppend Label1, "ID: " & PrettyID(CStr(Player("code")))
    LblAppend Label1, "地区: " & Player("country")
    LblAppend Label1, "上次上线于: " & Player("last_active_pretty")
    Label2 = "玩过的关卡: " & Player("courses_played")
    LblAppend Label2, "通过的关卡: " & Player("courses_cleared")
    LblAppend Label2, "通过的世界: " & Player("unique_super_world_clears")
    LblAppend Label2, "用过的命数: " & Player("courses_attempted")
    LblAppend Label2, "死掉的命数: " & Player("courses_deaths")
    LblAppend Label2, "首插关卡数: " & Player("first_clears")
    LblAppend Label2, "世界纪录数: " & Player("world_records")
    LblAppend Label2, ""
    LblAppend Label2, "简单团纪录: " & Player("easy_highscore")
    LblAppend Label2, "普通团纪录: " & Player("normal_highscore")
    LblAppend Label2, "困难团纪录: " & Player("expert_highscore")
    LblAppend Label2, "极难团纪录: " & Player("super_expert_highscore")
    LblAppend Label2, ""
    LblAppend Label2, "上传关卡数: " & Player("uploaded_levels")
    LblAppend Label2, "收到的点赞: " & Player("likes")
    LblAppend Label2, "工匠点数: " & Player("maker_points")
    If Player("super_world_id") <> "" Then LblAppend Label2, "世界ID: " & Player("super_world_id")
    Label3 = "对战段位: " & Player("versus_rank_name")
    LblAppend Label3, "对战积分: " & Player("versus_rating")
    LblAppend Label3, ""
    LblAppend Label3, "总场数: " & Player("versus_plays")
    LblAppend Label3, "胜利场数: " & Player("versus_won")
    LblAppend Label3, "失败场数: " & Player("versus_lost")
    LblAppend Label3, "连胜场数: " & Player("versus_win_streak")
    LblAppend Label3, "连败场数: " & Player("versus_lose_streak")
    LblAppend Label3, "掉线场数: " & Player("versus_disconnected")
    LblAppend Label3, "击杀数: " & Player("versus_kills")
    LblAppend Label3, "被击杀数: " & Player("versus_killed_by_others")
    LblAppend Label3, "网络区服: " & Player("region_name")
    LblAppend Label3, ""
    LblAppend Label3, "合作场数: " & Player("coop_plays")
    LblAppend Label3, "合作通过数: " & Player("coop_clears")
    Label4 = "取得的奖牌: "
    For Each SingleBadge In Player("badges")
        LblAppend Label4, PrettyBadge(CStr(SingleBadge("type_name"))) & " (" & Replace(Replace(Replace(CStr(SingleBadge("rank_name")), "Gold", "金牌"), "Silver", "银牌"), "Bronze", "铜牌") & ")"
    Next SingleBadge
    If Label4.Caption = "取得的奖牌: " Then Label4.Caption = "该玩家没有奖牌"
    PlayerID = CStr(Player("code"))
    WorldID = CStr(Player("super_world_id"))
    Me.Show
    DoEvents
    Call URLDownloadToFile(0, Player("mii_image"), EnvTempDir & "\HSSTemp\Mii.png", 0, 0)
    Image1.Picture = stdpic.LoadPictureEx(EnvTempDir & "\HSSTemp\Mii.png")

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Label1.Height = Me.Height - 3645
    Label2.Height = Me.Height - 765
    Label3.Height = Me.Height - 765
    Label2.Width = 3015 + (Me.Width - 11805) / 3
    Label3.Width = 3015 + (Me.Width - 11805) / 3
    Label4.Width = 3015 + (Me.Width - 11805) / 3
    Label3.Left = 5640 + ((Me.Width - 11805) / 3)
    Label4.Left = 8880 + ((Me.Width - 11805) / 3) * 2
    Line2.X1 = 5400 + ((Me.Width - 11805) / 3)
    Line2.X2 = 5400 + ((Me.Width - 11805) / 3)
    Line3.X1 = 8640 + ((Me.Width - 11805) / 3) * 2
    Line3.X2 = 8640 + ((Me.Width - 11805) / 3) * 2
    Line1.Y2 = Me.Height - 540
    Line2.Y2 = Me.Height - 540
    Line3.Y2 = Me.Height - 540
    cmdCopyID.Top = Me.Height - 1140
    cmdWorldId.Top = Me.Height - 1140
End Sub

Private Function PrettyBadge(Badge As String) As String
    PrettyBadge = Replace(Badge, "Super Expert", "极难")
    PrettyBadge = Replace(PrettyBadge, "Expert", "困难")
    PrettyBadge = Replace(PrettyBadge, "Easy", "简单")
    PrettyBadge = Replace(PrettyBadge, "Normal", "普通")
    PrettyBadge = Replace(PrettyBadge, "Endless Challenge", "马力欧耐力挑战")
    PrettyBadge = Replace(PrettyBadge, "Multiplayer Versus", "多人对战")
    PrettyBadge = Replace(PrettyBadge, "Number of Clears", "通过关卡")
End Function

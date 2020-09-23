VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   0
      TabIndex        =   0
      Top             =   5475
      Width           =   8340
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "continue to next box"
         Height          =   465
         Left            =   4200
         TabIndex        =   13
         Top             =   975
         Width           =   1140
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   315
         LargeChange     =   5
         Left            =   2700
         Max             =   20
         TabIndex        =   11
         Top             =   1125
         Width           =   1365
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   5550
         TabIndex        =   9
         Top             =   450
         Width           =   2490
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   1
         Left            =   1425
         TabIndex        =   3
         Top             =   450
         Width           =   1065
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   450
         Width           =   1065
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   315
         LargeChange     =   16
         Left            =   2700
         Max             =   256
         TabIndex        =   1
         Top             =   450
         Width           =   2640
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   240
         Left            =   2700
         TabIndex        =   12
         Top             =   900
         Width           =   1440
      End
      Begin VB.Label Label5 
         Caption         =   "Test:"
         Height          =   240
         Left            =   5550
         TabIndex        =   10
         Top             =   225
         Width           =   1890
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "Overflow !!!"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   3525
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   240
         Left            =   2700
         TabIndex        =   6
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label Label2 
         Caption         =   "Vert"
         Height          =   240
         Index           =   1
         Left            =   1425
         TabIndex        =   5
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Horiz"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   3525
      TabIndex        =   8
      Top             =   4575
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLIENT_MARGIN As Long = 10 'px
Private Const SPACE_WIDTH As Long = 5 'px
Private Const LINE_SPACING As Long = 5 'px

Private Const MAX_WORDS_IN_BOX As Long = 256
Private Const MAX_ROWS_IN_BOX As Long = 256

Private Const ALIGN_MIDDLE As Long = 0
Private Const ALIGN_NEAR As Long = 1
Private Const ALIGN_FAR As Long = 2
Private Const ALIGN_JUSTIFY As Long = ALIGN_NEAR Or ALIGN_FAR


Private Type T_BOX
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type


Private Type T_WORD
    rtBox As T_BOX
End Type

Private m_rWord(MAX_WORDS_IN_BOX) As T_WORD
Private m_lNumWords As Long


Private Type T_SEGM
    lMinX As Long
    lMaxX As Long
    
    lWordIdx() As Long
    lNumWords As Long
    
    lTotWidth As Long
'    bFull As Boolean
End Type


Private Type T_ROW
    lNumber As Long
    lTop As Long
    lMaxHeight As Long
    
    rSegm() As T_SEGM
    lNumSegms As Long
End Type


Private m_lAlignX As Long
Private m_lAlignY As Long

Private m_bContinueToNextBox As Boolean


    ' Generatore di regioni
Private m_lClientWidth As Long
Private m_lClientHeight As Long

Private m_lMargin As Long

Private m_hRegion(1) As Long
Private m_lNumRegions As Long


    ' API stuffz
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Private Sub DefineRegions()
    ' Definisce le regioni da usare
    Me.Cls
    
    FlushRegions
    
    
Dim lIdx As Long
Dim lhBrush As Long
Dim ptTmp() As POINTAPI
Dim rtTmp As RECT
    
    ' Ricava il margine
    m_lMargin = HScroll2.Value
    Label6.Caption = "Margin=" & m_lMargin & " px"
    
    
    Select Case List2.ListIndex
        Case 0
            ' Due semplici regioni rettangolari (60% / 40%)
            m_lNumRegions = 2
            
            SetRect rtTmp, 0, 0, m_lClientWidth * 0.6, m_lClientHeight
            InflateRect rtTmp, -CLIENT_MARGIN, -CLIENT_MARGIN
            m_hRegion(0) = CreateRectRgn(rtTmp.Left, _
                                        rtTmp.Top, _
                                        rtTmp.Right, _
                                        rtTmp.Bottom)
            
            SetRect rtTmp, m_lClientWidth * 0.6, 0, m_lClientWidth, m_lClientHeight
            InflateRect rtTmp, -CLIENT_MARGIN, -CLIENT_MARGIN
            m_hRegion(1) = CreateRectRgn(rtTmp.Left, _
                                        rtTmp.Top, _
                                        rtTmp.Right, _
                                        rtTmp.Bottom)
    
    
        Case 1
            ' Una regione ellittica
            m_lNumRegions = 1
            
            SetRect rtTmp, 0, 0, m_lClientWidth, m_lClientHeight
            InflateRect rtTmp, -CLIENT_MARGIN, -CLIENT_MARGIN
            m_hRegion(0) = CreateEllipticRgn(rtTmp.Left, _
                                            rtTmp.Top, _
                                            rtTmp.Right, _
                                            rtTmp.Bottom)
            
        Case 2
            ' doppio trapezio
            m_lNumRegions = 2
            
            ReDim ptTmp(3) As POINTAPI
            
            ptTmp(0).x = CLIENT_MARGIN
            ptTmp(0).y = CLIENT_MARGIN
            
            ptTmp(1).x = m_lClientWidth * 0.65 - CLIENT_MARGIN / 2
            ptTmp(1).y = CLIENT_MARGIN
            
            ptTmp(2).x = m_lClientWidth * 0.35 - CLIENT_MARGIN / 2
            ptTmp(2).y = m_lClientHeight - CLIENT_MARGIN * 2
            
            ptTmp(3).x = m_lClientWidth * 0.1 + CLIENT_MARGIN / 2
            ptTmp(3).y = m_lClientHeight - CLIENT_MARGIN * 2
            
            m_hRegion(0) = CreatePolygonRgn(ptTmp(0), _
                                            UBound(ptTmp) + 1, _
                                            2)
            
            ptTmp(0).x = m_lClientWidth * 0.65 + CLIENT_MARGIN / 2
            ptTmp(0).y = CLIENT_MARGIN
            
            ptTmp(1).x = m_lClientWidth * 0.9 - CLIENT_MARGIN / 2
            ptTmp(1).y = CLIENT_MARGIN
            
            ptTmp(2).x = m_lClientWidth - CLIENT_MARGIN * 2
            ptTmp(2).y = m_lClientHeight - CLIENT_MARGIN * 2
            
            ptTmp(3).x = m_lClientWidth * 0.35 + CLIENT_MARGIN / 2
            ptTmp(3).y = m_lClientHeight - CLIENT_MARGIN * 2
            
            m_hRegion(1) = CreatePolygonRgn(ptTmp(0), _
                                            UBound(ptTmp) + 1, _
                                            2)
            
            
        Case 3
            ' lettere "MV"
            m_lNumRegions = 2
            
            SetRect rtTmp, 0, 0, m_lClientWidth / 2, m_lClientHeight
            InflateRect rtTmp, -CLIENT_MARGIN, -CLIENT_MARGIN
            
            ReDim ptTmp(11) As POINTAPI
            
            ptTmp(0).x = rtTmp.Left
            ptTmp(0).y = rtTmp.Top
            
            ptTmp(1).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.25
            ptTmp(1).y = rtTmp.Top
            
            ptTmp(2).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.5
            ptTmp(2).y = rtTmp.Top + (rtTmp.Bottom - rtTmp.Top) * 0.25
            
            ptTmp(3).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.75
            ptTmp(3).y = rtTmp.Top
            
            ptTmp(4).x = rtTmp.Right
            ptTmp(4).y = rtTmp.Top
            
            ptTmp(5).x = rtTmp.Right
            ptTmp(5).y = rtTmp.Bottom
            
            ptTmp(6).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.75
            ptTmp(6).y = rtTmp.Bottom
            
            ptTmp(7).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.75
            ptTmp(7).y = rtTmp.Top + (rtTmp.Bottom - rtTmp.Top) * 0.5
            
            ptTmp(8).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.5
            ptTmp(8).y = rtTmp.Bottom
            
            ptTmp(9).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.25
            ptTmp(9).y = rtTmp.Top + (rtTmp.Bottom - rtTmp.Top) * 0.5
            
            ptTmp(10).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.25
            ptTmp(10).y = rtTmp.Bottom
            
            ptTmp(11).x = rtTmp.Left
            ptTmp(11).y = rtTmp.Bottom
            
            m_hRegion(0) = CreatePolygonRgn(ptTmp(0), _
                                            UBound(ptTmp) + 1, _
                                            2)
            
            SetRect rtTmp, m_lClientWidth / 2, 0, m_lClientWidth, m_lClientHeight
            InflateRect rtTmp, -CLIENT_MARGIN, -CLIENT_MARGIN
            
            ReDim ptTmp(6) As POINTAPI
            
            ptTmp(0).x = rtTmp.Left
            ptTmp(0).y = rtTmp.Top
            
            ptTmp(1).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.25
            ptTmp(1).y = rtTmp.Top
            
            ptTmp(2).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.5
            ptTmp(2).y = rtTmp.Top + (rtTmp.Bottom - rtTmp.Top) * 0.5
            
            ptTmp(3).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.75
            ptTmp(3).y = rtTmp.Top
            
            ptTmp(4).x = rtTmp.Right
            ptTmp(4).y = rtTmp.Top
            
            ptTmp(5).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.625
            ptTmp(5).y = rtTmp.Bottom
            
            ptTmp(6).x = rtTmp.Left + (rtTmp.Right - rtTmp.Left) * 0.375
            ptTmp(6).y = rtTmp.Bottom
            
            m_hRegion(1) = CreatePolygonRgn(ptTmp(0), _
                                            UBound(ptTmp) + 1, _
                                            2)
    End Select
    
    
    ' Rendering regioni
    For lIdx = 0 To m_lNumRegions - 1
        ' Sfondo bianco
        lhBrush = CreateSolidBrush(vbWhite)
        FillRgn Me.hdc, m_hRegion(lIdx), lhBrush
        DeleteObject lhBrush
        
        ' Contorno nero
        lhBrush = CreateSolidBrush(vbBlack)
        FrameRgn Me.hdc, m_hRegion(lIdx), lhBrush, 1, 1
        DeleteObject lhBrush
    Next
End Sub


Private Sub FlushRegions()
    ' Distruzione oggetti istanziati
Dim lIdx As Long
    
    For lIdx = 0 To UBound(m_hRegion)
        If m_hRegion(lIdx) Then _
            DeleteObject m_hRegion(lIdx)
        m_hRegion(lIdx) = 0
    Next
End Sub


Private Function GetRegionBox(ByVal lRgnIdx As Long, ByRef rtBox As T_BOX) As Boolean
Dim rtTmp As RECT

    If lRgnIdx < m_lNumRegions Then
        If m_hRegion(lRgnIdx) Then
            GetRgnBox m_hRegion(lRgnIdx), rtTmp
            
            InflateRect rtTmp, -m_lMargin, -m_lMargin
            
            rtBox.Left = rtTmp.Left
            rtBox.Top = rtTmp.Top
            rtBox.Width = rtTmp.Right - rtTmp.Left
            rtBox.Height = rtTmp.Bottom - rtTmp.Top
            
            GetRegionBox = True
        End If
    End If
End Function


Private Sub RegionSegment(ByVal lRgnIdx As Long, ByRef rRow As T_ROW)
    ' Restituisce un vettore di segmenti individuati alla coordinata Y
Dim lXPos As Long
Dim rtTmp As RECT
Dim bInner As Boolean
    
    GetRgnBox m_hRegion(lRgnIdx), rtTmp
        
    rRow.lNumSegms = 0
    ReDim rRow.rSegm(0)
    
    For lXPos = rtTmp.Left To rtTmp.Right
        If PtInRegion(m_hRegion(lRgnIdx), lXPos, rRow.lTop + rRow.lMaxHeight / 2) Then
            If bInner Then
                rRow.rSegm(rRow.lNumSegms - 1).lMaxX = lXPos - m_lMargin
            Else
                ReDim Preserve rRow.rSegm(rRow.lNumSegms)
                rRow.rSegm(rRow.lNumSegms).lMinX = lXPos + m_lMargin
                rRow.lNumSegms = rRow.lNumSegms + 1
            
                bInner = True
            End If
        Else
            bInner = False
        End If
    Next
End Sub



Private Sub Solve()
    ' Risolve il problema
    m_bContinueToNextBox = CBool(Check1.Value)
    
    Label3.Caption = "Words=" & m_lNumWords
    
    If m_lNumWords = 0 Then _
        Exit Sub
    
Dim lIdx As Long
Dim lTotLen As Long
Dim lTmp As Long

Dim lWordIdx As Long
Dim lStartIdx As Long
Dim lDirection As Long
Dim lLastBoxIdx As Long
Dim bContinueToNextBox As Boolean

Dim fpRefX As Single
Dim fpRefY As Single
Dim fpSpaceWidth As Single
Dim bInsertable As Boolean

Dim fpRowDistance As Single
Dim lFirstRowMaxHeight As Long

Dim rRow(MAX_ROWS_IN_BOX) As T_ROW 'la cella zero non e' mai usata
Dim lRowsUsed As Long
Dim lCurrRow As Long

Dim lUpperRowNum As Long
Dim lLowerRowNum As Long
Dim lCurrRowNum As Long
Dim lPrevRowNum As Long
Dim lRowCmd As Long

Dim lUpperSegmIdx As Long
Dim lLowerSegmIdx As Long
Dim lSegm As Long

Dim lVectIdx() As Long
Dim lPtrIx As Long
Dim lTotWords As Long

Dim lEsito(-1 To 1) As Long
Dim bOverflow(-1 To 1) As Boolean
Dim bEOF As Boolean

Dim bSafeMarker(MAX_ROWS_IN_BOX) As Boolean

Dim lRegionIdx As Long
Dim rtRgnBox As T_BOX


    ' Calcola lunghezza totale
    For lIdx = 0 To m_lNumWords - 1
        lTotLen = lTotLen + m_rWord(lIdx).rtBox.Width
    Next
    
    
    Select Case m_lAlignY
        Case ALIGN_NEAR, ALIGN_JUSTIFY
            bContinueToNextBox = m_bContinueToNextBox
    End Select
    
    
    
Label_Next_Region:
    Erase rRow
    Erase bSafeMarker

    If Not GetRegionBox(lRegionIdx, rtRgnBox) Then
        bOverflow(-1) = True
        GoTo Label_Check_Overflow
    End If
    
    
    ' Stabilisce punto di riferimento VERT
    Select Case m_lAlignY
        Case ALIGN_MIDDLE
            fpRefY = rtRgnBox.Top + rtRgnBox.Height / 2
        
            ' cerca indice della parola corispondente alla lunghezza media
            lTmp = 0
            For lIdx = 0 To m_lNumWords - 1
                lTmp = lTmp + m_rWord(lIdx).rtBox.Width
                If lTmp > (lTotLen / 2) Then
                    lStartIdx = lIdx
                    Exit For
                End If
            Next
            
        Case ALIGN_NEAR, ALIGN_JUSTIFY
            fpRefY = rtRgnBox.Top
            lDirection = 1
            
        Case ALIGN_FAR
            fpRefY = rtRgnBox.Top + rtRgnBox.Height
            lDirection = -1
    End Select
    
            
            
    ' Inserimento delle parole
    Select Case m_lAlignY
        Case ALIGN_MIDDLE
            Do
                lDirection = -1
                lWordIdx = lStartIdx
                bSafeMarker(lStartIdx) = True
                GoSub Label_Reset_Rows
            
                ' inserimento parole
                For lIdx = 0 To (m_lNumWords * 2)
                    lWordIdx = lWordIdx + lDirection * lIdx
                
                    If lWordIdx >= 0 And lWordIdx < m_lNumWords Then
Label_Reinsert_Word:
                        If lEsito(lDirection) Then
                            lRowCmd = 10 * lDirection
                        Else
                            lRowCmd = lDirection
                        End If
                        
                        GoSub Label_Insert_Word
                        
                        Select Case lEsito(lDirection)
                            Case &H100 'necessario ricalcolo
                                Exit For
                            Case &H200 'overflow
                                bOverflow(lDirection) = True
                            Case 1 'aggiungere riga
                                GoTo Label_Reinsert_Word
                        End Select
                    Else
                        lEsito(lDirection) = 0
                    End If
                    
                    lDirection = -lDirection
                Next
                
                ' Controllo overflow upper/lower
                If bOverflow(-1) And bOverflow(1) Then
                    Exit Do
                ElseIf bOverflow(-1) Then
                    lStartIdx = lStartIdx - 1
                    If lStartIdx < 0 Then
                        bOverflow(1) = True
                        Exit Do
                    ElseIf bSafeMarker(lStartIdx) Then
                        bOverflow(1) = True 'evita lo stallo
                        Exit Do
                    Else
                        lEsito(-1) = 1 'serve solo per far ciclare il "do"
                    End If
                    
                ElseIf bOverflow(1) Then
                    lStartIdx = lStartIdx + 1
                    If lStartIdx >= m_lNumWords Then
                        bOverflow(-1) = True
                        Exit Do
                    ElseIf bSafeMarker(lStartIdx) Then
                        bOverflow(-1) = True 'evita lo stallo
                        Exit Do
                    Else
                        lEsito(1) = 1 'serve solo per far ciclare il "do"
                    End If
                End If
                
            Loop While lEsito(lDirection) Or lEsito(-lDirection)
            
            If Not bOverflow(-1) And Not bOverflow(1) Then
                bEOF = True
                lRowCmd = lDirection
                GoSub Label_Insert_Word
                
                lRowCmd = -lDirection
                GoSub Label_Insert_Word
            End If
            
            
        Case ALIGN_NEAR, ALIGN_JUSTIFY, ALIGN_FAR
            Do
                GoSub Label_Reset_Rows
            
                ' inserimento parole
                For lIdx = lLastBoxIdx To m_lNumWords - 1
                    Select Case m_lAlignY
                        Case ALIGN_MIDDLE
                        
                        Case ALIGN_NEAR, ALIGN_JUSTIFY
                            lWordIdx = lIdx
                        Case ALIGN_FAR
                            lWordIdx = (m_lNumWords - 1) - lIdx
                    End Select
                    
                    If lEsito(lDirection) Then
                        lRowCmd = 10 * lDirection
                    Else
                        lRowCmd = lDirection
                    End If
                    
                    GoSub Label_Insert_Word
                    
                    Select Case lEsito(lDirection)
                        Case &H100 'necessario ricalcolo
                            Exit For
                        Case &H200 'overflow
                            bOverflow(lDirection) = True
                            Exit For
                        Case 1 'aggiungere riga
                            lIdx = lIdx - 1
                    End Select
                Next
                
                If lIdx = m_lNumWords Then _
                    bContinueToNextBox = False
                
                If bOverflow(lDirection) And bContinueToNextBox Then
                    ' Continua nella prossima regione
                    lLastBoxIdx = lIdx
                    bOverflow(lDirection) = False
                    Exit Do
                End If
            Loop While (lEsito(lDirection) <> 0) And Not bOverflow(lDirection)
            
            If Not bOverflow(lDirection) Then
                ' Aggiustamento segmenti toccati per ultimi
                bEOF = True
                lRowCmd = lDirection
                GoSub Label_Insert_Word
                
                
                If m_lAlignY = ALIGN_JUSTIFY And lRowsUsed > 1 Then
                    ' Distribuisce righe sull'intera altezza
                    For lTmp = (lRowsUsed - 1) To 0 Step -1
                        If rRow(lTmp).lNumber = lUpperRowNum Then _
                            Exit For
                    Next
                    
                    lFirstRowMaxHeight = rRow(lTmp).lMaxHeight
                    fpRowDistance = (rtRgnBox.Height - lFirstRowMaxHeight) / (lRowsUsed - 1)
                    
                    For lCurrRowNum = (lUpperRowNum + 1) To lLowerRowNum
                        For lTmp = (lRowsUsed - 1) To 0 Step -1
                            If rRow(lTmp).lNumber = lCurrRowNum Then _
                                Exit For
                        Next
                        
                        For lSegm = 0 To rRow(lTmp).lNumSegms - 1
                            For lIdx = 0 To rRow(lTmp).rSegm(lSegm).lNumWords - 1
                                lPtrIx = rRow(lTmp).rSegm(lSegm).lWordIdx(lIdx)
                                
                                m_rWord(lPtrIx).rtBox.Top = fpRefY + lFirstRowMaxHeight + fpRowDistance * (lCurrRowNum - lUpperRowNum) - m_rWord(lPtrIx).rtBox.Height
                            Next
                        Next
                    Next
                End If
            End If
    
            If bContinueToNextBox Then
                lRegionIdx = lRegionIdx + 1
                GoTo Label_Next_Region
            End If
    End Select
            
    
Label_Check_Overflow:
    If bOverflow(-1) Or bOverflow(1) Then
        bOverflow(0) = True
    Else
        ' Rivelatore di overflow
        lTmp = 0
        For lIdx = 0 To lRowsUsed - 1
            lTmp = lTmp + rRow(lIdx).lMaxHeight
            If lIdx Then _
                lTmp = lTmp + LINE_SPACING
        Next
        
        bOverflow(0) = (lTmp > rtRgnBox.Height)
    End If
    
    
    ' Posizionamento parole
    For lIdx = 0 To Label1.Count - 1
        Label1(lIdx).Move m_rWord(lIdx).rtBox.Left, _
                            m_rWord(lIdx).rtBox.Top
        Label1(lIdx).Visible = (lIdx < m_lNumWords)
    Next
    
    Label4.Visible = bOverflow(0)
    Exit Sub
    
    
    
Label_Reset_Rows:
    lEsito(lDirection) = 0
    lEsito(-lDirection) = 0

    bOverflow(-1) = False
    bOverflow(1) = False

    lLowerRowNum = 1000
    lUpperRowNum = lLowerRowNum
    lCurrRowNum = lLowerRowNum
    
    bEOF = False
    
    ' azzera uso righe, ma non altezza max
    lRowsUsed = 0
    For lIdx = 0 To UBound(rRow)
        rRow(lIdx).lNumber = 0
        rRow(lIdx).lTop = 0
        rRow(lIdx).lNumSegms = 0
    Next
    
    Return
    
    
    
Label_Insert_Word:
    lEsito(lDirection) = 0
    
    Select Case lRowCmd
        Case -10 'aggiunge verso l'alto
            lUpperRowNum = lUpperRowNum - 1
            
        Case 10 'aggiunge verso il basso
            lLowerRowNum = lLowerRowNum + 1
    End Select
    
    
    If lRowCmd < 0 Then
        lPrevRowNum = lUpperRowNum + 1
        lCurrRowNum = lUpperRowNum
        lSegm = lUpperSegmIdx
    ElseIf lRowCmd > 0 Then
        lPrevRowNum = lLowerRowNum - 1
        lCurrRowNum = lLowerRowNum
        lSegm = lLowerSegmIdx
    End If
        
    
    For lTmp = (lRowsUsed - 1) To 0 Step -1
        If rRow(lTmp).lNumber = lCurrRowNum Then _
            Exit For
    Next
    
    If lTmp < 0 Then
        ' Crea una nuova riga
        lCurrRow = lRowsUsed
        lRowsUsed = lRowsUsed + 1
        
        rRow(lCurrRow).lNumber = lCurrRowNum
        If rRow(lCurrRow).lMaxHeight < m_rWord(lWordIdx).rtBox.Height Then _
            rRow(lCurrRow).lMaxHeight = m_rWord(lWordIdx).rtBox.Height
    
        For lTmp = (lRowsUsed - 2) To 0 Step -1
            If rRow(lTmp).lNumber = lPrevRowNum Then _
                Exit For
        Next
        
        If lTmp < 0 Then
            If lRowCmd > 0 Then
                rRow(lCurrRow).lTop = fpRefY
            ElseIf lRowCmd < 0 Then
                rRow(lCurrRow).lTop = fpRefY - rRow(lCurrRow).lMaxHeight
            End If
        Else
            Select Case lRowCmd
                Case -10 'aggiunge verso l'alto
                    rRow(lCurrRow).lTop = rRow(lTmp).lTop - rRow(lCurrRow).lMaxHeight - LINE_SPACING
                    If rRow(lCurrRow).lTop < 0 Then
                        lEsito(lDirection) = &H200 'overflow verso l'alto
                        Return
                    End If
                    
                Case 10 'aggiunge verso il basso
                    rRow(lCurrRow).lTop = rRow(lTmp).lTop + rRow(lTmp).lMaxHeight + LINE_SPACING
                    If (rRow(lCurrRow).lTop + rRow(lTmp).lMaxHeight) > rtRgnBox.Height Then
                        lEsito(lDirection) = &H200 'overflow verso il basso
                        Return
                    End If
            End Select
        End If
        
        RegionSegment lRegionIdx, rRow(lCurrRow)


        ' Trova il segmento piu' opportuno
        Select Case m_lAlignX
            Case ALIGN_MIDDLE
            
            Case ALIGN_NEAR, ALIGN_JUSTIFY
                lSegm = 0
                
            Case ALIGN_FAR
                lSegm = rRow(lCurrRow).lNumSegms - 1
        End Select
    Else
        lCurrRow = lTmp
    End If

    If bEOF Then _
        GoTo Label_Adjust_Segment
    

Label_Next_Segment:
    If lSegm >= rRow(lCurrRow).lNumSegms Then _
        lSegm = -1

    If lSegm < 0 Then
        'esaurito estremo del segmento
        lEsito(lDirection) = 1
        Return
    End If
    
    
    If lRowCmd < 0 Then
        lUpperSegmIdx = lSegm
    ElseIf lRowCmd > 0 Then
        lLowerSegmIdx = lSegm
    End If


    ' Valuta occupazione
    With rRow(lCurrRow).rSegm(lSegm)
        If .lTotWidth Then
            bInsertable = (.lTotWidth + m_rWord(lWordIdx).rtBox.Width + SPACE_WIDTH) <= (.lMaxX - .lMinX)
        Else
            bInsertable = (m_rWord(lWordIdx).rtBox.Width) <= (.lMaxX - .lMinX)
        End If
    End With
    
    If bInsertable Then
        ' Aggiunge al segmento corrente
        With rRow(lCurrRow).rSegm(lSegm)
            If .lTotWidth Then _
                .lTotWidth = .lTotWidth + SPACE_WIDTH
            .lTotWidth = .lTotWidth + m_rWord(lWordIdx).rtBox.Width
            
            ReDim Preserve .lWordIdx(.lNumWords)
            .lWordIdx(.lNumWords) = lWordIdx
            .lNumWords = .lNumWords + 1
        End With
        
        If rRow(lCurrRow).lMaxHeight < m_rWord(lWordIdx).rtBox.Height Then
            rRow(lCurrRow).lMaxHeight = m_rWord(lWordIdx).rtBox.Height
            lEsito(lDirection) = &H100 'ripetere tutto
            Return
        End If
    
    Else
        ' Non ci sta piu' nel segmento corrente
        GoSub Label_Adjust_Segment
        
        Select Case lRowCmd
            Case Is < 0
                lSegm = lSegm - 1
            
            Case Is > 0
                lSegm = lSegm + 1
        End Select
        
        GoTo Label_Next_Segment
    End If
    
    Return
    
    
    
Label_Adjust_Segment:
    ' Completa allineamento orizzontale per tutti i segmenti della riga
    With rRow(lCurrRow)
        If .lNumSegms Then
            If .rSegm(lSegm).lNumWords Then
                ShellSort lVectIdx, _
                            .rSegm(lSegm).lWordIdx, _
                            .rSegm(lSegm).lNumWords, _
                            False
                fpSpaceWidth = SPACE_WIDTH
                
                Select Case m_lAlignX
                    Case ALIGN_MIDDLE
                        fpRefX = (.rSegm(lSegm).lMaxX + .rSegm(lSegm).lMinX - .rSegm(lSegm).lTotWidth) / 2
                        For lTmp = 0 To (.rSegm(lSegm).lNumWords - 1)
                            lPtrIx = .rSegm(lSegm).lWordIdx(lVectIdx(lTmp))
                            m_rWord(lPtrIx).rtBox.Left = fpRefX
                            m_rWord(lPtrIx).rtBox.Top = .lTop + .lMaxHeight - m_rWord(lPtrIx).rtBox.Height
                            fpRefX = fpRefX + m_rWord(lPtrIx).rtBox.Width + fpSpaceWidth
                        Next
                        
                    Case ALIGN_NEAR
                        fpRefX = .rSegm(lSegm).lMinX
                        For lTmp = 0 To (.rSegm(lSegm).lNumWords - 1)
                            lPtrIx = .rSegm(lSegm).lWordIdx(lVectIdx(lTmp))
                            m_rWord(lPtrIx).rtBox.Left = fpRefX
                            m_rWord(lPtrIx).rtBox.Top = .lTop + .lMaxHeight - m_rWord(lPtrIx).rtBox.Height
                            fpRefX = fpRefX + m_rWord(lPtrIx).rtBox.Width + fpSpaceWidth
                        Next
                        
                    Case ALIGN_FAR
                        fpRefX = .rSegm(lSegm).lMaxX
                        For lTmp = (.rSegm(lSegm).lNumWords - 1) To 0 Step -1
                            lPtrIx = .rSegm(lSegm).lWordIdx(lVectIdx(lTmp))
                            fpRefX = fpRefX - m_rWord(lPtrIx).rtBox.Width
                            m_rWord(lPtrIx).rtBox.Left = fpRefX
                            m_rWord(lPtrIx).rtBox.Top = .lTop + .lMaxHeight - m_rWord(lPtrIx).rtBox.Height
                            fpRefX = fpRefX - fpSpaceWidth
                        Next
                        
                    Case ALIGN_JUSTIFY
                        fpRefX = .rSegm(lSegm).lMinX
                        
                        If (.rSegm(lSegm).lNumWords > 1) And Not bEOF Then _
                            fpSpaceWidth = (.rSegm(lSegm).lMaxX - .rSegm(lSegm).lMinX - .rSegm(lSegm).lTotWidth + (.rSegm(lSegm).lNumWords - 1) * SPACE_WIDTH) / (.rSegm(lSegm).lNumWords - 1)
                        
                        For lTmp = 0 To (.rSegm(lSegm).lNumWords - 1)
                            lPtrIx = .rSegm(lSegm).lWordIdx(lVectIdx(lTmp))
                            m_rWord(lPtrIx).rtBox.Left = fpRefX
                            m_rWord(lPtrIx).rtBox.Top = .lTop + .lMaxHeight - m_rWord(lPtrIx).rtBox.Height
                            fpRefX = fpRefX + m_rWord(lPtrIx).rtBox.Width
                            fpRefX = fpRefX + fpSpaceWidth
                        Next
                End Select
            End If
        End If
    End With
    
    Return
End Sub


Private Sub ShellSort(ByRef lVect() As Long, _
                        ByRef lValue() As Long, _
                        ByVal lCount As Long, _
                        ByVal bDescending As Boolean _
                        )
    ' Riordino tramite shell-sort
Dim lInc As Long, lIdx1 As Long, lIdx2 As Long
Dim lTempVect As Long
Dim lTemp1 As Long, sTemp1 As String
Dim lTemp2 As Long, sTemp2 As String
    
    ReDim lVect(lCount - 1) As Long
    
    For lIdx1 = 0 To lCount - 1
        lVect(lIdx1) = lIdx1
    Next
    
    If lCount < 2 Then _
        Exit Sub
    
    
    lInc = 1
    Do
        lInc = 3 * lInc + 1
    Loop Until lInc > (UBound(lVect) + 1)
    
    Do
        lInc = lInc \ 3
        
        For lIdx1 = lInc To UBound(lVect)
            lIdx2 = lIdx1
            
            lTempVect = lVect(lIdx1)
            lTemp1 = lValue(lTempVect)
            
            Do
                If bDescending Then
                    If lValue(lVect(lIdx2 - lInc)) >= lTemp1 Then _
                        Exit Do
                Else
                    If lValue(lVect(lIdx2 - lInc)) <= lTemp1 Then _
                        Exit Do
                End If
                
                lVect(lIdx2) = lVect(lIdx2 - lInc)
                lIdx2 = lIdx2 - lInc
            Loop Until lIdx2 < lInc
            
            lVect(lIdx2) = lTempVect
        Next
    Loop Until lInc = 1
End Sub


Private Sub Form_Load()
Dim lIdx As Long

    Me.BackColor = QBColor(8)
    Me.ScaleMode = vbPixels

    For lIdx = 0 To 255
        m_rWord(lIdx).rtBox.Width = 8 + 100 * Rnd
        m_rWord(lIdx).rtBox.Height = 16 + IIf(Rnd > 0.9, 8, 0)
        
        If lIdx Then _
            Load Label1(lIdx)
        Label1(lIdx).Caption = lIdx
        Label1(lIdx).BackColor = vbYellow
        Label1(lIdx).Width = m_rWord(lIdx).rtBox.Width
        Label1(lIdx).Height = m_rWord(lIdx).rtBox.Height
        Label1(lIdx).ToolTipText = "Size=" & m_rWord(lIdx).rtBox.Width & "x" & m_rWord(lIdx).rtBox.Height & " px"
        Label1(lIdx).Visible = False
    Next
    
    
    List1(0).AddItem "Middle"
    List1(0).AddItem "Near"
    List1(0).AddItem "Far"
    List1(0).AddItem "Justify"
    List1(0).ListIndex = 1
    
    
    List1(1).AddItem "Middle"
    List1(1).AddItem "Near"
    List1(1).AddItem "Far"
    List1(1).AddItem "Justify"
    List1(1).ListIndex = 1
    
    
    List2.AddItem "double rect 60/40"
    List2.AddItem "ellipse"
    List2.AddItem "double polygon"
    List2.AddItem "'MV' letters"
    List2.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FlushRegions
End Sub

Private Sub Form_Resize()

    m_lClientWidth = Me.ScaleWidth
    If m_lClientWidth < 200 Then _
        m_lClientWidth = 200
        
    m_lClientHeight = Me.ScaleHeight
    If m_lClientHeight < 300 Then _
        m_lClientHeight = 300

    Frame1.Top = m_lClientHeight - Frame1.Height
    Frame1.Width = m_lClientWidth
    
    m_lClientHeight = m_lClientHeight - Frame1.Height
    
    DefineRegions
    
    Solve
End Sub

Private Sub HScroll1_Change()
    m_lNumWords = HScroll1.Value
    Solve
End Sub

Private Sub HScroll2_Change()
    Form_Resize
End Sub

Private Sub List1_Click(Index As Integer)
    ' Decide allineamento
    If Index Then
        m_lAlignY = List1(Index).ListIndex
    Else
        m_lAlignX = List1(Index).ListIndex
    End If
    
    Solve
End Sub

Private Sub List2_Click()
    Form_Resize
End Sub

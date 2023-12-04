Attribute VB_Name = "ModImageLoader"
Option Explicit
'***************************************************************************************
'*                   Module openGL pour chargement de texture
'***************************************************************************************
' v0.1 - 04/11/2009
' v0.2 - 21/07/2014
'***************************************************************************************
' http://arkham46.developpez.com
'***************************************************************************************
#If VBA7 Then
DefLngPtr A-Z
Private Const PtrNull As LongPtr = 0
#Else
DefLng A-Z
Private Const PtrNull As Long = 0
#End If
#If VBA7 Then
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" _
                                          (ByVal hdc As LongPtr, pBitmapInfo As BitmapInfo, ByVal un As Long, _
                                           lplpVoid As LongPtr, ByVal handle As LongPtr, ByVal dw As Long) As LongPtr
Private Declare PtrSafe Function GetObjectBmp Lib "gdi32" Alias "GetObjectA" _
       (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (Destination As Any, source As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function GetDC Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "User32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" _
                                (ByVal destdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
                                 ByVal nHeight As Long, ByVal srcdc As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, _
                                 ByVal dwRop As Long) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, LInput As GdiplusStartupInput, Optional ByVal lOutPut As LongPtr = 0) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef Image As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageDimension Lib "gdiplus" _
                                               (ByVal Image As LongPtr, ByRef width As Single, _
                                                ByRef height As Single) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As LongPtr, Rect As Rect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As bitmapData) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As LongPtr, lockedBitmapData As bitmapData) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
Private Declare PtrSafe Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As LongPtr, ByVal rfType As Long) As Long
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As LongPtr
    SuppressExternalCodecs As LongPtr
End Type
#Else
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CreateDIBSection Lib "gdi32" _
                                          (ByVal hdc As Long, pBitmapInfo As BitmapInfo, ByVal un As Long, _
                                           lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetObjectBmp Lib "gdi32" Alias "GetObjectA" _
       (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, source As Any, ByVal length As Long)
Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" _
                                (ByVal destdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
                                 ByVal nHeight As Long, ByVal srcdc As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                 ByVal dwRop As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, LInput As GdiplusStartupInput, Optional ByVal lOutPut As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef Image As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" _
                                               (ByVal Image As Long, ByRef width As Single, _
                                                ByRef height As Single) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As Long, Rect As Rect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As bitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As Long, lockedBitmapData As bitmapData) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
#End If
Private Type Rect
    Left As Long
    Top As Long
    right As Long
    Bottom As Long
End Type
Private Type bitmapData
    width As Long
    height As Long
    stride As Long
    PixelFormat As Long
    #If VBA7 Then
    scan0 As LongPtr
    #Else
    scan0 As Long
    #End If
    Reserved As Long
End Type
Private Type bitmap
   bmType As Long
   bmWidth As Long: bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   #If VBA7 Then
    bmBits As LongPtr
   #Else
    bmBits As Long
   #End If
End Type
Private Type BitmapInfoHeader
    biSize As Long
    biWidth As Long: biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long: biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type DIBSECTION
    dsBm As bitmap
    dsBmih As BitmapInfoHeader
    dsBitfields(2) As Long
    #If VBA7 Then
    dshSection As LongPtr
    #Else
    dshSection As Long
    #End If
    dsOffset As Long
End Type
Private Type BitmapInfo
    bmiHeader As BitmapInfoHeader
    bmiColors(0 To 255) As Long
End Type

Public Type tpImage
    texID As Long
    lNbByte As Long
    texture As Long
    Image() As Byte
    width As Long
    height As Long
    format As Long
End Type

Private Const SRCCOPY = &HCC0020
Private Const DIB_RGB_COLORS As Long = &H0
Private Const BI_RGB As Long = &H0

Private gGdipToken     ' Identifiant pour GdiPlus

'---------------------------------------------------------------------------------------
' Charge une texture à partir d'un fichier image
'---------------------------------------------------------------------------------------
Public Function TextureAddFromFile(pFile As String, Optional pTransparentColor As Long = -1) As tpImage
Dim lImage As Object
Dim lTexture As Long
Dim lBitmap As bitmap
Dim lNbByte As Long
Dim lFormat As Long
Dim lBufferOut() As Byte
Dim lBufferIn() As Byte
Dim lX As Long, lY As Long
Dim lhDC
Dim lhDCIn, lOldBmpIn
Dim lhDCRef
Dim lDIBPTR
Dim lhDIB
Dim lhOldDIB
Dim lBI As BitmapInfo
    ' Charge l'image
    On Error Resume Next
    Set lImage = LoadPicture(pFile)
    On Error GoTo 0
    ' Si format non reconnu => utilise gdiplus
    If lImage Is Nothing Then
        TextureAddFromFile = TextureAddFromFileGdip(pFile, pTransparentColor)
        Exit Function
    End If
    ' Lecture des informations de l'image
    #If Win64 Then
    Call GetObjectBmp(lImage, LenB(lBitmap), lBitmap)
    #Else
    Call GetObjectBmp(lImage, Len(lBitmap), lBitmap)
    #End If
    ' Définition des paramètres fonction de la profondeur de couleur
    Select Case lBitmap.bmBitsPixel
        Case 24
            lNbByte = 3
            lFormat = GL_RGB
        Case 32
            lNbByte = 4
            lFormat = GL_RGBA
        Case Else
            ' Taille de l'en-tete
            #If Win64 Then
                lBI.bmiHeader.biSize = LenB(lBI.bmiHeader)
            #Else
                lBI.bmiHeader.biSize = Len(lBI.bmiHeader)
            #End If
            ' Taille de l'image = taille de l'objet
            lBI.bmiHeader.biWidth = lBitmap.bmWidth
            lBI.bmiHeader.biHeight = lBitmap.bmHeight
            ' Compression RGB (pas de compression en fait)
            lBI.bmiHeader.biCompression = BI_RGB
            ' 32bits
            lBI.bmiHeader.biBitCount = 32
            ' Nombre de plans : toujours = 1 (d'apres MSDN)
            lBI.bmiHeader.biPlanes = 1
            ' Calcul de la taille de l'image
            lBI.bmiHeader.biSizeImage = lBI.bmiHeader.biHeight * (((lBI.bmiHeader.biWidth * lBI.bmiHeader.biBitCount) + 31) \ 32) * 4
            ' DC de reference
            lhDCRef = GetDC(0)
            ' Device context
            lhDC = CreateCompatibleDC(lhDCRef)
            ' Libere le DC
            ReleaseDC 0, lhDCRef
            ' Creation d'une nouvelle section DIB qui va contenir le bitmap
            lhDIB = CreateDIBSection(lhDC, lBI, DIB_RGB_COLORS, lDIBPTR, 0, 0)
            ' On selectionne ce DIB dans le contexte d'affichage en prenant soin de conserver l'ancien DIB
            lhOldDIB = SelectObject(lhDC, lhDIB)
            ' Création device context pour image source et sélection du bitmap
            lhDCIn = CreateCompatibleDC(lhDC)
            lOldBmpIn = SelectObject(lhDCIn, lImage.handle)
            ' Dessine l'image source dans le DIB 32 bits
            BitBlt lhDC, 0, 0, lBitmap.bmWidth, lBitmap.bmHeight, lhDCIn, 0, 0, SRCCOPY
            ' Supprime les objets qui ne sont plus utiles
            SelectObject lhDCIn, lOldBmpIn
            DeleteDC lhDCIn
            ' Pointeur des données
            lBitmap.bmBits = lDIBPTR
            ' Format d'image
            lNbByte = 4
            lFormat = GL_RGBA
    End Select
    
    ' Se uma cor de transparência for especificada
    If pTransparentColor <> -1 Then
        ' Ajoute un canal alpha si nécessaire, et rend transparent les textels de couleur pTransparentColor
        ReDim lBufferOut(1 To 4, 1 To lBitmap.bmWidth, 1 To lBitmap.bmHeight) As Byte
    Else
        ReDim lBufferOut(1 To lNbByte, 1 To lBitmap.bmWidth, 1 To lBitmap.bmHeight) As Byte
    End If
    ReDim lBufferIn(1 To lNbByte, 1 To lBitmap.bmWidth, 1 To lBitmap.bmHeight) As Byte
    RtlMoveMemory lBufferIn(1, 1, 1), ByVal lBitmap.bmBits, lNbByte * lBitmap.bmWidth * lBitmap.bmHeight
    ' Inverse les valeurs pour passer de BGR à RGB
    For lX = 1 To lBitmap.bmWidth
        For lY = 1 To lBitmap.bmHeight
            lBufferOut(1, lX, lY) = lBufferIn(3, lX, lY)
            lBufferOut(2, lX, lY) = lBufferIn(2, lX, lY)
            lBufferOut(3, lX, lY) = lBufferIn(1, lX, lY)
            If lNbByte >= 4 Then lBufferOut(4, lX, lY) = lBufferIn(4, lX, lY)
            ' Si une couleur de transparence est spécifiée
            If pTransparentColor <> -1 Then
                If lBufferIn(3, lX, lY) = 0 And lBufferIn(2, lX, lY) = 0 And lBufferIn(1, lX, lY) = 0 Then 'pTransparentColor
                    lBufferOut(4, lX, lY) = 0
                Else
                    lBufferOut(4, lX, lY) = &HFF
                End If
'                lBufferOut(1, lX, lY) = lBufferIn(1, lX, lY)
'                lBufferOut(2, lX, lY) = lBufferIn(2, lX, lY)
'                lBufferOut(3, lX, lY) = lBufferIn(3, lX, lY)
            End If
        Next
    Next

    ' Si une couleur de transparence est spécifiée
    If pTransparentColor <> -1 Then
        lNbByte = 4
        lFormat = GL_RGBA
    End If

    ' Libère les objets gdi
    If lBitmap.bmBitsPixel < 24 Then
        DeleteObject SelectObject(lhDC, lhOldDIB)
        DeleteDC lhDC
    End If
    ' Libère l'image
    TextureAddFromFile.Image = lBufferOut
    Set lImage = Nothing
    ' Renvoit le numéro de la texture créée
    TextureAddFromFile.texture = lTexture
    TextureAddFromFile.width = lBitmap.bmWidth
    TextureAddFromFile.height = lBitmap.bmHeight
    TextureAddFromFile.format = lFormat
    TextureAddFromFile.lNbByte = lNbByte
    
End Function

Private Function TextureAddFromFileGdip(pFile As String, Optional pTransparentColor As Long = -1) As tpImage
Dim lLib
Dim lGdiPSI As GdiplusStartupInput ' Structure d'initialisation de gdiplus
Dim lBitmap
Dim lrect As Rect
Dim lWidth As Single
Dim lHeight As Single
Dim bmpData As bitmapData
Dim lTexture As Long
Dim lX As Long, lY As Long
Const PixelFormat32bppARGB = &H26200A
    
    ' Charge la librarie gdiplus
    lLib = LoadLibrary("gdiplus.dll")
    lGdiPSI.GdiplusVersion = 1
    Call GdiplusStartup(gGdipToken, lGdiPSI)
    If (GdipLoadImageFromFile(StrPtr(pFile), lBitmap) = 0) Then
        GdipImageRotateFlip lBitmap, 6 ' Rotate180FlipX
        GdipGetImageDimension lBitmap, lWidth, lHeight
        lrect.right = lWidth
        lrect.Bottom = lHeight
        If GdipBitmapLockBits(lBitmap, lrect, &H1, PixelFormat32bppARGB, bmpData) = 0 Then
            ReDim lBufferIn(1 To 4, 1 To bmpData.width, 1 To bmpData.height) As Byte
            RtlMoveMemory lBufferIn(1, 1, 1), ByVal bmpData.scan0, bmpData.stride * bmpData.height
            Call GdipBitmapUnlockBits(lBitmap, bmpData)
            ' Inverse les valeurs pour passer de BGRA à RGBA
            ReDim lBufferOut(1 To 4, 1 To bmpData.width, 1 To bmpData.height) As Byte
            For lX = 1 To bmpData.width
                For lY = 1 To bmpData.height
                    lBufferOut(1, lX, lY) = lBufferIn(3, lX, lY)
                    lBufferOut(2, lX, lY) = lBufferIn(2, lX, lY)
                    lBufferOut(3, lX, lY) = lBufferIn(1, lX, lY)
                    lBufferOut(4, lX, lY) = lBufferIn(4, lX, lY)
                    ' Si une couleur de transparence est spécifiée
                    If pTransparentColor <> -1 Then
                        If RGB(lBufferIn(3, lX, lY), lBufferIn(2, lX, lY), lBufferIn(1, lX, lY)) = pTransparentColor Then
                            lBufferOut(4, lX, lY) = 0
                        Else
                            lBufferOut(4, lX, lY) = 255
                        End If
                    End If
                Next
            Next
        End If
        If lBitmap <> PtrNull Then GdipDisposeImage lBitmap
    End If
    If gGdipToken <> 0 Then GdiplusShutdown gGdipToken: gGdipToken = 0
    
    ' Renvoit le numéro de la texture créée
    TextureAddFromFileGdip.texture = lTexture
    TextureAddFromFileGdip.Image = lBufferOut
    TextureAddFromFileGdip.width = lWidth
    TextureAddFromFileGdip.height = lHeight
    TextureAddFromFileGdip.format = GL_RGBA
    TextureAddFromFileGdip.lNbByte = 4
End Function


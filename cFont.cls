VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private base&
Private fontTexture&, lColor(2) As Single

Public Function Initialize()
    fontTexture = TextureCache.Images("fontTexture").texId
    BuildFont
    lColor(0) = 1: lColor(1) = 1: lColor(2) = 1
End Function

Public Function finalize()
    KillFont
End Function

Public Function SetColor(clr As Variant)
    lColor(0) = clr(0): lColor(1) = clr(1): lColor(2) = clr(2)
End Function

Private Function BuildFont()
Const factor! = 1 / 16
Dim cx#, cy#                                                '' Holds Our X and Y Character Coord
Dim i&

    base = glGenLists(256)                                  '' Creating 256 Display Lists
    glBindTexture GL_TEXTURE_2D, fontTexture                '' Select Our Font Texture

    For i = 0 To 255                                        '' Loop Through All 256 Lists
        cx = (i Mod 16) / 16#                               '' X Position Of Current Character
        cy = (i \ 16) / 16#                                 '' Y Position Of Current Character
        glNewList base + i, GL_COMPILE                      '' Start Building A List
            glBegin GL_QUADS                                '' Use A Quad For Each Character
                glTexCoord2f cx, 1 - cy - factor            '' Texture Coord (Bottom Left)
                glVertex2i 0, 0                             '' Vertex Coord (Bottom Left)
                glTexCoord2f cx + factor, 1 - cy - factor   '' Texture Coord (Bottom Right)
                glVertex2i 16, 0                            '' Vertex Coord (Bottom Right)
                glTexCoord2f cx + factor, 1 - cy            '' Texture Coord (Top Right)
                glVertex2i 16, 16                           '' Vertex Coord (Top Right)
                glTexCoord2f cx, 1 - cy                     '' Texture Coord (Top Left)
                glVertex2i 0, 16                            '' Vertex Coord (Top Left)
            glEnd
            glTranslated 8, 0, 0                            '' Move To The Right Of The Character
        glEndList
    Next i
	
	'A = 65
    'a = 97
    'a - A = 32

End Function

Private Function KillFont()
    glDeleteLists base, 256
End Function

Public Sub glPrint(s As String, sset As Integer, X As Single, Y As Single, Optional scal! = 1)
Dim b() As Byte, i&, length#

    ReDim b(Len(s))
    If (sset > 1) Then sset = 1
    glBindTexture GL_TEXTURE_2D, fontTexture
    glPushMatrix
        glLoadIdentity
        glTranslated X, Y, 0
        glListBase base - 32 + 128 * sset
        For i = 1 To Len(s)
            b(i - 1) = Asc(Mid$(s, i, 1))
        Next
        b(Len(s)) = 0
        Call glScalef(scal, scal, scal)
        Call glColor3fv(lColor(0))
        glCallLists Len(s), GL_BYTE, b(0)
        Call glColor3f(1, 1, 1)
    glPopMatrix
End Sub

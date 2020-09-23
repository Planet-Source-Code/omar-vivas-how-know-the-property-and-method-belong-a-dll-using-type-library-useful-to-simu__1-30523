Attribute VB_Name = "BrowseFolder"
'
'//////////////////////////////////////////////////////////////////////////////
'/////               ESTE CÓDIGO INSERTALO EN UN MÓDULO BAS               /////
'//////////////////////////////////////////////////////////////////////////////
'
'------------------------------------------------------------------------------
' Módulo con las declaraciones y funciones para BrowseForFolder     (12/May/99)
'
' ©Guillermo 'guille' Som, 1999
'------------------------------------------------------------------------------
Option Explicit

'//////////////////////////////////////////////////////////////////////////////
' Variables, constantes y funciones para usar con BrowseForFolder   (12/May/99)
'//////////////////////////////////////////////////////////////////////////////
'
Private sFolderIni As String
'
Private Const WM_USER = &H400&
Public Const MAX_PATH = 260&
'
' Tipo para usar con SHBrowseForFolder
Private Type BrowseInfo
    hWndOwner               As Long             ' hWnd del formulario
    pIDLRoot                As Long             ' Especifica el pID de la carpeta inicial
    pszDisplayName          As String           ' Nombre del item seleccionado
    lpszTitle               As String           ' Título a mostrar encima del árbol
    ulFlags                 As Long             '
    lpfnCallback            As Long             ' Función CallBack
    lParam                  As Long             ' Información extra a pasar a la función Callback
    iImage                  As Long             '
End Type
'
'// Browsing for directory.
Public Const BIF_RETURNONLYFSDIRS = &H1&       '// For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN = &H2&      '// For starting the Find Computer
Public Const BIF_STATUSTEXT = &H4&
Public Const BIF_RETURNFSANCESTORS = &H8&
Public Const BIF_EDITBOX = &H10&
Public Const BIF_VALIDATE = &H20&              '// insist on valid result (or CANCEL)
'
Public Const BIF_BROWSEFORCOMPUTER = &H1000&   '// Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000&    '// Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000&  '// Browsing for Everything
'
'// message from browser
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_VALIDATEFAILED = 3          '// lParam:szPath ret:1(cont),0(EndDialog)
'Public Const BFFM_VALIDATEFAILEDW = 4&         '// lParam:wzPath ret:1(cont),0(EndDialog)
'
'// messages to browser
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BFFM_ENABLEOK = (WM_USER + 101)
Public Const BFFM_SETSELECTION = (WM_USER + 102)
'Public Const BFFM_SETSELECTIONW = (WM_USER + 103&)
'Public Const BFFM_SETSTATUSTEXTW = (WM_USER + 104&)
'
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        (lpbi As BrowseInfo) As Long
'
Private Declare Sub CoTaskMemFree Lib "OLE32.DLL" _
        (ByVal hMem As Long)
'
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long


Public Function BrowseFolderCallbackProc(ByVal hWndOwner As Long, _
                                        ByVal uMSG As Long, _
                                        ByVal lParam As Long, _
                                        ByVal pData As Long) As Long
    ' Llamada CallBack para usar con la función BrowseForFolder     (12/May/99)
    Dim szDir As String

    On Local Error Resume Next

    Select Case uMSG
    '--------------------------------------------------------------------------
    ' Este mensaje se enviará cuando se inicia el diálogo,
    ' entonces es cuando hay que indicar el directorio de inicio.
    Case BFFM_INITIALIZED
        ' El path de inicio será el directorio indicado,
        ' si no se ha asignado, usar el directorio actual
        If Len(sFolderIni) Then
            szDir = sFolderIni & Chr$(0)
        Else
            szDir = CurDir$ & Chr$(0)
        End If
        ' WParam  será TRUE  si se especifica un path.
        '         será FALSE si se especifica un pIDL.
        Call SendMessage(hWndOwner, BFFM_SETSELECTION, 1&, ByVal szDir)
    '--------------------------------------------------------------------------
    ' Este mensaje se produce cuando se cambia el directorio
    ' Si nuestro form está subclasificado para recibir mensajes,
    ' puede interceptar el mensaje BFFM_SETSTATUSTEXT
    ' para mostrar el directorio que se está seleccionando.
    Case BFFM_SELCHANGED
        szDir = String$(MAX_PATH, 0)
        ' Notifica a la ventana del directorio actualmente seleccionado,
        ' (al menos en teoría, ya que no lo hace...)
        If SHGetPathFromIDList(lParam, szDir) Then
            'Debug.Print szDir
            Call SendMessage(hWndOwner, BFFM_SETSTATUSTEXT, 0&, ByVal szDir)
        End If
        Call CoTaskMemFree(lParam)
    End Select

    Err = 0
    BrowseFolderCallbackProc = 0

'------------------------------------------------------------------------------
' Este es el código de C en el que está basada esta función Callback
' Código obtenido de la MSDN Library de Microsoft:
' HOWTO: Browse for Folders from the Current Directory
' Article ID: Q179378
'
'         TCHAR szDir[MAX_PATH];
'
'         switch(uMsg) {
'            case BFFM_INITIALIZED: {
'               if GetCurrentDirectory(sizeof(szDir)/sizeof(TCHAR),
'                                      szDir)) {
'                  // WParam is TRUE since you are passing a path.
'                  // It would be FALSE if you were passing a pidl.
'                  SendMessage(hwnd,BFFM_SETSELECTION,TRUE,(LPARAM)szDir);
'               }
'               break;
'            }
'            case BFFM_SELCHANGED: {
'               // Set the status window to the currently selected path.
'               if (SHGetPathFromIDList((LPITEMIDLIST) lp ,szDir)) {
'                  SendMessage(hwnd,BFFM_SETSTATUSTEXT,0,(LPARAM)szDir);
'               }
'               break;
'            }
'           default:
'               break;
'         }
'         return 0;
'------------------------------------------------------------------------------
End Function


Public Function rtnAddressOf(lngProc As Long) As Long
    ' Devuelve la dirección pasada como parámetro
    ' Esto se usará para asignar a una variable la dirección de una función
    ' o procedimiento.
    ' Por ejemplo, si en un tipo definido se asigna a una variable la dirección
    ' de una función o procedimiento
    rtnAddressOf = lngProc
End Function


Public Function BrowseForFolder(ByVal hWndOwner As Long, ByVal sPrompt As String, _
                Optional sInitDir As String = "", _
                Optional ByVal lFlags As Long = BIF_RETURNONLYFSDIRS) As String
    ' Muestra el diálogo de selección de directorios de Windows
    ' Si todo va bien, devuelve el directorio seleccionado
    ' Si se cancela, se devuelve una cadena vacía y se produce el error 32755
    '
    ' Los parámetros de entrada:
    '   El hWnd de la ventana
    '   El título a mostrar
    '   Opcionalmente el directorio de inicio
    '   En lFlags se puede especificar lo que se podrá seleccionar:
    '       BIF_BROWSEINCLUDEFILES, etc.
    '       por defecto es: BIF_RETURNONLYFSDIRS
    '
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo

    On Local Error Resume Next

    With udtBI
        .hWndOwner = hWndOwner
        ' Título a mostrar encima del árbol de selección
        .lpszTitle = sPrompt & vbNullChar
        ' Que es lo que debe devolver esta función
        .ulFlags = lFlags
        '.ulFlags = lFlags Or BIF_RETURNONLYFSDIRS
        '
        ' Si se especifica el directorio por el que se empezará...
        If Len(sInitDir) Then
            ' Asignar la variable que contendrá el directorio de inicio
            sFolderIni = sInitDir
            ' Indicar la función Callback a usar.
            ' Como hay que asignar esa dirección a una variable,
            ' se usa una función "intermedia" que devuelve el valor
            ' del parámetro pasado... es decir: ¡la dirección de la función!
            .lpfnCallback = rtnAddressOf(AddressOf BrowseFolderCallbackProc)
        End If
    End With
    Err = 0
    On Local Error GoTo 0

    ' Mostramos el cuadro de diálogo
    lpIDList = SHBrowseForFolder(udtBI)
    '
    If lpIDList Then
        ' Si se ha seleccionado un directorio...
        '
        ' Obtener el path
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        ' Quitar los caracteres nulos del final
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    Else
        ' Si se ha pulsado en cancelar...
        '
        ' Devolver una cadena vacía y asignar un error
        sPath = ""
        With Err
            .Source = "MBrowseFolder::BrowseForFolder"
            .Number = 32755
            .Description = "Cancelada la operación de BrowseForFolder"
        End With
    End If

    BrowseForFolder = sPath
End Function



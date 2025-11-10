Option Explicit

'----------------------------------------------------------------------
' SCRIPT DE INVENTARIO DE SOFTWARE BÁSICO
' Revisa el registro de desinstalación para encontrar software.
' NOTA: "Instalado" no significa "En funcionamiento" o "Licenciado".
'----------------------------------------------------------------------

' --- DEFINE EL SOFTWARE A BUSCAR ---
Dim arrSoftwareList
arrSoftwareList = Array( _
    "Cisco Umbrella", _
    "CrowdStrike", _
    "Qualys", _
    "Ivanti", _
    "Microsoft 365", _
    "Office 365", _
    "Teams", _
    "ActivTrak", _
    "AutoElevate" _
)
' ----------------------------------


' --- INICIALIZACIÓN ---
Dim dictSoftware, strComputer, objReg, strReport, strSoftware
Set dictSoftware = CreateObject("Scripting.Dictionary")

' Poblar el diccionario con el estado por defecto
For Each strSoftware In arrSoftwareList
    If Not dictSoftware.Exists(LCase(strSoftware)) Then
        dictSoftware.Add LCase(strSoftware), "NO INSTALADO"
    End If
Next

strComputer = "." ' El punto significa "este equipo"

' Conectarse al proveedor de registro de WMI
On Error Resume Next
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
If Err.Number <> 0 Then
    WScript.Echo "No se pudo conectar al proveedor de registro WMI. Ejecuta el script como administrador."
    WScript.Quit
End If
On Error GoTo 0

' --- RUTAS DEL REGISTRO A REVISAR ---
Const HKLM = &H80000002
Const HKCU = &H80000001
Dim arrPaths(2)
arrPaths(0) = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
arrPaths(1) = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
arrPaths(2) = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" ' Para HKCU

' --- EJECUTAR BÚSQUEDA ---
CheckRegistryWMI objReg, HKLM, arrPaths(0), dictSoftware
CheckRegistryWMI objReg, HKLM, arrPaths(1), dictSoftware
CheckRegistryWMI objReg, HKCU, arrPaths(2), dictSoftware


' --- GENERAR REPORTE ---
strReport = "--- REPORTE DE SOFTWARE INSTALADO ---" & vbCrLf & vbCrLf

' Manejo especial para Office (puede ser M365 o O365)
If dictSoftware.Item("microsoft 365") = "INSTALADO" Or dictSoftware.Item("office 365") = "INSTALADO" Then
    strReport = strReport & "Paquete Office 365: INSTALADO" & vbCrLf
Else
    strReport = strReport & "Paquete Office 365: NO INSTALADO" & vbCrLf
End If

' Reportar el resto del software
For Each strSoftware In dictSoftware.Keys
    ' Evitar duplicar el reporte de Office
    If strSoftware <> "microsoft 365" And strSoftware <> "office 365" Then
        ' Capitalizar el nombre para el reporte
        strReport = strReport & Capitalize(strSoftware) & ": " & dictSoftware.Item(strSoftware) & vbCrLf
    End If
Next

' --- LÍNEA CORREGIDA ---
strReport = strReport & vbCrLf & "NOTA: 'INSTALADO' solo confirma que el programa" & vbCrLf
strReport = strReport & "existe en el registro. No significa que esta" & vbCrLf ' <-- Aquí estaba el error
strReport = strReport & "activo, licenciado o 'EN FUNCIONAMIENTO'."
' --- FIN DE LA CORRECCIÓN ---

' Mostrar el reporte final
WScript.Echo strReport

' --- Limpieza ---
Set dictSoftware = Nothing
Set objReg = Nothing

' --- FIN DEL SCRIPT ---


' --- SUB-RUTINAS ---

Sub CheckRegistryWMI(objReg, hKey, strKeyPath, dict)
    Dim arrSubKeys, strSubKey, strDisplayNamePath, strDisplayName, lcaseDisplayName, strSoftwareKey
    
    ' Enumerar todas las sub-claves de la ruta de desinstalación
    objReg.EnumKey hKey, strKeyPath, arrSubKeys
    
    If IsNull(arrSubKeys) Then Exit Sub
    
    On Error Resume Next
    For Each strSubKey In arrSubKeys
        strDisplayNamePath = strKeyPath & "\" & strSubKey
        
        ' Leer el valor "DisplayName" de cada sub-clave
        objReg.GetStringValue hKey, strDisplayNamePath, "DisplayName", strDisplayName
        
        If Err.Number = 0 And Not IsNull(strDisplayName) Then
            lcaseDisplayName = LCase(strDisplayName)
            
            ' Comparar el DisplayName con nuestra lista de software
            For Each strSoftwareKey In dict.Keys
                If InStr(lcaseDisplayName, strSoftwareKey) > 0 Then
                    dict.Item(strSoftwareKey) = "INSTALADO"
                End If
            Next
        End If
        Err.Clear
    Next
    On Error GoTo 0
End Sub

Function Capitalize(str)
    If Len(str) > 0 Then
        Capitalize = UCase(Left(str, 1)) & Mid(str, 2)
    Else
        Capitalize = ""
    End If
End Function

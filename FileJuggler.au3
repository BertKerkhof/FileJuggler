#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=FileJuggler.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Juggler file functions
#AutoIt3Wrapper_Res_Fileversion=1.3.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 2019-11-07 Apache 2.0 license
#AutoIt3Wrapper_Res_SaveSource=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/reel /tc 2
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /rm
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; The latest version of this source is published at GitHub in the
; BertKerkhof repository "FileJuggler".

#include-once
#include <aDrosteArray.au3>; GitHub published by Bert Kerkhof

; Author: kerkhof.bert@gmail.com
; Tested with AutoIT v3.3.14.5 interpreter/compiler and win10


; FileJuggler =========================================================

; Juggling with file name elements:
;   + FileName         : Trims location from full path/filename
;   + FileExt          : Get the extension from a filename
;   + FileBase         : Trim extension from file location/file name
;   + FileLocation     : Get file location from filename
;   + FileDrive        : Get drive from full path
;   + DriveGetVolume   : Get CD|DVD|Blu-Ray volume name
;   + FileRootFolder   : Get main folder from full path
;   + FileSnoopPath    : Removes the user folder from path
;   + PathRevision     : Creates a new filename to prevent overwriting
;   + FileRevNum       : Returns revision number from file name
;   + aRevFilter       : Removes revisions

; Open system files and Folders:
;   + aFilter          : Select files with specified title and extension
;   + aFilesFromFolder : Create array of filenames from folder
;   + aSelectFolder    : Open dialog for the user to select folder
;   + aSelectFiles     : Open dialog for the user to select files
;
; Other:
;   + CheckPresence    : Checks the presence of utilities
;   + FileMatch        : Checks whether a filename matches a file mask
;   + DirTreeMove      : Moves files and directories
;
; Test and example(s):
;   + TestDialog       : Demonstrates use of SelectFolders and SelectFiles


; Juggle with path- and file name fragments ; =========================

; #FUNCTION#
; Name ..........: FileName
; Description ...: Trims location from the filename
; Syntax ........: FileName($sPath)
; Parameters ....: $sPath ... Name of the file.
; Returns .......: Trimmed filename including extension
; Author ........: Bert Kerkhof

Func FileName($sPath)
  Local $iPos = StringInStr($sPath, '\', 0, -1) + 1 ; Search from right
  Return StringMid($sPath, $iPos)
EndFunc   ;==>FileName

; #FUNCTION#
; Name ..........: FileExt
; Description ...: Get the extension from filename / full path
; Syntax ........: FileExt($sPath)
; Parameters ....: $sPath ... Name of the file.
; Returns .......: Extention
; Author ........: Bert Kerkhof

Func FileExt($sPath)
  Local $sShort = FileName($sPath)
  Local $iPos = StringInStr($sShort, '.', 0, -1) + 1 ; Search from right
  Return Lif($iPos = 1, '', StringMid($sShort, $iPos))
EndFunc   ;==>FileExt

; #FUNCTION#
; Name ..........: FileBase
; Description ...: Trim extension from full path / filename
; Syntax ........: FileBase($sPath)
; Parameters ....: $sPath ... Name of the file.
; Returns .......: Location\filename
; Author ........: Bert Kerkhof

Func FileBase($sPath) ; Trim extension from full path / filename
  Local $sShort = FileName($sPath)
  Local $iPos = StringInStr($sShort, '.', 0, -1) - 1 ; Search from right
  If $iPos > 0 Then $sShort = StringLeft($sShort, $iPos)
  Return FileLocation($sPath) & $sShort
EndFunc   ;==>FileBase

; #FUNCTION#
; Name ..........: FileLocation
; Description ...: Get path from filename, including trailing backslash
; Syntax ........: FileLocation($sPath)
; Parameters ....: $sPath ... Full path and name of the file
; Returns .......: Path
; Author ........: Bert Kerkhof

Func FileLocation($sPath)
  Local $iPos = StringInStr($sPath, '\', 0, -1)
  Return StringLeft($sPath, $iPos)
EndFunc   ;==>FileLocation

; #FUNCTION#
; Name ..........: FileDrive
; Description ...: Get drive-letter from file-name plus trailing backslash
; Syntax ........: FileDrive($sPath)
; Parameters ....: $sPath ... File [path and] name.
; Returns .......: Success : Drive designator containing trailing backslash.
;                  For example 'C:\'
;                : Failure: Empty string
; Author ........: Bert Kerkhof

Func FileDrive($sPath)
  If StringMid($sPath, 2, 2) == ':\' Then Return StringLeft($sPath, 3)
  If StringMid($sPath, 2, 1) == ':' Then Return StringLeft($sPath, 2) & '\'
  Return ''
EndFunc   ;==>FileDrive

; #FUNCTION#
; Name ..........: DriveGetVolume
; Description ...: Get CD/DVD/Blu-Ray volume name.
; Syntax ........: DriveGetVolume($sFileName)
; Parameter .....: $sFileName ... [Path and] name of a file.
; Returns .......: Success : Returns volume name including trailing
;                            backslash
;                  Failure : Empty string
; Author ........: Bert Kerkhof

Func DriveGetVolume($sFileName) ; Get CD|DVD|Blu-Ray volume name
  Return DriveGetLabel(FileDrive($sFileName)) & '\'
EndFunc   ;==>DriveGetVolume

; #FUNCTION#
; Name ..........: FileRootFolder
; Description ...: Get RootFolder from filename, including trailing
;                  backslash
; Syntax ........: FileRootFolder($sFileName)
; Parameter .....: $sPath ... [Path and] name of the file
; Returns .......: Folder (String)
; Author ........: Bert Kerkhof

Func FileRootFolder($sFileName)
  If StringMid($sFileName, 2, 2) == ':\' Then $sFileName = StringMid($sFileName, 4)
  If StringMid($sFileName, 2, 1) == ':' Then $sFileName = StringMid($sFileName, 3)
  Local $Pos = StringInStr($sFileName, '\')
  If $Pos Then
    Return StringLeft($sFileName, $Pos - 1)
  Else
    Return ''
  EndIf
EndFunc   ;==>FileRootFolder

; #FUNCTION#
; Name ..........: SnoopPath
; Description ...: Removes the folder from the full path.
; Syntax ........: SnoopPath($sPath)
; Parameters ....: $sPath ... [Full] name of the file.
; Returns .......: Snooped file name
; Author ........: Bert Kerkhof

Func SnoopPath($sPath) ; Removes the user folder from path
  Local $Uprofile = @UserProfileDir & '\'
  Local $Len = StringLen($Uprofile) ; first remove user profile part:
  If StringLeft($sPath, $Len) = $Uprofile Then $sPath = StringMid($sPath, $Len + 1)
  Local $iPos = StringInStr($sPath, '\') ; second remove one root folder:
  If $iPos Then $sPath = StringMid($sPath, $iPos + 1)
  Return $sPath
EndFunc   ;==>SnoopPath

; #FUNCTION#
; Name ..........: FileRevNum
; Description ...: Returns revision number from filename.
; Syntax ........: FileRevNum($sFileTitle)
; Parameters ....: $sFileTitle ... Optionally may include FileLocation / Extension
; Return value ..: Number
; Author ........: Bert Kerkhof

Func FileRevNum($sFileTitle)
  Local $aResult = StringRegExp($sFileTitle, "\((\d{1,3})\)\.\w{1,10}\z", 1)
  Return IsArray($aResult) ? $aResult[0] : 0
EndFunc   ;==>FileRevNum

; #FUNCTION#
; Name ..........: PathRevision
; Description ...: Creates a new file-name to prevent overwriting.
; Syntax ........: PathRevision($sFullPath)
; Parameters ....: $sFullPath ... The path and name of the file.
;                                 Including location, filename and extension
; Return value ..: A new full path with revision number between round
;                  brackets
; Author ........: Bert Kerkhof

Func PathRevision($sFullPath)
  Local $sBase = FileBase($sFullPath), $sExt = FileExt($sFullPath)
  Local $N = FileRevNum($sFullPath)
  If StringLen($sExt) Then $sExt = '.' & $sExt
  While FileExists($sFullPath)
    $N += 1
    $sFullPath = $sBase & "(" & $N & ")" & $sExt
  WEnd
  Return $sFullPath
EndFunc   ;==>PathRevision

; #FUNCTION#
; Name ..........: aRevFilter
; Description ...: Removes revision numbers from filenames.
; Syntax ........: aRevFilter($aFiles)
; Parameters ....: $aFiles ... An array of file-names.
; Returns .......: Array of filtered file-names
; Author ........: Bert Kerkhof

Func aRevFilter($aFiles)
  Local $aFiltered = aNew()
  For $I = 1 To $aFiles[0]
    If FileRevNum($aFiles[$I]) = 0 Then aAdd($aFiltered, $aFiles[$I])
  Next
  Return $aFiltered
EndFunc   ;==>aRevFilter

; #FUNCTION#
; Name ..........: aFilter
; Description ...: Select files that start with the specified title
;                  (if stated) and have a specified extensions (if stated).
; Syntax ........: aFilter($aArray[, $rExt = ''[, $sTitle = ''[, $nSortKey = 0[, $Ldescending = False]]]])
; Parameters ....: $aArray ...... Array of unknowns.
;                  $rExt ........ [optional] an unknown value.
;                                 Default is ''.
;                  $sTitle ...... [optional] a string value. Default is ''.
;                  $nSortKey .... [optional] a general number value. Default is 0.
;                                   0: NoSort  1: FileName  2: FileLocation
;                  $Ldescending .. [optional] sets the sort-order.
;                                  Default is False.
; Return value ..: 2dim array with for each file: [1] path and [2] name
; Author ........: Bert Kerkhof

Func aFilter($aArray, $rExt = '', $sTitle = '', $nSortKey = 0, $Ldescending = False)
  ; $nSortKey
  Local $aExt = aRecite($rExt), $aResult = aNew(), $aaFiles = aNew()
  For $I = 1 To $aArray[0]
    If $rExt > '' And aSearch($aExt, FileExt($aArray[$I])) = 0 Then ContinueLoop
    If $sTitle > '' And StringLeft(FileName($aArray[$I]), StringLen($sTitle)) <> $sTitle Then ContinueLoop
    aAdd($aResult, $aArray[$I]) ; Filename includes location
  Next
  If $nSortKey = 0 Then Return $aResult
  Local $sFile, $sDateTime
  ; Sort on (1) FileName (2) FileLocation (3) DateTime
  For $I = 1 To $aResult[0]
    $sFile = FileName($aResult[$I])
    $sDateTime = ($nSortKey = 3) ? FileGetTime($aResult[$I], 0, 1) : ''
    ; Modification datetime:
    aAdd($aaFiles, Array($sFile, $aResult[$I], $sDateTime))
  Next
  ; Sort on FileDateTime:
  aCombSort($aaFiles, $nSortKey, $Ldescending)
  For $I = 1 To $aaFiles[0]
    $aResult[$I] = ($aaFiles[$I])[2]
  Next
  Return $aResult
EndFunc   ;==>aFilter

; #FUNCTION#
; Name ..........: aFilesFromFolder
; Description ...: Return file-names that match a filter.
; Syntax ........: aFilesFromFolder($sFolder[, $rExt = ''[, $sTitle = ''[, $nSortKey = 0[, $Ldescending = False]]]])
; Parameters ....: $sFolder ...... Folder specification
;                  $rExt ......... [optional] an unknown value.
;                                  Default is ''.
;                  $sTitle ....... [optional] a string value. Default is ''.
;                  $nSortKey ..... [optional] a general number value.
;                                    0: NoSort
;                                    1: FileName
;                                    2: FileLocation
;                                    3: DateTime
;                                  Default is zero (NoSort)
;                  $Ldescending .. [optional] an unknown value.
;                                  Default is False.
; Return value ..: 2dim array with for each file: [1] path and [2] name
; Author ........: Bert Kerkhof

Func aFilesFromFolder($sFolder, $rExt = '', $sTitle = '', $nSortKey = 0, $Ldescending = False)
  Local $aExt = aRecite($rExt), $aaFiles = aNew()
  Local $hSearch = FileFindFirstFile($sFolder & $sTitle & '*.*')
  Local $sDateTime
  While True
    Local $sFile = FileFindNextFile($hSearch)
    If @error Then ExitLoop ; No more files in folder
    If @extended Then ContinueLoop ; IsFolder
    If $rExt > '' And aSearch($aExt, FileExt($sFile)) = 0 Then ContinueLoop
    $sDateTime = ($nSortKey = 3) ? FileGetTime($sFile, 0, 1) : ''
    ; Modification datetime
    aAdd($aaFiles, Array($sFile, $sFolder & $sFile, $sDateTime))
  WEnd
  FileClose($hSearch)
   ; Sort on FileDateTime
  If $nSortKey Then aCombSort($aaFiles, $nSortKey, $Ldescending)
  Local $aResult = aNew($aaFiles[0])
  For $I = 1 To $aaFiles[0]
    $aResult[$I] = ($aaFiles[$I])[2]
  Next
  Return $aResult ; One-dimensional.
EndFunc   ;==>aFilesFromFolder


; User interaction ====================================================

; #FUNCTION#
; Name ..........: aSelectFolder
; Description ...: Opens a Choose Folder dialog.
; Syntax ........: aSelectFolder([$sOption = ""])
; Parameter .....: $sOption. Choose:
;                    "Subfolders"  to include subfolders
; Usage .........: FileChangeDir($sFolder) ; Sets the inital folder
;                  Local $aFolders = aSelectFolders('SubFolders')
; Returns .......: Array of full folder paths, with trailing slash
;                  If no selection is made, an empty one-based
;                  array is returned and the error flag is set.
; Author ........: Bert Kerkhof
; Comment .......: The initial folder is the current folder.

Func aSelectFolder($sOption = "")
  Local $iOption = aSearch(aRecite("|Subfolders"), $sOption)
  If $iOption = 0 Then MsgBox(64, "Error in aSelectFolder", "Option not found: " & $sOption)

  Local $N = 0, $sFolder = @WorkingDir, $aFolders = aNew()
  Local $sMess = "Select folder" & Lif($iOption = 2, " including subfolders")
  $sFolder = FileSelectFolder($sMess, $sFolder)
  If @error Then Return SetError(9, 0, $aFolders)
  FileChangeDir($sFolder) ; Update current folder
  If StringRight($sFolder, 1) <> '\' Then $sFolder &= '\'
  aAdd($aFolders, $sFolder) ; The main folder
  Local $hSearch, $sFile
  While $N < $aFolders[0] And $iOption = 2 ; Recurse
    $N += 1
    $sFolder = $aFolders[$N]
    $hSearch = FileFindFirstFile($sFolder & '*.*')
    While True
      $sFile = FileFindNextFile($hSearch)
      If @error Then ExitLoop ; No more files in folder
      If @extended Then aAdd($aFolders, $sFolder & $sFile & '\')
    WEnd
    FileClose($hSearch)
  WEnd
  Return $aFolders
EndFunc   ;==>aSelectFolder

; #FUNCTION#
; Name ..........: aSelectFiles
; Description ...: Opens a dialog to choose a file.
; Syntax ........: aSelectFiles([$rExt = ""[, $sOption = ""]])
; Parameter 1 ...: $rExt is a list of extensions to filter files
; Parameter 2 ...: $rOption ........ Choose a mode: "SubFolders|Files|File"
;                    Default ....... Files from the current folder are
;                                    chosen.
;                    "Subfolders" .. Includes files from subfolders
;                    "File" ........ Choose one file
;                    "Files" ....... Choose one or more files
;                  File and Files override the Subfolders option
; Usage .........: FileChangeDir($sFolder)
;                  Local $aFiles = aSelectFiles("mp4|m4v|mkv", "Subfolders")
;                  Local $bFiles = aSelectFiles("mp4|m4v|mkv", "Files")
; Returns .......: Array of full paths of the chosen files
;                  If no selection is made, an empty one-based
;                  array is returned and the error flag is set.
; Author ........: Bert Kerkhof
; Comment .......: The initial folder to choose from is the current folder.

Func aSelectFiles($rExt = "", $sOption = "")
  Local $iOption = aSearch(aRecite("|Subfolders|File|Files"), $sOption)
  If $iOption = 0 Then
    MsgBox(64, "Error in aSelectFiles", "Option not found: '" & $sOption & "'")
    $sOption = "" ; Default
  EndIf
  Local $aExt = aRecite($rExt), $aFiles = aNew()
  If $iOption < 3 Then ; Choose files from current folder:
    Local $aFolders = aSelectFolder($sOption)
    If @error Then Return SetError(9, 0, $aFolders)
    Local $sFolder, $hSearch, $sFile
    For $I = 1 To $aFolders[0]
      $sFolder = $aFolders[$I]
      $hSearch = FileFindFirstFile($sFolder & '\*.*')
      While True
        $sFile = FileFindNextFile($hSearch)
        If @error Then ExitLoop ; No more files in folder
        If @extended Then ContinueLoop ; IsFolder
        If aSearch($aExt, FileExt($sFile)) = 0 Then ContinueLoop
        aAdd($aFiles, $sFolder & $sFile)
      WEnd
      FileClose($hSearch)
    Next
    Return $aFiles
  Endif
  ; Choose one or more files:
  Local $nOption = ($iOption = 3) ? 0 : 4
  Local $sSpec = '( ' & sRecite($aExt, '; ', '*.') & ')'
  Local $sInvitation = "Select file" & Lif($nOption, "s")
  Local $sString = FileOpenDialog($sInvitation, @WorkingDir, $sSpec, $nOption)
  If @error Then Return SetError(9, 0, $aFiles)
  Local $aIn = StringSplit($sString, '|')
  If $aIn[0] > 1 Then
    $sFolder = $aIn[1] & '\'
    aDelete($aIn, 1)
  Else
    $sFolder = FileLocation($aIn[1])
    $aIn[1] = FileName($aIn[1])
  EndIf
  For $I = $aIn[0] To 1 Step -1 ; Add files that match $aExt :
    If aSearch($aExt, FileExt($aIn[$I])) = 0 And $rExt <> '*' Then ContinueLoop
    aAdd($aFiles, $sFolder & $aIn[$I])
  Next
  FileChangeDir(FileLocation($aFiles[1])) ; Update current folder
  Return $aFiles ; One dimensional: full path
EndFunc   ;==>aSelectFiles

; #FUNCTION#
; Name ..........: CheckPresence
; Description ...: Check the presence of an external utilitie.
; Syntax ........: CheckPresence($rUtil)
; Parameters ....: $rUtil ... [Path and] name of the needed utility.
; Returns .......: None
; Author ........: Bert Kerkhof

Func CheckPresence($rUtil)
  Local $Err, $aUtil = aRecite($rUtil)
  Local $sUtil, $Support
  For $I = 1 To $aUtil[0]
    $sUtil = $aUtil[$I]
    If Not FileExists(@ScriptDir & '\Utilities\' & $sUtil) Then
      $Support = "For operation the " & FileBase($sUtil) & " utility is needed." & @CRLF & @CRLF & _
          "Place '" & $sUtil & " in : " & @CRLF & @ScriptDir & '\Utilities\'
      MsgBox(64, FileBase(@ScriptName), $Support)
      $Err = True
    EndIf
  Next
  If $Err Then Exit 9
EndFunc   ;==>CheckPresence

; #FUNCTION#
; Name ..........: FileMatch
; Description ...: Checks wether a file matches a file mask.
; Syntax ........: FileMatch($sPath, $sMask)
; Parameters ....: sFname .. Filename to check
;                  sMask ... File mask may contain parts of the filename,
;                              '*' substitutes zero or more characters
; Returns .......: True .... File matches with the mask
;                  False ... File does not match
; Author ........: Bert Kerkhof

Func FileMatch($sPath, $sMask)
  Local $Len, $Found = -1, $Pos = 1, $Match = True
  For $I = 1 To 5
    $sPath = StringReplace($sPath, '??', '?')
  Next
  Local $aM = StringSplit($sMask, '*?')
  $Len = StringLen($aM[1])
  For $I = 1 To $aM[0]
    If $Match And StringLen($aM[$I]) > 0 Then
      $Found = StringInStr($sPath, $aM[$I], 0, 1, $Pos, $Len)
      $Match = $Found > 0
      $Pos += $Found + StringLen($aM[$I])
    EndIf
    $Len = 9999
  Next
  Return $Match
EndFunc   ;==>FileMatch

; #FUNCTION#
; Name ..........: DirTreeMove
; Description ...: Moves a directory including subdirectories and files.
; Syntax ........: DirTreeMove($sSourceFolder, $sTargetFolder[, $Overwrite = 0])
; Parameters ....: cSourceFolder .. Source directory
;                : cTargetFolder .. Target directory
;                : Flag ........... If set then target files
;                                   with a same name are overwritten
; Operation .....: All files and subfolders in the cSource folder
;                  are moved, folder structure is always copied
; Returns .......: cSourceFolder exist not failure: 0, Success: 1
; Comment        : Uses the technique of recursion
; Returns .......: None
; Author ........: Bert Kerkhof

Func DirTreeMove($sSourceFolder, $sTargetFolder, $Overwrite = 0)
  Local $hFind, $sFile, $I, $aSubFolder = aNew()
  ; Create necessary destination directory structure
  $Overwrite = 8 + Lif($Overwrite, 1, 0)
  ; cSourceFolder exitst not:
  If StringInStr(FileGetAttrib($sSourceFolder), 'D') == 0 Then Return 0
   ; Create destination directory structure just in case it exists not.
  DirCreate($sTargetFolder)
  $hFind = FileFindFirstFile($sSourceFolder & '\*.*')
  While $hFind <> -1
    $sFile = FileFindNextFile($hFind)
    If @error == 1 Then ExitLoop
    If @extended Then
      aAdd($aSubFolder, $sFile)
    Else
      FileMove($sSourceFolder & '\' & $sFile, $sTargetFolder & '\' & $sFile, $Overwrite)
    EndIf
  WEnd
  FileClose($hFind)
  ; Recurse at last to reduce number of open files :
  For $I = 1 To $aSubFolder[0]
    DirTreeMove($sSourceFolder & '\' & $aSubFolder[$I], $sTargetFolder & '\' & $aSubFolder[$I], $Overwrite)
  Next
  DirRemove($sSourceFolder)
  Return 1
EndFunc   ;==>DirTreeMove


; Helper routines =====================================================

; #FUNCTION#
; Name ..........: Fatal
; Description ...: Notify the user.
; Syntax ........: Fatal($sMess)
; Parameters ....: $sMess .... Informative message.
; Returns .......: None
; Author ........: Bert Kerkhof

Func Fatal($sMess)
  $sMess = 'Fatal error' & @CRLF & @CRLF & $sMess & @CRLF & @CRLF
  MsgBox(64, FileBase(FileName(@ScriptDir)), $sMess & 'Please notify the programmer')
  Exit
EndFunc   ;==>Fatal


; Tests and example: ==================================================

; #FUNCTION#
; Name ..........: TestFileDialog
; Description ...: Demonstrates use of aSelectFolders and aSelectFiles.
; Syntax ........: TestFileDialog()
; Parameters ....: None
; Returns .......: None
; Author ........: Bert Kerkhof

Func TestFileDialog()
  FileChangeDir(@UserProfileDir & "\Videos") ; Set current folder
  Local $aFiles = aSelectFiles("mp4|mkv")
  If @error Then Exit
  MsgBox(64, 'Selected from main folder', sRecite($aFiles, @CRLF & @CRLF))
  $aFiles = aSelectFiles("mp4|mkv", "Subfolders")
  If @error Then Exit
  MsgBox(64, 'Selected from folder and subfolders', sRecite($aFiles, @CRLF & @CRLF))
  $aFiles = aSelectFiles("mp4|mkv", "File")
  If @error Then Exit
  MsgBox(64, 'Selected max one file', sRecite($aFiles, @CRLF & @CRLF))
  $aFiles = aSelectFiles("mp4|mkv", "Files")
  If @error Then Exit
  MsgBox(64, 'Selected file(s)', sRecite($aFiles, @CRLF & @CRLF))
EndFunc   ;==>TestFileDialog
; TestFileDialog()

; End =================================================================

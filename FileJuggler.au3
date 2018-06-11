#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=FileJuggler.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Juggler file functions
#AutoIt3Wrapper_Res_Fileversion=1.3
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 2018-06-03 Apache 2.0 license
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/reel
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include-once
#include <aDrosteArray.au3> ; Published on GitHub repository "aDrosteArray"

; Module FileJuggler ================================================================
; Tested with AutoIT v3.3.14.5 interpreter/compiler

; Juggling with file names elements :
;        + FileName            : Trims location (and backslash) from full path / filename
;        + FileExt             : Get the extension from a filename
;        + FileBase            : Trim extension from file location / file name
;        + FileLocation        : Get file location from filename, includes trailing slash
;        + FileDrive           : Get drive from full path
;        + DriveGetVolume      : Get CD|DVD|Blu-Ray volume name
;        + FileRootFolder      : Get main folder from full path
;        + FileSnoopPath       : Removes the user folder from path
;        + PathRevision        : Creates a new filename to prevent overwriting
;        + FileRevNum          : Returns revision number from file name
;        + aRevFilter          : Removes revisions

; Open system files and Folders :
;        + aFilter             : Select files with specified title and extension
;        + aFilesFromFolder    : Create array of filenames from folder
;        + aSelectFolder       : Open dialog for the user to select folder;
;        + aSelectFiles        : Open dialog for the user to select files
;        + FileVolume          : Fill serial number in for multiple file
;        + FileAddenda         : Organize files with multiple parts
;
; Other :
;        + CheckPresence       : Checks the presence of utilities
;        + FileMatch           : Checks whether a filename matches a file mask
;        + DirTreeMove         : Moves files and directories
;
; Test and example(s) :
;        + FileaAddendaExample : Demonstrates use of FileAddenda function
;        + TestDialog          : Demonstrates use of SelectFolders and SelectFiles
;
; Tested with ...: AutoIt v3.3.14.5
; Author ........: kerkhof.bert@gmail.com

#Region ; Juggle with path- and file name fragments ; =========================

Func FileName($cPath) ; Trims location from the filename
  ; Syntax.........: FileName(cFname)
  ; Parameters ....: cFname - Name of file
  ; Return values .: Trimmed filename including extension
  Local $Pos = StringInStr($cPath, '\', 0, -1) + 1 ; Search from right
  ; If StringRight($cPath,1)='\' Then $cPath = StringTrimRight($cPath, 1)
  Return StringMid($cPath, $Pos)
EndFunc   ;==>FileName

Func FileExt($cPath) ; Get the extension from filename / full path
  ; Syntax.........: FileExt(cFname)
  ; Parameters ....: cFname - Name of file
  ; Return values .: Extension
  Local $ShortName = FileName($cPath)
  Local $Pos = StringInStr($ShortName, '.', 0, -1) + 1 ; Search from right
  Return Lif($Pos = 1, '', StringMid($ShortName, $Pos))
EndFunc   ;==>FileExt

Func FileBase($cPath) ; Trim extension from full path / filename
  ; Syntax.........: FileBase(cFname)
  ; Parameters ....: cFname - Name of file
  ; Return values .: Location\Filename
  Local $ShortName = FileName($cPath)
  Local $Pos = StringInStr($ShortName, '.', 0, -1) - 1 ; Search from right
  If $Pos > 0 Then $ShortName = StringLeft($ShortName, $Pos)
  Return FileLocation($cPath) & $ShortName
EndFunc   ;==>FileBase

Func FileLocation($cPath) ; Get path from filename, includes trailing backslash
  ; Syntax.........: FileLocation(cFname)
  ; Parameters ....: cFname - Name of file
  ; Return values .: Path
  Local $Pos = StringInStr($cPath, '\', 0, -1)
  Return StringLeft($cPath, $Pos)
EndFunc   ;==>FileLocation

Func FileDrive($cPath) ; Get drive letter from filename plus trailing backslash
  ; Syntax.........: FileDrive(cFname)
  ; Parameters ....: cFname - Name of file
  ; Return values .: Success : Drive designator containing trailing backslash.
  ;                  For example 'C:\'
  ;                : Failure: Empty string
  If StringMid($cPath, 2, 2) == ':\' Then Return StringLeft($cPath, 3)
  If StringMid($cPath, 2, 1) == ':' Then Return StringLeft($cPath, 2) & '\'
  Return ''
EndFunc   ;==>FileDrive

Func DriveGetVolume($cPath) ; Get CD|DVD|Blu-Ray volume name
  ; Syntax.........: DiveGetVolume($cPath)
  ; Parameters ....: cFname - Path of file
  ; Return values .: Succes : Returns volume name including trailing backslash
  ;                : Failure : Empty string
  Return DriveGetLabel(FileDrive($cPath)) & '\'
EndFunc   ;==>DriveGetVolume

Func FileRootFolder($cPath) ; Get RootFolder from filename, includes trailing backslash
  ; Syntax.........: FileRootFolder(cFname)
  ; Parameters ....: cFname - Name of file
  ; Return values .: FileRootFolder
  If StringMid($cPath, 2, 2) == ':\' Then $cPath = StringMid($cPath, 4)
  If StringMid($cPath, 2, 1) == ':' Then $cPath = StringMid($cPath, 3)
  Local $Pos = StringInStr($cPath, '\')
  If $Pos Then
    Return StringLeft($cPath, $Pos - 1)
  Else
    Return ''
  EndIf
EndFunc   ;==>FileRootFolder

Func SnoopPath($cPath) ; Removes the user folder from path
  ; Syntax.........: SnoopPath(cFname)
  ; Parameters ....: cFname - Name of file
  ; Return values .: Succes: Snooped file name
  Local $Uprofile = @UserProfileDir & '\'
  Local $Len = StringLen($Uprofile) ; first remove user profile part:
  If StringLeft($cPath, $Len) = $Uprofile Then $cPath = StringMid($cPath, $Len + 1)
  Local $Pos = StringInStr($cPath, '\') ; second remove one root folder:
  If $Pos Then $cPath = StringMid($cPath, $Pos + 1)
  Return $cPath
EndFunc   ;==>SnoopPath

Func PathRevision($cFullPath) ; Creates a new filename to prevent overwriting
  ; Syntax.........: PathRevision($FileName)
  ; Parameters ....: $cFullPath
  ;                  Includes location, filename and extension
  ; Return values .: A new path with revision number between round brackets
  ; Author ........: kerkhof.bert@gmail.com

  If Not FileExists($cFullPath) Then Return $cFullPath
  Local $I, $cExt = FileExt($cFullPath)
  If StringLen($cExt) Then $cExt = '.' & $cExt
  For $I = 1 To 999
    Local $NewName = FileBase($cFullPath) & '(' & $I & ')' & $cExt
    If Not FileExists($NewName) Then ExitLoop
  Next
  Return $NewName
EndFunc   ;==>PathRevision

Func FileRevNum($cFileTitle) ; Optionally may include FileLocation / Extension
  ; Returns revision number from file name
  Local $Short = FileBase(FileName($cFileTitle))
  Local $Pos1 = StringInStr($Short, '(', 0, -1) + 1 ; Search from right
  If $Pos1 = 1 Or StringRight($Short, 1) <> ')' Then Return 0
  Local $Rev = StringTrimRight(StringMid($Short, $Pos1), 1)
  If Not StringIsDigit($Rev) Then Return 0
  Return Lif($Rev > 999, 0, $Rev)
EndFunc   ;==>FileRevNum

Func aRevFilter($aFiles)
  ; Removes revisions
  Local $I, $aFiltered = aNewArray()
  For $I = 1 To $aFiles[0]
    If FileRevNum($aFiles[$I]) = 0 Then aAdd($aFiltered, $aFiles[$I])
  Next
  Return $aFiltered
EndFunc   ;==>aRevFilter

; Select files that start with the specified title (if stated) And ..
;   have one of the specified extensions (if stated)
Func aFilter($aArray, $rExt = '', $sTitle = '', $nSortKey = 0, $Descending = 0)
  ; $nSortKey [0] NoSort [1] FileName [2] FileLocation
  Local $I, $aResult = aNewArray(), $aExt = aRecite($rExt), $aaFiles = aNewArray()
  For $I = 1 To $aArray[0]
    If $rExt > '' And aSearch($aExt, FileExt($aArray[$I])) = 0 Then ContinueLoop
    If $sTitle > '' And StringLeft(FileName($aArray[$I]), StringLen($sTitle)) <> $sTitle Then ContinueLoop
    aAdd($aResult, $aArray[$I]) ; Filename includes location
  Next
  If $nSortKey = 0 Then Return $aResult
  Local $sFile, $sDateTime
  For $I = 1 To $aResult[0] ; Sort on (1) FileName (2) FileLocation (3) DateTime
    $sFile = FileName($aResult[$I])
    $sDateTime = ($nSortKey = 3) ? FileGetTime($aResult[$I], 0, 1) : ''
    aAdd($aaFiles, aConcat($sFile, $aResult[$I], $sDateTime)) ; Modification datetime
  Next
  aCombSort($aaFiles, $nSortKey, $Descending) ; Sort on FileDateTime
  For $I = 1 To $aaFiles[0]
    $aResult[$I] = ($aaFiles[$I])[2]
  Next
  Return $aResult
EndFunc   ;==>aFilter

Func aFilesFromFolder($sFolder, $rExt = '', $sTitle = '', $nSortKey = 0, $Descending = 0)
  ; $nSortKey [0] NoSort [1] FileName [2] FileLocation [3] DateTime
  Local $aaFiles = aNewArray(), $aExt = aRecite($rExt)
  Local $hSearch = FileFindFirstFile($sFolder & $sTitle & '*.*')
  Local $sDateTime
  While True
    Local $sFile = FileFindNextFile($hSearch)
    If @error Then ExitLoop ; No more files in folder
    If @extended Then ContinueLoop ; IsFolder
    If $rExt > '' And aSearch($aExt, FileExt($sFile)) = 0 Then ContinueLoop
    $sDateTime = ($nSortKey = 3) ? FileGetTime($sFile, 0, 1) : ''
    aAdd($aaFiles, aConcat($sFile, $sFolder & $sFile, $sDateTime)) ; Modification datetime
  WEnd
  FileClose($hSearch)
  If $nSortKey Then aCombSort($aaFiles, $nSortKey, $Descending) ; Sort on FileDateTime
  Local $I, $aResult = aNewArray($aaFiles[0])
  For $I = 1 To $aaFiles[0]
    $aResult[$I] = ($aaFiles[$I])[2]
  Next
  Return $aResult ; One-dimensional.
  ; Each element contains the complete file path (FileLocation & FileName)
EndFunc   ;==>aFilesFromFolder

#EndRegion ; Juggle with path- and file name fragments ; ======================

#Region ; Helper routines for aSelectFiles ====================================

Func Garant(ByRef $Param, $Backup) ; Fill in if value is missing
  Return ($Param = 0 Or $Param = '') ? $Backup : $Param
EndFunc   ;==>Garant

Func cGareel($aOption, $rWishList)
  ; Transform a wish list to the single key that matters for the work at hand
  Local $I, $iPos, $aWishList = aRecite($rWishList)
  Local $cWishList
  For $I = 1 To $aWishList[0]
    $cWishList = $aWishList[$I]
    $iPos = aSearch($aOption, $cWishList)
    If $iPos Then Return $aOption[$iPos]
  Next
  Return $aOption[1] ; Default
EndFunc   ;==>cGareel

#EndRegion ; Helper routines for aSelectFiles =================================

#Region ; =====================================================================

Func aSelectFolder(ByRef $cFolder, $rOption = '', $ParentWin = 0)
  ; Description ...: Opens a Choose Folder dialog.
  ; Parameter 1....: $cFolder : Output the chosen folder.
  ; Parameter 2 ...: $rOptions to choose the start folder / to exclude subfolders:
  ;                  ExcludeSubFolders : inhibit inclusion of subfolders
  ;                  Possible startfolders :
  ;                  'Documents' (default) / 'Pictures' / 'Music' / 'Videos' / 'Desktop' / 'Favorites'
  ; Usage .........: Local $cFolder = ''
  ;                  Local $aFolders = aSelectFolders($cFolder, 'Burst|Video')
  ; Returns .......: Array of full folder paths, with trailing slash
  ;                  The chosen $cFolder is passed by reference
  ; Author ........: kerkhof.bert@gmail.com

  Local $aFolderChoice = aRecite('Documents|Pictures|Music|Videos|Favorites|Desktop')
  Local $cCateg = cGareel($aFolderChoice, $rOption)
  $cFolder = Garant($cFolder, @UserProfileDir & '\' & $cCateg & '\')
  Local $Lrecurse = Not aSearch(aRecite($rOption), 'ExcludeSubFolders')
  ; Local $cType = (StringLower(StringRight($cCateg, 1) = 's') ? StringTrimRight($cCateg, 1) : $cCateg)
  Local $cInvitation = 'Select folder'
  Local $N = 0, $aFolders = aNewArray()
  $cFolder = FileSelectFolder($cInvitation, $cFolder, 1, '', $ParentWin)
  If @error Then Return SetError(9, 0, $aFolders)
  If StringRight($cFolder, 1) <> '\' Then $cFolder &= '\'
  aAdd($aFolders, $cFolder) ; The main folder
  Local $sFolder, $hSearch, $cFile
  While $N < $aFolders[0] And $Lrecurse
    $N += 1
    $sFolder = $aFolders[$N]
    $hSearch = FileFindFirstFile($sFolder & '*.*')
    While True
      $cFile = FileFindNextFile($hSearch)
      If @error Then ExitLoop ; No more files in folder
      If @extended Then aAdd($aFolders, $sFolder & $cFile & '\')
    WEnd
    FileClose($hSearch)
  WEnd
  Return $aFolders
EndFunc   ;==>aSelectFolder

Func aSelectFiles(ByRef $cFolder, $rOption = '', $rExt = '', $ParentWin = 0)
  ; Description ...: Opens a Choose File dialog.
  ; Parameter 1....: $cFolders : Output the chosen folder.
  ; Parameter 2 ...: $rOptions : to choose a dialog mode and a start folder.
  ;                  Dialog mode :  Burst / Multi / Single / ExcludeSubFolders
  ;                  'Burst' (default) to choose a complete folder and subfolders
  ;                  'Burst|ExludeSubFolders' inhibits the inclusion of subfolders
  ;                  'Multi' implies to choose one or more files from one folder
  ;                  'Single' choose a single file
  ;                  StartFolder :
  ;                  'Documents' (default) / 'Pictures' / 'Music' / 'Videos' / 'Desktop' / 'Favorites'
  ; Parameter 3 ...: $rExt is a list of extensions to filter files
  ; Usage .........: Local $cFolder = ''
  ;                  Local $aFiles = aSelectFiles($cFolder, 'Burst|Video', 'mp4|m4v|mkv')
  ; Returns .......: Array of full paths of the chosen files
  ;                  The chosen folder is passed by reference
  ; Author ........: kerkhof.bert@gmail.com

  Local $aFolderChoice = aRecite("Documents|Pictures|Music|Videos|Favorites|Desktop")
  Local $aDialogMode = aRecite('Burst|Multi|Single')
  Local $aOption = aRecite($rOption), $aExt = aRecite($rExt)
  Local $J, $allOptions, $iOption

  If Not @Compiled Then
    $allOptions = aFlat($aDialogMode, $aFolderChoice, 'ExcludeSubFolders')
    For $J = 1 To $aOption[0]
      $iOption = aSearch($allOptions, $aOption[$J])
      If $iOption = 0 Then MsgBox(64, 'Error in aSelectFiles', 'Option not found: ' & $aOption[$J])
    Next
  EndIf

  Local $cDialogMode = cGareel($aDialogMode, $rOption)
  Local $cCateg = cGareel($aFolderChoice, $rOption)
  Local $aFiles = aNewArray()
  $cFolder = Garant($cFolder, @UserProfileDir & '\' & $cCateg)
  Local $aFolders, $sFolder, $hSearch, $cFile
  If $cDialogMode = 'Burst' Then
    ; OpenFolder dialog :
    $aFolders = aSelectFolder($cFolder, $rOption, $ParentWin)
    If @error Then Return SetError(9, 0, $aFolders)
    For $I = 1 To $aFolders[0]
      $sFolder = $aFolders[$I]
      $hSearch = FileFindFirstFile($sFolder & '\*.*')
      While True
        $cFile = FileFindNextFile($hSearch)
        If @error Then ExitLoop ; No more files in folder
        If @extended Then ContinueLoop ; IsFolder
        If aSearch($aExt, FileExt($cFile)) = 0 Then ContinueLoop
        aAdd($aFiles, $sFolder & $cFile)
      WEnd
      FileClose($hSearch)
    Next
    Return $aFiles ; One-dimensional. Locaton\FileName in each element
  EndIf

  ; OpenFile dialog :
  Local $nOption = ($cDialogMode = 'Single') ? 0 : 4
  Local $cSpec = '( ' & sRecite($aExt, '; ', '*.') & ')'
  Local $cInvitation = 'Select file' & Lif($cDialogMode = 'Multi', 's')
  Local $cString = FileOpenDialog($cInvitation, $cFolder, $cSpec, $nOption, '', $ParentWin)
  If @error Then Return SetError(9, 0, $aFiles)
  Local $aIn = StringSplit($cString, '|')
  If $aIn[0] > 1 Then
    $cFolder = $aIn[1] & '\'
    aDelete($aIn, 1)
  Else
    $cFolder = FileLocation($aIn[1])
    $aIn[1] = FileName($aIn[1])
  EndIf
  For $I = $aIn[0] To 1 Step -1 ; Add files that match $aExt :
    If aSearch($aExt, FileExt($aIn[$I])) = 0 And $rExt <> '*' Then ContinueLoop
    aAdd($aFiles, $cFolder & $aIn[$I])
  Next
  Return $aFiles ; One dimensional: full path
EndFunc   ;==>aSelectFiles

#EndRegion ; ==================================================================

#Region ; Multiple file parts system ==========================================

Func FileVolume(Const $sAddendum, Const $Number)
  ; Fill in a serial number for a multiple file parts postfix
  ; Syntax.........: Volume(sAddendum, nNumber)
  ; Parameters ....: sAddendum - Postfix match string to recognize file with multiple parts
  ;                              the match string has one or more # characters, stand for numbers 0..9
  ;                  nNumber   - Serial number of file part
  ; Return values .: sVolume   - Postfix with the specified serial number
  ; Author ........: kerkhof.bert@gmail.com
  Local $Parts = StringSplit($sAddendum, '#')
  Return $Parts[1] & StringLeft('00000', $Parts[0] - 2) & String($Number) & $Parts[$Parts[0]]
EndFunc   ;==>FileVolume

Global $Addenda = aConcat('.-vol-##', '_#') ; Default addenda postfixes

Func FileAddenda(ByRef $aF, $Addenda) ; Organize files with multiple parts
  ; Syntax.........: FileAddenda(ByRef aFiles, aAddenda)
  ; Parameters ....: aFiles   : Array of filenames to process
  ;                  aAddenda : Array of possible multi-volume postfixes :
  ;                  + Default: ['.-vol-##', '_#']
  ;                  + The # character stands for numbers 0..9

  ; Return value ..: aMulti array with as much elements as aFiles :
  ;                  + For files with addenda, the corresponding element contains array of part names, alphabet ordered
  ;                  + For files without addenda, the element contains an array with one filename
  ;                    aFiles is a passed by reference array that will be modified when the function is called :
  ;                  + For files with addenda, this array contains a basename, extension included
  ;                  + The addenda positions in this array delete. So if there are addenda, this array will shrink
  ; Author ........: kerkhof.bert@gmail.com

  Local $Multi2D[$aF[0] + 1], $Found, $FoundMore, $ShortName, $Rechts, $Knip, $PostFiles, $Serial
  $Multi2D = aNewArray($aF[0])
  Local $Parts, $sFileName
  For $F = 1 To $aF[0]
    If StringLen($aF[$F]) > 0 Then
      ; CutPath($aF[$F], $sDir, $sFileName, $sBase, $sExt)
      $PostFiles = aConcat(FileName($aF[$F])) ; Single file
      For $A = 1 To $Addenda[0]
        $Rechts = FileVolume($Addenda[$A], 2) & '.' & FileExt($aF[$F])
        $Knip = StringLen($aF[$F]) - StringLen($Rechts)
        If StringRight($aF[$F], StringLen($Rechts)) == $Rechts Then
          $ShortName = StringLeft($aF[$F], $Knip) & FileExt($aF[$F])
          $Parts = aConcat(FileBase($ShortName) & '.' & FileExt($ShortName), FileBase($ShortName) & $Rechts)
          $Found = aSearch($aF, $ShortName)
          If $Found == 0 Then ; Search for long name:
            $Parts[1] = FileBase($ShortName) & FileVolume($Addenda[$A], 1) & FileExt($ShortName)
            $Found = aSearch($aF, FileLocation($ShortName) & '\' & FileBase($ShortName))
          EndIf
          If $Found Then ; Multi-File pair found
            $aF[$F] = $ShortName ; Fill in short name, includes extension
            $aF[$Found] = '' ; Prevent processing second file as single
            For $Serial = 3 To 9999
              $sFileName = FileBase($ShortName) & FileVolume($Addenda[$A], $Serial) & FileExt($ShortName)
              $FoundMore = aSearch($aF, FileLocation($aF[$F]) & '\' & $sFileName)
              If $FoundMore == 0 Then ExitLoop
              $aF[$FoundMore] = '' ; Prevent processing as single
              aAdd($Parts, $sFileName)
            Next
            $PostFiles = $Parts
            ExitLoop
          EndIf
        EndIf
      Next
      $Multi2D[$F] = $PostFiles
    EndIf
  Next
  For $F = $aF[0] To 1 Step -1 ; Shorten :
    If StringLen($aF[$F]) == 0 Then
      aDelete($aF, $F)
      aDelete($Multi2D, $F)
    EndIf
  Next
  Return $Multi2D
EndFunc   ;==>FileAddenda

Func CheckPresence($rUtil) 
  ; Description : Check the presence of multiple external utilities :
  Local $I, $Err, $aUtil = aRecite($rUtil)
  Local $cUtil, $Support
  For $I = 1 To $aUtil[0]
    $cUtil = $aUtil[$I]
    If Not FileExists(@ScriptDir & '\Utilities\' & $cUtil) Then
      $Support = "For operation the " & FileBase($cUtil) & " utility is needed." & @CRLF & @CRLF & _
          "Place '" & $cUtil & " in : " & @CRLF & @ScriptDir & '\Utilities\'
      MsgBox(64, FileBase(@ScriptName), $Support)
      $Err = True
    EndIf
  Next
  If $Err Then Exit 9
EndFunc   ;==>CheckPresence

Func FileMatch($cPath, $Mask) ; Checks wether a file matches a file mask
  ; Syntax.........: FileMatch(sFname, sMask)
  ; Parameters ....: sFname - Filename to check
  ;                  sMask  -  File mask may contain parts of the filename, '*' substitutes zero or more characters
  ; Return values .: True   - File matches with the mask
  ;                  False  - File does not match
  ; Author ........: kerkhof.bert@gmail.com

  Local $I, $Len, $Found = -1, $Pos = 1, $Match = True
  For $I = 1 To 5
    $cPath = StringReplace($cPath, '??', '?')
  Next
  Local $aM = StringSplit($Mask, '*?')
  $Len = StringLen($aM[1])
  For $I = 1 To $aM[0]
    If $Match And StringLen($aM[$I]) > 0 Then
      $Found = StringInStr($cPath, $aM[$I], 0, 1, $Pos, $Len)
      $Match = $Found > 0
      $Pos += $Found + StringLen($aM[$I])
    EndIf
    $Len = 9999
  Next
  Return $Match
EndFunc   ;==>FileMatch

Func DirTreeMove($cSourceFolder, $cTargetFolder, $Overwrite = 0) ; Moves a directory including subdirectories and files
  ; Syntax.........: DirTreeMove(cSourceFolder, cTargetFolder, Flag)
  ; Parameters ....: cSourceFolder - Source directory
  ;                : cTargetFolder - Target directory
  ;                : Flag          - If set then target files with a same name are overwritten
  ; Operation      : All files and subfolders in the cSource folder are moved,
  ;                : folder structure is always copied
  ; Return values .: cSourceFolder exist not failure: 0, Success: 1
  ; Comment        : Uses the technique of recursion
  ; Author ........: kerkhof.bert@gmail.com

  Local $hFind, $cFile, $I, $aSubFolder = aNewArray()
  $Overwrite = 8 + Lif($Overwrite, 1, 0) ; Create necessary destination directory structure
  If StringInStr(FileGetAttrib($cSourceFolder), 'D') == 0 Then Return 0 ; cSourceFolder exitst not
  DirCreate($cTargetFolder) ; Create destination directory structure just in case it exists not
  $hFind = FileFindFirstFile($cSourceFolder & '\*.*')
  While $hFind <> -1
    $cFile = FileFindNextFile($hFind)
    If @error == 1 Then ExitLoop
    If @extended Then
      aAdd($aSubFolder, $cFile)
    Else
      FileMove($cSourceFolder & '\' & $cFile, $cTargetFolder & '\' & $cFile, $Overwrite)
    EndIf
  WEnd
  FileClose($hFind) ; Recurse at last to reduce number of open files :
  For $I = 1 To $aSubFolder[0]
    DirTreeMove($cSourceFolder & '\' & $aSubFolder[$I], $cTargetFolder & '\' & $aSubFolder[$I], $Overwrite)
  Next
  DirRemove($cSourceFolder)
  Return 1
EndFunc   ;==>DirTreeMove

; Return part of source name:
Global Enum $zTITLE, $zRES, $zTRACKNAME, $zLANG, $zEXT
Func zElem(Const $sSource, Const $N) ; [1] sTitle [2] TrackName [3] Language [4] Extension
  Local $I, $aKey = StringSplit(FileName($sSource), '.')
  ; aElem is type string:
  Local $aElem = aRecite('|||'), $nCopy = Min(2, $aKey[0] - 2)
  For $I = 4 - $nCopy To 4
    $aElem[$I] = $aKey[$aKey[0] + $I - 4]
  Next
  Return $N ? $aElem[$N] : sTitleCase(sRecite(aLeft($aKey, Max(1, $aKey[0] - 4)), ' '))
EndFunc   ;==>zElem

#EndRegion ; Multiple file parts system =======================================

#Region ; Helper routines =====================================================

Func Fatal($sMess)
  $sMess = 'Fatal error' & @CRLF & @CRLF & $sMess & @CRLF & @CRLF
  MsgBox(64, FileBase(FileName(@ScriptDir)), $sMess & 'Please notify the programmer')
  Exit
EndFunc   ;==>Fatal

#EndRegion ; Helper routines ==================================================

#Region ; Tests and examples: =================================================

Func FileAddendaExample() ; Demonstrates use of FileJuggler functions
  ; Syntax.........: FileAdendaExample()
  ; Parameters ....: none
  ; Return values .: none

  Local $Multi2D, $Multi, $F
  Local $sFolder
  Local $aF = aSelectFiles($sFolder, 'Videos', 'mpg|mpeg')
  If $aF[0] == 0 Then Exit
  $Multi2D = FileAddenda($aF, $Addenda) ; Detect files and multiple parts
  For $F = 1 To $aF[0]
    $Multi = $Multi2D[$F]
    MsgBox(48, 'FileJuggler', 'basename: "' & $aF[$F] & '"')
    MsgBox(48, 'FileJuggler', 'filenames: ' & sRecite($Multi, ' ', True))
  Next
EndFunc   ;==>FileAddendaExample

Func TestDialog() ; Demonstrates use of aSelectFolders and aSelectFiles
  Local $cFolder = ''
  Local $aFiles = aSelectFolder($cFolder, 'Videos')
  If @error Then Exit
  MsgBox(64, 'Selected: ' & $cFolder, sRecite($aFiles, @CRLF & @CRLF))
  $aFiles = aSelectFiles($cFolder, 'Videos', 'mp4|mkv')
  If @error Then Exit
  MsgBox(64, 'Selected: ' & $cFolder, sRecite($aFiles, @CRLF & @CRLF))
EndFunc   ;==>TestDialog

; TestDialog()

Func TestaFilesFromFolder()
  Local $sManroot = @ScriptDir & '\'
  Local $aFiles = aFilesFromFolder($sManroot, 'rtf')
  MsgBox(64, 'TestaFilesFromFolder', $aFiles[0])
EndFunc   ;==>TestaFilesFromFolder

; TestaFilesFromFolder()

#EndRegion ; Tests and examples ===============================================

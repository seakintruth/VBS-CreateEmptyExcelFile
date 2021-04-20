'
' Dependencies: ZipAFolder, moveFile, writeToFile, Environ, OpenWithExplorer, FileExists, FolderExists
' Version: 1.0.0
' by jeremy.gerdes@navy.mil 
' CC0 1.0 Universal (CC0 1.0) Public Domain Dedication
' [TODO] reset the file attributes of date - Created, Modified, and Accessed to 12/31/1979 11:00 PM (for eastern time zone) to all files
' Usage example
' BuildEmptyExcelFile GetCurrentFileFolder() & "\" & "anEmptyExcelFile.xlsx"
Public Sub BuildEmptyExcelFile(strNewExcelFile)
    Dim strLocalAppTmp
    strLocalAppTmp = Environ("LocalAppData") & "\" & "tmp"
    MkDir strLocalAppTmp
    Dim strTmpZipPath
    strTmpZipPath = strLocalAppTmp & "\" & "buildEmptyExcelFile"
    MkDir strTmpZipPath
    WriteToFile strTmpZipPath & "\" & "[Content_Types].xml" ,  _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
        "<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/><Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/><Override PartName=""/xl/theme/theme1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.theme+xml""/><Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/><Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/><Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/></Types>", _
        False, True
    MkDir strTmpZipPath & "\" & "_rels"
    WriteToFile _
        strTmpZipPath & "\" & "_rels" & "\" & ".rels" , _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/></Relationships>", _
        False, True
    MkDir strTmpZipPath & "\" & "docProps"
    WriteToFile _
        strTmpZipPath & "\" & "docProps" & "\" & "app.xml" , _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes""><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=""2"" baseType=""variant""><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=""1"" baseType=""lpstr""><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><Company>HPES NMCI NGEN</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>", _
        False, True
    WriteToFile _
        strTmpZipPath & "\" & "docProps" & "\" & "core.xml" , _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:creator>jeremy.gerdes</dc:creator><cp:lastModifiedBy>jeremy.gerdes</cp:lastModifiedBy><dcterms:created xsi:type=""dcterms:W3CDTF"">2021-04-09T02:10:55Z</dcterms:created><dcterms:modified xsi:type=""dcterms:W3CDTF"">2021-04-09T02:11:30Z</dcterms:modified></cp:coreProperties>", _
        False, True
    MkDir strTmpZipPath & "\" & "xl"
    MkDir strTmpZipPath & "\" & "xl" & "\" & "theme"
    WriteToFile _
        strTmpZipPath & "\" & "xl" & "\" & "styles.xml", _
"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac x16r2"" xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"" xmlns:x16r2=""http://schemas.microsoft.com/office/spreadsheetml/2015/02/main""><fonts count=""1"" x14ac:knownFonts=""1""><font><sz val=""11""/><color theme=""1""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font></fonts><fills count=""2""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill></fills><borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs><cellXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/></cellXfs><cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>" & _
"<dxfs count=""0""/><tableStyles count=""0"" defaultTableStyle=""TableStyleMedium2"" defaultPivotStyle=""PivotStyleLight16""/><extLst><ext uri=""{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"" xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main""><x14:slicerStyles defaultSlicerStyle=""SlicerStyleLight1""/></ext><ext uri=""{9260A510-F301-46a8-8635-F512D64BE5F5}"" xmlns:x15=""http://schemas.microsoft.com/office/spreadsheetml/2010/11/main""><x15:timelineStyles defaultTimelineStyle=""TimeSlicerStyleLight1""/></ext></extLst></styleSheet>", _
        False, True
    WriteToFile _
        strTmpZipPath & "\" & "xl" & "\" & "workbook.xml", _
"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _ 
"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x15"" xmlns:x15=""http://schemas.microsoft.com/office/spreadsheetml/2010/11/main""><fileVersion appName=""xl"" lastEdited=""6"" lowestEdited=""6"" rupBuild=""14420""/><workbookPr defaultThemeVersion=""164011""/><mc:AlternateContent xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006""><mc:Choice Requires=""x15""><x15ac:absPath url=""\\snnsvr045\NNSY Perm\T&amp;I Lab\NNSY IT Assistant\Documentation\"" xmlns:x15ac=""http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac""/></mc:Choice></mc:AlternateContent><bookViews><workbookView xWindow=""0"" yWindow=""210"" windowWidth=""15495"" windowHeight=""6435""/></bookViews><sheets><sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/></sheets><calcPr calcId=""162913""/><extLst>" & _ 
"<ext uri=""{140A7094-0E35-4892-8432-C4D2E57EDEB5}"" xmlns:x15=""http://schemas.microsoft.com/office/spreadsheetml/2010/11/main""><x15:workbookPr chartTrackingRefBase=""1""/></ext></extLst></workbook>", _
        False, True
    MkDir strTmpZipPath & "\" & "xl" & "\" & "_rels"
    WriteToFile _
        strTmpZipPath & "\" & "xl" & "\" & "_rels" & "\" & "workbook.xml.rels", _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""theme/theme1.xml""/><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/></Relationships>", _
        False, True
    MkDir strTmpZipPath & "\" & "xl" & "\" & "theme"
    WriteToFile _
        strTmpZipPath & "\" & "xl" & "\" & "theme" & "\" & "theme1.xml", _
"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
"<a:theme xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" name=""Office Theme""><a:themeElements><a:clrScheme name=""Office""><a:dk1><a:sysClr val=""windowText"" lastClr=""000000""/></a:dk1><a:lt1><a:sysClr val=""window"" lastClr=""FFFFFF""/></a:lt1><a:dk2><a:srgbClr val=""44546A""/></a:dk2><a:lt2><a:srgbClr val=""E7E6E6""/></a:lt2><a:accent1><a:srgbClr val=""5B9BD5""/></a:accent1><a:accent2><a:srgbClr val=""ED7D31""/></a:accent2><a:accent3><a:srgbClr val=""A5A5A5""/></a:accent3><a:accent4><a:srgbClr val=""FFC000""/></a:accent4><a:accent5><a:srgbClr val=""4472C4""/></a:accent5><a:accent6><a:srgbClr val=""70AD47""/></a:accent6><a:hlink><a:srgbClr val=""0563C1""/></a:hlink><a:folHlink><a:srgbClr val=""954F72""/></a:folHlink></a:clrScheme><a:fontScheme name=""Office""><a:majorFont><a:latin typeface=""Calibri Light"" panose=""020F0302020204030204""/><a:ea typeface=""""/><a:cs typeface=""""/><a:font script=""Jpan"" typeface=""??????å? Light""/>" & _
"<a:font script=""Hang"" typeface=""?? ??""/><a:font script=""Hans"" typeface=""???? Light""/><a:font script=""Hant"" typeface=""?¼????w""/><a:font script=""Arab"" typeface=""Times New Roman""/><a:font script=""Hebr"" typeface=""Times New Roman""/><a:font script=""Thai"" typeface=""Tahoma""/><a:font script=""Ethi"" typeface=""Nyala""/><a:font script=""Beng"" typeface=""Vrinda""/><a:font script=""Gujr"" typeface=""Shruti""/><a:font script=""Khmr"" typeface=""MoolBoran""/><a:font script=""Knda"" typeface=""Tunga""/><a:font script=""Guru"" typeface=""Raavi""/><a:font script=""Cans"" typeface=""Euphemia""/><a:font script=""Cher"" typeface=""Plantagenet Cherokee""/><a:font script=""Yiii"" typeface=""Microsoft Yi Baiti""/><a:font script=""Tibt"" typeface=""Microsoft Himalaya""/><a:font script=""Thaa"" typeface=""MV Boli""/><a:font script=""Deva"" typeface=""Mangal""/><a:font script=""Telu"" typeface=""Gautami""/><a:font script=""Taml"" typeface=""Latha""/><a:font script=""Syrc"" typeface=""Estrangelo Edessa""/>" & _
"<a:font script=""Orya"" typeface=""Kalinga""/><a:font script=""Mlym"" typeface=""Kartika""/><a:font script=""Laoo"" typeface=""DokChampa""/><a:font script=""Sinh"" typeface=""Iskoola Pota""/><a:font script=""Mong"" typeface=""Mongolian Baiti""/><a:font script=""Viet"" typeface=""Times New Roman""/><a:font script=""Uigh"" typeface=""Microsoft Uighur""/><a:font script=""Geor"" typeface=""Sylfaen""/></a:majorFont><a:minorFont><a:latin typeface=""Calibri"" panose=""020F0502020204030204""/><a:ea typeface=""""/><a:cs typeface=""""/><a:font script=""Jpan"" typeface=""??????å?""/><a:font script=""Hang"" typeface=""?? ??""/><a:font script=""Hans"" typeface=""????""/><a:font script=""Hant"" typeface=""?¼????w""/><a:font script=""Arab"" typeface=""Arial""/><a:font script=""Hebr"" typeface=""Arial""/><a:font script=""Thai"" typeface=""Tahoma""/><a:font script=""Ethi"" typeface=""Nyala""/><a:font script=""Beng"" typeface=""Vrinda""/><a:font script=""Gujr"" typeface=""Shruti""/><a:font script=""Khmr"" typeface=""DaunPenh""/>" & _
"<a:font script=""Knda"" typeface=""Tunga""/><a:font script=""Guru"" typeface=""Raavi""/><a:font script=""Cans"" typeface=""Euphemia""/><a:font script=""Cher"" typeface=""Plantagenet Cherokee""/><a:font script=""Yiii"" typeface=""Microsoft Yi Baiti""/><a:font script=""Tibt"" typeface=""Microsoft Himalaya""/><a:font script=""Thaa"" typeface=""MV Boli""/><a:font script=""Deva"" typeface=""Mangal""/><a:font script=""Telu"" typeface=""Gautami""/><a:font script=""Taml"" typeface=""Latha""/><a:font script=""Syrc"" typeface=""Estrangelo Edessa""/><a:font script=""Orya"" typeface=""Kalinga""/><a:font script=""Mlym"" typeface=""Kartika""/><a:font script=""Laoo"" typeface=""DokChampa""/><a:font script=""Sinh"" typeface=""Iskoola Pota""/><a:font script=""Mong"" typeface=""Mongolian Baiti""/><a:font script=""Viet"" typeface=""Arial""/><a:font script=""Uigh"" typeface=""Microsoft Uighur""/><a:font script=""Geor"" typeface=""Sylfaen""/></a:minorFont></a:fontScheme>" & _
"<a:fmtScheme name=""Office""><a:fillStyleLst><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:lumMod val=""110000""/><a:satMod val=""105000""/><a:tint val=""67000""/></a:schemeClr></a:gs><a:gs pos=""50000""><a:schemeClr val=""phClr""><a:lumMod val=""105000""/><a:satMod val=""103000""/><a:tint val=""73000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:lumMod val=""105000""/><a:satMod val=""109000""/><a:tint val=""81000""/></a:schemeClr></a:gs></a:gsLst><a:lin ang=""5400000"" scaled=""0""/></a:gradFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:satMod val=""103000""/><a:lumMod val=""102000""/><a:tint val=""94000""/></a:schemeClr></a:gs><a:gs pos=""50000""><a:schemeClr val=""phClr""><a:satMod val=""110000""/><a:lumMod val=""100000""/><a:shade val=""100000""/></a:schemeClr></a:gs>" & _
"<a:gs pos=""100000""><a:schemeClr val=""phClr""><a:lumMod val=""99000""/><a:satMod val=""120000""/><a:shade val=""78000""/></a:schemeClr></a:gs></a:gsLst><a:lin ang=""5400000"" scaled=""0""/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=""6350"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:prstDash val=""solid""/><a:miter lim=""800000""/></a:ln><a:ln w=""12700"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:prstDash val=""solid""/><a:miter lim=""800000""/></a:ln><a:ln w=""19050"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:prstDash val=""solid""/><a:miter lim=""800000""/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst>" & _
"<a:outerShdw blurRad=""57150"" dist=""19050"" dir=""5400000"" algn=""ctr"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""63000""/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:solidFill><a:schemeClr val=""phClr""><a:tint val=""95000""/><a:satMod val=""170000""/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:tint val=""93000""/><a:satMod val=""150000""/><a:shade val=""98000""/><a:lumMod val=""102000""/></a:schemeClr></a:gs><a:gs pos=""50000""><a:schemeClr val=""phClr""><a:tint val=""98000""/><a:satMod val=""130000""/><a:shade val=""90000""/><a:lumMod val=""103000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""63000""/><a:satMod val=""120000""/></a:schemeClr></a:gs></a:gsLst>" & _
"<a:lin ang=""5400000"" scaled=""0""/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=""{05A4C25C-085E-4340-85A3-A5531E510DB2}""><thm15:themeFamily xmlns:thm15=""http://schemas.microsoft.com/office/thememl/2012/main"" name=""Office Theme"" id=""{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}"" vid=""{4A3C46E8-61CC-4603-A589-7422A47A8E4A}""/></a:ext></a:extLst></a:theme>", _
        False, True
    MkDir strTmpZipPath & "\" & "xl" & "\" & "worksheets"
    WriteToFile _
        strTmpZipPath & "\" & "xl" & "\" & "worksheets" & "\" & "sheet1.xml", _
"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac"" xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""><dimension ref=""A1""/><sheetViews><sheetView tabSelected=""1"" workbookViewId=""0""/></sheetViews><sheetFormatPr defaultRowHeight=""15"" x14ac:dyDescent=""0.25""/><sheetData/><pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/></worksheet>", _
        False, True
    ZipAFolder strTmpZipPath, _
        strLocalAppTmp  & "\" & "EmptyExcelFile.zip"
    moveFile strLocalAppTmp  & "\" & "EmptyExcelFile.zip", strNewExcelFile
End Sub 

'Dependancies NONE
'Version 1.0.0
'By jeremy.gerdes@navy.mil
Public Sub moveFile(strSourcePath, strDestinationPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Does the source file exist?
    If FileExists(strSourcePath) Then
        'If the destination is a folder then move the file into the folder preserving the source name
        If FolderExists(strDestinationPath) Then
             strDestinationPath = strDestinationPath & "\" & fso.GetFileName(strSourcePath)
        End If
        'If the destination allready exists, attempt to delete it... this move method allways over writes
        If FileExists(strDestinationPath) Then
           fso.DeleteFile strDestinationPath, True
        End If
        'If the destination path doesn't exist attempt to make it
        If Not FolderExists(fso.GetParentFolderName(strDestinationPath)) Then
            MkDir (fso.GetParentFolderName(strDestinationPath))
        End If
        'Move the file
        fso.moveFile strSourcePath, strDestinationPath
    End If
End Sub

Public Sub WriteToFile(ByRef strFileName, ByRef strContent, ByRef fOpenFile, ByRef fOverwrite)
    Dim tf ' As Object
    Dim FSO ' As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(strFileName) Then
        If fOverwrite Then
            FSO.DeleteFile strFileName
        End If
    End If
    Set tf = FSO.OpenTextFile(strFileName, 8, True)
    tf.WriteLine strContent
    tf.Close
    If fOpenFile Then
        OpenWithExplorer strFileName
    End If
    'Clean up
    Set tf = Nothing
    Set FSO = Nothing
End Sub

Sub ZipAFolder (sFolder, zipFile)
    'From /a/15143587/1146659
    With CreateObject("Scripting.FileSystemObject")
        zipFile = .GetAbsolutePathName(zipFile)
        sFolder = .GetAbsolutePathName(sFolder)

        With .CreateTextFile(zipFile, True)
            .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, chr(0))
        End With
    End With

    With CreateObject("Shell.Application")
        .NameSpace(zipFile).CopyHere .NameSpace(sFolder).Items

        Do Until .NameSpace(zipFile).Items.Count = _
                 .NameSpace(sFolder).Items.Count
            WScript.Sleep 200
        Loop
    End With

End Sub

Public Sub OpenWithExplorer(ByRef strFilePath)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("Explorer.exe " & strFilePath)
    Set wshShell = Nothing
End Sub

Public Function Environ(ByRef strName)
    'Replaces VBA.Envrion Public Function with wscript version for use in all VB engines
    Dim wshShell: Set wshShell = CreateObject("WScript.Shell")
    Dim strResult: strResult = wshShell.ExpandEnvironmentStrings("%" & strName & "%")
    'wshShell.ExpandEnvironmentStrings behaves differently than VBA.Environ when no environment variable is found,
    '  conforming all results to return nothing if no result was found, like VBA.Environ
    If strResult = "%" & strName & "%" Then
        Environ = vbNullString
    Else
        Environ = strResult
    End If
    'cleanup
    Set wshShell = Nothing
End Function

' Return true if file exists and false if file does not exist
Public Function FileExists(ByVal strPath) ' As String) As Boolean
Dim FSO 'As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FileExists = FSO.FileExists(strPath)
    ' Clean up
    Set FSO = Nothing
End Function

Public Function FolderExists(ByVal strPath) 'As String) As Boolean
Dim FSO ' As Object
    ' Note I used to use the vba.Dir Public Function but using that Public Function
    ' will lock the folder and prevent it from being deleted.
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = FSO.FolderExists(strPath)
    ' Clean up
    Set FSO = Nothing
End Function

Function MkDir(strPath)
    ' Version: 1.0.3
    ' Dependancies: NONE
    ' Returns: True if no errors, i.e. folder path allready existed, or was able to be created without errors
    ' Usage Example: MkDir Environ("temp") & "\" & "opsRunner"
    ' Emulates linux 'MkDir -p' command:  creates folders without complaining if it allready exists
    ' Superceeds the the VBA.MkDir function, but requires that drive be included in strPath
    ' By jeremy.gerdes@navy.mil
    Dim fso ' As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    If Not fso.FolderExists(strPath) Then
        On Error Resume Next
        Dim fRestore ' As Boolean
        fRestore = False
        'Handle Network Paths
        If Left(strPath, 2) = "\\" Then
            strPath = Right(strPath, Len(strPath) - 2)
            fRestore = True
        End If
        Dim arryPaths 'As Variant
        arryPaths = Split(strPath, "\")
        'Restore Server file path prefix
        If fRestore Then
            arryPaths(0) = "\\" & arryPaths(0)
        End If
        Dim intDir ' As Integer
        Dim strBuiltPath ' As String
        For intDir = LBound(arryPaths) To UBound(arryPaths)
            strBuiltPath = strBuiltPath & arryPaths(intDir) & "\"
            If Not fso.FolderExists(strBuiltPath) Then
                fso.CreateFolder strBuiltPath
            End If
        Next
    End If
    MkDir = (Err.Number = 0)
    'cleanup
    Set fso = Nothing
End Function

'Dependencies: NONE
'--- This function modified from: https://www.reddit.com/r/vba/comments/aom8xs/how_to_tell_if_script_is_running_in_vba_vbscript/ei0z2ll?utm_source=share&utm_medium=web2x&context=3
'--- Returns a string containing which script engine this is running in. Modified to test that wscript has a property 'Version' to allow for wscript emulation in the VBIDE
'--- Will return either "VBS","VBA", or "HTA".
Function ScriptEngine()
    On Error Resume Next
    ScriptEngine = "VBA"
    Dim tmp
    tmp = wscript.Version
    If Err.Number = 0 Then
        ScriptEngine = "VBS"
    End If
    Err.Clear
    ReDim window(0)
    If Err.Number = 501 Then
        ScriptEngine = "HTA"
    End If
    On Error GoTo 0
End Function

'Dependencies: ScriptEngine
Function GetCurrentFileFolder()
    Select Case ScriptEngine()
        Case "VBS"
            GetCurrentFileFolder = Left(wscript.ScriptFullName, Len(wscript.ScriptFullName) - Len(wscript.ScriptName) - 1)
        Case "VBA"
            Select Case Application.Name
                Case "Microsoft Word"
                    GetCurrentFileFolder = ThisDocument.Path
                Case "Microsoft Access"
                    GetCurrentFileFolder = CurrentProject.Path
                Case "Microsoft Excel"
                    GetCurrentFileFolder = ThisWorkbook.Path
            End Select
            'Not going to bother checking for other VBA contexts like powerpoint, visio, ms project or autocad.
        Case "HTA"
            '----------------------------------------------------------------------
            ' - There are several methods to get the current directory of the HTA -
            '----------------------------------------------------------------------
            'From testing don't use the following
            'This method drops the server path for network paths
            'strPath =  Left(Document.Location.pathname, InStrRev(Document.Location.pathname, "\") - 1) & "\Main.hta"
            'This method works fine if the HTA directly executed,
            'but if explorer.exe or cscript.exe executes the hta this method returns the %windir% dictory
            'Dim objScripShell
            'Set objScripShell = CreateObject("WScript.Shell")
            'strPath =  objScripShell.CurrentDirectory & "\Main.hta"
            Dim strPath
            strPath = jsUrlDecode(Document.Location.href)
            strPath = Replace(strPath, "/", "\")
            strPath = Left(strPath, InStrRev(strPath, "\") - 1)

            'URLs that begin with a drive letter will begin with 'file:\\\' Check this first
            If Left(strPath, 8) = "file:\\\" Then
            strPath = Right(strPath, Len(strPath) - 8)
            End If

            'URLs that begin with a server name will begin with 'file:'
            If Left(strPath, 5) = "file:" Then
            strPath = Right(strPath, Len(strPath) - 5)
            End If
            GetCurrentFileFolder = strPath
    End Select
End Function

BuildEmptyExcelFile GetCurrentFileFolder & "\" & "anEmptyExcelFile.xlsx"

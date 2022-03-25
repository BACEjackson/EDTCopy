Imports System.Data
Imports System.Data.OleDb
Imports ADODB

Module Module1
    Sub Main()
        Dim Region As String = ""
        Dim RegionInt As Integer = 0
        Dim ITEMName As String = ""
        Dim PrjCode As String = ""
        Dim OutDirectories(7) As String
        Dim UserAndPW(3, 2) As String
        Dim SheetMetal As String = "NO"

        ParseCommandLineArgs(Region, ITEMName, PrjCode, SheetMetal)
        GetSettings(OutDirectories, UserAndPW)

        RegionInt = Convert.ToInt16(Region)

        If Region > 3 Then
            Write_Item_to_DB("C:\EDT\Transfer\CHN\" & PrjCode & "\" & ITEMName & "\" & ITEMName & ".itm", "CHN")
            CopyCHN(OutDirectories(5), OutDirectories(4), OutDirectories(7), OutDirectories(6), PrjCode, ITEMName, SheetMetal, UserAndPW)
            Region = Region - 4
        End If
        If Region > 1 And SheetMetal <> "MC" Then
            Write_Item_to_DB("C:\EDT\Transfer\BEL\" & PrjCode & "\" & ITEMName & "\" & ITEMName & ".itm", "BEL")
            CopyBEL(OutDirectories(2), PrjCode, ITEMName, SheetMetal, UserAndPW)
            Region = Region - 2
        End If
        If Region = 1 Then
            Write_Item_to_DB("C:\EDT\Transfer\USA\" & PrjCode & "\" & ITEMName & "\" & ITEMName & ".itm", "USA")
            CopyUSA(OutDirectories(1), PrjCode, ITEMName, SheetMetal, UserAndPW)
        End If

    End Sub

    Sub ParseCommandLineArgs(ByRef Region As String, ByRef ITEMName As String, ByRef PrjCode As String, ByRef SheetMetal As String)
        Dim inputArgument As String = "/input="
        Dim inputName As String = ""
        Dim ArgumentArry As String()
        Region = ""
        ITEMName = ""


        'Region: USA = 1,BEL = 2,CHN = 4
        'the combination of different regions is the sum of the numbers ie: 3 is US and BEL, 5 is US and CHN, 7 is US,BEL and CHN

        For Each s As String In My.Application.CommandLineArgs
            If s.ToLower.StartsWith(inputArgument) Then
                inputName = s.Remove(0, inputArgument.Length)
            End If
        Next
        'If inputName = "" Then MsgBox("No Aurgements Passed")

        'inputName = "1|TEST09M5|EJJ-TEST|YES"

        ArgumentArry = Split(inputName, "|")

        Region = ArgumentArry(0)
        ITEMName = Trim(ArgumentArry(1))
        PrjCode = ArgumentArry(2)
        SheetMetal = ArgumentArry(3)

    End Sub

    Sub GetSettings(ByRef OutDirectories() As String, UserAndPW(,) As String)
        Dim SettingFile As String = "C:\Vault2020\Designs\CAD Support\Tools\BACToolscfg.txt"
        Dim SettingArray() As String
        Dim linCnt As Integer = 1
        Dim i As Integer = 1
        Try
            Dim FileReader = My.Computer.FileSystem.OpenTextFileReader(SettingFile)

            Do Until FileReader.EndOfStream
                ReDim Preserve SettingArray(i)
                SettingArray(i) = FileReader.ReadLine()
                i = i + 1
            Loop
            FileReader.Close()
            linCnt = i
        Catch ex As Exception
            MsgBox("Cannot Find Setting File.")
            Exit Sub
        End Try

        For i = 1 To linCnt - 1
            If Left(SettingArray(i), 8) = "USEDTdir" Then OutDirectories(1) = Mid(SettingArray(i), 12, Len(SettingArray(i)) - 11)
            If Left(SettingArray(i), 9) = "BELEDTdir" Then OutDirectories(2) = Mid(SettingArray(i), 13, Len(SettingArray(i)) - 12)
            If Left(SettingArray(i), 9) = "CHNERPdir" Then OutDirectories(4) = Mid(SettingArray(i), 13, Len(SettingArray(i)) - 12)
            If Left(SettingArray(i), 10) = "CHNSPRTdir" Then OutDirectories(5) = Mid(SettingArray(i), 14, Len(SettingArray(i)) - 13)
            If Left(SettingArray(i), 9) = "DALCAMdir" Then OutDirectories(6) = Mid(SettingArray(i), 13, Len(SettingArray(i)) - 12)
            If Left(SettingArray(i), 9) = "DALPPDdir" Then OutDirectories(7) = Mid(SettingArray(i), 13, Len(SettingArray(i)) - 12)
            If Left(SettingArray(i), 7) = "USA IP:" Then UserAndPW(1, 2) = Mid(SettingArray(i), 8, Len(SettingArray(i)) - 7)
            If Left(SettingArray(i), 7) = "BEL IP:" Then UserAndPW(2, 2) = Mid(SettingArray(i), 8, Len(SettingArray(i)) - 7)
            If Left(SettingArray(i), 7) = "CHN IP:" Then UserAndPW(3, 2) = Mid(SettingArray(i), 8, Len(SettingArray(i)) - 7)
            If Left(SettingArray(i), 9) = "USA USER:" And Len(SettingArray(i)) > 9 Then
                UserAndPW(1, 0) = Mid(SettingArray(i), 10, Len(SettingArray(i)) - 9)
            End If
            If Left(SettingArray(i), 13) = "USA PASSWORD:" And Len(SettingArray(i)) > 13 Then
                UserAndPW(1, 1) = Mid(SettingArray(i), 14, Len(SettingArray(i)) - 13)
            End If
            If Left(SettingArray(i), 9) = "BEL USER:" And Len(SettingArray(i)) > 9 Then
                UserAndPW(2, 0) = Mid(SettingArray(i), 10, Len(SettingArray(i)) - 9)
            End If
            If Left(SettingArray(i), 13) = "BEL PASSWORD:" And Len(SettingArray(i)) > 13 Then
                UserAndPW(2, 1) = Mid(SettingArray(i), 14, Len(SettingArray(i)) - 13)
            End If
            If Left(SettingArray(i), 9) = "CHN USER:" And Len(SettingArray(i)) > 9 Then
                UserAndPW(3, 0) = Mid(SettingArray(i), 10, Len(SettingArray(i)) - 9)
            End If
            If Left(SettingArray(i), 13) = "CHN PASSWORD:" And Len(SettingArray(i)) > 13 Then
                UserAndPW(3, 1) = Mid(SettingArray(i), 14, Len(SettingArray(i)) - 13)
            End If

        Next

    End Sub
    Private Sub CopyUSA(TargetDir As String, ProjectCode As String, ItemFile As String, SheetMetal As String, USERPW(,) As String)
        Dim SourceFolder As String = "C:\EDT\Transfer\USA\" & ProjectCode & "\" & ItemFile & "\"
        Dim TargetERP As String = TargetDir & ProjectCode & "\"
        Dim TargetPPD As String = TargetDir & ProjectCode & "\PPD\"
        Dim TargetCAM As String = TargetDir & ProjectCode & "\CAM\"
        Dim TargetIPT As String = TargetDir & ProjectCode & "\IPT\"
        Dim FileRoot As String = ""
        Dim MCFolder As String = ""
        Dim ErrorDesc As String = ""

        Try
            FileIO.FileSystem.CurrentDirectory = TargetERP
        Catch ex As Exception

            Try
                FileIO.FileSystem.CurrentDirectory = USERPW(1, 2)
            Catch ex2 As Exception
                Try
                    Shell("NET USE \\" & USERPW(1, 2) & " /USER:" & USERPW(1, 0) & " " & USERPW(1, 1))
                Catch ex1 As Exception
                    MsgBox("Could not log into " & TargetDir)
                    Exit Sub
                End Try
            End Try

            Try
                FileIO.FileSystem.CurrentDirectory = TargetDir
                FileIO.FileSystem.CreateDirectory(ProjectCode)
                FileIO.FileSystem.CurrentDirectory = TargetERP
                FileIO.FileSystem.CreateDirectory("PPD")
                FileIO.FileSystem.CreateDirectory("CAM")
                FileIO.FileSystem.CreateDirectory("IPT")
            Catch ext As Exception
                MsgBox("Could not access " & TargetDir)
                Exit Sub
            End Try
        End Try

        Try
            For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder)
                FileRoot = FileIO.FileSystem.GetName(foundfile)
                ErrorDesc = "Trying to copy " & FileRoot
                My.Computer.FileSystem.CopyFile(foundfile, TargetERP & FileRoot, True)
            Next

            If Mid(SourceFolder, Len(SourceFolder) - 2, 2) = "M3" Then
                MCFolder = Left(SourceFolder, Len(SourceFolder) - 3) & "MC\"
                Try
                    For Each foundfile As String In My.Computer.FileSystem.GetFiles(MCFolder)
                        FileRoot = FileIO.FileSystem.GetName(foundfile)
                        ErrorDesc = "Trying to copy " & FileRoot
                        My.Computer.FileSystem.CopyFile(foundfile, TargetERP & FileRoot, True)
                    Next
                    ErrorDesc = "Trying to delete " & MCFolder
                    My.Computer.FileSystem.DeleteDirectory(MCFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)
                Catch ex As Exception

                End Try
            End If

            If SheetMetal = "YES" Or SheetMetal = "FLAT" Then
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\CAM")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    ErrorDesc = "Trying to copy " & FileRoot
                    My.Computer.FileSystem.CopyFile(foundfile, TargetCAM & FileRoot, True)
                Next
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\IPT")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    ErrorDesc = "Trying to copy " & FileRoot
                    My.Computer.FileSystem.CopyFile(foundfile, TargetIPT & FileRoot, True)
                Next
            End If

            For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\PPD")
                FileRoot = FileIO.FileSystem.GetName(foundfile)
                ErrorDesc = "Trying to copy " & FileRoot
                My.Computer.FileSystem.CopyFile(foundfile, TargetPPD & FileRoot, True)
            Next

            ErrorDesc = "Trying to delete " & SourceFolder
            My.Computer.FileSystem.DeleteDirectory(SourceFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)

        Catch ex As Exception
            MsgBox("Could not Copy USA Drawing Data.")
            Exit Sub
        End Try

    End Sub
    Private Sub CopyBEL(TargetDir As String, ProjectCode As String, ItemFile As String, SheetMetal As String, USERPW(,) As String)
        Dim SourceFolder As String = "C:\EDT\Transfer\BEL\" & ProjectCode & "\" & ItemFile & "\"
        Dim TargetERP As String = TargetDir & ProjectCode & "\"
        Dim TargetPPD As String = TargetDir & ProjectCode & "\PPD\"
        Dim TargetCAM As String = TargetDir & ProjectCode & "\CAM\"
        Dim TargetSAT As String = TargetDir & ProjectCode & "\SAT\"
        Dim TargetIPT As String = TargetDir & ProjectCode & "\IPT\"
        Dim FileRoot As String = ""
        'Dim tempText As String()
        'Dim writeText As String()
        Dim ReadLine As String = ""
        Dim i As Integer = 1
        Dim h As Integer = 1
        'Dim linCnt As Integer
        'Dim found As Boolean
        Dim Enc As System.Text.Encoding = System.Text.Encoding.ASCII

        Try
            FileIO.FileSystem.CurrentDirectory = TargetERP
        Catch ex As Exception

            Try
                FileIO.FileSystem.CurrentDirectory = USERPW(2, 2)
            Catch ex2 As Exception
                Try
                    Shell("NET USE \\" & USERPW(2, 2) & " /USER:" & USERPW(2, 0) & " " & USERPW(2, 1))
                Catch ex1 As Exception
                    MsgBox("Could not log into " & TargetDir)
                    Exit Sub
                End Try
            End Try

            Try
                FileIO.FileSystem.CurrentDirectory = TargetDir
                FileIO.FileSystem.CreateDirectory(ProjectCode)
                FileIO.FileSystem.CurrentDirectory = TargetERP
                FileIO.FileSystem.CreateDirectory("PPD")
                FileIO.FileSystem.CreateDirectory("CAM")
                FileIO.FileSystem.CreateDirectory("SAT")
                FileIO.FileSystem.CreateDirectory("IPT")
            Catch ext As Exception
                MsgBox("Could not access " & TargetDir)
                Exit Sub
            End Try
        End Try

        Try

            For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder)
                FileRoot = FileIO.FileSystem.GetName(foundfile)
                My.Computer.FileSystem.CopyFile(foundfile, TargetERP & FileRoot, True)
            Next

            If SheetMetal <> "MC" Then
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\PPD")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    My.Computer.FileSystem.CopyFile(foundfile, TargetPPD & FileRoot, True)
                Next
            End If
            If SheetMetal = "YES" Or SheetMetal = "FLAT" Then
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\CAM")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    My.Computer.FileSystem.CopyFile(foundfile, TargetCAM & FileRoot, True)
                Next

                If SheetMetal = "YES" Then
                    For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\SAT")
                        FileRoot = FileIO.FileSystem.GetName(foundfile)
                        My.Computer.FileSystem.CopyFile(foundfile, TargetSAT & FileRoot, True)
                    Next
                End If

                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\IPT")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    My.Computer.FileSystem.CopyFile(foundfile, TargetIPT & FileRoot, True)
                Next
            End If

            My.Computer.FileSystem.DeleteDirectory(SourceFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)

        Catch ex As Exception
            MsgBox("Could not Copy Belgium Drawing Data.")
            Exit Sub
        End Try
    End Sub
    Private Sub CopyCHN(TargetDir As String, ERPDir As String, DALPPDdir As String, DALCAMdir As String,
                        ProjectCode As String, ItemFile As String, SheetMetal As String, USERPW(,) As String)
        Dim SourceFolder As String = "C:\EDT\Transfer\CHN\" & ProjectCode & "\" & ItemFile & "\"
        Dim TargetERP As String = ERPDir & ProjectCode & "\"
        'Dim TargetPPD As String = TargetDir & ProjectCode & "\PPDCN\"
        'Dim TargetCAM As String = TargetDir & ProjectCode & "\CAMCN\"
        'Dim TargetIPT As String = TargetDir & ProjectCode & "\IPTCN\"
        'Dim FileRoot As String = ""

        'Sept-15, 2015; Change output to match US PPD.
        Dim TargetPPD As String = TargetDir & "\PPD-EDT\" & Left(ItemFile, 2) & "\" & Mid(ItemFile, 3, 2) & "\"
        Dim TargetCAM As String = TargetDir & "\CAM-EDT\" & Left(ItemFile, 2) & "\" & Mid(ItemFile, 3, 2) & "\"
        Dim TargetIPT As String = TargetDir & "\IPT-EDT\" & Left(ItemFile, 2) & "\" & Mid(ItemFile, 3, 2) & "\"
        Dim TargetDALPPD As String = DALPPDdir & Left(ItemFile, 2) & "\" & Mid(ItemFile, 3, 2) & "\"
        Dim TargetDALCAM As String = DALCAMdir & Left(ItemFile, 2) & "\" & Mid(ItemFile, 3, 2) & "\"
        Dim FileRoot As String = ""

        Try
            FileIO.FileSystem.CurrentDirectory = TargetERP
        Catch ex As Exception
            Try
                FileIO.FileSystem.CurrentDirectory = USERPW(1, 2)
            Catch ex2 As Exception
                Try
                    Shell("NET USE \\" & USERPW(1, 2) & " /USER:" & USERPW(1, 0) & " " & USERPW(1, 1))
                Catch ex1 As Exception
                    MsgBox("Could not log into " & TargetDir)
                    Exit Sub
                End Try
            End Try

            Try
                FileIO.FileSystem.CurrentDirectory = ERPDir
                FileIO.FileSystem.CreateDirectory(ProjectCode)
            Catch ext As Exception
                MsgBox("Could not access " & TargetDir)
                Exit Sub
            End Try
        End Try

        Try
            FileIO.FileSystem.CurrentDirectory = TargetDir
        Catch ex As Exception

            Try
                FileIO.FileSystem.CurrentDirectory = USERPW(3, 2)
            Catch ex2 As Exception
                Try
                    Shell("NET USE \\" & USERPW(3, 2) & " /USER:" & USERPW(3, 0) & " " & USERPW(3, 1))
                Catch ex1 As Exception
                    MsgBox("Could not log into " & TargetDir)
                    Exit Sub
                End Try
            End Try

        End Try

        Try
            FileIO.FileSystem.CurrentDirectory = TargetPPD
        Catch ex As Exception
            Try
                FileIO.FileSystem.CurrentDirectory = TargetDir & "\PPD-EDT\" & Left(ItemFile, 2) & "\"
                FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
            Catch ex2 As Exception
                Try
                    FileIO.FileSystem.CurrentDirectory = TargetDir & "\PPD-EDT\"
                    FileIO.FileSystem.CreateDirectory(Left(ItemFile, 2))
                    FileIO.FileSystem.CurrentDirectory = TargetDir & "\PPD-EDT\" & Left(ItemFile, 2) & "\"
                    FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
                Catch ext As Exception
                    MsgBox("Could not access " & TargetPPD)
                    Exit Sub
                End Try
            End Try
        End Try

        Try
            FileIO.FileSystem.CurrentDirectory = TargetCAM
        Catch ex As Exception
            Try
                FileIO.FileSystem.CurrentDirectory = TargetDir & "\CAM-EDT\" & Left(ItemFile, 2) & "\"
                FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
            Catch ex2 As Exception
                Try
                    FileIO.FileSystem.CurrentDirectory = TargetDir & "\CAM-EDT\"
                    FileIO.FileSystem.CreateDirectory(Left(ItemFile, 2))
                    FileIO.FileSystem.CurrentDirectory = TargetDir & "\CAM-EDT\" & Left(ItemFile, 2) & "\"
                    FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
                Catch ext As Exception
                    MsgBox("Could not access " & TargetCAM)
                    Exit Sub
                End Try
            End Try
        End Try

        Try
            FileIO.FileSystem.CurrentDirectory = TargetIPT
        Catch ex As Exception
            Try
                FileIO.FileSystem.CurrentDirectory = TargetDir & "\IPT-EDT\" & Left(ItemFile, 2) & "\"
                FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
            Catch ex2 As Exception
                Try
                    FileIO.FileSystem.CurrentDirectory = TargetDir & "\IPT-EDT\"
                    FileIO.FileSystem.CreateDirectory(Left(ItemFile, 2))
                    FileIO.FileSystem.CurrentDirectory = TargetDir & "\IPT-EDT\" & Left(ItemFile, 2) & "\"
                    FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
                Catch ext As Exception
                    MsgBox("Could not access " & TargetIPT)
                    Exit Sub
                End Try
            End Try
        End Try

        'Try
        '    FileIO.FileSystem.CurrentDirectory = TargetDALPPD
        'Catch ex As Exception
        '    Try
        '        FileIO.FileSystem.CurrentDirectory = DALPPDdir & Left(ItemFile, 2) & "\"
        '        FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
        '    Catch ex2 As Exception
        '        Try
        '            FileIO.FileSystem.CurrentDirectory = DALPPDdir
        '            FileIO.FileSystem.CreateDirectory(Left(ItemFile, 2))
        '            FileIO.FileSystem.CurrentDirectory = DALPPDdir & Left(ItemFile, 2) & "\"
        '            FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
        '        Catch ext As Exception
        '            MsgBox("Could not access " & TargetDALPPD)
        '            Exit Sub
        '        End Try
        '    End Try
        'End Try

        'Try
        '    FileIO.FileSystem.CurrentDirectory = TargetDALCAM
        'Catch ex As Exception
        '    Try
        '        FileIO.FileSystem.CurrentDirectory = DALCAMdir & Left(ItemFile, 2) & "\"
        '        FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
        '    Catch ex2 As Exception
        '        Try
        '            FileIO.FileSystem.CurrentDirectory = DALCAMdir
        '            FileIO.FileSystem.CreateDirectory(Left(ItemFile, 2))
        '            FileIO.FileSystem.CurrentDirectory = DALCAMdir & Left(ItemFile, 2) & "\"
        '            FileIO.FileSystem.CreateDirectory(Mid(ItemFile, 3, 2))
        '        Catch ext As Exception
        '            MsgBox("Could not access " & TargetDALCAM)
        '            Exit Sub
        '        End Try
        '    End Try
        'End Try


        Try
            For Each foundfile As String In FileIO.FileSystem.GetFiles(SourceFolder)
                FileRoot = FileIO.FileSystem.GetName(foundfile)
                My.Computer.FileSystem.CopyFile(foundfile, TargetERP & FileRoot, True)
            Next

            If SheetMetal <> "MC" Then
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\PPD")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    My.Computer.FileSystem.CopyFile(foundfile, TargetPPD & FileRoot, True)
                    'My.Computer.FileSystem.CopyFile(foundfile, TargetDALPPD & FileRoot, True)
                Next
            End If

            If SheetMetal = "YES" Or SheetMetal = "FLAT" Then
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\CAM")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    My.Computer.FileSystem.CopyFile(foundfile, TargetCAM & FileRoot, True)
                    'My.Computer.FileSystem.CopyFile(foundfile, TargetDALCAM & FileRoot, True)
                Next
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(SourceFolder & "\IPT")
                    FileRoot = FileIO.FileSystem.GetName(foundfile)
                    My.Computer.FileSystem.CopyFile(foundfile, TargetIPT & FileRoot, True)
                Next
            End If

            My.Computer.FileSystem.DeleteDirectory(SourceFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)

        Catch ex As Exception
            MsgBox("Could not Copy China Drawing Data.")
            Exit Sub
        End Try

    End Sub
    Sub Write_Item_to_DB(Item_text_file As String, RegionCode As String)
        Dim Item_Text As String = ""

        Try
            Dim FileReader = My.Computer.FileSystem.OpenTextFileReader(Item_text_file)
            Do Until FileReader.EndOfStream
                Item_Text = FileReader.ReadLine()
            Loop
            FileReader.Close()
        Catch ex As Exception
            Exit Sub
        End Try

        Dim EDT_DB As ADODB.Connection
        Dim ConnString As String
        Dim Insert_string As String

        Insert_string = "INSERT INTO EDT_INFO (PartNo, Description, ItemKey, ItemType, DrawingSize, Revision, SourceCode, UOM," &
                        "CutLength, CutWidth, Make, Gauge, PurchaseCat, FunctionCat, AccessoryCat, ProductCat, ECONum, [User], RegionCode, Release_Date) VALUES ('" _
            & Left(Item_Text, 18) & "','" _
            & Mid(Item_Text, 19, 30) & "','" _
            & Mid(Item_Text, 49, 5) & "','" _
            & Mid(Item_Text, 54, 4) & "','" _
            & Mid(Item_Text, 59, 2) & "','" _
            & Mid(Item_Text, 61, 2) & "','" _
            & Mid(Item_Text, 63, 1) & "','" _
            & Mid(Item_Text, 64, 3) & "','" _
            & Mid(Item_Text, 67, 15) & "','" _
            & Mid(Item_Text, 83, 15) & "','" _
            & Mid(Item_Text, 99, 3) & "','" _
            & Mid(Item_Text, 102, 2) & "','" _
            & Mid(Item_Text, 104, 18) & "','" _
            & Mid(Item_Text, 122, 2) & "','" _
            & Mid(Item_Text, 124, 2) & "','" _
            & Mid(Item_Text, 126, 2) & "','" _
            & Mid(Item_Text, 131, 10) & "','" _
            & Mid(Item_Text, 141, 14) & "','" _
            & RegionCode & "','" & Now.Date & "')"

        ConnString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=\\jessupnfs.bac.amsted\OUTPUT\ENG\EDT_log.accdb;Persist Security Info=False"
        Try
            EDT_DB = New ADODB.Connection
            EDT_DB.ConnectionString = ConnString
            EDT_DB.Open()
            EDT_DB.Execute(Insert_string)
            EDT_DB.Close()
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub
End Module

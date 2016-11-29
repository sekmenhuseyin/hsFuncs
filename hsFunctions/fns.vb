Imports Effortless.Net.Encryption
''' <summary>
''' Frequently  used functions cobined in a single dll
''' </summary>
Public Class fns
    Private Shared alphabet() As String = {"£", "锋", "ਖ", "ਕ", "메", "导", "ن", "%", "马", "哦", "子", "!", "\", "s", "吉", "ն", "ք", "펜", "ط", "ש", "ਦ", "아", "ե", "땅", "快", "力", "+", "과", "疯", "ц", "'", "电", "ד", "z", "女", "孔", "带", "x", "ס", "爱", "$", "鼓", "ո", "ו", "կ", "ز", ",", "坦", "冻", "د", "艾", "침", "ղ", "얼", "ਅ", "ر", "녀", "յ", "류", "笔", "娜", "h", "y", "ը", "п", "բ", "ט", "迪", "ئ", "脑", "ب", "ռ", "פ", "크", "弗", "ş", "男", "n", "ğ", "}", "ਗ", "教", "提", "ћ", "€", "տ", "람", "时", "克", "贝", "f", "v", "դ", "结", "6", "콩", "a", "ਲ", "1", "勒", "屁", "^", "д", "k", "吾", "o", "ר", "ث", "ش", "ਫ", "غ", "图", "ت", "饶", "|", "筆", "9", "维", "ָ", "购", "p", "拉", "ր", "午", "花", "ص", "屄", "3", "贼", "ਤ", "소", "g", "晚", "밤", "亲", "շ", "ם", "و", "报", "术", "姆", "订", "片", "影", "能", "铁", "5", "_", "ı", "u", "일", "西", "索", "比", "d", "ਧ", "ւ", "和", "蕪", "ا", "2", "₺", "ਜ", "脚", "丝", "t", "א", "#", "ս", "德", "վ", "굴", "m", "院", "堅", "儿", "ի", "0", "校", "գ", "斯", "ਰ", "報", "ਬ", "=", "w", "豆", "ح", "ض", "尺", "上", "ؤ", "ي", "ջ", "高", "开", ";", "的", "平", "家", "é", "ע", "칼", "览", "լ", "水", "ա", "字", "服", "탱", "ч", "ק", "저", "肖", "ਨ", "性", "م", "白", "后", "面", "ਹ", "ਾ", "展", ">", "자", "室", "q", "c", "ء", "س", "果", "զ", "*", "ਿ", "名", "ਸ", "b", "&", "i", "?", "մ", "ه", "4", "孩", "小", "ف", "锈", "色", "天", "刻", "j", "ੀ", "/", "ਮ", "남", "冰", "魔", "伊", "ق", "른", "[", "ظ", "人", "ਚ", "{", "л", "]", "7", "头", "r", "夜", "月", "剧", "l", "з", "ö", "像", "չ", "마", "½", "前", "生", "ל", "线", "ç", "и", "尔", ")", "e", "녁", "手", "ج", "ג", "(", "穿", "裙", "师", "ل", ".", "ך", "<", "ü", "8", "情", "월", "诶", "杰", "别", "恕", "ى", "ع", "خ", "견", "荒", "卡", "נ", "演", "学"}
    Private Shared Generator As Random = New Random(Now.Millisecond)
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32
    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    ''' <summary>
    ''' Convert calendar types
    ''' </summary>
    Public Shared Function ConvertDateCalendar(ByVal DateConv As DateTime, ByVal Calendar As String, ByVal DateLangCulture As String) As String
        Dim DTFormat As DateTimeFormatInfo
        DateLangCulture = DateLangCulture.ToLower()
        '' We can't have the hijri date writen in English. We will get a runtime error - LAITH - 11/13/2005 1:01:45 PM -
        If Calendar = "Hijri" AndAlso DateLangCulture.StartsWith("en-") Then
            DateLangCulture = "ar-sa"
        End If
        '' Set the date time format to the given culture - LAITH - 11/13/2005 1:04:22 PM -
        DTFormat = New CultureInfo(DateLangCulture, False).DateTimeFormat
        '' Set the calendar property of the date time format to the given calendar - LAITH - 11/13/2005 1:04:52 PM -
        Select Case Calendar
            Case "Hijri"
                DTFormat.Calendar = New HijriCalendar()
                Exit Select
            Case "Gregorian"
                DTFormat.Calendar = New GregorianCalendar()
                Exit Select
            Case Else
                Return ""
        End Select
        '' We format the date structure to whatever we want - LAITH - 11/13/2005 1:05:39 PM -
        DTFormat.ShortDatePattern = "dd/MM/yyyy"
        'https://msdn.microsoft.com/en-us/library/8tfzyc64(v=vs.110).aspx
        Return (DateConv.[Date].ToString("D", DTFormat))
    End Function
    Public Enum enumCalendars
        Gregorian
        Hijri
    End Enum
    ''' <summary>
    ''' All input must be validated against a whitelist of acceptable value ranges.
    ''' </summary>
    Public Shared Function ValidateInputs(Str As String) As String
        Str = Str.Replace("!", "")
        Str = Str.Replace("'", "")
        Str = Str.Replace("(", "")
        Str = Str.Replace(")", "")
        Str = Str.Replace(";", "")
        Str = Str.Replace(":", "")
        Str = Str.Replace("@", "")
        Str = Str.Replace("&", "")
        Str = Str.Replace("=", "")
        Str = Str.Replace("+", "")
        Str = Str.Replace("$", "")
        Str = Str.Replace(",", "")
        Str = Str.Replace("/", "")
        Str = Str.Replace("?", "")
        Str = Str.Replace("%", "")
        Str = Str.Replace("#", "")
        Str = Str.Replace("[", "")
        Str = Str.Replace("]", "")
        Return Str.Trim
    End Function
    ''' <summary>
    ''' clear non latin characters from a string
    ''' </summary>
    Public Shared Function ValidateNonLatins(Str As String) As String
        Str = Str.Replace("Ö", "O")
        Str = Str.Replace("ö", "o")
        Str = Str.Replace("İ", "I")
        Str = Str.Replace("ı", "i")
        Str = Str.Replace("Ş", "S")
        Str = Str.Replace("ş", "s")
        Str = Str.Replace("Ğ", "G")
        Str = Str.Replace("ğ", "g")
        Str = Str.Replace("Ü", "U")
        Str = Str.Replace("ü", "u")
        Str = Str.Replace("Ç", "C")
        Str = Str.Replace("ç", "c")
        Return Str.Trim
    End Function
    ''' <summary>
    ''' Auto adjust dropdownwidth for a combobox, according to longest combobox item
    ''' </summary>
    Public Shared Sub AdjustComboBoxWidth(senderComboBox As ComboBox)
        Dim width As Integer = senderComboBox.DropDownWidth
        Dim g As Graphics = senderComboBox.CreateGraphics()
        Dim font As Font = senderComboBox.Font

        Dim vertScrollBarWidth As Integer = If((senderComboBox.Items.Count > senderComboBox.MaxDropDownItems), SystemInformation.VerticalScrollBarWidth, 0)

        Dim newWidth As Integer
        For Each s As String In senderComboBox.Items
            newWidth = CInt(g.MeasureString(s, font).Width) + vertScrollBarWidth
            If width < newWidth Then
                width = newWidth
            End If
        Next
        'change width
        senderComboBox.DropDownWidth = width
    End Sub
    ''' <summary>
    ''' Count a character in a string
    ''' </summary>
    Public Shared Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Return value.Count(Function(c As Char) c = ch)
    End Function
    ''' <summary>
    ''' Finds 
    ''' </summary>
    Public Shared Function GetIPAddress() As String
        Dim ipAddresss = String.Empty
        Dim strHostName As String = Dns.GetHostName()
        Dim iphe As IPHostEntry = Dns.GetHostEntry(strHostName)
        'check all ips
        For Each ipheal As IPAddress In iphe.AddressList
            If ipheal.AddressFamily = Sockets.AddressFamily.InterNetwork Then ipAddresss = ipheal.ToString()
        Next
        Return ipAddresss
    End Function
    ''' <summary>
    ''' Converts first letters of each word to uppercase
    ''' </summary>
    Public Shared Function TitleCase(str As String) As String
        Return StrConv(Trim(str), VbStrConv.ProperCase)
    End Function
    ''' <summary>
    ''' Removes an item from array
    ''' </summary>
    Public Shared Sub RemoveItemFromArray(Of T)(ByRef arrayName As T(), ByVal index As Integer)
        'stackoverflow.com/questions/3448103/how-can-i-delete-an-item-from-an-array-in-vb-net
        Dim uBound = arrayName.GetUpperBound(0)
        Dim lBound = arrayName.GetLowerBound(0)
        Dim arrLen = uBound - lBound

        If index < lBound OrElse index > uBound Then
            Throw New ArgumentOutOfRangeException(
            String.Format("Index must be from {0} to {1}.", lBound, uBound))

        Else
            'create an array 1 element less than the input array
            Dim outArr(arrLen - 1) As T
            'copy the first part of the input array
            Array.Copy(arrayName, 0, outArr, 0, index)
            'then copy the second part of the input array
            Array.Copy(arrayName, index + 1, outArr, index, uBound - index)

            arrayName = outArr
        End If
    End Sub
    ''' <summary>
    ''' To paste in datagridview
    ''' </summary>
    Public Shared Sub PasteUnboundRecords(dgv As DataGridView)
        Try
            Dim rowLines As String() = Clipboard.GetText(TextDataFormat.Text).Split(New String(0) {vbCr & vbLf}, StringSplitOptions.None)
            Dim currentRowIndex As Integer = (If(dgv.CurrentRow IsNot Nothing, dgv.CurrentRow.Index, 0))
            Dim currentColumnIndex As Integer = (If(dgv.CurrentCell IsNot Nothing, dgv.CurrentCell.ColumnIndex, 0))
            Dim currentColumnCount As Integer = dgv.Columns.Count

            'eksik satırları ekle
            If rowLines.Count > dgv.RowCount - 1 Then
                For index = dgv.RowCount To rowLines.Count - 1
                    dgv.Rows.Add()
                Next
            End If
            'yeni satır eklenemesin
            dgv.AllowUserToAddRows = False
            For rowLine As Integer = 0 To rowLines.Length - 1

                If rowLine = rowLines.Length - 1 AndAlso String.IsNullOrEmpty(rowLines(rowLine)) Then
                    Exit For
                End If

                Dim columnsData As String() = rowLines(rowLine).Split(New String(0) {vbTab}, StringSplitOptions.None)
                If (currentColumnIndex + columnsData.Length) > dgv.Columns.Count Then
                    For columnCreationCounter As Integer = 0 To ((currentColumnIndex + columnsData.Length) - currentColumnCount) - 1
                        If columnCreationCounter = rowLines.Length - 1 Then
                            Exit For
                        End If
                    Next
                End If
                If dgv.Rows.Count > (currentRowIndex + rowLine) Then
                    For columnsDataIndex As Integer = 0 To columnsData.Length - 1
                        If currentColumnIndex + columnsDataIndex <= dgv.Columns.Count - 1 Then
                            dgv.Rows(currentRowIndex + rowLine).Cells(currentColumnIndex + columnsDataIndex).Value = columnsData(columnsDataIndex)
                        End If
                    Next
                Else
                    Dim pasteCells As String() = New String(dgv.Columns.Count - 1) {}
                    For cellStartCounter As Integer = currentColumnIndex To dgv.Columns.Count - 1
                        If columnsData.Length > (cellStartCounter - currentColumnIndex) Then
                            pasteCells(cellStartCounter) = columnsData(cellStartCounter - currentColumnIndex)
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Gets all menus and submenus in a single list
    ''' </summary>
    Public Shared Sub GetSubMenus2List(menu As MenuStrip, menues As List(Of ToolStripItem))
        For Each t As ToolStripItem In menu.Items
            GetMenues(t, menues) 'alt menüleri de alsın diye iç içe bir döngü
        Next
    End Sub
    Private Shared Sub GetMenues(ByVal Current As ToolStripItem, ByRef menues As List(Of ToolStripItem))
        If Current.Name.StartsWith("ToolStripMenuItem") = False Then menues.Add(Current) 'not the seperators
        If TypeOf (Current) Is ToolStripMenuItem Then
            For Each menu As ToolStripItem In DirectCast(Current, ToolStripMenuItem).DropDownItems
                GetMenues(menu, menues)
            Next
        End If
    End Sub
    ''' <summary>
    ''' Truky validate an email address
    ''' </summary>
    Public Shared Function isEmail(email As String) As Boolean
        'info from: en.wikipedia.org/wiki/Email_address
        'ilk kontroller
        If Len(email) > 320 Then Return False '320 karakterden uzun olamaz
        If InStr(email, "..") > 0 Then Return False 'iki nokta asla yanyana olamaz
        If InStr(email, "@") = 0 Then Return False 'en az bir @ işareti olmalı
        If email.Length <> email.Trim().Length Then Return False 'başta ve sonda boşluk olamaz
        If Left(email, 1) = "." Or Right(email, 1) = "." Then Return False 'başta ve sonda nokta olamaz
        'find position of @
        Dim qPos As Integer
        For i As Integer = 0 To email.Length - 1
            If email.Substring(i, 1) = "@" Then qPos = i
        Next
        'local and domain parts
        Dim localPart As String = email.Substring(0, qPos)
        Dim domainPart As String = email.Substring(qPos + 1)
        Dim forbidden As String = "%# (),:;<>@[\]" : Dim tmp As String
        ''''''''''''''local part control
        If Len(localPart) > 64 Then Return False 'local part uzunluk kontrolü
        'eğer komple tırnak içindeyse
        If Left(localPart, 1) = """" And Right(localPart, 1) = """" Then
        Else 'tırnak yoksa normal kurallar
            If InStr(localPart, "@") > 0 Then Return False '@ işareti olmamalı
            If InStr(localPart, """") > 0 Then Return False '"" işareti olmamalı
            If Right(localPart, 1) = "." Then Return False 'sonda nokta olamaz
            'yasak karakter kullanımı
            tmp = localPart
            For i As Integer = 0 To forbidden.Length - 1 : tmp = tmp.Replace(forbidden.Substring(i, 1), "") : Next
            If tmp <> localPart Then Return False
        End If
        '''''''''''domain part control
        'uzunluk kontrol
        If Len(domainPart) < 3 Or Len(domainPart) > 64 Then Return False
        If InStr(domainPart, """") > 0 Then Return False '"" işareti olmamalı
        If InStr(domainPart, "*") > 0 Then Return False '* işareti olmamalı
        If Left(domainPart, 1) = "." Then Return False 'başta nokta olamaz
        'yasak karakter kullanımı
        tmp = domainPart
        For i As Integer = 0 To forbidden.Length - 1 : tmp = tmp.Replace(forbidden.Substring(i, 1), "") : Next
        If tmp <> domainPart Then Return False
        'if passes every control then return true
        Return True
    End Function
    ''' <summary>
    ''' Gets website sourcedata from a given url
    ''' </summary>
    Public Shared Function GetWebResponse(url As String) As String
        Dim request As WebRequest = WebRequest.Create(url)
        ' If required by the server, set the credentials.
        request.Credentials = CredentialCache.DefaultCredentials
        ' Get the response.
        Dim response As WebResponse = request.GetResponse()
        ' Display the status.
        'Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
        ' Get the stream containing content returned by the server.
        Dim dataStream As Stream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim reader As New StreamReader(dataStream)
        ' Read the content.
        Dim responseFromServer As String = reader.ReadToEnd()
        ' Clean up the streams and the response.
        reader.Close()
        response.Close()
        Return responseFromServer
    End Function
    ''' <summary>
    ''' Randomizes an array
    ''' </summary>
    Public Shared Sub RandomizeArray(ByVal items() As String)
        Dim max_index As Integer = items.Length - 1
        Dim rnd As New Random
        For i As Integer = 0 To max_index - 1
            ' Pick an item for position i.
            Dim j As Integer = rnd.Next(i, max_index + 1)
            ' Swap them.
            Dim temp As String = items(i)
            items(i) = items(j)
            items(j) = temp
        Next i
    End Sub
    ''' <summary>
    ''' Calculate time difference
    ''' </summary>
    Public Shared Function DeltaTime(var1 As DateTime, var2 As DateTime) As TimeSpan
        Dim duration As TimeSpan
        If var2 > var1 Then duration = var2 - var1 Else duration = var1 - var2
        Return duration
    End Function
    ''' <summary>
    ''' Creates a random number
    ''' </summary>
    Public Shared Function GetRandomNumber(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Return Generator.Next(Min, Max)
    End Function
    ''' <summary>
    ''' returns file size in readable format
    ''' </summary>
    Public Shared Function GetFileSize(ByVal TheFile As String) As String
        If TheFile.Length = 0 Then Return ""
        If Not File.Exists(TheFile) Then Return ""
        '---
        Dim TheSize As ULong = My.Computer.FileSystem.GetFileInfo(TheFile).Length
        Dim SizeType As String = ""
        Dim DoubleBytes As Double
        '---
        Try
            Select Case TheSize
                Case Is >= 1099511627776
                    DoubleBytes = CDbl(TheSize / 1099511627776) 'TB
                    Return FormatNumber(DoubleBytes, 2) & " TB"
                Case 1073741824 To 1099511627775
                    DoubleBytes = CDbl(TheSize / 1073741824) 'GB
                    Return FormatNumber(DoubleBytes, 2) & " GB"
                Case 1048576 To 1073741823
                    DoubleBytes = CDbl(TheSize / 1048576) 'MB
                    Return FormatNumber(DoubleBytes, 2) & " MB"
                Case 1024 To 1048575
                    DoubleBytes = CDbl(TheSize / 1024) 'KB
                    Return FormatNumber(DoubleBytes, 2) & " KB"
                Case 0 To 1023
                    DoubleBytes = TheSize ' bytes
                    Return FormatNumber(DoubleBytes, 2) & " bytes"
                Case Else
                    Return ""
            End Select
        Catch
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Write information to a INI file
    ''' </summary>
    Public Shared Sub INI_Write(ByVal iniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamVal As String)
        Dim Result As Integer = WritePrivateProfileString(Section, ParamName, ParamVal, iniFileName)
    End Sub
    ''' <summary>
    ''' Read information from a INI file
    ''' </summary>
    Public Shared Function INI_Read(ByVal IniFileName As String, ByVal Section As String, ByVal ParamName As String, Optional ByVal ParamDefault As String = "") As String
        Dim ParamVal As String = Space$(1024)
        Dim LenParamVal As Integer = GetPrivateProfileString(Section, ParamName, ParamDefault, ParamVal, Len(ParamVal), IniFileName)
        Return Left$(ParamVal, LenParamVal)
    End Function
    ''' <summary>
    ''' Unzip using Shell32
    ''' </summary>
    Public Shared Sub UnZip(outputFolder As String, inputZip As String)
        Dim shObj As Object = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"))
        'codeproject.com/Tips/257193/Easily-zip-unzip-files-using-Windows-Shell32
        'msdn.microsoft.com/en-us/library/windows/desktop/bb787866(v=vs.85).aspx
        'Declare the folder where the items will be extracted.
        Dim output As Object = shObj.NameSpace((outputFolder))
        'Declare the input zip file.
        Dim input As Object = shObj.NameSpace((inputZip))
        'Extract the items from the zip file.
        '4=Do not display a progress dialog box.
        '16=Respond with "Yes to All" for any dialog box that is displayed.
        output.CopyHere((input.Items), 20) '20=4+16
    End Sub
    ''' <summary>
    ''' Encrypt with salt
    ''' </summary>
    Public Shared Function Encrypt(txt As String, salt As String) As String
        If txt = "" Then Return ""
        If salt = "" Then salt = Application.CompanyName
        If Len(salt) < 8 Then Return ""
        Return Strings.Encrypt(txt, Application.CompanyName, salt, String.Empty.PadLeft(32, "#"), Bytes.KeySize.Size256)
    End Function
    ''' <summary>
    ''' Decrypt with salt
    ''' </summary>
    Public Shared Function Decrypt(txt As String, salt As String) As String
        If txt = "" Then Return ""
        If salt = "" Then salt = Application.CompanyName
        If Len(salt) < 8 Then Return ""
        Return Strings.Decrypt(txt, Application.CompanyName, salt, String.Empty.PadLeft(32, "#"), Bytes.KeySize.Size256)
    End Function
    ''' <summary>
    ''' Generates hash with most secure way as possible
    ''' </summary>
    Public Shared Function GenerateHash(txt As String, salt As String) As String
        If txt = "" Then Return ""
        If salt = "" Then salt = Application.CompanyName
        Return Hash.Create(HashType.SHA512, txt, salt, True)
    End Function
    ''' <summary>
    ''' Encrpyts without salt
    ''' </summary>
    Public Shared Function EncryptAscii(txt As String) As String
        If txt = "" Then Return ""
        Dim ascii, shift, uzunluk As Integer : Dim crypted As String
        'rastgele shift sayısını üret
        shift = GetRandomNumber(-99, +99)
        If shift < -9 Then
            crypted = alphabet(Asc("-")) & alphabet(Asc(shift.ToString().Substring(1, 1))) & alphabet(Asc(shift.ToString().Substring(2, 1)))
        ElseIf shift < 0 Then
            crypted = alphabet(Asc("-")) & alphabet(Asc("0")) & alphabet(Asc(shift.ToString().Substring(1, 1)))
        ElseIf shift < 10 Then
            crypted = alphabet(Asc("+")) & alphabet(Asc("0")) & alphabet(Asc(shift))
        Else
            crypted = alphabet(Asc("+")) & alphabet(Asc(shift.ToString().Substring(0, 1))) & alphabet(Asc(shift.ToString().Substring(1, 1)))
        End If
        'her karakteri alphabetteki harf ile değiştir
        uzunluk = txt.Length
        For index = 0 To uzunluk - 1
            'ilk önce seçili harfin ascii kodunu buluyoruz
            ascii = Asc(txt.Substring(index, 1)) + shift
            If ascii < 0 Then ascii += alphabet.Length Else If ascii >= alphabet.Length Then ascii -= alphabet.Length
            crypted &= alphabet(ascii)
            'şimdide ratgele harfi ekliyoruz
            Dim tmp As Integer = GetRandomNumber(0, alphabet.Length - 1)
            crypted &= alphabet(tmp)
        Next
        Return crypted
    End Function
    ''' <summary>
    ''' Decrypts without salt
    ''' </summary>
    Public Shared Function DecryptAscii(txt As String) As String
        If txt = "" Then Return ""
        Dim ascii, shift, uzunluk As Integer : Dim decrypted As String = ""
        'ilk önce shifti bul
        For i = 0 To 2
            For j = 0 To alphabet.Length - 1
                If txt.Substring(i, 1) = alphabet(j) Then ascii = j : Exit For
            Next
            If ascii < 0 Then ascii += alphabet.Length Else If ascii >= alphabet.Length Then ascii -= alphabet.Length
            decrypted &= Chr(ascii)
        Next
        shift = Convert.ToInt32(decrypted) : decrypted = ""
        txt = txt.Remove(0, 3)
        uzunluk = txt.Length
        For i = 0 To uzunluk - 1
            If i Mod 2 = 0 Then
                For j = 0 To alphabet.Length - 1
                    If txt.Substring(i, 1) = alphabet(j) Then ascii = j - shift : Exit For
                Next
                If ascii < 0 Then ascii += alphabet.Length Else If ascii >= alphabet.Length Then ascii -= alphabet.Length
                decrypted &= Chr(ascii)
            End If
        Next
        Return decrypted
    End Function
    ''' <summary>
    ''' Upload file via FTP
    ''' </summary>
    Public Shared Function FtpFileUpload(localFilePath As String, remoteFilePath As String, FtpAddress As String, FtpUser As String, FtpPass As String) As Boolean
        Try
            Dim Ftp As FtpWebRequest = DirectCast(WebRequest.Create(FtpAddress & remoteFilePath), FtpWebRequest)
            Ftp.Proxy = Nothing
            Ftp.Credentials = New NetworkCredential(FtpUser, FtpPass)
            Ftp.Method = WebRequestMethods.Ftp.UploadFile
            Ftp.UseBinary = True
            Dim fs As FileStream = File.OpenRead(localFilePath)
            Dim buffer As Byte() = New Byte(CInt(fs.Length) - 1) {}
            fs.Read(buffer, 0, buffer.Length)
            fs.Close()
            'klasör yoksa hata veriyor
            Dim ftpstream As Stream = Ftp.GetRequestStream()
            ftpstream.Write(buffer, 0, buffer.Length)
            ftpstream.Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Download file via FTP
    ''' </summary>
    Public Shared Function FtpFileDownload(remoteFilePath As String, localFilePath As String, FtpAddress As String, FtpUser As String, FtpPass As String) As Boolean
        Try
            Dim Ftp As FtpWebRequest = DirectCast(WebRequest.Create(FtpAddress & remoteFilePath), FtpWebRequest)
            Ftp.Proxy = Nothing
            Ftp.Credentials = New NetworkCredential(FtpUser, FtpPass)
            Ftp.Method = WebRequestMethods.Ftp.DownloadFile
            Dim response As FtpWebResponse = DirectCast(Ftp.GetResponse(), FtpWebResponse)
            Dim responseStream As Stream = Ftp.GetResponse.GetResponseStream()
            Dim fileStream As New FileStream(localFilePath, FileMode.OpenOrCreate)
            Dim buffer As [Byte]() = New [Byte](29) {}
            Dim read As Integer = 0
            Do
                read = responseStream.Read(buffer, 0, buffer.Length)
                fileStream.Write(buffer, 0, read)
            Loop While read <> 0
            responseStream.Close()
            fileStream.Flush()
            fileStream.Close()
            Return True
        Catch ex As WebException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Delete a file in FTP
    ''' </summary>
    Public Shared Function FtpFileDelete(remoteFilePath As String, FtpAddress As String, FtpUser As String, FtpPass As String) As Boolean
        Try
            Dim Ftp As FtpWebRequest = DirectCast(WebRequest.Create(FtpAddress & remoteFilePath), FtpWebRequest)
            Ftp.Proxy = Nothing
            Ftp.Credentials = New NetworkCredential(FtpUser, FtpPass)
            Ftp.Method = WebRequestMethods.Ftp.DeleteFile
            Dim response As FtpWebResponse = DirectCast(Ftp.GetResponse(), FtpWebResponse)
            response = CType(Ftp.GetResponse(), FtpWebResponse)
            response.Close()
            Return True
        Catch ex As WebException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Rename file in FTP
    ''' </summary>
    Public Shared Function FtpFileRename(ByVal remoteFilePath As String, ByVal remoteFileNewPath As String, FtpAddress As String, FtpUser As String, FtpPass As String) As Boolean
        Try
            Dim reqFTP As FtpWebRequest = Nothing
            Dim ftpStream As Stream = Nothing
            reqFTP = DirectCast(FtpWebRequest.Create(New Uri(FtpAddress & remoteFilePath)), FtpWebRequest)
            reqFTP.Method = WebRequestMethods.Ftp.Rename
            reqFTP.RenameTo = remoteFileNewPath
            reqFTP.UseBinary = True
            reqFTP.KeepAlive = False
            reqFTP.Credentials = New NetworkCredential(FtpUser, FtpPass)
            Dim response As FtpWebResponse = DirectCast(reqFTP.GetResponse(), FtpWebResponse)
            ftpStream = response.GetResponseStream()
            ftpStream.Close()
            response.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
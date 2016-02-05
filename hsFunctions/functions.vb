Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Public Class functions
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32
    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Shared alphabet() As String = {"£", "锋", "ਖ", "ਕ", "메", "导", "ن", "%", "马", "哦", "子", "!", "\", "s", "吉", "ն", "ք", "펜", "ط", "ש", "ਦ", "아", "ե", "땅", "快", "力", "+", "과", "疯", "ц", "'", "电", "ד", "z", "女", "孔", "带", "x", "ס", "爱", "$", "鼓", "ո", "ו", "կ", "ز", ",", "坦", "冻", "د", "艾", "침", "ղ", "얼", "ਅ", "ر", "녀", "յ", "류", "笔", "娜", "h", "y", "ը", "п", "բ", "ט", "迪", "ئ", "脑", "ب", "ռ", "פ", "크", "弗", "ş", "男", "n", "ğ", "}", "ਗ", "教", "提", "ћ", "€", "տ", "람", "时", "克", "贝", "f", "v", "դ", "结", "6", "콩", "a", "ਲ", "1", "勒", "屁", "^", "д", "k", "吾", "o", "ר", "ث", "ش", "ਫ", "غ", "图", "ت", "饶", "|", "筆", "9", "维", "ָ", "购", "p", "拉", "ր", "午", "花", "ص", "屄", "3", "贼", "ਤ", "소", "g", "晚", "밤", "亲", "շ", "ם", "و", "报", "术", "姆", "订", "片", "影", "能", "铁", "5", "_", "ı", "u", "일", "西", "索", "比", "d", "ਧ", "ւ", "和", "蕪", "ا", "2", "₺", "ਜ", "脚", "丝", "t", "א", "#", "ս", "德", "վ", "굴", "m", "院", "堅", "儿", "ի", "0", "校", "գ", "斯", "ਰ", "報", "ਬ", "=", "w", "豆", "ح", "ض", "尺", "上", "ؤ", "ي", "ջ", "高", "开", ";", "的", "平", "家", "é", "ע", "칼", "览", "լ", "水", "ա", "字", "服", "탱", "ч", "ק", "저", "肖", "ਨ", "性", "م", "白", "后", "面", "ਹ", "ਾ", "展", ">", "자", "室", "q", "c", "ء", "س", "果", "զ", "*", "ਿ", "名", "ਸ", "b", "&", "i", "?", "մ", "ه", "4", "孩", "小", "ف", "锈", "色", "天", "刻", "j", "ੀ", "/", "ਮ", "남", "冰", "魔", "伊", "ق", "른", "[", "ظ", "人", "ਚ", "{", "л", "]", "7", "头", "r", "夜", "月", "剧", "l", "з", "ö", "像", "չ", "마", "½", "前", "生", "ל", "线", "ç", "и", "尔", ")", "e", "녁", "手", "ج", "ג", "(", "穿", "裙", "师", "ل", ".", "ך", "<", "ü", "8", "情", "월", "诶", "杰", "别", "恕", "ى", "ع", "خ", "견", "荒", "卡", "נ", "演", "学"}
    Shared Generator As Random = New Random(Now.Millisecond)
    ''' <summary>
    ''' combobox auto change dropdownwidth
    ''' </summary>
    ''' <param name="senderComboBox"></param>
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
    ''' karakter sayımı
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="ch"></param>
    ''' <returns></returns>
    Public Shared Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Return value.Count(Function(c As Char) c = ch)
    End Function
    ''' <summary>
    ''' find ip
    ''' </summary>
    ''' <param name="local_IP"></param>
    ''' <returns></returns>
    Public Shared Function GetIPAddress(Optional local_IP As String = "172.22") As String
        Dim hostname As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName())
        Dim ip As IPAddress() = hostname.AddressList
        For Each adres In ip
            If adres.ToString().StartsWith(local_IP) = True Then Return adres.ToString() : Exit For
        Next
        Return ip(0).ToString()
    End Function
    ''' <summary>
    ''' make titlecase
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Shared Function TitleCase(str As String) As String
        Return StrConv(Trim(str), VbStrConv.ProperCase)
    End Function
    ''' <summary>
    ''' dgv'ye yapıştır kabul ettirir
    ''' </summary>
    ''' <param name="dgv"></param>
    Public Shared Sub PasteUnboundRecords(dgv As DataGridView)
        Try
            Dim rowLines As String() = Clipboard.GetText(TextDataFormat.Text).Split(New String(0) {vbCr & vbLf}, StringSplitOptions.None)
            Dim currentRowIndex As Integer = (If(dgv.CurrentRow IsNot Nothing, dgv.CurrentRow.Index, 0))
            Dim currentColumnIndex As Integer = (If(dgv.CurrentCell IsNot Nothing, dgv.CurrentCell.ColumnIndex, 0))
            Dim currentColumnCount As Integer = dgv.Columns.Count

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
            'Log Exception
        End Try
    End Sub
    ''' <summary>
    ''' encrypt text with mda5
    ''' </summary>
    ''' <param name="Input"></param>
    ''' <returns></returns>
    Public Shared Function Md5Convert(Input As String) As String
        If IsDBNull(Input) = True Then Return ""
        Using md5Hash As MD5 = MD5.Create()
            ' Convert the input string to a byte array and compute the hash. 
            Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(Input))
            ' Create a new Stringbuilder to collect the bytes and create a string. 
            Dim sBuilder As New StringBuilder()
            ' Loop through each byte of the hashed data and format each one as a hexadecimal string. 
            Dim x As Integer
            For x = 0 To data.Length - 1
                sBuilder.Append(data(x).ToString("x2"))
            Next x
            ' Return the hexadecimal string. 
            Return sBuilder.ToString()
        End Using
    End Function
    ''' <summary>
    ''' tüm menüleri tek bir listeye atmak için kullanılır
    ''' </summary>
    ''' <param name="menu"></param>
    ''' <param name="menues"></param>
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
    ''' write to ini file
    ''' </summary>
    ''' <param name="iniFileName"></param>
    ''' <param name="Section"></param>
    ''' <param name="ParamName"></param>
    ''' <param name="ParamVal"></param>
    Public Shared Sub INI_Write(ByVal iniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamVal As String)
        Dim Result As Integer = WritePrivateProfileString(Section, ParamName, ParamVal, iniFileName)
    End Sub
    ''' <summary>
    ''' read from ini file
    ''' </summary>
    ''' <param name="IniFileName"></param>
    ''' <param name="Section"></param>
    ''' <param name="ParamName"></param>
    ''' <param name="ParamDefault"></param>
    ''' <returns></returns>
    Public Shared Function INI_Read(ByVal IniFileName As String, ByVal Section As String, ByVal ParamName As String, Optional ByVal ParamDefault As String = "") As String
        Dim ParamVal As String = Space$(1024)
        Dim LenParamVal As Long = GetPrivateProfileString(Section, ParamName, ParamDefault, ParamVal, Len(ParamVal), IniFileName)
        Return Left$(ParamVal, LenParamVal)
    End Function
    ''' <summary>
    ''' validate email address
    ''' </summary>
    ''' <param name="email"></param>
    ''' <returns></returns>
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
    ''' 
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Shared Function isNumeric(value As String) As Boolean
        Return isNumeric(value)
    End Function
    ''' <summary>
    ''' gets website sourcedata from a given url
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
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
    ''' randomizes an array
    ''' </summary>
    ''' <param name="items"></param>
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
    ''' creates a random number
    ''' </summary>
    ''' <param name="Min"></param>
    ''' <param name="Max"></param>
    ''' <returns></returns>
    Public Shared Function GetRandomNumber(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Return Generator.Next(Min, Max)
    End Function
    ''' <summary>
    ''' şifrele
    ''' asc olarak 32 ve sonrasını şifreleyecek
    ''' 0-9: 48-57, +:43, -:45
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <returns></returns>
    Public Shared Function Encrypt(txt As String) As String
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
    ''' şifreyi çöz
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <returns></returns>
    Public Shared Function Decrypt(txt As String) As String
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
End Class

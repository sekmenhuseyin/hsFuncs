Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Public Class fnStandart
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32
    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    ''' <summary>
    ''' combobox auto change dropdownwidth
    ''' </summary>
    ''' <param name="senderComboBox"></param>
    Public Sub AdjustComboBoxWidth(senderComboBox As ComboBox)
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
    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Return value.Count(Function(c As Char) c = ch)
    End Function
    ''' <summary>
    ''' find ip
    ''' </summary>
    ''' <param name="local_IP"></param>
    ''' <returns></returns>
    Public Function GetIPAddress(Optional local_IP As String = "172.22") As String
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
    Public Function TitleCase(str As String) As String
        Return StrConv(Trim(str), VbStrConv.ProperCase)
    End Function
    ''' <summary>
    ''' dgv'ye yapıştır kabul ettirir
    ''' </summary>
    ''' <param name="dgv"></param>
    Public Sub PasteUnboundRecords(dgv As DataGridView)
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
    Public Function Md5Convert(Input As String) As String
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
    Public Sub GetSubMenus2List(menu As MenuStrip, menues As List(Of ToolStripItem))
        For Each t As ToolStripItem In menu.Items
            GetMenues(t, menues) 'alt menüleri de alsın diye iç içe bir döngü
        Next
    End Sub
    Private Sub GetMenues(ByVal Current As ToolStripItem, ByRef menues As List(Of ToolStripItem))
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
    Public Sub INI_Write(ByVal iniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamVal As String)
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
    Public Function INI_Read(ByVal IniFileName As String, ByVal Section As String, ByVal ParamName As String, Optional ByVal ParamDefault As String = "") As String
        Dim ParamVal As String = Space$(1024)
        Dim LenParamVal As Long = GetPrivateProfileString(Section, ParamName, ParamDefault, ParamVal, Len(ParamVal), IniFileName)
        Return Left$(ParamVal, LenParamVal)
    End Function
    ''' <summary>
    ''' validate email address
    ''' </summary>
    ''' <param name="email"></param>
    ''' <returns></returns>
    Public Function isEmail(email As String) As Boolean
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
    Public Function isNumeric(value As String) As Boolean
        Return isNumeric(value)
    End Function
    ''' <summary>
    ''' gets website sourcedata from a given url
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    Private Function GetWebResponse(url As String) As String
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
End Class

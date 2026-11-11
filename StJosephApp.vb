Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports System.Diagnostics
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

' =====================================================================
'  ST. JOSEPH VOCATIONAL INSTITUTE - Admissions Portal
'  Pages: 1 Login / Sign-up  ->  2 Bio-data  ->  3 Education (+ uploads)
'         ->  4 Sponsorship  ->  5 Review & Submit  ->  6 Payment & Receipt
'
'  Needs two pictures next to the .exe (the .vbproj copies them for you):
'     background.jpg   (shown blurred behind every page)
'     logo.png         (school badge)
' =====================================================================

Public Module Program
    <STAThread>
    Public Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Dim first As New LoginForm()
        first.Show()
        Application.Run()
    End Sub
End Module

' ---------------------------------------------------------------------
'  Shared look & feel, picture loading and blur
' ---------------------------------------------------------------------
Public Module Theme
    Public ReadOnly Gold As Color = Color.FromArgb(255, 200, 60)
    Public ReadOnly Navy As Color = Color.FromArgb(18, 30, 58)
    Private _bg As Bitmap
    Private _logo As Bitmap
    Private _logoTried As Boolean

    ' Looks for a picture next to the .exe, then in the current folder,
    ' then in parent folders (so it also works straight from the project folder).
    Public Function FindAsset(fileName As String) As String
        Dim starts As New List(Of String)
        starts.Add(AppDomain.CurrentDomain.BaseDirectory)
        starts.Add(Application.StartupPath)
        starts.Add(Directory.GetCurrentDirectory())
        For Each s As String In starts
            Dim cur As String = s
            For i As Integer = 0 To 5
                If String.IsNullOrEmpty(cur) Then Exit For
                Dim candidate As String = Path.Combine(cur, fileName)
                If File.Exists(candidate) Then Return candidate
                Dim parent As DirectoryInfo = Directory.GetParent(cur)
                If parent Is Nothing Then Exit For
                cur = parent.FullName
            Next
        Next
        Return Nothing
    End Function

    Public Function BlurredBackground() As Bitmap
        If _bg Is Nothing Then
            Try
                Dim f As String = FindAsset("background.jpg")
                If f IsNot Nothing Then
                    Using src As New Bitmap(f)
                        _bg = MakeBlur(src)
                    End Using
                End If
            Catch
                _bg = Nothing
            End Try
            If _bg Is Nothing Then _bg = FallbackGradient()
        End If
        Return _bg
    End Function

    Public Function Logo() As Bitmap
        If Not _logoTried Then
            _logoTried = True
            Try
                Dim f As String = FindAsset("logo.png")
                If f IsNot Nothing Then
                    Using tmp As New Bitmap(f)
                        _logo = New Bitmap(tmp)
                    End Using
                End If
            Catch
                _logo = Nothing
            End Try
        End If
        Return _logo
    End Function

    Private Function ResizeImage(src As Image, w As Integer, h As Integer) As Bitmap
        Dim bmp As New Bitmap(Math.Max(1, w), Math.Max(1, h))
        Using g As Graphics = Graphics.FromImage(bmp)
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.PixelOffsetMode = PixelOffsetMode.HighQuality
            g.DrawImage(src, 0, 0, bmp.Width, bmp.Height)
        End Using
        Return bmp
    End Function

    ' Shrink a lot, then scale back up in two steps = smooth blur
    Private Function MakeBlur(src As Bitmap) As Bitmap
        Dim scale As Double = Math.Min(1.0, 1400.0 / src.Width)
        Dim w As Integer = Math.Max(32, CInt(src.Width * scale))
        Dim h As Integer = Math.Max(32, CInt(src.Height * scale))
        Using tiny As Bitmap = ResizeImage(src, w \ 22, h \ 22)
            Using mid As Bitmap = ResizeImage(tiny, w \ 6, h \ 6)
                Return ResizeImage(mid, w, h)
            End Using
        End Using
    End Function

    Private Function FallbackGradient() As Bitmap
        Dim bmp As New Bitmap(1200, 800)
        Using g As Graphics = Graphics.FromImage(bmp)
            Using br As New LinearGradientBrush(New Rectangle(0, 0, 1200, 800), Color.FromArgb(25, 55, 109), Color.FromArgb(120, 70, 150), 45.0F)
                g.FillRectangle(br, 0, 0, 1200, 800)
            End Using
        End Using
        Return bmp
    End Function

    ' radius = diameter of the corner arcs
    Public Function RoundedPath(r As Rectangle, radius As Integer) As GraphicsPath
        Dim p As New GraphicsPath()
        p.AddArc(r.X, r.Y, radius, radius, 180, 90)
        p.AddArc(r.Right - radius - 1, r.Y, radius, radius, 270, 90)
        p.AddArc(r.Right - radius - 1, r.Bottom - radius - 1, radius, radius, 0, 90)
        p.AddArc(r.X, r.Bottom - radius - 1, radius, radius, 90, 90)
        p.CloseFigure()
        Return p
    End Function
End Module

' ---------------------------------------------------------------------
'  Data collected across the forms
' ---------------------------------------------------------------------
Public Module AppData
    Public ReadOnly Data As New Dictionary(Of String, String)
    Public ReadOnly Docs As New List(Of String)
    Public ReadOnly AppFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "StJosephVI")
    Public ReadOnly UserFile As String = Path.Combine(AppFolder, "users.dat")

    ' ---- EDIT THESE: fee and where applicants must send the money ----
    Public Const AppFee As Decimal = 50000D
    Public Const CurrencyCode As String = "UGX"
    Public Const MobileMoneyNumber As String = "0700 000 000"
    Public Const MobileMoneyName As String = "St. Joseph Vocational Institute"
    Public Const BankDetails As String = "Bank: YOUR BANK NAME" & vbLf & "Account name: St. Joseph Vocational Institute" & vbLf & "Account no.: 0000000000"

    ' Set once the application has been submitted
    Public RefNo As String = ""
    ' True while the student is fixing something from the review / receipt page
    Public ReturnToReview As Boolean

    Public Function FeeText() As String
        Return CurrencyCode & " " & AppFee.ToString("N0", Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Function Saved(key As String) As String
        Dim s As String = Nothing
        If Data.TryGetValue(key, s) Then Return s
        Return ""
    End Function

    Private Sub Section(sb As StringBuilder, title As String, keys As String())
        sb.AppendLine(title)
        sb.AppendLine(New String("-"c, title.Length))
        For Each k As String In keys
            sb.AppendLine("  " & k & ": " & Saved(k))
        Next
        sb.AppendLine()
    End Sub

    Public Function BuildSummary() As String
        Dim sb As New StringBuilder()
        Section(sb, "ACCOUNT", {"Username"})
        Section(sb, "PERSONAL BIO-DATA", {"Full name", "Date of birth", "Gender", "Nationality", "National ID / Passport no.", "Phone number", "Email address", "Home address", "Next of kin (name & phone)"})
        Section(sb, "EDUCATION BACKGROUND", {"Highest level of education", "Last school / institution", "Year completed", "Grade / result", "Course applying for"})
        sb.AppendLine("  Documents attached: " & Docs.Count.ToString())
        For Each d As String In Docs
            sb.AppendLine("    - " & Path.GetFileName(d))
        Next
        sb.AppendLine()
        Section(sb, "SPONSORSHIP", {"Sponsorship type", "Sponsor name", "Relationship", "Sponsor phone", "Sponsor email", "Sponsor address", "Sponsor occupation / employer", "Notes"})
        sb.AppendLine("PAYMENT")
        sb.AppendLine("-------")
        sb.AppendLine("  Application fee: " & FeeText())
        For Each k As String In {"Payment status", "Payment method", "Payment details", "Paid on"}
            Dim v As String = Saved(k)
            If v.Length = 0 AndAlso k = "Payment status" Then v = "Not yet paid"
            If v.Length > 0 Then sb.AppendLine("  " & k & ": " & v)
        Next
        sb.AppendLine()
        Return sb.ToString()
    End Function

    ' Full printable document = header + reference + summary
    Public Function BuildDocument() As String
        Dim sb As New StringBuilder()
        sb.AppendLine("ST. JOSEPH VOCATIONAL INSTITUTE")
        sb.AppendLine("Student Admission Application")
        sb.AppendLine(New String("="c, 40))
        sb.AppendLine("Reference: " & If(RefNo.Length > 0, RefNo, "DRAFT (not yet submitted)"))
        sb.AppendLine("Printed: " & Date.Now.ToString("yyyy-MM-dd HH:mm"))
        sb.AppendLine()
        sb.Append(BuildSummary())
        Return sb.ToString()
    End Function

    Public Function ApplicationFolder() As String
        Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "StJosephVI_Applications", RefNo)
    End Function

    ' (Re)writes application.txt, application.pdf and copies the uploaded documents
    Public Sub WriteApplication()
        Dim folder As String = ApplicationFolder()
        Directory.CreateDirectory(folder)
        Dim doc As String = BuildDocument()
        File.WriteAllText(Path.Combine(folder, "application.txt"), doc, Encoding.UTF8)
        File.WriteAllBytes(Path.Combine(folder, "application.pdf"), BuildPdf(doc))
        For Each f As String In Docs
            If File.Exists(f) Then
                Dim dest As String = Path.Combine(folder, Path.GetFileName(f))
                If Not String.Equals(Path.GetFullPath(f), Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase) Then
                    File.Copy(f, dest, True)
                End If
            End If
        Next
    End Sub

    Public Function WrapLines(text As String, width As Integer) As List(Of String)
        Dim lines As New List(Of String)
        For Each raw As String In text.Replace(vbCrLf, vbLf).Split(ChrW(10))
            Dim s As String = raw
            If s.Length = 0 Then
                lines.Add("")
            Else
                While s.Length > width
                    lines.Add(s.Substring(0, width))
                    s = "    " & s.Substring(width)
                End While
                lines.Add(s)
            End If
        Next
        Return lines
    End Function

    Private Function PdfSafe(text As String) As String
        Dim sb As New StringBuilder()
        For Each ch As Char In text
            Dim code As Integer = AscW(ch)
            If code = 10 Then
                sb.Append(ch)
            ElseIf code = 9 Then
                sb.Append("  ")
            ElseIf code >= 32 AndAlso code <= 126 Then
                sb.Append(ch)
            ElseIf code <> 13 Then
                sb.Append("?"c)
            End If
        Next
        Return sb.ToString()
    End Function

    ' Tiny dependency-free PDF writer (plain Courier text, A4, auto page breaks)
    Public Function BuildPdf(text As String) As Byte()
        Dim lines As List(Of String) = WrapLines(PdfSafe(text), 80)
        Const perPage As Integer = 60
        Dim pageCount As Integer = Math.Max(1, (lines.Count + perPage - 1) \ perPage)
        Dim total As Integer = 3 + 2 * pageCount

        Dim sb As New StringBuilder()
        Dim offs As New List(Of Integer)
        sb.Append("%PDF-1.4" & vbLf)

        Dim addObj As Action(Of String) = Sub(body As String)
                                              offs.Add(sb.Length)
                                              sb.Append(offs.Count.ToString() & " 0 obj" & vbLf & body & vbLf & "endobj" & vbLf)
                                          End Sub

        Dim kids As New StringBuilder()
        For i As Integer = 0 To pageCount - 1
            kids.Append((4 + 2 * i).ToString() & " 0 R ")
        Next
        addObj("<< /Type /Catalog /Pages 2 0 R >>")
        addObj("<< /Type /Pages /Kids [" & kids.ToString() & "] /Count " & pageCount.ToString() & " >>")
        addObj("<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>")

        For i As Integer = 0 To pageCount - 1
            Dim st As New StringBuilder()
            st.Append("BT" & vbLf & "/F1 10 Tf" & vbLf & "12 TL" & vbLf & "50 802 Td" & vbLf)
            Dim first As Integer = i * perPage
            Dim last As Integer = Math.Min(lines.Count, first + perPage) - 1
            For j As Integer = first To last
                Dim esc As String = lines(j).Replace("\", "\\").Replace("(", "\(").Replace(")", "\)")
                st.Append("(" & esc & ") '" & vbLf)
            Next
            st.Append("ET")
            Dim content As String = st.ToString()
            addObj("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> >> /Contents " & (5 + 2 * i).ToString() & " 0 R >>")
            addObj("<< /Length " & content.Length.ToString() & " >>" & vbLf & "stream" & vbLf & content & vbLf & "endstream")
        Next

        Dim xrefPos As Integer = sb.Length
        sb.Append("xref" & vbLf & "0 " & (total + 1).ToString() & vbLf)
        sb.Append("0000000000 65535 f " & vbLf)
        For Each o As Integer In offs
            sb.Append(o.ToString("D10") & " 00000 n " & vbLf)
        Next
        sb.Append("trailer" & vbLf & "<< /Size " & (total + 1).ToString() & " /Root 1 0 R >>" & vbLf)
        sb.Append("startxref" & vbLf & xrefPos.ToString() & vbLf & "%%EOF")
        Return Encoding.ASCII.GetBytes(sb.ToString())
    End Function
End Module

' ---------------------------------------------------------------------
'  Download (PDF / text) and print
' ---------------------------------------------------------------------
Public Module Exporter
    Public Sub DownloadApplication(owner As IWin32Window)
        Dim baseName As String = If(RefNo.Length > 0, RefNo, "SJVI-Draft")
        Using dlg As New SaveFileDialog With {
            .Title = "Download application", .FileName = baseName,
            .Filter = "PDF document (*.pdf)|*.pdf|Text file (*.txt)|*.txt",
            .DefaultExt = "pdf", .AddExtension = True, .OverwritePrompt = True}
            If dlg.ShowDialog(owner) <> DialogResult.OK Then Return
            Try
                Dim ext As String = Path.GetExtension(dlg.FileName).ToLowerInvariant()
                If ext = ".txt" Then
                    File.WriteAllText(dlg.FileName, BuildDocument(), Encoding.UTF8)
                Else
                    File.WriteAllBytes(dlg.FileName, BuildPdf(BuildDocument()))
                End If
                MessageBox.Show(owner, "Saved to:" & vbCrLf & dlg.FileName, "Download complete",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(owner, "Could not save the file: " & ex.Message, "St. Joseph Vocational Institute",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Using
    End Sub

    Public Sub PrintApplication(owner As IWin32Window)
        Dim q As New Queue(Of String)(WrapLines(BuildDocument(), 85))
        Using f As New Font("Courier New", 10.0F)
            Using doc As New PrintDocument()
                AddHandler doc.PrintPage, Sub(s As Object, e As PrintPageEventArgs)
                                              Dim lh As Single = f.GetHeight(e.Graphics)
                                              Dim y As Single = e.MarginBounds.Top
                                              While q.Count > 0 AndAlso y + lh <= e.MarginBounds.Bottom
                                                  e.Graphics.DrawString(q.Dequeue(), f, Brushes.Black, CSng(e.MarginBounds.Left), y)
                                                  y += lh
                                              End While
                                              e.HasMorePages = q.Count > 0
                                          End Sub
                Using dlg As New PrintDialog With {.Document = doc, .UseEXDialog = True}
                    If dlg.ShowDialog(owner) = DialogResult.OK Then
                        Try
                            doc.Print()
                        Catch ex As Exception
                            MessageBox.Show(owner, "Could not print: " & ex.Message, "St. Joseph Vocational Institute",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    End If
                End Using
            End Using
        End Using
    End Sub
End Module

' ---------------------------------------------------------------------
'  Base form: blurred photo, dark veil, glass card, logo, progress bar
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class BaseForm
    Inherits Form

    Protected Const CardX As Integer = 110
    Protected Const CardY As Integer = 116
    Protected Const CardW As Integer = 780
    Protected Const CardH As Integer = 530

    Protected ReadOnly Card As New Panel()
    Protected ReadOnly TitleLbl As New Label()
    Private ReadOnly _step As Integer
    Private ReadOnly _showLogo As Boolean
    Private _canvas As Bitmap
    Private _switching As Boolean

    Public Sub New(pageTitle As String, stepNo As Integer, Optional showHeaderLogo As Boolean = True)
        _step = stepNo
        _showLogo = showHeaderLogo
        AutoScaleMode = AutoScaleMode.None
        Text = "St. Joseph Vocational Institute - Admissions"
        ClientSize = New Size(1000, 670)
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        DoubleBuffered = True
        BackColor = Navy
        Font = New Font("Segoe UI", 10.0F)

        Controls.Add(New Label With {
            .Text = "ST. JOSEPH VOCATIONAL INSTITUTE",
            .Font = New Font("Georgia", 26.0F, FontStyle.Bold),
            .ForeColor = Gold, .BackColor = Color.Transparent,
            .TextAlign = ContentAlignment.MiddleCenter, .AutoSize = False,
            .Bounds = New Rectangle(0, 12, 1000, 46)})
        Controls.Add(New Label With {
            .Text = "S T U D E N T   A D M I S S I O N   P O R T A L",
            .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
            .ForeColor = Color.White, .BackColor = Color.Transparent,
            .TextAlign = ContentAlignment.MiddleCenter, .AutoSize = False,
            .Bounds = New Rectangle(0, 60, 1000, 26)})

        ' Transparent card: the glass itself is painted into the background picture
        Card.Bounds = New Rectangle(CardX, CardY, CardW, CardH)
        Card.BackColor = Color.Transparent
        Controls.Add(Card)

        TitleLbl.Text = pageTitle
        TitleLbl.Font = New Font("Georgia", 18.0F, FontStyle.Bold)
        TitleLbl.ForeColor = Color.White
        TitleLbl.BackColor = Color.Transparent
        TitleLbl.AutoSize = True
        TitleLbl.Location = New Point(30, 16)
        Card.Controls.Add(TitleLbl)

        Card.Controls.Add(New Panel With {.BackColor = Gold, .Bounds = New Rectangle(30, 54, 70, 3)})
        Card.Controls.Add(New Label With {
            .Text = "Step " & stepNo.ToString() & " of 6",
            .ForeColor = Color.FromArgb(215, 222, 240), .BackColor = Color.Transparent,
            .AutoSize = False, .TextAlign = ContentAlignment.MiddleRight,
            .Bounds = New Rectangle(560, 24, 190, 24)})

        _canvas = BuildCanvas()
    End Sub

    ' ---- painting ---------------------------------------------------
    Private Function BuildCanvas() As Bitmap
        Dim w As Integer = ClientSize.Width
        Dim h As Integer = ClientSize.Height
        Dim bmp As New Bitmap(w, h)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.PixelOffsetMode = PixelOffsetMode.HighQuality

            ' 1. blurred photo, scaled to cover the whole form
            Dim bg As Bitmap = BlurredBackground()
            Dim scale As Double = Math.Max(w / CDbl(bg.Width), h / CDbl(bg.Height))
            Dim sw As Integer = CInt(Math.Ceiling(bg.Width * scale))
            Dim sh As Integer = CInt(Math.Ceiling(bg.Height * scale))
            g.DrawImage(bg, (w - sw) \ 2, (h - sh) \ 2, sw, sh)

            ' 2. dark veil so the words stay readable
            Using veil As New SolidBrush(Color.FromArgb(120, 6, 12, 30))
                g.FillRectangle(veil, 0, 0, w, h)
            End Using
            Using hdr As New LinearGradientBrush(New Rectangle(0, 0, w, 130), Color.FromArgb(215, 6, 12, 30), Color.FromArgb(0, 6, 12, 30), LinearGradientMode.Vertical)
                g.FillRectangle(hdr, 0, 0, w, 130)
            End Using

            ' 3. soft shadow + glass card
            Dim cardRect As New Rectangle(CardX, CardY, CardW, CardH)
            For i As Integer = 7 To 1 Step -1
                Dim r As Rectangle = Rectangle.Inflate(cardRect, i * 2, i * 2)
                Using p As GraphicsPath = RoundedPath(r, 36 + i * 4)
                    Using br As New SolidBrush(Color.FromArgb(14, 0, 0, 0))
                        g.FillPath(br, p)
                    End Using
                End Using
            Next
            Using p As GraphicsPath = RoundedPath(cardRect, 36)
                Using br As New SolidBrush(Color.FromArgb(200, 12, 22, 46))
                    g.FillPath(br, p)
                End Using
                Using pen As New Pen(Color.FromArgb(170, Gold), 1.6F)
                    g.DrawPath(pen, p)
                End Using
            End Using

            ' 4. progress bar
            For i As Integer = 1 To 6
                Dim c As Color = If(i <= _step, Gold, Color.FromArgb(90, 255, 255, 255))
                Using br As New SolidBrush(c)
                    g.FillRectangle(br, CardX + (i - 1) * 131, 98, 125, 5)
                End Using
            Next

            ' 5. school badge in the header
            If _showLogo Then
                Dim lg As Bitmap = Logo()
                If lg IsNot Nothing Then g.DrawImage(lg, 26, 6, 86, 86)
            End If
        End Using
        Return bmp
    End Function

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        Dim c As Bitmap = _canvas
        If c Is Nothing Then
            ' form is closing / being disposed: just paint a plain colour
            e.Graphics.Clear(Navy)
            Return
        End If
        e.Graphics.DrawImageUnscaled(c, 0, 0)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso _canvas IsNot Nothing Then
            Dim c As Bitmap = _canvas
            _canvas = Nothing
            c.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' ---- navigation -------------------------------------------------
    Protected Sub NavigateTo(target As BaseForm)
        _switching = True
        target.StartPosition = FormStartPosition.Manual
        target.Location = Location
        target.Show()
        Close()
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        MyBase.OnFormClosed(e)
        If Not _switching Then Application.Exit()
    End Sub

    ' ---- control helpers -------------------------------------------
    Protected Function AddLabel(txt As String, x As Integer, y As Integer) As Label
        Dim l As New Label With {
            .Text = txt, .ForeColor = Color.White, .BackColor = Color.Transparent,
            .AutoSize = True, .Location = New Point(x, y)}
        Card.Controls.Add(l)
        Return l
    End Function

    Protected Function AddText(caption As String, x As Integer, y As Integer, w As Integer, Optional isPassword As Boolean = False) As TextBox
        Dim lbl As Label = AddLabel(caption, x, y)
        Dim t As New TextBox With {
            .Location = New Point(x, y + 24), .Width = w,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.FromArgb(245, 247, 252),
            .Font = New Font("Segoe UI", 10.5F), .Tag = lbl}
        If isPassword Then t.UseSystemPasswordChar = True
        Card.Controls.Add(t)
        Return t
    End Function

    Protected Function AddCombo(caption As String, x As Integer, y As Integer, w As Integer, ParamArray items() As String) As ComboBox
        Dim lbl As Label = AddLabel(caption, x, y)
        Dim c As New ComboBox With {
            .Location = New Point(x, y + 24), .Width = w,
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(245, 247, 252),
            .Font = New Font("Segoe UI", 10.5F), .Tag = lbl}
        c.Items.AddRange(items)
        Card.Controls.Add(c)
        Return c
    End Function

    Protected Function AddButton(txt As String, x As Integer, y As Integer, w As Integer, primary As Boolean) As Button
        Dim b As New Button With {
            .Text = txt, .Bounds = New Rectangle(x, y, w, 40),
            .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand,
            .Font = New Font("Segoe UI", 10.5F, FontStyle.Bold)}
        b.FlatAppearance.BorderSize = 0
        If primary Then
            b.BackColor = Gold
            b.ForeColor = Navy
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 220, 120)
        Else
            b.BackColor = Color.FromArgb(70, 85, 120)
            b.ForeColor = Color.White
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(95, 112, 150)
        End If
        Card.Controls.Add(b)
        Return b
    End Function

    ' Move a field (and its caption) to a new row
    Protected Sub Place(c As Control, y As Integer)
        c.Top = y + 24
        DirectCast(c.Tag, Label).Top = y
    End Sub

    Protected Sub ShowField(c As Control, visible As Boolean)
        c.Visible = visible
        DirectCast(c.Tag, Label).Visible = visible
    End Sub

    Protected Sub Fill(c As Control, key As String)
        Dim v As String = Saved(key)
        If v.Length > 0 Then c.Text = v
    End Sub

    Protected Sub Save(key As String, c As Control)
        Data(key) = c.Text.Trim()
    End Sub

    Protected Sub Warn(msg As String)
        MessageBox.Show(Me, msg, "St. Joseph Vocational Institute", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Protected Function Check(ParamArray required() As Control) As Boolean
        For Each c As Control In required
            If String.IsNullOrWhiteSpace(c.Text) Then
                Warn("Please complete all fields marked with *.")
                c.Focus()
                Return False
            End If
        Next
        Return True
    End Function

    Protected Shared Function IsEmail(s As String) As Boolean
        s = s.Trim()
        Dim at As Integer = s.IndexOf("@"c)
        Return at > 0 AndAlso Not s.Contains(" ") AndAlso s.IndexOf("."c, at) > at + 1 AndAlso Not s.EndsWith(".")
    End Function

    Protected Shared Function IsPhone(s As String) As Boolean
        Dim digits As Integer = 0
        For Each ch As Char In s
            If Char.IsDigit(ch) Then digits += 1
        Next
        Return digits >= 9
    End Function
End Class

' ---------------------------------------------------------------------
'  1. LOGIN / SIGN UP
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class LoginForm
    Inherits BaseForm

    Private txtUser As TextBox
    Private txtEmail As TextBox
    Private txtPass As TextBox
    Private txtConfirm As TextBox
    Private chkShow As CheckBox
    Private btnMain As Button
    Private lnkSwitch As LinkLabel
    Private _signup As Boolean

    Public Sub New()
        MyBase.New("Student Login", 1, False)

        ' left side: big badge + welcome text
        Dim lg As Bitmap = Logo()
        If lg IsNot Nothing Then
            Card.Controls.Add(New PictureBox With {
                .Image = lg, .SizeMode = PictureBoxSizeMode.Zoom,
                .BackColor = Color.Transparent, .Bounds = New Rectangle(40, 80, 240, 240)})
        End If
        Card.Controls.Add(New Label With {
            .Text = "Welcome!", .Font = New Font("Georgia", 17.0F, FontStyle.Bold),
            .ForeColor = Gold, .BackColor = Color.Transparent,
            .AutoSize = False, .TextAlign = ContentAlignment.MiddleCenter,
            .Bounds = New Rectangle(20, 335, 280, 36)})
        Card.Controls.Add(New Label With {
            .Text = "Log in to continue your application, or create an account to get started.",
            .ForeColor = Color.FromArgb(225, 230, 242), .BackColor = Color.Transparent,
            .AutoSize = False, .TextAlign = ContentAlignment.TopCenter,
            .Bounds = New Rectangle(30, 375, 260, 70)})
        Card.Controls.Add(New Panel With {.BackColor = Color.FromArgb(80, 95, 130), .Bounds = New Rectangle(320, 85, 1, 380)})

        ' right side: fields
        txtUser = AddText("Username", 355, 80, 380)
        txtEmail = AddText("Email address", 355, 150, 380)
        txtPass = AddText("Password", 355, 220, 380, True)
        txtConfirm = AddText("Confirm password", 355, 290, 380, True)
        chkShow = New CheckBox With {
            .Text = "Show password", .ForeColor = Color.White, .BackColor = Color.Transparent,
            .AutoSize = True, .Left = 355, .Cursor = Cursors.Hand}
        Card.Controls.Add(chkShow)
        btnMain = AddButton("Login", 355, 300, 380, True)

        lnkSwitch = New LinkLabel With {
            .AutoSize = False, .TextAlign = ContentAlignment.MiddleCenter,
            .LinkColor = Gold, .ActiveLinkColor = Color.White, .VisitedLinkColor = Gold,
            .BackColor = Color.Transparent, .Width = 380, .Height = 28, .Left = 355}
        Card.Controls.Add(lnkSwitch)

        AddHandler btnMain.Click, AddressOf OnMainClick
        AddHandler lnkSwitch.LinkClicked, Sub(s, e) SetMode(Not _signup)
        AddHandler chkShow.CheckedChanged, AddressOf OnShowPassword
        AcceptButton = btnMain
        SetMode(False)
    End Sub

    Private Sub OnShowPassword(sender As Object, e As EventArgs)
        txtPass.UseSystemPasswordChar = Not chkShow.Checked
        txtConfirm.UseSystemPasswordChar = Not chkShow.Checked
    End Sub

    Private Sub SetMode(signup As Boolean)
        _signup = signup
        ShowField(txtEmail, signup)
        ShowField(txtConfirm, signup)
        Place(txtPass, If(signup, 220, 150))
        chkShow.Top = If(signup, 345, 205)
        btnMain.Top = If(signup, 385, 245)
        lnkSwitch.Top = btnMain.Top + 52
        TitleLbl.Text = If(signup, "Create Your Account", "Student Login")
        btnMain.Text = If(signup, "Create Account", "Login")
        lnkSwitch.Text = If(signup, "Already registered? Login", "New student? Create an account")
        txtPass.Clear()
        txtConfirm.Clear()
        txtUser.Focus()
    End Sub

    Private Shared Function Hash(user As String, pw As String) As String
        Dim salt As Byte() = Encoding.UTF8.GetBytes("SJVI|" & user.ToLowerInvariant())
        Using kdf As New Rfc2898DeriveBytes(pw, salt, 100000, HashAlgorithmName.SHA256)
            Return Convert.ToBase64String(kdf.GetBytes(32))
        End Using
    End Function

    Private Shared Function FindUser(user As String) As String()
        If Not File.Exists(UserFile) Then Return Nothing
        For Each line As String In File.ReadAllLines(UserFile)
            Dim p As String() = line.Split("|"c)
            If p.Length >= 3 AndAlso String.Equals(p(0), user, StringComparison.OrdinalIgnoreCase) Then Return p
        Next
        Return Nothing
    End Function

    Private Sub OnMainClick(sender As Object, e As EventArgs)
        Dim user As String = txtUser.Text.Trim()
        Dim pw As String = txtPass.Text

        Try
            If _signup Then
                If Not Check(txtUser, txtEmail, txtPass, txtConfirm) Then Return
                If user.Contains("|") OrElse txtEmail.Text.Contains("|") Then
                    Warn("The '|' character is not allowed.") : Return
                End If
                If Not IsEmail(txtEmail.Text) Then
                    Warn("Please enter a valid email address.") : txtEmail.Focus() : Return
                End If
                If pw.Length < 6 Then
                    Warn("Password must be at least 6 characters.") : txtPass.Focus() : Return
                End If
                If pw <> txtConfirm.Text Then
                    Warn("Passwords do not match.") : txtConfirm.Focus() : Return
                End If
                If FindUser(user) IsNot Nothing Then
                    Warn("That username is already taken.") : txtUser.Focus() : Return
                End If
                Directory.CreateDirectory(AppFolder)
                File.AppendAllText(UserFile, user & "|" & txtEmail.Text.Trim() & "|" & Hash(user, pw) & Environment.NewLine)
                MessageBox.Show(Me, "Account created successfully. Let's start your application!", "Welcome",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                If Not Check(txtUser, txtPass) Then Return
                Dim rec As String() = FindUser(user)
                If rec Is Nothing OrElse rec(2) <> Hash(user, pw) Then
                    Warn("Incorrect username or password.") : Return
                End If
            End If
        Catch ex As Exception
            Warn("Something went wrong: " & ex.Message)
            Return
        End Try

        Data("Username") = user
        NavigateTo(New BioDataForm())
    End Sub
End Class

' ---------------------------------------------------------------------
'  2. BIO-DATA
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class BioDataForm
    Inherits BaseForm

    Private txtName, txtNat, txtId, txtPhone, txtEmail, txtAddr, txtKin As TextBox
    Private dtpDob As DateTimePicker
    Private cboGender As ComboBox

    Public Sub New()
        MyBase.New("Personal Bio-Data", 2)

        txtName = AddText("Full name *", 30, 70, 340)
        AddLabel("Date of birth *", 30, 140)
        dtpDob = New DateTimePicker With {
            .Format = DateTimePickerFormat.Short, .Location = New Point(30, 164),
            .Width = 340, .MaxDate = Date.Today, .Value = Date.Today.AddYears(-18),
            .Font = New Font("Segoe UI", 10.5F)}
        Card.Controls.Add(dtpDob)
        cboGender = AddCombo("Gender *", 30, 210, 340, "Male", "Female", "Prefer not to say")
        txtNat = AddText("Nationality *", 30, 280, 340)
        txtId = AddText("National ID / Passport no.", 30, 350, 340)

        txtPhone = AddText("Phone number *", 410, 70, 340)
        txtEmail = AddText("Email address", 410, 140, 340)
        txtAddr = AddText("Home address *", 410, 210, 340)
        txtKin = AddText("Next of kin (name & phone) *", 410, 280, 340)

        Dim btnBack As Button = AddButton("‹  Back", 30, 470, 150, False)
        Dim btnNext As Button = AddButton(If(ReturnToReview, "Save changes  ›", "Next  ›"), 570, 470, 180, True)
        AddHandler btnBack.Click, AddressOf OnBack
        AddHandler btnNext.Click, AddressOf OnNext

        Fill(txtName, "Full name") : Fill(cboGender, "Gender") : Fill(txtNat, "Nationality")
        Fill(txtId, "National ID / Passport no.") : Fill(txtPhone, "Phone number")
        Fill(txtEmail, "Email address") : Fill(txtAddr, "Home address")
        Fill(txtKin, "Next of kin (name & phone)")
        Dim d As Date
        If Date.TryParse(Saved("Date of birth"), d) AndAlso d <= Date.Today Then dtpDob.Value = d
    End Sub

    Private Sub OnBack(sender As Object, e As EventArgs)
        If ReturnToReview Then
            NavigateTo(New SubmissionForm())
        Else
            NavigateTo(New LoginForm())
        End If
    End Sub

    Private Sub OnNext(sender As Object, e As EventArgs)
        If Not Check(txtName, cboGender, txtNat, txtPhone, txtAddr, txtKin) Then Return
        If Not IsPhone(txtPhone.Text) Then
            Warn("Please enter a valid phone number.") : txtPhone.Focus() : Return
        End If
        If txtEmail.Text.Trim().Length > 0 AndAlso Not IsEmail(txtEmail.Text) Then
            Warn("Please enter a valid email address.") : txtEmail.Focus() : Return
        End If
        Save("Full name", txtName)
        Data("Date of birth") = dtpDob.Value.ToString("yyyy-MM-dd")
        Save("Gender", cboGender) : Save("Nationality", txtNat)
        Save("National ID / Passport no.", txtId) : Save("Phone number", txtPhone)
        Save("Email address", txtEmail) : Save("Home address", txtAddr)
        Save("Next of kin (name & phone)", txtKin)
        If ReturnToReview Then
            NavigateTo(New SubmissionForm())
        Else
            NavigateTo(New EducationForm())
        End If
    End Sub
End Class

' ---------------------------------------------------------------------
'  3. EDUCATION BACKGROUND + DOCUMENT UPLOAD
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class EducationForm
    Inherits BaseForm

    Private cboLevel, cboCourse As ComboBox
    Private txtSchool, txtYear, txtGrade As TextBox
    Private lst As ListBox

    Public Sub New()
        MyBase.New("Education Background", 3)

        cboLevel = AddCombo("Highest level of education *", 30, 70, 340,
                            "Primary", "Lower Secondary (O-Level)", "Upper Secondary (A-Level)",
                            "Certificate", "Diploma", "Degree")
        txtSchool = AddText("Last school / institution *", 30, 140, 340)
        txtYear = AddText("Year completed *", 30, 210, 340)
        txtGrade = AddText("Grade / result", 30, 280, 340)
        ' Edit this list to match the courses your institute actually offers
        cboCourse = AddCombo("Course applying for *", 30, 350, 340,
                             "Carpentry & Joinery", "Tailoring & Garment Cutting", "Motor Vehicle Mechanics",
                             "Electrical Installation", "Plumbing", "Building & Concrete Practice",
                             "Hairdressing & Beauty", "Computer Applications (ICT)", "Catering & Hospitality")

        AddLabel("Academic documents (PDF, JPG, PNG, DOC)", 410, 70)
        Dim btnUpload As Button = AddButton("＋  Upload documents…", 410, 98, 340, False)
        lst = New ListBox With {
            .Location = New Point(410, 148), .Size = New Size(340, 150),
            .FormattingEnabled = True, .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.FromArgb(245, 247, 252), .Font = New Font("Segoe UI", 10.0F)}
        Card.Controls.Add(lst)
        Dim btnRemove As Button = AddButton("Remove selected", 410, 308, 340, False)

        Dim btnBack As Button = AddButton("‹  Back", 30, 470, 150, False)
        Dim btnNext As Button = AddButton(If(ReturnToReview, "Save changes  ›", "Next  ›"), 570, 470, 180, True)

        AddHandler lst.Format, Sub(s, e) e.Value = Path.GetFileName(CStr(e.ListItem))
        AddHandler btnUpload.Click, AddressOf OnUpload
        AddHandler btnRemove.Click, Sub(s, e)
                                        If lst.SelectedIndex >= 0 Then lst.Items.RemoveAt(lst.SelectedIndex)
                                    End Sub
        AddHandler btnBack.Click, AddressOf OnBack
        AddHandler btnNext.Click, AddressOf OnNext

        Fill(cboLevel, "Highest level of education") : Fill(txtSchool, "Last school / institution")
        Fill(txtYear, "Year completed") : Fill(txtGrade, "Grade / result")
        Fill(cboCourse, "Course applying for")
        For Each f As String In Docs
            lst.Items.Add(f)
        Next
    End Sub

    Private Sub OnBack(sender As Object, e As EventArgs)
        If ReturnToReview Then
            NavigateTo(New SubmissionForm())
        Else
            NavigateTo(New BioDataForm())
        End If
    End Sub

    Private Sub OnUpload(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog With {
            .Title = "Select academic documents", .Multiselect = True,
            .Filter = "Documents|*.pdf;*.jpg;*.jpeg;*.png;*.doc;*.docx|All files|*.*"}
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                For Each f As String In dlg.FileNames
                    If Not lst.Items.Contains(f) Then lst.Items.Add(f)
                Next
            End If
        End Using
    End Sub

    Private Sub OnNext(sender As Object, e As EventArgs)
        If Not Check(cboLevel, txtSchool, txtYear, cboCourse) Then Return
        Dim y As Integer
        If Not Integer.TryParse(txtYear.Text.Trim(), y) OrElse y < 1950 OrElse y > Date.Today.Year Then
            Warn("Please enter a valid year (e.g. 2022).")
            txtYear.Focus()
            Return
        End If
        Save("Highest level of education", cboLevel) : Save("Last school / institution", txtSchool)
        Save("Year completed", txtYear) : Save("Grade / result", txtGrade)
        Save("Course applying for", cboCourse)
        Docs.Clear()
        For Each it As Object In lst.Items
            Docs.Add(CStr(it))
        Next
        If ReturnToReview Then
            NavigateTo(New SubmissionForm())
        Else
            NavigateTo(New SponsorshipForm())
        End If
    End Sub
End Class

' ---------------------------------------------------------------------
'  4. SPONSORSHIP
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class SponsorshipForm
    Inherits BaseForm

    Private cboType As ComboBox
    Private txtName, txtRel, txtPhone, txtEmail, txtAddr, txtOcc, txtNotes As TextBox

    Public Sub New()
        MyBase.New("Sponsorship Details", 4)

        cboType = AddCombo("Who will pay the fees? *", 30, 70, 340,
                           "Self-sponsored", "Parent / Guardian", "Relative", "Employer",
                           "Church / NGO", "Government bursary", "Scholarship", "Other")
        txtName = AddText("Sponsor's full name *", 30, 140, 340)
        txtRel = AddText("Relationship to student *", 30, 210, 340)
        txtPhone = AddText("Sponsor's phone *", 30, 280, 340)
        txtEmail = AddText("Sponsor's email", 30, 350, 340)

        txtAddr = AddText("Sponsor's address", 410, 70, 340)
        txtOcc = AddText("Occupation / employer", 410, 140, 340)
        AddLabel("Additional notes", 410, 210)
        txtNotes = New TextBox With {
            .Multiline = True, .AcceptsReturn = True, .ScrollBars = ScrollBars.Vertical,
            .Location = New Point(410, 234), .Size = New Size(340, 110),
            .BorderStyle = BorderStyle.FixedSingle, .BackColor = Color.FromArgb(245, 247, 252),
            .Font = New Font("Segoe UI", 10.5F)}
        Card.Controls.Add(txtNotes)

        Dim btnBack As Button = AddButton("‹  Back", 30, 470, 150, False)
        Dim btnNext As Button = AddButton(If(ReturnToReview, "Save changes  ›", "Review  ›"), 570, 470, 180, True)
        AddHandler btnBack.Click, AddressOf OnBack
        AddHandler btnNext.Click, AddressOf OnNext
        AddHandler cboType.SelectedIndexChanged, Sub(s, e) UpdateSponsorFields()

        Fill(cboType, "Sponsorship type") : Fill(txtName, "Sponsor name") : Fill(txtRel, "Relationship")
        Fill(txtPhone, "Sponsor phone") : Fill(txtEmail, "Sponsor email") : Fill(txtAddr, "Sponsor address")
        Fill(txtOcc, "Sponsor occupation / employer") : Fill(txtNotes, "Notes")
        UpdateSponsorFields()
    End Sub

    Private Sub OnBack(sender As Object, e As EventArgs)
        If ReturnToReview Then
            NavigateTo(New SubmissionForm())
        Else
            NavigateTo(New EducationForm())
        End If
    End Sub

    Private Function IsSelf() As Boolean
        Return cboType.Text = "Self-sponsored"
    End Function

    Private Sub UpdateSponsorFields()
        For Each t As TextBox In {txtName, txtRel, txtPhone, txtEmail, txtAddr, txtOcc}
            t.Enabled = Not IsSelf()
        Next
    End Sub

    Private Sub OnNext(sender As Object, e As EventArgs)
        If Not Check(cboType) Then Return
        If Not IsSelf() Then
            If Not Check(txtName, txtRel, txtPhone) Then Return
            If Not IsPhone(txtPhone.Text) Then
                Warn("Please enter a valid phone number for the sponsor.") : txtPhone.Focus() : Return
            End If
            If txtEmail.Text.Trim().Length > 0 AndAlso Not IsEmail(txtEmail.Text) Then
                Warn("Please enter a valid sponsor email address.") : txtEmail.Focus() : Return
            End If
        End If
        Save("Sponsorship type", cboType)
        If IsSelf() Then
            For Each k As String In {"Sponsor name", "Relationship", "Sponsor phone", "Sponsor email", "Sponsor address", "Sponsor occupation / employer"}
                Data(k) = ""
            Next
        Else
            Save("Sponsor name", txtName) : Save("Relationship", txtRel)
            Save("Sponsor phone", txtPhone) : Save("Sponsor email", txtEmail) : Save("Sponsor address", txtAddr)
            Save("Sponsor occupation / employer", txtOcc)
        End If
        Save("Notes", txtNotes)
        NavigateTo(New SubmissionForm())
    End Sub
End Class

' ---------------------------------------------------------------------
'  5. REVIEW, EDIT & SUBMIT
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class SubmissionForm
    Inherits BaseForm

    Private chk As CheckBox

    Public Sub New()
        MyBase.New("Review & Submit", 5)
        ReturnToReview = False

        Card.Controls.Add(New TextBox With {
            .Multiline = True, .ReadOnly = True, .ScrollBars = ScrollBars.Vertical,
            .Bounds = New Rectangle(30, 70, 720, 270),
            .BorderStyle = BorderStyle.FixedSingle, .BackColor = Color.FromArgb(245, 247, 252),
            .Font = New Font("Consolas", 10.0F), .Text = BuildSummary()})

        ' edit shortcuts + download
        Dim btnEditBio As Button = AddButton("✎  Edit bio-data", 30, 350, 170, False)
        Dim btnEditEdu As Button = AddButton("✎  Edit education", 213, 350, 170, False)
        Dim btnEditSpon As Button = AddButton("✎  Edit sponsorship", 397, 350, 170, False)
        Dim btnDownload As Button = AddButton("⬇  Download", 580, 350, 170, False)
        AddHandler btnEditBio.Click, Sub(s, e)
                                         ReturnToReview = True
                                         NavigateTo(New BioDataForm())
                                     End Sub
        AddHandler btnEditEdu.Click, Sub(s, e)
                                         ReturnToReview = True
                                         NavigateTo(New EducationForm())
                                     End Sub
        AddHandler btnEditSpon.Click, Sub(s, e)
                                          ReturnToReview = True
                                          NavigateTo(New SponsorshipForm())
                                      End Sub
        AddHandler btnDownload.Click, Sub(s, e) DownloadApplication(Me)

        chk = New CheckBox With {
            .Text = "I declare that the information provided is true and complete.",
            .ForeColor = Color.White, .BackColor = Color.Transparent,
            .AutoSize = True, .Location = New Point(30, 405)}
        Card.Controls.Add(chk)

        Dim btnBack As Button = AddButton("‹  Back", 30, 470, 150, False)
        Dim btnSubmit As Button = AddButton(If(RefNo.Length > 0, "Update Application  ✔", "Submit Application  ✔"), 520, 470, 230, True)
        AddHandler btnBack.Click, Sub(s, e) NavigateTo(New SponsorshipForm())
        AddHandler btnSubmit.Click, AddressOf OnSubmit
    End Sub

    Private Sub OnSubmit(sender As Object, e As EventArgs)
        If Not chk.Checked Then
            Warn("Please tick the declaration before submitting.")
            Return
        End If
        Try
            Dim updating As Boolean = RefNo.Length > 0
            If Not updating Then RefNo = "SJVI-" & Date.Now.ToString("yyyyMMdd-HHmmss")
            WriteApplication()
            MessageBox.Show(Me, If(updating, "Application updated successfully!", "Application submitted successfully!") & vbCrLf & vbCrLf &
                            "Reference number: " & RefNo & vbCrLf & "Saved to: " & ApplicationFolder(),
                            "Submitted", MessageBoxButtons.OK, MessageBoxIcon.Information)
            If Saved("Payment status").Length > 0 Then
                NavigateTo(New ReceiptForm())
            Else
                NavigateTo(New PaymentForm())
            End If
        Catch ex As Exception
            Warn("Could not save the application: " & ex.Message)
        End Try
    End Sub
End Class

' ---------------------------------------------------------------------
'  6a. PAYMENT (shown right after submission)
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class PaymentForm
    Inherits BaseForm

    Private cboMethod As ComboBox
    Private txtA, txtB As TextBox
    Private lblHow As Label

    Public Sub New()
        MyBase.New("Application Fee Payment", 6)

        InfoRow("Application reference", RefNo, 70)
        InfoRow("Applicant", Saved("Full name"), 140)
        InfoRow("Course", Saved("Course applying for"), 210)
        AddLabel("Amount due", 30, 280)
        Card.Controls.Add(New Label With {
            .Text = FeeText(), .Font = New Font("Georgia", 22.0F, FontStyle.Bold),
            .ForeColor = Gold, .BackColor = Color.Transparent,
            .AutoSize = True, .Location = New Point(30, 306)})
        If Saved("Payment status").Length > 0 Then
            Card.Controls.Add(New Label With {
                .Text = "A payment is already recorded for this application. Confirming again will replace it.",
                .ForeColor = Color.FromArgb(225, 230, 242), .BackColor = Color.Transparent,
                .AutoSize = False, .Bounds = New Rectangle(30, 365, 340, 50)})
        End If

        cboMethod = AddCombo("Payment method *", 410, 70, 340, "MTN Mobile Money", "Airtel Money", "Bank deposit", "Cash at school office")
        txtA = AddText("Mobile money number used *", 410, 140, 340)
        txtB = AddText("Transaction ID *", 410, 210, 340)
        lblHow = New Label With {
            .ForeColor = Color.FromArgb(225, 230, 242), .BackColor = Color.Transparent,
            .AutoSize = False, .Bounds = New Rectangle(410, 285, 340, 130)}
        Card.Controls.Add(lblHow)

        Dim btnEdit As Button = AddButton("‹  Review / Edit", 30, 470, 170, False)
        Dim btnLater As Button = AddButton("Pay later", 215, 470, 130, False)
        Dim btnPay As Button = AddButton("Confirm Payment  ✔", 520, 470, 230, True)
        AddHandler btnEdit.Click, Sub(s, e) NavigateTo(New SubmissionForm())
        AddHandler btnLater.Click, Sub(s, e) NavigateTo(New ReceiptForm())
        AddHandler btnPay.Click, AddressOf OnConfirm
        AddHandler cboMethod.SelectedIndexChanged, Sub(s, e) UpdateMethod()
        cboMethod.SelectedIndex = 0
    End Sub

    Private Sub InfoRow(caption As String, value As String, y As Integer)
        AddLabel(caption, 30, y)
        Card.Controls.Add(New Label With {
            .Text = value, .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
            .ForeColor = Color.White, .BackColor = Color.Transparent,
            .AutoSize = False, .AutoEllipsis = True,
            .Bounds = New Rectangle(30, y + 24, 340, 28)})
    End Sub

    Private Function IsMobile() As Boolean
        Return cboMethod.Text = "MTN Mobile Money" OrElse cboMethod.Text = "Airtel Money"
    End Function

    Private Sub UpdateMethod()
        Dim capA As Label = DirectCast(txtA.Tag, Label)
        Dim capB As Label = DirectCast(txtB.Tag, Label)
        txtA.Clear()
        txtB.Clear()
        Select Case cboMethod.Text
            Case "MTN Mobile Money", "Airtel Money"
                ShowField(txtA, True)
                capA.Text = "Mobile money number used *"
                capB.Text = "Transaction ID *"
                lblHow.Text = "Send " & FeeText() & " to " & MobileMoneyNumber & " (" & MobileMoneyName & ") using " &
                              cboMethod.Text & ", then enter the transaction ID from the confirmation SMS."
            Case "Bank deposit"
                ShowField(txtA, True)
                capA.Text = "Bank name used *"
                capB.Text = "Deposit slip / transaction no. *"
                lblHow.Text = "Deposit " & FeeText() & " to:" & vbCrLf & BankDetails
            Case Else
                ShowField(txtA, False)
                capB.Text = "Office receipt no. *"
                lblHow.Text = "Pay " & FeeText() & " at the school bursar's office and enter the receipt number you were given."
        End Select
        If txtA.Visible Then
            Place(txtB, 210)
            lblHow.Top = 285
        Else
            Place(txtB, 140)
            lblHow.Top = 215
        End If
    End Sub

    Private Sub OnConfirm(sender As Object, e As EventArgs)
        If Not Check(cboMethod) Then Return
        If txtA.Visible AndAlso Not Check(txtA) Then Return
        If Not Check(txtB) Then Return
        If IsMobile() AndAlso Not IsPhone(txtA.Text) Then
            Warn("Please enter a valid mobile money number.")
            txtA.Focus()
            Return
        End If

        Dim details As String
        Select Case cboMethod.Text
            Case "MTN Mobile Money", "Airtel Money"
                details = "Number " & txtA.Text.Trim() & ", Transaction ID " & txtB.Text.Trim()
            Case "Bank deposit"
                details = "Bank " & txtA.Text.Trim() & ", Ref " & txtB.Text.Trim()
            Case Else
                details = "Office receipt no. " & txtB.Text.Trim()
        End Select
        Data("Payment status") = "Submitted - awaiting verification by the school"
        Data("Payment method") = cboMethod.Text
        Data("Payment details") = details
        Data("Paid on") = Date.Now.ToString("yyyy-MM-dd HH:mm")

        Try
            WriteApplication()
        Catch ex As Exception
            Warn("Payment recorded, but the file could not be updated: " & ex.Message)
        End Try
        MessageBox.Show(Me, "Payment details recorded. The school will verify your payment." & vbCrLf & vbCrLf &
                        "Reference number: " & RefNo, "Payment", MessageBoxButtons.OK, MessageBoxIcon.Information)
        NavigateTo(New ReceiptForm())
    End Sub
End Class

' ---------------------------------------------------------------------
'  6b. RECEIPT: download, print, edit, finish
' ---------------------------------------------------------------------
<DesignerCategory("Code")>
Public Class ReceiptForm
    Inherits BaseForm

    Public Sub New()
        MyBase.New("Application Complete", 6)

        Card.Controls.Add(New TextBox With {
            .Multiline = True, .ReadOnly = True, .ScrollBars = ScrollBars.Vertical,
            .Bounds = New Rectangle(30, 70, 720, 270),
            .BorderStyle = BorderStyle.FixedSingle, .BackColor = Color.FromArgb(245, 247, 252),
            .Font = New Font("Consolas", 10.0F), .Text = BuildDocument()})

        Dim btnDownload As Button = AddButton("⬇  Download (PDF / text)", 30, 350, 230, False)
        Dim btnPrint As Button = AddButton("Print", 270, 350, 230, False)
        Dim btnFolder As Button = AddButton("Open saved folder", 510, 350, 240, False)

        Dim paid As Boolean = Saved("Payment status").Length > 0
        Card.Controls.Add(New Label With {
            .Text = "Payment status: " & If(paid, Saved("Payment status"), "Not yet paid"),
            .ForeColor = Gold, .BackColor = Color.Transparent, .AutoSize = True,
            .Location = New Point(30, 408)})

        Dim btnEdit As Button = AddButton("✎  Edit application", 30, 470, 190, False)
        Dim btnPay As Button = AddButton(If(paid, "Change payment", "Pay now  ›"), 235, 470, 170, False)
        Dim btnFinish As Button = AddButton("Finish  ✔", 570, 470, 180, True)

        AddHandler btnDownload.Click, Sub(s, e) DownloadApplication(Me)
        AddHandler btnPrint.Click, Sub(s, e) PrintApplication(Me)
        AddHandler btnFolder.Click, AddressOf OnOpenFolder
        AddHandler btnEdit.Click, Sub(s, e)
                                      ReturnToReview = True
                                      NavigateTo(New SubmissionForm())
                                  End Sub
        AddHandler btnPay.Click, Sub(s, e) NavigateTo(New PaymentForm())
        AddHandler btnFinish.Click, AddressOf OnFinish
    End Sub

    Private Sub OnOpenFolder(sender As Object, e As EventArgs)
        Try
            Process.Start(New ProcessStartInfo(ApplicationFolder()) With {.UseShellExecute = True})
        Catch ex As Exception
            Warn("Could not open the folder: " & ex.Message)
        End Try
    End Sub

    Private Sub OnFinish(sender As Object, e As EventArgs)
        Data.Clear()
        Docs.Clear()
        RefNo = ""
        ReturnToReview = False
        NavigateTo(New LoginForm())
    End Sub
End Class

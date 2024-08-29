Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$url = "https://hr.tpi.flextronics.com/wUZLFlow/Default.aspx"
$user_name = "tpiwiche"
$password = "Chii153kawa"

<#
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PowerShell HTML GUI" WindowStartupLocation="Manual">
    <Grid>
        <WebBrowser
            Name="WebBrowser"
            Margin="10,10,0,0"
            Width="415"
            Height="340"
            HorizontalAlignment="Left"
            VerticalAlignment="Top"
        />
    </Grid>
</Window>
'@
$reader = New-Object System.Xml.XmlNodeReader($xaml)
$Form = [Windows.Markup.XamlReader]::Load($reader)
$Form.Width = 415
$Form.Height = 340
$Form.Topmost = $True
$WebBrowser = $Form.FindName('WebBrowser')
$WebBrowser.Navigate($url)
$WebBrowser.Add_DocumentCompleted({
        $doc = $this.Document
        if ($doc) {
            $doc.getElementById("ContentPlaceHolder1_txtUserName").InnerText = $user_name
            $doc.getElementById("ContentPlaceHolder1_txtPassword").InnerText = $password
            $doc.getElementById("ContentPlaceHolder1_cmdOK").Click()
        }
    })
# $syncHash.Window = $Form
# $syncHash.Browser = $WebBrowser
$Form.ShowDialog()

#>

##############################################
#<#
$form = New-Object System.Windows.Forms.Form
$form.Text = "My form"
$form.Size = New-Object System.Drawing.Size @(1070, 860)
$web = New-Object System.Windows.Forms.WebBrowser
$web.Location = New-Object System.Drawing.Point(3, 3)
$web.MinimumSize = New-Object System.Drawing.Size(20, 20)
$web.Size = New-Object System.Drawing.Size(1050, 840)
$web.ScriptErrorsSuppressed = $true
$web.Navigate($url)
$web.Add_DocumentCompleted({
        $doc = $web.Document
        if ($doc) {
            $login_btn = $doc.getElementById("ContentPlaceHolder1_cmdOK").DomElement
            if ($login_btn) {
                $doc.getElementById("ContentPlaceHolder1_txtUserName").InnerText = $user_name
                $doc.getElementById("ContentPlaceHolder1_txtPassword").InnerText = $password
                $login_btn.Click()
                $doc.getElementById("TreeView1t18").DomElement.Click()
                #$doc.getElementById("ContentPlaceHolder1_Button1").DomElement.Click()
            }
        }

        Write-Host "Login complete" -ForegroundColor Cyan
        #Unregister-Event -SourceIdentifier "Add_DocumentCompleted"
        #$web.Delete_DocumentCompleted({})
    })
$form.Controls.Add($web)
$form.ShowDialog()

#>

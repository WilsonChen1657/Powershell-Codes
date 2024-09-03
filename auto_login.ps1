Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$url = "https://hr.tpi.flextronics.com/wUZLFlow/Default.aspx"
$user_name = "tpiwiche"
$password = "Chii153kawa"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Auto login HR E-Flow System"
$form.Size = New-Object System.Drawing.Size @(1040, 710)

$web_browser = New-Object System.Windows.Forms.WebBrowser
$web_browser.Dock = 'Fill'
$web_browser.ScriptErrorsSuppressed = $true
$web_browser.Navigate($url)

$login_handler = {
    param (
        [System.Windows.Forms.WebBrowser]$browser,
        [System.Windows.Forms.WebBrowserDocumentCompletedEventArgs]$e
    )
    if ($browser -is [System.Windows.Forms.WebBrowser]) {
        # Get the document
        $doc = $browser.Document
        if ($doc) {
            $login_btn = $doc.getElementById("ContentPlaceHolder1_cmdOK").DomElement
            if ($login_btn) {
                # Unregister the event
                $browser.Remove_DocumentCompleted($login_handler)
                $browser.Add_DocumentCompleted($sign_in_treeview_handler)
                # Login
                $doc.getElementById("ContentPlaceHolder1_txtUserName").InnerText = $user_name
                $doc.getElementById("ContentPlaceHolder1_txtPassword").InnerText = $password
                $login_btn.Click()
                Write-Host "Login complete" -ForegroundColor Cyan
            }
        }
    }
}

$sign_in_treeview_handler = {
    param (
        [System.Windows.Forms.WebBrowser]$browser,
        [System.Windows.Forms.WebBrowserDocumentCompletedEventArgs]$e
    )
    if ($browser -is [System.Windows.Forms.WebBrowser]) {
        $doc = $browser.Document
        if ($doc) {
            $browser.Remove_DocumentCompleted($sign_in_treeview_handler)
            $browser.Add_DocumentCompleted($sign_in_handler)
            # Click "Sign In / Sign Out" treeview item
            $doc.getElementById("TreeView1t18").DomElement.Click()
        }
    }
}

$sign_in_handler = {
    param (
        [System.Windows.Forms.WebBrowser]$browser,
        [System.Windows.Forms.WebBrowserDocumentCompletedEventArgs]$e
    )
    if ($browser -is [System.Windows.Forms.WebBrowser]) {
        $doc = $browser.Document
        if ($doc) {
            $browser.Remove_DocumentCompleted($sign_in_handler)
            $now = Get-Date
            $base_datetime = Get-Date -Hour 12 -Minute 0 -Second 0
            # Sign in time
            $in_time = $doc.getElementById("ContentPlaceHolder1_txtInTime").InnerText
            # Sign out time
            $out_time = $doc.getElementById("ContentPlaceHolder1_txtOutTime").InnerText
            if (($null -eq $in_time) -and ($now.TimeOfDay -lt $base_datetime.TimeOfDay)) {
                Write-Host "Sign in" -ForegroundColor Cyan
                # Click "Sign In"
                #$doc.getElementById("ContentPlaceHolder1_Button1").DomElement.Click()
            }
            elseif (($null -eq $out_time) -and ($now.TimeOfDay -gt $base_datetime.TimeOfDay)) {
                Write-Host "Sign out" -ForegroundColor Cyan
                # Click "Sign Out"
                #$doc.getElementById("ContentPlaceHolder1_Button2").DomElement.Click()
            }
        }
    }
}

$web_browser.Add_DocumentCompleted($login_handler)
$form.Controls.Add($web_browser)
$form.ShowDialog()

##############################################

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

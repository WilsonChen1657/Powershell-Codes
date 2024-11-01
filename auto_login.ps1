Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$url = "https://hr.tpi.flextronics.com/wUZLFlow/Default.aspx"
$user_name = "tpiwiche"
$password = "w12210210W"

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
            # Language dropdown
            $dropdown = $doc.getElementById("ContentPlaceHolder1_ddlLanguage").DomElement
            $option_zh_tw = $dropdown | Where-Object { $_.value -eq "zh-TW" }
            if ( $option_zh_tw.selected -eq $false) {
                $option_zh_tw.selected = $true
                # Trigger the change event
                $dropdown.fireEvent("onchange")
            }
            else {
                $login_btn = $doc.getElementById("ContentPlaceHolder1_cmdOK").DomElement
                if ($login_btn) {
                    # Unregister the login event
                    $browser.Remove_DocumentCompleted($login_handler)
                    # Register sign in treeview event
                    $browser.Add_DocumentCompleted($sign_in_treeview_handler)
                    # Fill in the login details
                    $doc.getElementById("ContentPlaceHolder1_txtUserName").InnerText = $user_name
                    $doc.getElementById("ContentPlaceHolder1_txtPassword").InnerText = $password
                    $login_btn.Click()
                    Write-Host "Login complete" -ForegroundColor Cyan
                }
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
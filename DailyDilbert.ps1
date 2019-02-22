function displayImage
{
    param($file)

    # Load in Windows forms assembly
    [void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")

    $image = [System.Drawing.Image]::Fromfile($file)
    [void][System.Windows.Forms.Application]::EnableVisualStyles()
    
    # Create a picture box using winforms
    $form = New-Object Windows.Forms.Form
    $form.Text = "Image Viewer"
    $form.Width = $image.Size.Width
    $form.Height =  $image.Size.Height
    $pictureBox = New-Object Windows.Forms.PictureBox
    $pictureBox.Width =  $image.Size.Width
    $pictureBox.Height =  $image.Size.Height

    # Add image and display
    $pictureBox.Image = $image
    $form.Controls.Add($pictureBox)
    $form.Add_Shown( { $form.Activate() } )
    $form.ShowDialog()
}


function getDilbert
{
    param($path)

    try 
    {
        # Dilbert website stores its daily cartoon strip on "http://dilbert.com/strip/<todays date>"
        # 
        # We parse the url and look for the source image. 
        $url = "http://dilbert.com/strip/$(Get-Date -UFormat "%Y-%m-%d")"
    	$dilbert = Invoke-WebRequest -Uri $url -UseBasicParsing
    	$images = $dilbert.images | Where-Object { $_.class -eq 'img-responsive img-comic' } | Select-Object src
        Invoke-WebRequest -uri $("http:" + $images.src) -OutFile $path | Out-Null
    }
    catch
    {
	write-host "Ooops! There was a problem with the web request!"
    $error.Exception
	write-host "Launching previous Dilbert..."
    }
}


<#
    Dilbert web page has recently enforced higher level of encryption, therefore we need to force 
    PowerShell to use TLS1.2 for all web requests
#>
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
Write-Host "Please wait while we retrieve Dilbert..."
$file = "$env:USERPROFILE\DailyDilbert.gif"
getDilbert $file

if (Test-Path -Path $file)
    # If we can't get today's dilbert image then it will launch the previous 
    # one from the root of your user profile.
    {
        displayImage $file
    }
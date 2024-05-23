    
# =============================================================================================================================
#
# Script Name:     Resize-Image.ps1
# Description:     This script takes a source file and creates multiple versions of it using the image files located in C:\Windows\Web as templates.
#                  It duplicates the names, and metadata per file and recreates a folder structure in a defined location for later use
#
#                  
#=============================================================================================================================
    
    
    
    $MicrosoftWallpaper = "C:\Windows\web"
    $SourceWallpaperFiles =  (Get-ChildItem -Path $MicrosoftWallpaper -Recurse | Where-Object { $_.Extension -eq '.jpg' } |Select-Object fullname).fullname
    $SourceWallpaperFolders =  (Get-ChildItem -Path $MicrosoftWallpaper -Recurse -Directory  |Select-Object fullname).fullname


    #Branding source files
    #Single media image used at present
    $File = ### <Add source file(s) location here here>  


    # File and Folder output destination
    #$OutputImageDestination = ### <Add destination output location here here>  

    if(!(Test-path $OutputImageDestination)) {
            New-Item -Path $OutputImageDestination  -ItemType Directory -Force  -ErrorAction Stop 
     }
 


     #Emulate Folder structure
     
     foreach($SourceWallpaperFolder in $SourceWallpaperFolders){
        Copy-Item $SourceWallpaperFolder -Destination $SourceWallpaperFolder.Replace('C:\Windows\web',$OutputImageDestination) -Force
     }

     add-type -AssemblyName System.Drawing

     foreach ($WallpaperFile in $SourceWallpaperFiles){
       
        # OUtput path for image to be saved to
        $OutputPath = $WallpaperFile.Replace('C:\Windows\web',$OutputImageDestination)
                
        # Open image file
        $img = [System.Drawing.Image]::FromFile((Get-Item $File))

        $image = [System.Drawing.Image]::FromFile((Get-Item $WallpaperFile))
        [Int]$OldHeight = $image.Height
        [Int]$OldWidth = $image.Width
        $HorizontalResolution  = [Int]$image.HorizontalResolution       #In DPI
        $VerticalResolution    = [Int]$image.VerticalResolution       #In DPI
        $PixelFormat =  [System.Drawing.Image]::GetPixelFormatSize($image.PixelFormat)

        $destination_rect = New-Object System.Drawing.Rectangle 0, 0, $OldWidth, $OldHeight
        $destination_image = New-Object System.Drawing.Bitmap $OldWidth, $OldHeight , 'Format24bppRgb'       
        $graphics = [System.Drawing.Graphics]::FromImage($destination_image)

        $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $wrapmode = [System.Drawing.Imaging.ImageAttributes]::new()
        $wrapmode.SetWrapMode([System.Drawing.Drawing2D.WrapMode]::TileFlipXY)

        $graphics.DrawImage($img, $destination_rect, 0, 0, $image.Width, $image.Height , [System.Drawing.GraphicsUnit]::Pixel, $wrapmode)

        $destination_image.Save($OutputPath)

        $destination_image.Dispose()
        $graphics.Dispose()
        $wrapmode.Dispose()
        $img.Dispose()
        $image.Dispose()
    }

   

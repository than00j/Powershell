$SysIcon = Get-SystemIcon -IconNo 18
$Bitmap = $SysIcon.ToBitmap()
$hBitmap = $Bitmap.GetHbitmap()
$wpfBitmap = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
                $hBitmap, 
                [System.IntPtr]::Zero, 
                [System.Windows.Int32Rect]::Empty, 
                [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())

$Window.Icon = $wpfBitmap

param(
  [ValidateSet('TEST','SELECTED','NVPRINT','STOP')][string]$Action = 'TEST',
  [string]$Port = 'COM1',
  [ValidateSet(9600,38400,115200)][int]$Baud = 9600,
  [string]$ImagePath = ''
)
$ErrorActionPreference = 'Stop'

function Open-PrinterPort {
  param([string]$Name,[int]$Rate)
  $sp = New-Object System.IO.Ports.SerialPort
  $sp.PortName = $Name
  $sp.BaudRate = $Rate
  $sp.Parity = [System.IO.Ports.Parity]::None
  $sp.DataBits = 8
  $sp.StopBits = [System.IO.Ports.StopBits]::One
  $sp.Handshake = [System.IO.Ports.Handshake]::None
  $sp.ReadTimeout = 500
  $sp.WriteTimeout = 3000
  $sp.Open()
  return $sp
}

function Send-Bytes {
  param($Sp,[byte[]]$Bytes)
  $Sp.Write($Bytes,0,$Bytes.Length)
}

function Send-Test {
  param($Sp)
  # ESC/POS 초기화 + 테스트 출력. 7403S 계열 직렬 프린터에서 일반적인 8N1 통신을 사용합니다.
  Send-Bytes $Sp ([byte[]](0x1B,0x40))
  $enc = [System.Text.Encoding]::GetEncoding(949)
  $text = "테스트 출력`r`n------------------------------`r`nPOS CARE PRO`r`n프린터 통신 테스트 정상`r`nCOM PORT : $Port`r`nBAUD     : $Baud`r`n------------------------------`r`n테스트 출력`r`n`r`n"
  Send-Bytes $Sp $enc.GetBytes($text)
  # 용지를 충분히 올려 사용자가 출력물을 바로 확인할 수 있도록 feed 합니다.
  Send-Bytes $Sp ([byte[]](0x1B,0x64,0x05))
  # ESC/POS 자동 커팅: 용지를 충분히 급지한 뒤 전체 컷 명령을 전송합니다.
  # 프린터가 자동 커터를 지원하지 않으면 이 명령은 무시될 수 있습니다.
  Send-Bytes $Sp ([byte[]](0x1D,0x56,0x00))
}

function Send-RasterImage {
  param($Sp,[string]$Path)
  if([string]::IsNullOrWhiteSpace($Path)){ throw 'NV image 파일을 선택해 주세요.' }
  if(!(Test-Path -LiteralPath $Path -PathType Leaf)){ throw '선택한 NV image 파일을 찾을 수 없습니다.' }
  Add-Type -AssemblyName System.Drawing
  $src = New-Object System.Drawing.Bitmap($Path)
  try {
    $maxWidth = 576
    $w = $src.Width
    $h = $src.Height
    if($w -gt $maxWidth){
      $h = [int][Math]::Round($h * ($maxWidth / [double]$w))
      $w = $maxWidth
    }
    $bmp = New-Object System.Drawing.Bitmap($w,$h,[System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    try {
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      try {
        $g.Clear([System.Drawing.Color]::White)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.DrawImage($src,0,0,$w,$h)
      } finally { $g.Dispose() }

      $bytesPerRow = [int][Math]::Ceiling($w / 8.0)
      $data = New-Object byte[] ($bytesPerRow * $h)
      for($y=0;$y -lt $h;$y++){
        for($x=0;$x -lt $w;$x++){
          $c = $bmp.GetPixel($x,$y)
          $gray = (0.299*$c.R + 0.587*$c.G + 0.114*$c.B)
          if($gray -lt 180){
            $idx = $y*$bytesPerRow + [int]($x/8)
            $data[$idx] = $data[$idx] -bor (1 -shl (7 - ($x % 8)))
          }
        }
      }
      Send-Bytes $Sp ([byte[]](0x1B,0x40))
      $header = New-Object byte[] 8
      $header[0]=0x1D;$header[1]=0x76;$header[2]=0x30;$header[3]=0x00
      $header[4]=[byte]($bytesPerRow -band 0xFF);$header[5]=[byte](($bytesPerRow -shr 8)-band 0xFF)
      $header[6]=[byte]($h -band 0xFF);$header[7]=[byte](($h -shr 8)-band 0xFF)
      Send-Bytes $Sp $header
      Send-Bytes $Sp $data
      Send-Bytes $Sp ([byte[]](0x0A,0x0A))
    } finally { $bmp.Dispose() }
  } finally { $src.Dispose() }
}

$sp = $null
try {
  $sp = Open-PrinterPort $Port $Baud
  switch($Action){
    'TEST'     { Send-Test $sp }
    'SELECTED' { Send-RasterImage $sp $ImagePath }
    'NVPRINT'  { Send-RasterImage $sp $ImagePath }
    'STOP'     { Send-Bytes $sp ([byte[]](0x1B,0x40)) }
  }
  $sp.Close()
  exit 0
} catch {
  if($sp){ try{$sp.Close()}catch{} }
  Write-Error $_.Exception.Message
  exit 2
}

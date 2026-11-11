# 이름표 인쇄 도우미 — 현황지 "이름표 출력"이 보낸 그림을 TSC TTP-247 드라이버로 직접 출력한다.
# 이유: 크롬 인쇄는 이 프린터 드라이버에서 안 나오고(빨간불), 이 프린터는 가로로 이어진 꽉 찬 검은 면(약 35mm 이상)을
#       못 찍는다. 그래서 검은 면을 50% 격자 망점으로 바꾸고 드라이버 기본 용지·1:1(203dpi)로 보낸다.
# 사용: PC 시작 시 자동 실행(시작프로그램의 NameTagHelper.vbs). 127.0.0.1:9123 에서만 듣는다.
param([string]$Printer = 'TSC TTP-247', [int]$Port = 9123)
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;using System.Drawing;using System.Drawing.Imaging;
public static class NtHalftone{
 // 검은 면을 격자 50%로 만들고, 가로 점이 많은 줄이 연달아 오면 절반으로 더 줄인다.
 public static Bitmap Run(Bitmap src,int limit){
  int W=src.Width,H=src.Height;var r=new Bitmap(W,H,PixelFormat.Format24bppRgb);
  int prev=0;
  for(int y=0;y<H;y++){
   var row=new bool[W];int n=0;
   for(int x=0;x<W;x++){var c=src.GetPixel(x,y);bool dark=(c.R+c.G+c.B)/3<170;bool on=dark&&(((x+y)&1)==0);row[x]=on;if(on)n++;}
   if(n>limit&&prev>limit){int k=0;n=0;for(int x=0;x<W;x++){if(row[x]){if((k++&1)==1)row[x]=false;else n++;}}}
   for(int x=0;x<W;x++)r.SetPixel(x,y,row[x]?Color.Black:Color.White);
   prev=n;}
  return r;}
}
"@
function Send-Print([Drawing.Bitmap]$bmp){
  $d = New-Object System.Drawing.Printing.PrintDocument
  $d.PrinterSettings.PrinterName = $Printer
  $d.add_PrintPage({ param($s,$e)
    $g=$e.Graphics; $g.PageUnit=[Drawing.GraphicsUnit]::Pixel
    $g.InterpolationMode='NearestNeighbor'; $g.PixelOffsetMode='Half'
    $g.DrawImage($bmp,(New-Object Drawing.Rectangle(0,0,800,800)))
  }.GetNewClosure())
  $d.Print()
}
$l = New-Object System.Net.HttpListener
$l.Prefixes.Add("http://127.0.0.1:$Port/")
$l.Start()
while ($l.IsListening) {
  $ctx = $l.GetContext(); $req=$ctx.Request; $res=$ctx.Response
  try {
    $res.Headers.Add('Access-Control-Allow-Origin','*')
    $res.Headers.Add('Access-Control-Allow-Methods','POST, GET, OPTIONS')
    $res.Headers.Add('Access-Control-Allow-Headers','Content-Type')
    $res.Headers.Add('Access-Control-Allow-Private-Network','true')
    $body = 'ok'
    if ($req.HttpMethod -eq 'OPTIONS') { $res.StatusCode = 204 }
    elseif ($req.HttpMethod -eq 'POST' -and $req.Url.AbsolutePath -eq '/print') {
      $sr = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8); $txt = $sr.ReadToEnd(); $sr.Dispose()
      $b64 = ($txt -replace '^.*base64,','') -replace '[\s"}]',''
      $ms = New-Object IO.MemoryStream(,[Convert]::FromBase64String($b64))
      $src = New-Object Drawing.Bitmap($ms)
      if ($src.Width -ne 800 -or $src.Height -ne 800) { $t = New-Object Drawing.Bitmap(800,800); $gg=[Drawing.Graphics]::FromImage($t); $gg.DrawImage($src,0,0,800,800); $gg.Dispose(); $src=$t }
      $out = [NtHalftone]::Run($src, 320)
      Send-Print $out
      $out.Dispose(); $src.Dispose()
    } elseif ($req.HttpMethod -eq 'GET') { $body = "nametag-helper $Printer" }
    else { $res.StatusCode = 404; $body = 'not found' }
    $buf=[Text.Encoding]::UTF8.GetBytes($body); $res.ContentLength64=$buf.Length; $res.OutputStream.Write($buf,0,$buf.Length)
  } catch { try { $res.StatusCode=500; $m=[Text.Encoding]::UTF8.GetBytes("error: $($_.Exception.Message)"); $res.OutputStream.Write($m,0,$m.Length) } catch {} }
  finally { try { $res.Close() } catch {} }
}

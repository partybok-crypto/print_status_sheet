# 이름표 인쇄 도우미 — 현황지 "이름표 출력"이 보낸 그림을 TSC TTP-247 드라이버로 직접 출력한다.
# 이유: 크롬 인쇄는 이 프린터 드라이버에서 안 나오고(빨간불), 이 프린터는 가로로 이어진 꽉 찬 검은 면(약 35mm 이상)을
#       못 찍는다. 그래서 검은 면을 40% 격자 망점으로 바꾸고 드라이버 기본 용지·1:1(203dpi)로 보낸다.
# 사용: PC 시작 시 자동 실행(시작프로그램의 NameTagHelper.vbs). 127.0.0.1:9123 에서만 듣는다.
param([string]$Printer = 'TSC TTP-247', [int]$Port = 9123, [int]$Budget = 100000)
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;using System.Drawing;using System.Drawing.Imaging;
public static class NtHalftone{
 // 검은 점이 budget 개보다 많은 줄은 균일하게 솎아 budget 개로 만든다(줄마다 위상을 달리해 세로 정렬을 막는다).
 // 점이 적은 줄은 그대로 둔다. 이유: 이 프린터는 한 줄에 켜지는 점 수에 한계가 있고(약 100~250), 넘으면 빨간불로 멈춘다.
 public static Bitmap Run(Bitmap src,int budget){
  int W=src.Width,H=src.Height;var r=new Bitmap(W,H,PixelFormat.Format24bppRgb);
  for(int y=0;y<H;y++){
   var row=new bool[W];int n=0;
   for(int x=0;x<W;x++){var c=src.GetPixel(x,y);bool dark=(c.R+c.G+c.B)/3<170;row[x]=dark;if(dark)n++;}
   if(n>budget){
    // 점을 3개씩 붙여서(런) 찍는다: 한 줄 점 수는 budget 그대로, 흩어진 1점보다 잉크가 잘 옮겨붙어 더 진하다.
    // 위아래 2줄은 같은 위상으로 묶어 3x2 덩어리를 만든다.
    int run=4;double period=(double)n*run/budget;if(period<run)period=run;
    double ph=(((y/2)*0.6180339887)%1.0)*period;int k=0;
    for(int x=0;x<W;x++){if(row[x]){double pos=(k+ph)%period;k++;if(pos>=run)row[x]=false;}}}
   for(int x=0;x<W;x++)r.SetPixel(x,y,row[x]?Color.Black:Color.White);}
  return r;}
}
"@
function Send-Print([Drawing.Bitmap]$bmp){
  $d = New-Object System.Drawing.Printing.PrintDocument
  $d.PrinterSettings.PrinterName = $Printer
  $paper = $d.PrinterSettings.PaperSizes | ? { $_.PaperName -eq '새 용지' } | select -First 1
  if ($paper) { $d.DefaultPageSettings.PaperSize = $paper }
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
      $out = [NtHalftone]::Run($src, $Budget)
      Send-Print $out
      $out.Dispose(); $src.Dispose()
    } elseif ($req.HttpMethod -eq 'GET') { $body = "nametag-helper $Printer" }
    else { $res.StatusCode = 404; $body = 'not found' }
    $buf=[Text.Encoding]::UTF8.GetBytes($body); $res.ContentLength64=$buf.Length; $res.OutputStream.Write($buf,0,$buf.Length)
  } catch { try { $res.StatusCode=500; $m=[Text.Encoding]::UTF8.GetBytes("error: $($_.Exception.Message)"); $res.OutputStream.Write($m,0,$m.Length) } catch {} }
  finally { try { $res.Close() } catch {} }
}








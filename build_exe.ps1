# ============================================================
#  build_exe.ps1 — Gera o executável standalone do Gerador Revinn
#  Inclui automaticamente as DLLs do GTK3 para que o usuário
#  final NÃO precise instalar nada além de copiar a pasta.
# ============================================================

$ErrorActionPreference = "Stop"

# --- 1. Localizar GTK3 ---
$gtkPaths = @(
    "C:\Program Files\GTK3-Runtime Win64\bin",
    "C:\Program Files (x86)\GTK3-Runtime Win64\bin",
    "C:\msys64\mingw64\bin",
    "C:\tools\msys64\mingw64\bin"
)

$gtkBin = $null
foreach ($p in $gtkPaths) {
    if (Test-Path $p) { $gtkBin = $p; break }
}

if (-not $gtkBin) {
    Write-Error "GTK3 Runtime nao encontrado. Instale em: https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases"
    exit 1
}

Write-Host "GTK3 encontrado em: $gtkBin" -ForegroundColor Green

# --- 2. Montar lista de --add-binary para todas as DLLs do GTK3 ---
$dllList = Get-ChildItem "$gtkBin\*.dll"
Write-Host "$($dllList.Count) DLLs do GTK3 serao incluidas" -ForegroundColor Green

# --- 3. Incluir subpastas importantes do GTK3 ---
$gtkRoot = Split-Path $gtkBin -Parent
$extraDataArgs = [System.Collections.Generic.List[string]]::new()

$subDirs = @(
    @{ src = "share\glib-2.0\schemas"; rel = "share/glib-2.0/schemas" },
    @{ src = "lib\gdk-pixbuf-2.0";     rel = "lib/gdk-pixbuf-2.0" },
    @{ src = "etc\pango";              rel = "etc/pango" },
    @{ src = "etc\gtk-3.0";            rel = "etc/gtk-3.0" }
)

foreach ($d in $subDirs) {
    $full = Join-Path $gtkRoot $d.src
    if (Test-Path $full) {
        $extraDataArgs.Add("--add-data")
        $extraDataArgs.Add("${full};$($d.rel)")
        Write-Host "  + incluindo: $full" -ForegroundColor Cyan
    }
}

# --- 4. Montar argumentos base ---
$baseArgs = [System.Collections.Generic.List[string]]@(
    "--noconfirm",
    "--onedir",
    "--windowed",
    "--name", "Gerador_Relatorios_Revinn",
    "--add-data", "Logo Revinn;Logo Revinn",
    "--collect-all", "customtkinter",
    "--collect-all", "weasyprint",
    "--collect-all", "pydyf",
    "--collect-all", "cffi",
    "--hidden-import", "PIL._tkinter_finder",
    "--hidden-import", "weasyprint.text.fonts"
)

# Adicionar cada DLL individualmente
foreach ($dll in $dllList) {
    $baseArgs.Add("--add-binary")
    $baseArgs.Add("$($dll.FullName);.")
}

# Adicionar subpastas
foreach ($arg in $extraDataArgs) {
    $baseArgs.Add($arg)
}

$baseArgs.Add("main.py")

# --- 5. Executar PyInstaller ---
Write-Host ""
Write-Host "Iniciando build..." -ForegroundColor Yellow

$pyinstaller = ".\venv\Scripts\pyinstaller.exe"
& $pyinstaller @baseArgs

if ($LASTEXITCODE -eq 0) {
    # --- 6. Gerar README (LEIA-ME) ao lado do executavel ---
    $distDir = "dist\Gerador_Relatorios_Revinn"
    $readmeText = @"
============================================================
  GERADOR DE RELATORIOS - REVINN ESTRATEGIAS
  Como executar a aplicacao
============================================================

COMO ABRIR
----------
1. Mantenha ESTA PASTA COMPLETA no computador (nao mova nem
   copie apenas o arquivo .exe).
2. De dois cliques em:

       Gerador_Relatorios_Revinn.exe

   O executavel precisa ficar DENTRO desta pasta, junto com a
   pasta "_internal" e demais arquivos. Ela contem as
   bibliotecas do WeasyPrint, GTK3 (DLLs), fontes, logo etc.
   necessarias para gerar o PDF.

INSTALACAO
----------
- Nao e necessario instalar Python, GTK3 ou qualquer outra
  dependencia. Tudo ja esta incluido nesta pasta.
- Para enviar a outra pessoa, compacte (ZIP) a pasta INTEIRA
  e peca que seja extraida por completo antes de abrir.

ATALHO
------
Para criar um atalho na Area de Trabalho: clique com o botao
direito em Gerador_Relatorios_Revinn.exe > Enviar para >
Area de trabalho (criar atalho). NAO copie o .exe sozinho.

USO BASICO
----------
1. Selecione os CSVs de Check-outs e Reservas Emitidas.
2. (Opcional) Selecione a planilha de Pickup (.xlsx).
3. Escolha a pasta de destino e clique em GERAR RELATORIOS.
4. Serao gerados: .xlsx, .docx e .pdf (Lamina de Performance).

PROBLEMAS COMUNS
----------------
- O programa nao abre / fecha sozinho: verifique se voce
  extraiu o ZIP completo e se esta executando o .exe de
  dentro da pasta original.
- Antivirus bloqueou: adicione a pasta as excecoes (o app nao
  e assinado digitalmente).
- PDF nao gerado: confirme que a pasta "_internal" nao foi
  apagada ou movida.

Revinn Estrategias - Todos os direitos reservados.
"@
    Set-Content -Path (Join-Path $distDir "LEIA-ME.txt") -Value $readmeText -Encoding UTF8

    Write-Host ""
    Write-Host "Build concluido com sucesso!" -ForegroundColor Green
    Write-Host "Executavel em: dist\Gerador_Relatorios_Revinn\Gerador_Relatorios_Revinn.exe"
    Write-Host "LEIA-ME.txt gerado em: $distDir" -ForegroundColor Green
    Write-Host ""
    Write-Host "Para distribuir, envie a PASTA COMPLETA:" -ForegroundColor Cyan
    Write-Host "dist\Gerador_Relatorios_Revinn\" -ForegroundColor White
    Write-Host "O destinatario NAO precisa instalar Python, GTK3 ou qualquer dependencia." -ForegroundColor Green
} else {
    Write-Host "Build falhou. Verifique as mensagens acima." -ForegroundColor Red
}

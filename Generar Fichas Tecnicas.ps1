# Regenera "Fichas Tecnicas Caujaral.html" a partir de los libros de recetas.
# Sin acentos a proposito: el archivo se guarda sin BOM y los acentos se corromperian.
# Imprime al final una linea de estado que la tarea programada usa para decidir:
#   ESTADO=SIN-CAMBIOS | ESTADO=ACTUALIZADO | ESTADO=ERROR

$ErrorActionPreference = "Stop"

$raiz   = Split-Path -Parent $MyInvocation.MyCommand.Path
$origen = "C:\Users\SergioValencia\OneDrive - CLUB LAGOS DE CAUJARAL\DOCUMENTOS 2026\LIBRO DE RECETAS"
$work   = Join-Path $raiz "trabajo"
$salida = Join-Path $raiz "Fichas Tecnicas Caujaral.html"
$hashFn = Join-Path $raiz "ultimo hash.txt"

try {
  Import-Module ImportExcel -ErrorAction Stop
  New-Item -ItemType Directory -Force -Path $work | Out-Null

  # Los libros se BUSCAN POR PATRON, no por nombre exacto: asi Sergio puede
  # renombrarlos, unirlos en uno solo o agregar otro, sin que haya que tocar el codigo.
  $patron = "LIBRO DE RECETAS AUTOMATIZADOS V.2*.xlsx"
  $encontrados = @(Get-ChildItem -Path $origen -Filter $patron -File -ErrorAction Stop |
                   Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)
  if ($encontrados.Count -eq 0) {
    throw ("No se encontro ningun libro que coincida con '" + $patron + "' en " + $origen)
  }

  # Se trabaja siempre sobre copias: el original puede estar abierto en Excel o sincronizando.
  $libros = @()
  $n = 0
  foreach ($a in $encontrados) {
    $n++
    $etiqueta = [System.IO.Path]::GetFileNameWithoutExtension($a.Name)
    $etiqueta = ($etiqueta -replace '^LIBRO DE RECETAS AUTOMATIZADOS V\.2\s*','').Trim()
    if ($etiqueta -eq "") { $etiqueta = "LIBRO $n" }
    $copia = "libro$n.xlsx"
    Copy-Item $a.FullName (Join-Path $work $copia) -Force
    $libros += @{ arch = $a.Name; copia = $copia; libro = $etiqueta }
  }
  "libros encontrados: {0}" -f (($libros | ForEach-Object { $_.arch }) -join " | ")

  function Esc([string]$s) {
    if ($null -eq $s) { return "" }
    $s = $s -replace '\\','\\\\'
    $s = $s -replace '"','\"'
    $s = $s -replace "`r",''
    $s = $s -replace "`n",'\n'
    $s = $s -replace "`t",' '
    return $s
  }
  function Num([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return 0 }
    $s = $s -replace '[^0-9,\.\-]',''
    $s = $s -replace ',',''
    $d = 0.0
    if ([double]::TryParse($s, [ref]$d)) { return $d }
    return 0
  }

  # Quita tildes y normaliza, para comparar encabezados sin pelear con la ortografia.
  function NormEnc([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $t = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $b = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
      if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
        [void]$b.Append($ch)
      }
    }
    return (($b.ToString().ToUpper()) -replace '\s+',' ').Trim()
  }

  # Las dos hojas BASE NO tienen la misma estructura: en BEBIDAS la columna H es
  # "TIPO DE MENAJE" y en ALIMENTOS es "Columna1", un sobrante de tabla de Excel.
  # Por eso las columnas se ubican por el ENCABEZADO (fila 3) y no por posicion:
  # asi cada libro puede diferir y la columna FOTO puede ir donde sea.
  function MapaBase($ws) {
    $m = @{}
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
      $h = NormEnc $ws.Cells[3,$c].Text
      if ($h -eq "" -or $h -like "COLUMNA*") { continue }
      if     ($h -eq "FORMULA")                                    { $m["formula"] = $c }
      elseif ($h -like "NOMBRE*")                                  { $m["nombre"]  = $c }
      elseif ($h -like "*CATEGORIA*")                              { $m["cat"]  = $c }
      elseif ($h -like "*METODO*")                                 { $m["met"]  = $c }
      elseif ($h -like "*UTENCILIO*" -or $h -like "*UTENSILIO*")   { $m["ute"]  = $c }
      elseif ($h -like "*MENAJE*")                                 { $m["men"]  = $c }
      elseif ($h -like "*FOTO*" -or $h -like "*IMAGEN*")           { $m["foto"] = $c }
      elseif ($h -like "ELABOR*")                                  { $m["el"]   = $c }
      elseif ($h -like "APROB*")                                   { $m["ap"]   = $c }
      elseif ($h -like "PARAMETRIZ*")                              { $m["pa"]   = $c }
    }
    return $m
  }

  # Los errores de Excel (#VALUE!, #N/A, #REF!...) se tratan como celda vacia:
  # una formula rota no puede terminar mostrandose como texto en la ficha del operario.
  function Celda($ws, $mapa, [string]$campo, [int]$fila) {
    if (-not $mapa.ContainsKey($campo)) { return "" }
    $v = $ws.Cells[$fila, $mapa[$campo]].Text.Trim()
    if ($v -match '^#(VALUE|VALOR|N/A|REF|NAME|NOMBRE|DIV/0|NUM|NULL|NULO|SPILL|CALC)') { return "" }
    return $v
  }

  # Lee el gramaje que viene escrito en el propio nombre del articulo.
  # Regla de Sergio para licores (grupos 2-12 / 2-13 / 2-14): la cantidad parametrizada
  # es la fraccion de botella, asi que cantidad x los ML del nombre da los ML servidos.
  # 0,045 de "TEQUILA OLMECA REPOSADO 1000 ML (667)" = 45 ml.
  # El numero entre parentesis es el PESO DE LA BOTELLA VACIA: sirve para pesar inventario,
  # NO entra en esta conversion.
  function Presentacion([string]$nombre) {
    $t = $nombre.ToUpper()
    $t = $t -replace '\(\s*\d+(?:[\.,]\d+)?\s*\)', ' '   # fuera el peso de botella
    if ($t -match '(\d+(?:[\.,]\d+)?)\s*X?\s*ML')              { return @{ v = [double](($Matches[1]) -replace ',','.');        u = "ML" } }
    if ($t -match '(\d+(?:[\.,]\d+)?)\s*(?:GR|GRM|GRAMOS)\b')  { return @{ v = [double](($Matches[1]) -replace ',','.');        u = "GR" } }
    if ($t -match '(\d+(?:[\.,]\d+)?)\s*(?:KG|KL)\b')          { return @{ v = [double](($Matches[1]) -replace ',','.') * 1000; u = "GR" } }
    if ($t -match '(\d+(?:[\.,]\d+)?)\s*(?:LT)\b')             { return @{ v = [double](($Matches[1]) -replace ',','.') * 1000; u = "ML" } }
    return $null
  }

  # ---- ingredientes (hoja "Query 1") ----
  # EL CATALOGO SE LEE DE UN SOLO LIBRO: el primero en orden alfabetico (el consolidado).
  # Query 1 y Query2 son exports del sistema, y mezclar dos exports de fechas distintas
  # DUPLICA ingredientes cuando el sistema renombra un articulo: aparecian juntos
  # "RC SALSA TARTARA" (export viejo) y "SB SALSA TARTARA" (export nuevo) en la misma receta.
  # Los libros secundarios solo aportan su hoja BASE.
  $principal = @($libros[0])
  if ($libros.Count -gt 1) {
    "  catalogo (Query 1 y Query2): solo desde {0}" -f $libros[0].libro
  }
  $ing = @{}; $hdr = @{}
  $costoEstandar = 0
  $costoDeUltimaCompra = 0
  foreach ($l in $principal) {
    $p  = Open-ExcelPackage -Path (Join-Path $work $l.copia)
    $ws = $p.Workbook.Worksheets["Query 1"]
    if ($null -eq $ws -or $null -eq $ws.Dimension) {
      Close-ExcelPackage $p -NoSave
      throw ("El libro principal (" + $l.libro + ") no tiene la hoja 'Query 1'. Sin catalogo no se puede publicar.")
    }
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
      $cod = $ws.Cells[$r,1].Text.Trim()
      if ($cod -eq "") { continue }
      $art = $ws.Cells[$r,3].Text.Trim()
      if ($art -eq "") { continue }
      $pre = $ws.Cells[$r,4].Text.Trim()
      # El costo sale de "CostoEstandar" (col 5), pero desde el export del 06/08/2026
      # el sistema lo manda en CERO en todas las filas. Cuando eso pasa se toma
      # "CostoUltimaCompra" (col 7), que si trae valor. Sin este respaldo la ficha
      # mostraba todos los costos en $ 0.
      $cst = Num $ws.Cells[$r,5].Text
      if ($cst -le 0) {
        $cst = Num $ws.Cells[$r,7].Text
        if ($cst -gt 0) { $script:costoDeUltimaCompra++ }
      } else { $script:costoEstandar++ }
      $cnt = Num $ws.Cells[$r,6].Text
      if (-not $ing.ContainsKey($cod)) {
        $ing[$cod] = New-Object System.Collections.ArrayList
        $hdr[$cod] = $ws.Cells[$r,2].Text.Trim()
      }
      $key = "$art|$pre"; $dup = $false
      foreach ($x in $ing[$cod]) { if ($x.k -eq $key) { $dup = $true; break } }
      if (-not $dup) { [void]$ing[$cod].Add([pscustomobject]@{k=$key; a=$art; u=$pre; q=$cnt; c=$cst}) }
    }
    Close-ExcelPackage $p -NoSave
  }

  # ---- cabeceras (hoja "Query2") ----
  $meta = @{}
  foreach ($l in $principal) {
    $p  = Open-ExcelPackage -Path (Join-Path $work $l.copia)
    $ws = $p.Workbook.Worksheets["Query2"]
    if ($null -eq $ws -or $null -eq $ws.Dimension) {
      "  Query2 de {0}: no existe o esta vacia, se omite" -f $l.libro
      Close-ExcelPackage $p -NoSave
      continue
    }
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
      $cod = $ws.Cells[$r,5].Text.Trim()
      if ($cod -eq "") { continue }
      $meta[$cod] = [pscustomobject]@{
        rend = Num $ws.Cells[$r,2].Text
        um   = $ws.Cells[$r,3].Text.Trim()
        cost = Num $ws.Cells[$r,4].Text
        nom  = $ws.Cells[$r,1].Text.Trim()
      }
    }
    Close-ExcelPackage $p -NoSave
  }

  # Nombre de receta reducido a lo esencial, para poder emparejar aunque cambien
  # tildes, mayusculas, signos o el prefijo "RC"/"RECETA" que usa el sistema.
  function NormNombre([string]$s) {
    $x = NormEnc $s
    $x = $x -replace '[^A-Z0-9 ]',' '
    $x = ($x -replace '\s+',' ').Trim()
    $x = $x -replace '^(RC|RECETA|SUB|SB|AD)\s+',''
    return $x.Trim()
  }

  # Indice nombre -> codigo. Sirve de RED DE SEGURIDAD: la columna FORMULA de la hoja
  # BASE es un VLOOKUP contra Query2, y cuando se re-exportan las consultas los nombres
  # dejan de coincidir y devuelve #N/A. Sin esto, cada re-export borraba de la ficha
  # decenas de metodos que en realidad seguian escritos.
  $porNombre = @{}
  foreach ($c in $meta.Keys) {
    $n = NormNombre $meta[$c].nom
    if ($n -ne "" -and -not $porNombre.ContainsKey($n)) { $porNombre[$n] = $c }
  }
  foreach ($c in $hdr.Keys) {
    $n = NormNombre $hdr[$c]
    if ($n -ne "" -and -not $porNombre.ContainsKey($n)) { $porNombre[$n] = $c }
  }

  # ---- ficha operativa (hoja "BASE"): categoria, metodo, utensilios, menaje ----
  $rescatadas = 0
  $huerfanas = New-Object System.Collections.ArrayList
  # Si un mismo codigo aparece en varios libros, MANDA EL PRIMERO en orden alfabetico
  # (hoy: ALIMENTOS Y BEBIDAS, el centralizado). Asi, cuando una receta quedo copiada
  # en dos libros, gana la que Sergio esta editando y no una version vieja del otro.
  $base = @{}
  $idxLibro = 0
  foreach ($l in $libros) {
    $idxLibro++
    $p  = Open-ExcelPackage -Path (Join-Path $work $l.copia)
    $ws = $p.Workbook.Worksheets["BASE"]
    # La hoja BASE puede no estar, o estar vacia, si Sergio decide concentrar toda
    # la informacion operativa en un solo libro. Eso no es un error: se avisa y se sigue.
    if ($null -eq $ws -or $null -eq $ws.Dimension) {
      "  BASE de {0}: no existe o esta vacia, se omite" -f $l.libro
      Close-ExcelPackage $p -NoSave
      continue
    }
    $mapa = MapaBase $ws
    $hallado = ($mapa.Keys | Sort-Object) -join ", "
    "  BASE de {0}: columnas por encabezado -> {1}" -f $l.libro, $(if($hallado){$hallado}else{"NINGUNA"})
    if (-not $mapa.ContainsKey("formula")) {
      "  aviso: la hoja BASE de {0} no tiene columna FORMULA en la fila 3, se omite" -f $l.libro
      Close-ExcelPackage $p -NoSave
      continue
    }
    if (-not $mapa.ContainsKey("foto")) {
      "  aviso: {0} todavia no tiene columna FOTO en la hoja BASE" -f $l.libro
    }

    for ($r = 4; $r -le $ws.Dimension.End.Row; $r++) {
      $cod = Celda $ws $mapa "formula" $r
      $met = Celda $ws $mapa "met" $r
      $nom = Celda $ws $mapa "nombre" $r

      # si el VLOOKUP no dio codigo, se busca la receta por su nombre
      if ($cod -eq "" -or $cod -like "*N/A*") {
        $llave = NormNombre $nom
        if ($llave -ne "" -and $porNombre.ContainsKey($llave)) {
          $cod = $porNombre[$llave]
          if ($met -ne "") { $rescatadas++ }
        }
      }
      if ($cod -eq "" -or $cod -like "*N/A*") {
        if ($met -ne "" -and $nom -ne "") { [void]$huerfanas.Add($nom) }
        continue
      }
      $prev = $base[$cod]
      if ($null -ne $prev) {
        if ($prev.idx -lt $idxLibro) {
          # Ya lo puso un libro anterior. Ese manda en lo que tenga escrito, pero
          # los campos que dejo VACIOS si se rellenan desde este libro: asi no se
          # pierde lo que solo exista en el libro secundario.
          if ($prev.cat -eq "") { $prev.cat = Celda $ws $mapa "cat"  $r }
          if ($prev.met -eq "") { $prev.met = $met }
          if ($prev.ute -eq "") { $prev.ute = Celda $ws $mapa "ute"  $r }
          if ($prev.men -eq "") { $prev.men = Celda $ws $mapa "men"  $r }
          if ($prev.foto -eq "") { $prev.foto = Celda $ws $mapa "foto" $r }
          if ($prev.el -eq "")  { $prev.el  = Celda $ws $mapa "el"   $r }
          if ($prev.ap -eq "")  { $prev.ap  = Celda $ws $mapa "ap"   $r }
          if ($prev.pa -eq "")  { $prev.pa  = Celda $ws $mapa "pa"   $r }
          continue
        }
        # dentro del mismo libro, gana la fila con el metodo mas completo
        if ($prev.met.Length -ge $met.Length) { continue }
      }
      $base[$cod] = [pscustomobject]@{
        idx   = $idxLibro
        cat   = Celda $ws $mapa "cat"  $r
        met   = $met
        ute   = Celda $ws $mapa "ute"  $r
        men   = Celda $ws $mapa "men"  $r
        foto  = Celda $ws $mapa "foto" $r
        el    = Celda $ws $mapa "el"   $r
        ap    = Celda $ws $mapa "ap"   $r
        pa    = Celda $ws $mapa "pa"   $r
        libro = $l.libro
      }
    }
    Close-ExcelPackage $p -NoSave
  }

  if ($ing.Count -lt 700) { throw ("Solo se leyeron " + $ing.Count + " recetas; se esperaban 800 o mas. No se publica.") }

  # Da igual en cual libro viva la hoja BASE, pero si no aparece en ninguno
  # la ficha se quedaria sin metodo ni utensilios: ahi si hay que parar.
  $conMet = ($base.Values | Where-Object { $_.met -ne "" }).Count
  if ($base.Count -eq 0) { throw "Ninguno de los libros aporto filas de la hoja BASE. No se publica." }
  if ($conMet -eq 0)     { throw "Ninguna receta quedo con metodo de preparacion. Revisa la hoja BASE. No se publica." }

  # ---- equivalencias de los articulos que vienen en UND ----
  # El archivo "Equivalencias UND.csv" se puede editar a mano: lo que Sergio escriba ahi
  # MANDA sobre lo que se deduce del nombre. Si no existe, se crea prellenado.
  $eqFn = Join-Path $raiz "Equivalencias UND.csv"
  $manual = @{}
  if (Test-Path $eqFn) {
    foreach ($f in (Import-Csv -Path $eqFn -Encoding UTF8)) {
      $art = $f.Articulo
      if ([string]::IsNullOrWhiteSpace($art)) { continue }
      $v = Num $f.Equivale
      if ($v -le 0) { continue }
      $u = ("" + $f.Unidad).Trim().ToUpper()
      if ($u -ne "ML" -and $u -ne "GR") { $u = "GR" }
      $manual[$art.Trim().ToUpper()] = @{ v = $v; u = $u }
    }
  }

  # equivalencia efectiva por articulo UND: primero la tabla, si no el nombre
  function EqDe([string]$art) {
    $k = $art.Trim().ToUpper()
    if ($manual.ContainsKey($k)) { return $manual[$k] }
    return (Presentacion $art)
  }

  # gramos (o ml, que se manejan 1 a 1) de una linea, para UNA receta
  function GramosLinea([string]$art, [string]$um, [double]$cant) {
    switch (("" + $um).Trim().ToUpper()) {
      "KL"  { return $cant * 1000 }
      "LT"  { return $cant * 1000 }
      "GR"  { return $cant }
      "GRM" { return $cant }
      "ML"  { return $cant }
    }
    $e = EqDe $art
    if ($null -ne $e) { return $cant * $e.v }
    return $null
  }

  # si la tabla no existe todavia, se deja lista con TODOS los articulos UND,
  # ya prellenada con lo que se pudo deducir del nombre
  if (-not (Test-Path $eqFn)) {
    $undArt = @{}
    foreach ($cod in $ing.Keys) {
      foreach ($x in $ing[$cod]) {
        if (("" + $x.u).Trim().ToUpper() -like "UND*") {
          $k = $x.a.Trim()
          if (-not $undArt.ContainsKey($k)) { $undArt[$k] = 0 }
          $undArt[$k]++
        }
      }
    }
    $filas = foreach ($a in ($undArt.Keys | Sort-Object)) {
      $p = Presentacion $a
      [pscustomobject]@{
        Articulo  = $a
        Equivale  = $(if ($p) { $p.v } else { "" })
        Unidad    = $(if ($p) { $p.u } else { "" })
        Origen    = $(if ($p) { "deducido del nombre" } else { "FALTA" })
        Apariciones = $undArt[$a]
      }
    }
    $filas | Sort-Object Origen, @{e={-$_.Apariciones}} | Export-Csv -Path $eqFn -NoTypeInformation -Encoding UTF8
    "tabla de equivalencias creada: {0}" -f $eqFn
  }

  # ---- fotos de los platos ----
  # En la hoja BASE, columna I "FOTO", va el NOMBRE del archivo (ej: "arroz con pollo.jpg").
  # Las imagenes viven en la subcarpeta "Fotos". Se reducen y se empotran en la pagina,
  # porque el visor bloquea las imagenes que vengan de afuera.
  Add-Type -AssemblyName System.Drawing
  $dirFotos = Join-Path $raiz "Fotos"
  New-Item -ItemType Directory -Force -Path $dirFotos | Out-Null
  $cacheFoto = @{}
  $pesoFotos = 0

  function FotoDataUri([string]$nombre) {
    if ([string]::IsNullOrWhiteSpace($nombre)) { return "" }
    $nombre = $nombre.Trim()
    if ($cacheFoto.ContainsKey($nombre)) { return $cacheFoto[$nombre] }

    $ruta = $null
    if (Test-Path -LiteralPath $nombre) {
      $ruta = $nombre                              # ruta completa
    } else {
      $cand = Join-Path $dirFotos $nombre
      if (Test-Path -LiteralPath $cand) {
        $ruta = $cand
      } else {
        foreach ($ext in @(".jpg",".jpeg",".png",".webp")) {   # sin extension
          $c2 = Join-Path $dirFotos ($nombre + $ext)
          if (Test-Path -LiteralPath $c2) { $ruta = $c2; break }
        }
      }
    }
    if ($null -eq $ruta) {
      # Hay un nombre escrito en la columna FOTO pero no existe el archivo:
      # o es un error de digitacion, o quedo texto que no es una foto.
      # OJO: aqui va Write-Host y no una cadena suelta. Cualquier salida sin capturar
      # dentro de una funcion se suma al valor de retorno y termina metida en el JSON.
      Write-Host ("  aviso: en la columna FOTO dice '" + $nombre + "' pero no hay ese archivo en la carpeta Fotos")
      $cacheFoto[$nombre] = ""
      return ""
    }

    try {
      $img = [System.Drawing.Image]::FromFile($ruta)
      $anchoMax = 560
      $w = $img.Width; $h = $img.Height
      if ($w -gt $anchoMax) { $h = [int][math]::Round($h * $anchoMax / $w); $w = $anchoMax }
      $bmp = New-Object System.Drawing.Bitmap($w, $h)
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $g.DrawImage($img, 0, 0, $w, $h)
      $cod = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" }
      $par = New-Object System.Drawing.Imaging.EncoderParameters(1)
      $par.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 72L)
      $ms = New-Object System.IO.MemoryStream
      $bmp.Save($ms, $cod, $par)
      $uri = "data:image/jpeg;base64," + [Convert]::ToBase64String($ms.ToArray())
      $script:pesoFotos += $uri.Length
      $g.Dispose(); $bmp.Dispose(); $ms.Dispose(); $img.Dispose()
      $cacheFoto[$nombre] = $uri
      return $uri
    } catch {
      Write-Host ("  aviso: no se pudo leer la foto '" + $nombre + "' (" + $_.Exception.Message + ")")
      $cacheFoto[$nombre] = ""
      return ""
    }
  }

  # ---- armar el JSON a mano (ConvertTo-Json de PS 5.1 destroza los arreglos anidados) ----
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"generado":"'); [void]$sb.Append((Get-Date -Format "yyyy-MM-dd")); [void]$sb.Append('","recetas":[')
  $primero = $true
  foreach ($cod in ($ing.Keys | Sort-Object)) {
    $m = $meta[$cod]; $b = $base[$cod]; $nom = $hdr[$cod]
    if ($null -ne $m -and $m.nom -ne "") { $nom = $m.nom }
    if (-not $primero) { [void]$sb.Append(',') }
    $primero = $false
    [void]$sb.Append('{"c":"');   [void]$sb.Append((Esc $cod))
    [void]$sb.Append('","n":"');  [void]$sb.Append((Esc $nom))
    [void]$sb.Append('","um":"'); [void]$sb.Append((Esc $(if($m){$m.um}else{""})))
    [void]$sb.Append('","rd":');  [void]$sb.Append($(if($m){[math]::Round($m.rend,3)}else{0}))
    [void]$sb.Append(',"cs":');   [void]$sb.Append($(if($m){[math]::Round($m.cost,2)}else{0}))
    [void]$sb.Append(',"cat":"'); [void]$sb.Append((Esc $(if($b){$b.cat}else{""})))
    [void]$sb.Append('","lb":"'); [void]$sb.Append((Esc $(if($b){$b.libro}else{""})))
    [void]$sb.Append('","mp":"'); [void]$sb.Append((Esc $(if($b){$b.met}else{""})))
    [void]$sb.Append('","ut":"'); [void]$sb.Append((Esc $(if($b){$b.ute}else{""})))
    [void]$sb.Append('","mn":"'); [void]$sb.Append((Esc $(if($b){$b.men}else{""})))
    [void]$sb.Append('","fo":"'); [void]$sb.Append($(if($b){ FotoDataUri $b.foto }else{""}))
    [void]$sb.Append('","el":"'); [void]$sb.Append((Esc $(if($b){$b.el}else{""})))
    [void]$sb.Append('","ap":"'); [void]$sb.Append((Esc $(if($b){$b.ap}else{""})))
    [void]$sb.Append('","pa":"'); [void]$sb.Append((Esc $(if($b){$b.pa}else{""})))
    [void]$sb.Append('","i":[')
    $p2 = $true
    foreach ($x in $ing[$cod]) {
      if (-not $p2) { [void]$sb.Append(',') }
      $p2 = $false
      $g = GramosLinea $x.a $x.u $x.q
      [void]$sb.Append('["'); [void]$sb.Append((Esc $x.a))
      [void]$sb.Append('","'); [void]$sb.Append((Esc $x.u))
      [void]$sb.Append('",');  [void]$sb.Append([math]::Round($x.q,4))
      [void]$sb.Append(',');   [void]$sb.Append([math]::Round($x.c,2))
      [void]$sb.Append(',');   [void]$sb.Append($(if ($null -eq $g) { "null" } else { [math]::Round($g,3) }))
      [void]$sb.Append(']')
    }
    [void]$sb.Append(']}')
  }
  [void]$sb.Append(']}')
  $json = $sb.ToString()

  # ---- huella del contenido, ignorando la fecha de generacion ----
  $paraHash = $json -replace '"generado":"[^"]*"', ''
  $md5 = [System.Security.Cryptography.MD5]::Create()
  $hash = [BitConverter]::ToString($md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($paraHash))).Replace("-","")
  $hashPrevio = ""
  if (Test-Path $hashFn) { $hashPrevio = (Get-Content $hashFn -Raw).Trim() }

  $conMetodo = ($ing.Keys | Where-Object { $null -ne $base[$_] -and $base[$_].met -ne "" }).Count
  $lineas = 0; foreach ($k in $ing.Keys) { $lineas += $ing[$k].Count }

  if ($rescatadas -gt 0) {
    "  {0} recetas con metodo se reconectaron POR NOMBRE (su columna FORMULA daba #N/A)" -f $rescatadas
  }
  if ($huerfanas.Count -gt 0) {
    "  {0} recetas tienen metodo escrito pero no se pudo emparejar con ninguna formula:" -f $huerfanas.Count
    foreach ($h in ($huerfanas | Sort-Object -Unique)) { "     - " + $h }
  }

  "  costos: {0} de costo estandar, {1} de ultima compra" -f $costoEstandar, $costoDeUltimaCompra
  $sinCosto = 0
  foreach ($k in $ing.Keys) { foreach ($x in $ing[$k]) { if ($x.c -le 0) { $sinCosto++ } } }
  if ($lineas -gt 0 -and $sinCosto -gt ($lineas * 0.5)) {
    throw ("La mitad o mas de las lineas quedaron sin costo (" + $sinCosto + " de " + $lineas + "). Revisa las columnas de costo del export. No se publica.")
  }

  $conFoto = ($cacheFoto.Values | Where-Object { $_ -ne "" }).Count
  "recetas={0} lineas={1} conMetodo={2} conFoto={3} pesoFotos={4}KB" -f `
    $ing.Count, $lineas, $conMetodo, $conFoto, [math]::Round($pesoFotos/1KB,0)

  if ($hash -eq $hashPrevio -and (Test-Path $salida)) {
    "ESTADO=SIN-CAMBIOS"
    exit 0
  }

  # ---- freno de mano: caida brusca contra la corrida anterior ----
  # Si Sergio esta editando el libro y la corrida lo agarra a medio camino, las cifras
  # se desploman. Antes de eso se prefiere NO publicar y avisar: el libro que ven en la
  # cocina no puede quedar mutilado por un guardado a destiempo.
  $conteoFn = Join-Path $raiz "ultimo conteo.txt"
  if (Test-Path $conteoFn) {
    $previo = @{}
    foreach ($linea in (Get-Content $conteoFn)) {
      $par = $linea -split '=', 2
      if ($par.Count -eq 2) { $previo[$par[0].Trim()] = [int]$par[1].Trim() }
    }
    $alertas = @()
    if ($previo.ContainsKey("lineas") -and $previo["lineas"] -gt 0 -and $lineas -lt ($previo["lineas"] * 0.85)) {
      $alertas += ("lineas de ingrediente: " + $previo["lineas"] + " -> " + $lineas)
    }
    if ($previo.ContainsKey("conMetodo") -and $previo["conMetodo"] -gt 0 -and $conMetodo -lt ($previo["conMetodo"] * 0.7)) {
      $alertas += ("recetas con metodo: " + $previo["conMetodo"] + " -> " + $conMetodo)
    }
    if ($previo.ContainsKey("recetas") -and $previo["recetas"] -gt 0 -and $ing.Count -lt ($previo["recetas"] * 0.9)) {
      $alertas += ("recetas: " + $previo["recetas"] + " -> " + $ing.Count)
    }
    if ($alertas.Count -gt 0) {
      "CAIDA BRUSCA respecto a la corrida anterior:"
      foreach ($a in $alertas) { "  - " + $a }
      "Lo mas probable: el libro se guardo a medias o se esta editando ahora."
      "No se toco el archivo publicado. Revisa el Excel y vuelve a correr."
      "ESTADO=REVISAR"
      exit 2
    }
  }

  # ---- inyectar en la plantilla ----
  $tpl = [System.IO.File]::ReadAllText((Join-Path $raiz "plantilla.html"), [System.Text.Encoding]::UTF8)
  if ($tpl -notmatch '__DATOS__') { throw "La plantilla no tiene el marcador __DATOS__." }
  $html = $tpl.Replace('__DATOS__', ($json -replace '</','<\/'))

  $logoFn = Join-Path $raiz "logo.txt"
  if (Test-Path $logoFn) {
    $html = $html.Replace('__LOGO__', ([System.IO.File]::ReadAllText($logoFn, [System.Text.Encoding]::UTF8)).Trim())
  } else {
    $html = $html.Replace('__LOGO__', '')
  }
  [System.IO.File]::WriteAllText($salida, $html, (New-Object System.Text.UTF8Encoding($false)))
  [System.IO.File]::WriteAllText($hashFn, $hash, (New-Object System.Text.UTF8Encoding($false)))

  # Copia al repositorio, si existe. Sin esto la version de Railway se queda
  # atras sin que se note: el generador vive en OneDrive y el repo en Proyectos.
  $repo = "C:\Users\SergioValencia\Proyectos\fichas-tecnicas-caujaral"
  if (Test-Path $repo) {
    [System.IO.File]::WriteAllText((Join-Path $repo "Fichas Tecnicas Caujaral.html"), $html, (New-Object System.Text.UTF8Encoding($false)))
    Copy-Item (Join-Path $raiz "plantilla.html") (Join-Path $repo "plantilla.html") -Force
    Copy-Item (Join-Path $raiz "Generar Fichas Tecnicas.ps1") (Join-Path $repo "Generar Fichas Tecnicas.ps1") -Force
    if (Test-Path $eqFn) { Copy-Item $eqFn (Join-Path $repo "Equivalencias UND.csv") -Force }

    # Publicar en la web. Decision de Sergio del 06/08/2026: que cada cambio del
    # libro salga solo, sin tener que empujarlo a mano. Railway redespliega al
    # detectar el push. Si algo falla aqui NO se tumba la corrida: la pagina ya
    # quedo generada y el artefacto se publica igual.
    try {
      $tieneRemoto = $false
      $remotos = & git -C $repo remote 2>$null
      if ($LASTEXITCODE -eq 0 -and $remotos) { $tieneRemoto = $true }

      if (-not $tieneRemoto) {
        "  repositorio sin remoto todavia: no se publica en la web"
      } else {
        & git -C $repo add -A 2>&1 | Out-Null
        $pendiente = & git -C $repo status --porcelain 2>$null
        if (-not $pendiente) {
          "  repositorio sin cambios: no hay nada que publicar"
        } else {
          $mensaje = "Actualizacion del libro de recetas " + (Get-Date -Format "yyyy-MM-dd HH:mm") +
                     " - " + $ing.Count + " recetas, " + $conMetodo + " con metodo"
          & git -C $repo -c user.name="Sergio Valencia" -c user.email="svalencia536@gmail.com" commit -m $mensaje 2>&1 | Out-Null
          $salidaPush = & git -C $repo push 2>&1
          if ($LASTEXITCODE -eq 0) {
            "  publicado en la web: commit y push hechos, Railway redespliega solo"
          } else {
            "  aviso: el commit quedo hecho pero el push fallo. Queda pendiente de subir."
            foreach ($l in $salidaPush) { "     " + $l }
          }
        }
      }
    } catch {
      "  aviso: no se pudo publicar en la web ({0}). La pagina si quedo generada." -f $_.Exception.Message
    }
  }
  # se guardan las cifras de esta corrida para comparar en la siguiente
  $conteoFn = Join-Path $raiz "ultimo conteo.txt"
  $lineasConteo = @(
    ("recetas=" + $ing.Count),
    ("lineas=" + $lineas),
    ("conMetodo=" + $conMetodo)
  )
  [System.IO.File]::WriteAllLines($conteoFn, $lineasConteo, (New-Object System.Text.UTF8Encoding($false)))

  "archivo={0}" -f $salida
  "ESTADO=ACTUALIZADO"
  exit 0
}
catch {
  "ERROR: {0}" -f $_.Exception.Message
  "ESTADO=ERROR"
  exit 1
}

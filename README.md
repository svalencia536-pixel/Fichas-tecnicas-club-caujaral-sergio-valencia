# Fichas técnicas de cocina y bar — Club Lagos de Caujaral

Consulta de recetas para los operarios de cocina y bar. Se busca el plato o el trago
por su nombre y sale la ficha completa: ingredientes con cantidad y conversión a gramos
o mililitros, método de preparación, utensilios y el menaje en que se sirve.

## Cómo funciona

```
LIBRO DE RECETAS AUTOMATIZADOS V.2 *.xlsx   (en OneDrive, lo alimenta Sergio)
                 |
    Generar Fichas Tecnicas.ps1             (lee, valida y arma los datos)
                 |
          plantilla.html                    (el diseño, con el marcador __DATOS__)
                 |
     Fichas Tecnicas Caujaral.html          (página autónoma, lista para publicar)
```

El generador toma **tres hojas** del libro:

| Hoja | Qué aporta |
|---|---|
| `Query 1` | Ingredientes de cada receta, con cantidad, unidad y costo |
| `Query2` | Rendimiento, unidad de medida y costo de la receta |
| `BASE` | Categoría, método, utensilios, menaje, foto y firmas |

Las columnas de `BASE` se ubican **por el nombre del encabezado** (fila 3), no por su
posición, así que se pueden mover de sitio sin romper nada.

## Correrlo

```powershell
& ".\Generar Fichas Tecnicas.ps1"
```

Imprime al final una línea de estado:

| Estado | Qué significa |
|---|---|
| `ESTADO=ACTUALIZADO` | Se regeneró la página. Hay que republicar. |
| `ESTADO=SIN-CAMBIOS` | El Excel no cambió desde la última corrida. |
| `ESTADO=REVISAR` | Las cifras cayeron de golpe. **No publicar**: casi siempre el libro se guardó a medias. |
| `ESTADO=ERROR` | Falló algo. El mensaje dice qué. |

## Los frenos que trae

No publica si algo huele mal, porque esta página la usa la cocina:

- Menos de 700 recetas leídas.
- Ninguna receta con método de preparación.
- Más de la mitad de las líneas sin costo.
- Caída brusca contra la corrida anterior: líneas bajo el 85 %, recetas bajo el 90 %
  o métodos bajo el 70 %. Las cifras de referencia quedan en `ultimo conteo.txt`.

## Cosas que ya nos mordieron

- **`Query 1` y `Query2` se leen de un solo libro.** Mezclar dos exports de fechas
  distintas duplica ingredientes cuando el sistema renombra un artículo: aparecían
  juntos `RC SALSA TARTARA` y `SB SALSA TARTARA` en la misma receta.
- **La columna `FORMULA` de `BASE` es un `VLOOKUP` contra `Query2`.** Cada re-export
  la rompe y devuelve `#N/A`. Por eso el generador tiene un rescate: si no hay código,
  empareja la receta por su nombre normalizado.
- **El costo estándar puede venir en cero.** Pasa cuando el reporte de Zeus se corre sin
  el parámetro *Período de costo estándar*. El generador cae entonces a
  `CostoUltimaCompra` y lo dice en la ficha.
- **En PowerShell, cualquier salida suelta dentro de una función se suma al valor de
  retorno.** Un aviso mal escrito metió `System.Object[]` dentro de los datos. Para
  avisos dentro de funciones va `Write-Host`.

## Publicación en Railway

`server.js` entrega la página ya generada. No consulta nada en vivo: cada despliegue
sirve la versión que venga en el repositorio.

**Para actualizar lo que ve la cocina:** correr el generador, `git commit`, `git push`.
Railway redespliega solo y el link no cambia.

### Clave de acceso

| Variable | Qué hace |
|---|---|
| `FICHAS_PASSWORD` | Si existe, la página pide usuario y clave. Si no, queda abierta. |
| `FICHAS_USER` | El usuario. Por defecto `caujaral`. |

**Sin `FICHAS_PASSWORD` cualquiera con el link ve las recetas y los costos**, porque van
dentro del propio archivo. La página se sirve con `noindex` y un `robots.txt` que la
excluye, pero eso solo frena a los buscadores, no a quien tenga la dirección.

`/salud` devuelve en JSON cuántas recetas hay, la fecha de corte de los datos y si está
protegida con clave.

## Archivos

| Archivo | Para qué |
|---|---|
| `Generar Fichas Tecnicas.ps1` | El proceso completo |
| `plantilla.html` | Diseño y comportamiento de la página |
| `Fichas Tecnicas Caujaral.html` | La página generada |
| `Equivalencias UND.csv` | **Editable a mano.** Cuánto pesa cada artículo que se pide por unidad |
| `logo.txt` | El logo del club como data-uri |
| `Fotos/` | Imágenes de los platos, referenciadas por la columna `FOTO` |

## Equivalencias por unidad

Los licores y los artículos con el gramaje en el nombre se resuelven solos:
`0,045` de `TEQUILA OLMECA REPOSADO 1000 ML` son **45 ml**. El número entre paréntesis
del nombre es el **peso de la botella vacía** y no entra en esa cuenta.

Lo que se cuenta por unidad de verdad —huevos, deditos de queso, porciones— se completa
a mano en `Equivalencias UND.csv`. Lo que se escriba ahí **manda** sobre lo deducido
del nombre.

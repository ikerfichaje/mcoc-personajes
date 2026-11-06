# MCOC Buscador

App web **offline** y pensada para el móvil para consultar los campeones de **MARVEL Batalla de Superhéroes** (*Marvel Contest of Champions*):

- **Campeones**: busca por nombre o por **efecto** (Veneno, Anular, Bloqueo de poder…) y elige dónde debe aparecer: inmune, lo provoca al rival, lo obtiene, pasivo, lo resiste o en cualquier parte del kit. Se pueden combinar varios efectos y filtrar por clase y rareza (7★/6★).
- **Ficha** de cada campeón: tiers (youtubers y MCOC Guide Blog), reliquia recomendada y cuál conviene, inmunidades, efectos, sinergias y mapa de habilidades. Toca cualquier efecto para ver quién más lo tiene.
- **Equipo**: junta hasta 5 campeones y mira sus inmunidades, los huecos («nadie es inmune a…»), los efectos que provocan, las sinergias entre ellos y con otros campeones, y sus reliquias.
- **Reliquias**: Battlecasts y Statcasts con los campeones que las usan, y **Mejores combos** campeón + reliquia para atacar o defender.

Sin dependencias ni compilación: HTML + CSS + JavaScript.

## Estructura

```
index.html, css/, js/app.js        la app
data/data.js                       datos generados desde el Excel
data/images.js                     mapa nombre → retrato
img/                               retratos (128 px)
sw.js, manifest.webmanifest, icons/   modo offline / instalable
scripts/
  excel-to-json.ps1    Excel → data/data.js (+ data.json)
  download-images.ps1  descarga los retratos de MCOCHUB
  serve.ps1            servidor local para probar
MCOC_campeones_nov2026.xlsx        fuente de los datos
```

## Uso en el ordenador

- Rápido: abre `index.html` con doble clic (funciona, pero sin modo instalable).
- Con servidor local: `powershell -ExecutionPolicy Bypass -File scripts\serve.ps1` y abre <http://localhost:8765/>.

## Instalarla en el móvil

El móvil necesita abrirla desde una dirección **https** una vez; después queda guardada y funciona sin conexión.
Se publica con **GitHub Pages desde la raíz de `main`** (*Settings → Pages → Deploy from a branch → main / (root)*). Luego, en el móvil:

- **Android (Chrome)**: ⋮ → «Instalar aplicación».
- **iPhone (Safari)**: Compartir → «Añadir a pantalla de inicio».

## Actualizar los datos

1. Sustituye el Excel (mismas hojas y columnas).
2. `powershell -ExecutionPolicy Bypass -File scripts\excel-to-json.ps1`
3. Si hay campeones nuevos: `powershell -ExecutionPolicy Bypass -File scripts\download-images.ps1`
4. Sube `VERSION` en `sw.js` (p. ej. `mcoc-v2`) para que los móviles descarguen la nueva versión.

## Fuentes y aviso legal

**Proyecto de fans, no afiliado a Kabam ni a Marvel.** MARVEL Batalla de Superhéroes, los nombres de los campeones y sus imágenes son propiedad de Kabam Games y Marvel.

- Kits, inmunidades y retratos: [MCOCHUB](https://mcochub.insaneskull.com) (la app enlaza a la ficha completa de cada campeón en lugar de copiar los textos).
- Tiers: [mcoc.app](https://mcoc.app) (Vega, Lagacy, MetalSonicDude, Seatin) y [MCOC Guide Blog](https://www.mcocguideblog.com) (CaptainRaghu). Son opiniones de sus autores.
- Reliquias recomendadas: [MCOC Scout](https://mcocscout.com/relics).
- La clasificación de efectos se ha hecho automáticamente a partir del texto en inglés de los kits y puede tener errores.

Los retratos (`app/img/`) son propiedad de Kabam / Marvel y se obtuvieron de MCOCHUB; se incluyen sin ánimo de lucro. Si los titulares lo piden, se retirarán (la app muestra las iniciales del campeón cuando falta una imagen).

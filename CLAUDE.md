# CLAUDE.md — App de campeones de MARVEL Batalla de Superhéroes (MCOC)

> Contexto heredado de una conversación previa en claude.ai (6 de noviembre de 2026).
> Léelo entero antes de empezar. El archivo `MCOC_campeones_nov2026.xlsx` de esta carpeta contiene los datos ya procesados.

## 1. Objetivo

Construir una app con los datos de los campeones de **MARVEL Batalla de Superhéroes** (nombre en español de *Marvel Contest of Champions*, MCOC, de Kabam) y publicarla en GitHub.
El usuario habla **español**: la interfaz, los textos y las respuestas deben ir en español. Los nombres de campeones se mantienen en inglés (así aparecen en el juego y en la comunidad).

Lo que el usuario pidió a lo largo de la conversación, en orden:
1. Todos los campeones con versión **7★**; si un campeón no tiene 7★, usar su versión **6★**.
2. Comparar la valoración (tier) de varias fuentes.
3. Añadir **habilidades, inmunidades y efectos pasivos** de cada campeón.
4. Pasarlo todo a Excel.

## 2. Datos incluidos: `MCOC_campeones_nov2026.xlsx`

Hojas (en este orden):

| Hoja | Contenido |
|---|---|
| **Leyenda** | Significado de columnas, escalas de tier y fuentes. |
| **Campeones** | 338 filas, una por campeón (actualizado el 12/11/2026, ver §9). |
| **Mapa de habilidades** | ~2.360 filas: una por cada parte del kit de cada campeón (firma, especiales, ataques fuertes…). |
| **Matriz de inmunidades** | Campeones × 24 inmunidades. Valores: `Sí` (permanente), `Cond.` (condicional), `¿?` (no confirmada), vacío (no / sin datos). Al final, 2 filas con `COUNTIF`. |
| **Resumen** | Recuentos por clase con fórmulas `COUNTIF(S)` sobre la hoja Campeones. |
| **Reliquias por campeón** | 337 filas: Battlecast recomendado, Statcast recomendado (retrato o nombre exacto) y «¿Cuál te conviene más?». |
| **Battlecasts** | 25 Battlecasts: clase, nº y lista de campeones que lo recomiendan (ordenados por tier), efectos conocidos. |
| **Statcasts** | 24 familias de Statcast (prefijo × clase) con sus sufijos + tabla de retratos y qué campeones los recomiendan. |
| **Guía de reliquias** | Explicación de Battlecast vs Statcast, cómo elegir, fuentes y limitaciones. |

Columnas de **Campeones** (A→R):
`Campeón | Clase | Rareza máx. | Tier youtubers | Guide Blog ataque | Guide Blog defensa | Inmunidades permanentes | Inmunidades condicionales | Inmunidades no confirmadas | Resistencias pasivas | Efectos que provoca al rival | Mejoras/efectos que obtiene | Otras mecánicas | Efectos pasivos (siempre activos) | Por sinergia | Habilidades clave (fichas parciales) | Ficha | Enlace texto completo`

- Listas dentro de una celda separadas por `, `; grupos separados por `; `. Vacío se representa con `—`.
- Clases en español: Cósmico, Mutante, Ciencia, Habilidad (Skill), Tecnología, Místico.
- `Rareza máx.`: `7★` (280 campeones) o `6★` (58).
- `Ficha`: `Completa` (kit completo de MCOCHUB, ~270) o `Parcial` (solo datos de otras fuentes).
- Inmunidades condicionales: `Efecto (dónde)`, p. ej. `Aturdimiento (Phasing)`, `X (firma)` = requiere despertar la firma.
- Columnas de **Mapa de habilidades**: `Campeón | Clase | Rareza máx. | Parte del kit | ¿Pasiva? | Efectos que aparecen`.

Para la app, lo más práctico es convertir el Excel a JSON (p. ej. con un script `scripts/excel-to-json` usando `xlsx`/SheetJS en Node o `pandas` en Python) y que la app lea el JSON.

## 3. Fuentes y lo que se aprendió de cada una

| Fuente | Uso | Notas |
|---|---|---|
| **mcocscout.com/7-star-champions** | Lista de campeones con 7★ + 4 habilidades clave de cada uno | HTML estático, se lee bien. Fiable para saber quién tiene 7★. |
| **mcoc.app/champions** | Roster completo (335), tier agregado de 4 youtubers (Vega, Lagacy, MetalSonicDude, Seatin) | La página de cada campeón dice “X is not available as a 7-star champion”. Las páginas de prestigio/inmunidades son JS (no se pueden leer sin navegador). |
| **mcoc.app/tier-list/{clase}** | Clase de cada campeón y tier por clase | Sirvió para confirmar clases dudosas (p. ej. Darkhawk, Psycho-Man, Doctor Octopus y Green Goblin son Tecnología; Iron Fist, Unstoppable Colossus y Loki son Místico). |
| **mcocguideblog.com/p/mcoc-offense-and-defense-tier-list.html** | Tier de ataque (S+…E) y defensa (S…D) de un autor (CaptainRaghu), 14/10/2026 | No valora Deathless ni Stellar-Forged. Nombres con alias (ver §6). |
| **MCOCHUB** (mcochub.insaneskull.com) | Texto completo de kits, inmunidades, sinergias | Los datos se obtuvieron del repo de GitHub **Shadey1/Mcoc-help**, archivo `data/champions/abilities.json` (importado de MCOCHUB el 2026-10-26, 270 campeones, solo los que tienen 7★). |
| **Shadey1/Mcoc-help** (otros archivos) | `data/immunities/_locks.json`, `data/champions/immunities-auntm.json` (incluye `passives`), `immunities-kabam.json` | Inmunidades cruzadas (auntm.ai congelado a mediados de 2024; Kabam Champion Spotlights). |

Otros hallazgos:
- Muchas “tier list 2026” que salen en Google son contenido antiguo reciclado (p. ej. dicen que Hercules es el mejor 7★, pero **Hercules no tiene 7★**). No usarlas.
- Varios S+ **no tienen 7★**: Aegon, Hercules, Cosmic Ghost Rider, Kate Bishop, Quake.
- Los cuatro “Victorious” (Human Torch, Invisible Woman, Mister Fantastic, Thing) son 7★ recientes y aún no tienen tier.
- El acceso a GitHub por API da error de límite de peticiones sin token; descargar repos con `https://codeload.github.com/OWNER/REPO/zip/refs/heads/main` funciona.

### Estructura de `abilities.json` (MCOCHUB vía Mcoc-help)
```json
{ "version": 1, "source": "MCOCHUB (https://mcochub.insaneskull.com)", "lastImported": "2026-10-26",
  "champions": { "<slug>": {
      "source": { "slug": "...", "url": "https://mcochub.insaneskull.com/champions/<slug>" },
      "pills": {
        "abilities":  [ { "name": "Poison" }, { "name": "Armor Break", "synergy": { "partners": ["manthing"], "note": "..." } } ],
        "immunities": [ { "name": "Poison Immunity" } ],
        "tags": [ "Villain", "Size: XL", "..." ] },
      "kit": {
        "signature": { "title": "TOXIC HUSK", "trigger": "", "lines": ["..."] },
        "cards": [ { "title": "PASSIVE" | "ALWAYS ACTIVE" | "SPECIAL ATTACK 1" | "<nombre propio>", "trigger": "", "lines": ["..."] } ] } } } }
```
Ojo: algunos `cards` son duplicados cuyo título es una frase del texto (artefacto del scraper): se descartan si todas sus líneas ya aparecieron antes.

## 4. Cómo se procesaron los datos (reglas heurísticas)

1. **Slug** de cada nombre: minúsculas, quitar `'’.()`, resto de caracteres no alfanuméricos → `-`. Excepciones: Captain Marvel (Movie)→`captain-marvel`, Kang→`kang-the-conqueror`, Spider-Woman (Jessica Drew)→`spider-woman`, Falcon (Joaquín Torres)→`falcon-joaquin-torres`, M.O.D.O.K.→`modok`.
2. **Inmunidades** (por frase del kit):
   - Se buscan frases con `immun` que casen `Immune to X`, `Immunity to X`, `X Immunity` o `X Immune`, y se extraen términos de una lista fija (Poison, Bleed, Incinerate, Shock, Coldsnap, Frostbite, Stun, Stagger, Nullify, Armor Break, Armor Shattered, Power Drain/Burn/Lock/Steal/Sting, Heal Block, Degeneration, Rupture, Reverse(d) Controls, Fate Seal, Neuroshock, Delirium, Ability Accuracy Reduction, etc.).
   - Se descartan frases tipo “ignore immunities”, “if the Opponent is immune”, “Buff Immunity Debuff”…
   - **Condicional** si: está en la firma, o el título de la tarjeta es de acción (SPECIAL/HEAVY/ATTACK/CRITICAL/BLOCK/WHEN/ON/CHARG/PARR/DASH/KNOCK), o antes de “immun” hay `while/during/if/when/after/each/per/until/against/…`, o justo después aparece `while/during/if/when/as long as/until/for N`.
   - Las inmunidades que da MCOCHUB (`pills.immunities`) u otras fuentes y **no** se encuentran en el texto van a “Inmunidades no confirmadas”.
3. **Resistencias**: `(T|t)akes? N% less damage from X` → “N% menos daño de X”; más bandas `resist` de otras fuentes (se descartan valores >100%).
4. **Provoca vs obtiene** (por cada `pills.abilities` sin sinergia):
   - Provoca si casa `inflict… (≤4 palabras) X`, `X Debuff`, `X the Opponent` u `Opponent is/becomes X`.
   - Obtiene si casa `gain/grant… (≤4 palabras) X` o `X Buff`.
   - Puede salir en ambas; si no casa nada → “Otras mecánicas”.
5. **Pasivos**: tarjetas con título PASSIVE, ALWAYS ACTIVE, FIGHT START, PRE-FIGHT ABILITY o PERSISTENT CHARGE.
6. **Mapa de habilidades**: para cada tarjeta, términos del glosario que aparecen en su texto.

### Errores conocidos de estas heurísticas
- Algunas falsas atribuciones (p. ej. Hulk aparece con Furia en “provoca”).
- Algunas condiciones salen con el nombre de la tarjeta en inglés (p. ej. `Aturdimiento (While Untouchable)`).
- Las fichas “Parcial” (58 campeones solo 6★ + Beast, Doctor Doom, Ghost, Goldpool, Platinumpool y los 4 Victorious) no tienen kit completo.
- Mejora posible: obtener los kits que faltan desde las páginas de MCOCHUB o la wiki de Fandom (marvel-contestofchampions.fandom.com) y reprocesar.

## 4b. Reliquias (añadido después)

Objetivo del usuario: saber **la mejor reliquia para cada campeón**, **qué beneficios da cada reliquia** y **cuál le conviene más y por qué**.

- **Emparejamientos campeón → reliquia:** de `mcocscout.com/relics` (HTML estático, cubre los 337). Formato: `Battlecast: X | Statcast: Y`. Algunos solo tienen uno de los dos; los campeones nuevos traen el nombre exacto del Statcast (p. ej. «Insulated Mystic Relic of Toughness»).
  Alias de Battlecast en esa web: `Ant Man`→Ant-Man, `Juggy`→Juggernaut, `Miss Marvel`→Ms. Marvel, `Spidey 2099`→Spider-Man 2099, `Venom The Duck`→The Cosmic Egg.
- **Statcasts:** se nombran «[Prefijo] [Clase] Relic of [Sufijo]». MCOC Scout recomienda casi siempre por el **retrato** del relic (un campeón por familia), no por nombre. Solo está confirmado que el retrato **Kushala = Insulated Mystic** (dato de un hilo del foro oficial). Retratos por clase: Ciencia (Thing, Human Torch, Captain America IW, Morbius), Habilidad (Bullseye, Karnak, Moondragon, Punisher), Mutante (Magneto, Apocalypse, Domino, Colossus), Tecnología (Prowler, Red Skull, Iron Man, Stark Spidey), Cósmico (Superior Iron Man, Heimdall, Hulkling, Terrax), Místico (Kushala, Doctor Doom, Unstoppable Colossus, The Destroyer). Prefijos por clase (de MCOCHUB): Cósmico = Armoured/Piercing/Shielded/Sturdy; Mutante = Armoured/Conductive/Expert/Sturdy; Místico = Armoured/Conductive/Insulated/Shielded; Ciencia = Impactful/Insulated/Shielded/Tough; Habilidad = Cruel/Insulated/Precise/Tough; Tecnología = Cruel/Impactful/Insulated/Sturdy. **Pendiente:** confirmar qué retrato corresponde a cada prefijo y qué estadísticas da cada prefijo/sufijo.
- **Battlecasts (25):** Thor, Black Panther, Hulk, Storm, Scarlet Witch, Green Goblin, Ghost Rider, Sentinel, Captain America, Mister Sinister, Gambit, Black Widow, Wolverine, Venom, Ms. Marvel, Iron Fist, Hulkbuster, Spider-Man 2099, Winter Soldier, Vision, Ant-Man, Valkyrie, Juggernaut, Gamora, The Cosmic Egg. La clase se dedujo de los campeones que lo usan. **Efectos detallados pendientes** salvo Thor (Fortaleza/«Valhallan Fortitude» y Aptitud, de la wiki de Fandom).
- **Funcionamiento general** (MCOC Guide Blog + Fandom): un relic por campeón; Battlecast = se activa con una condición en combate e invoca un ataque «Striker», runas personalizables, habilidad innata con duplicado, Embers suben habilidades y Alloys suben rango → orientado a ataque. Statcast = estadísticas pasivas todo el combate, **también en defensa**.
- **«¿Cuál te conviene más?»** (heurística propia): con los tiers de MCOC Guide Blog, ataque S+/S → Battlecast; defensa S/A → Statcast; ambas → las dos según el uso; si solo hay uno recomendado, ese.

### Mejores fuentes para completar reliquias (no accesibles desde claude.ai)
- **Champion Spotlights oficiales de Kabam** (`playcontestofchampions.com/news/champion-spotlight-...`): tienen una sección «Recommended Masteries, Stat Focus, Relics» con la reliquia recomendada y el porqué. Se pueden leer por la API WP-JSON del sitio (así lo hizo el repo Mcoc-help para las inmunidades).
- **MCOCHUB**: `/relics` (lista de 120 Statcasts + 25 Battlecasts con efectos, la página carga por JS), `/stat-focus`, y cada ficha de campeón incluye «Suggested Battlecasts» y «7* Stat Focus Suggestions».
- **MCOCHUB tiene API oficial** (`https://mcochub.insaneskull.com/api/v1/...`: champions, tags, abilities, immunities, aw) con clave gratuita bajo petición (LINE/Discord/email), límite 100 peticiones/hora y 1.000/día, y obligación de citar MCOCHUB con enlace. Distingue `full` (siempre) vs `partial` (condicional/sinergia con `note`). Es **la mejor fuente para la app**: más fiable que mis heurísticas y con permiso de uso. No incluye reliquias.

## 5. Campeones solo 6★ (sin versión 7★)
Abomination, Aegon, Agent Venom, Air-Walker, Annihilus, Black Bolt, Black Widow, Captain America, Captain America (WWII), Captain Marvel (Classic), Cosmic Ghost Rider, Cyclops (New Xavier School), Daredevil, Darkhawk, Doctor Octopus, Dormammu, Drax, Electro, Elektra, Green Goblin, Groot, Guillotine, Hela, Hercules, Hulk (Ragnarok), Ikaris, Iron Fist, Iron Fist (Immortal), Iron Patriot, Jubilee, Kate Bishop, King Groot, Loki, Magneto, Moon Knight, Morningstar, Ms. Marvel, Ms. Marvel (Kamala Khan), Nebula, Old Man Logan, Proxima Midnight, Psycho-Man, Psylocke, Quake, Rogue, Ronan, Silver Centurion, Squirrel Girl, Taskmaster, The Champion, Thor, Unstoppable Colossus, Vision (Age of Ultron), Vulture, Wasp, Winter Soldier, Wolverine, Yellowjacket.
(Clases verificadas con mcoc.app y MCOC Scout; Squirrel Girl es Habilidad, corregido en el Excel.)

## 6. Tablas de equivalencias útiles

### Alias de nombres en MCOC Guide Blog → nombre canónico
```json
{
"Dr. Bong": "Doctor Bong",
"Stark Spidey": "Spider-Man (Stark Enhanced)",
"She Hulk": "She-Hulk",
"Spider Punk": "Spider-Punk",
"Immortal Abomination": "Abomination (Immortal)",
"Immortal Hulk": "Hulk (Immortal)",
"Ant Man": "Ant-Man",
"Joaquin Torres (Falcon)": "Falcon (Joaquín Torres)",
"Spiderman 2099": "Spider-Man 2099",
"Yellow Jacket": "Yellowjacket",
"MODOK": "M.O.D.O.K.",
"Captain America (Classic)": "Captain America",
"Hit Monkey": "Hit-Monkey",
"Franken Castle": "Franken-Castle",
"Black Panther (Classic)": "Black Panther",
"Daredevil (Classic)": "Daredevil",
"Magneto OG (Red Suit)": "Magneto",
"Magneto (Marvel Now)": "Magneto (House of X)",
"Negasonic Warhead": "Negasonic Teenage Warhead",
"X-23": "Wolverine (X-23)",
"Hulk Buster": "Hulkbuster",
"Howard The Duck": "Howard the Duck",
"Infamous Iron Man": "Iron Man (Infamous)",
"Iron Man (IW)": "Iron Man (Infinity War)",
"Star Lord": "Star-Lord",
"Storm Pyramid X": "Storm (Pyramid X)",
"Spider-Man Symbiote": "Spider-Man (Symbiote)",
"Venom The Duck": "Venom the Duck",
"Kamala Khan": "Ms. Marvel (Kamala Khan)",
"Air Walker": "Air-Walker",
"Black Widow (CV)": "Black Widow (Claire Voyant)",
"Spiderman Supreme": "Spider-Man (Supreme)",
"Spider-Woman": "Spider-Woman (Jessica Drew)",
"Scarlet Witch (Classic)": "Scarlet Witch (Classic)"
}
```

### Glosario de efectos inglés → español (usado en el Excel)
```json
{
"Acid Burn": "Quemadura ácida",
"Adrenaline": "Adrenalina",
"Amplify": "Amplificar",
"Aptitude": "Aptitud",
"Armor": "Armadura",
"Armor Break": "Rotura de armadura",
"Armor Shattered": "Armadura destrozada",
"Armor Up": "Aumento de armadura",
"Atrophy": "Atrofia",
"Auto-Block": "Bloqueo automático",
"Battered": "Maltrecho",
"Berserk": "Berserker",
"Bleed": "Sangrado",
"Bleed Resistance": "Resistencia a sangrado",
"Bleed Vulnerability": "Vulnerabilidad a sangrado",
"Block Penetration": "Penetración de bloqueo",
"Breach": "Brecha",
"Buff Immunity Debuff": "Inmunidad a mejoras (perjuicio)",
"Bulwark": "Baluarte",
"Cauterize": "Cauterizar",
"Charm": "Encanto",
"Cleanse": "Limpiar",
"Coldsnap": "Ola de frío",
"Coldsnap Resistance": "Resistencia a ola de frío",
"Combo Detonation": "Detonación de combo",
"Combo Shield": "Escudo de combo",
"Combo Wave": "Onda de combo",
"Concussion": "Conmoción",
"Confidence": "Confianza",
"Corrosion": "Corrosión",
"Counterpunch": "Contragolpe",
"Cowardice": "Cobardía",
"Critical Damage": "Daño crítico",
"Critical Resistance": "Resistencia a críticos",
"Cruelty": "Crueldad",
"Crumble": "Desmoronar",
"Crush": "Aplastar",
"Cunning": "Astucia",
"Damage Reflection": "Reflejo de daño",
"Danger Sense": "Sentido del peligro",
"Daunted": "Intimidado (Daunted)",
"Debilitate": "Debilitar",
"Decelerate": "Desacelerar",
"Degeneration": "Degeneración",
"Degenration Resistance": "Resistencia a degeneración",
"Degenration Vulnerability": "Vulnerabilidad a degeneración",
"Delirium": "Delirio",
"Delirium Duration Resist": "Resistencia a duración de delirio",
"Destroy": "Destruir",
"Devolve": "Involucionar",
"Direct Damage": "Daño directo",
"Disarm": "Desarmar",
"Disintegration": "Desintegración",
"Disintegration Vulnerability": "Vulnerabilidad a desintegración",
"Disorient": "Desorientar",
"Dodge or Dodging": "Esquivar",
"Embolden": "Envalentonar",
"Endurance": "Aguante",
"Energize": "Energizar",
"Energy Damage": "Daño de energía",
"Energy Penetration": "Penetración de energía",
"Energy Resistance": "Resistencia a energía",
"Energy Vulnerability": "Vulnerabilidad a energía",
"Enervate": "Enervar",
"Ensnare": "Atrapar",
"Evade": "Evasión",
"Exhaust": "Agotar",
"Exhaustion": "Agotamiento",
"Expertise": "Pericia",
"Falter": "Vacilar",
"Fate Seal": "Sello del destino",
"Fatigue": "Fatiga",
"Fear": "Miedo",
"Fervor": "Fervor",
"Flourish": "Florecer",
"Force Field": "Campo de fuerza",
"Fragility": "Fragilidad",
"Frenzy": "Frenesí",
"Frostbite": "Congelación",
"Fumble": "Torpeza",
"Fury": "Furia",
"Glance": "Golpe de refilón",
"Glancing": "Golpe de refilón",
"Grit": "Coraje",
"Halt": "Detener",
"Heal Block": "Bloqueo de curación",
"Heavy Proficiency": "Dominio del ataque fuerte",
"Hinder": "Obstaculizar",
"Ignition": "Ignición",
"Immortality": "Inmortalidad",
"Impact": "Impacto",
"Implacable": "Implacable",
"Incinerate": "Incinerar",
"Incinerate Duration Resist": "Resistencia a duración de incinerar",
"Incinerate Resistance": "Resistencia a incinerar",
"Incinerate Vulnerability": "Vulnerabilidad a incinerar",
"Indestructible": "Indestructible",
"Ineptitude": "Ineptitud",
"Inexorable": "Inexorable",
"Infuriate": "Enfurecer",
"Injury": "Herida",
"Instinct": "Instinto",
"Intensify": "Intensificar",
"Intimidate": "Intimidar",
"Invalidate": "Invalidar",
"Invisibility": "Invisibilidad",
"Life Steal": "Robo de vida",
"Light Proficiency": "Dominio del ataque ligero",
"Marked": "Marcado",
"Misery": "Miseria",
"Miss": "Fallo (Miss)",
"Myriad": "Miríada",
"Neuroshock": "Neurodescarga",
"Neurotoxin": "Neurotoxina",
"Neutralize": "Neutralizar",
"No Critical": "Sin críticos",
"Non Contact Attacks": "Ataques sin contacto",
"Nova Flame": "Llama nova",
"Nullify": "Anular",
"Obfuscate": "Ofuscar",
"Overload": "Sobrecarga",
"Panic": "Pánico",
"Parry Non-contact": "Parada sin contacto",
"Perfect Block": "Bloqueo perfecto",
"Petrify": "Petrificar",
"Phasing or Phase": "Atravesar (Phasing)",
"Physical Damage": "Daño físico",
"Physical Penetration": "Penetración física",
"Physical Resistance": "Resistencia física",
"Physical Vulnerability": "Vulnerabilidad física",
"Pierce": "Perforar",
"Plasma": "Plasma",
"Poison": "Veneno",
"Poison Duration Resist": "Resistencia a duración de veneno",
"Poison Resistance": "Resistencia a veneno",
"Power Burn": "Quema de poder",
"Power Detonation": "Detonación de poder",
"Power Dissolve": "Disolución de poder",
"Power Drain": "Drenaje de poder",
"Power Efficiency": "Eficiencia de poder",
"Power Gain": "Ganancia de poder",
"Power Leech": "Sanguijuela de poder",
"Power Lock": "Bloqueo de poder",
"Power Reroute": "Redirección de poder",
"Power Steal": "Robo de poder",
"Power Sting": "Aguijón de poder",
"Precision": "Precisión",
"Protection": "Protección",
"Provoke": "Provocar",
"Prowess": "Destreza (Prowess)",
"Pulverize": "Pulverizar",
"Purify": "Purificar",
"Pursuit": "Persecución",
"Rage": "Ira",
"Reckless": "Temerario",
"Regeneration": "Regeneración",
"Reinforce": "Reforzar",
"Rejuvenate": "Rejuvenecer",
"Resistance Up": "Aumento de resistencia",
"Resonance": "Resonancia",
"Retribution": "Represalia",
"Reverb": "Reverberación",
"Reversal": "Inversión",
"Reverse Controls": "Controles invertidos",
"Root": "Enraizar",
"Rupture": "Ruptura",
"Rupture Resistance": "Resistencia a ruptura",
"Sabotage": "Sabotaje",
"Safeguard": "Salvaguarda",
"Seared": "Chamuscado",
"Self-Repair": "Autorreparación",
"Shock": "Descarga",
"Shock Resistance": "Resistencia a descarga",
"Sleep": "Sueño",
"Slow": "Lentitud",
"Slow Resistance": "Resistencia a lentitud",
"Soul Barb": "Púa de alma",
"Special Concussion": "Conmoción especial",
"Special Inexorable": "Inexorable especial",
"Special Lock": "Bloqueo de especiales",
"Special Power Drain": "Drenaje de poder especial",
"Spectre": "Espectro",
"Stagger": "Tambaleo",
"Steadfast": "Firmeza",
"Steady": "Estable",
"Stifle": "Sofocar",
"Stimulate": "Estimular",
"Stun": "Aturdimiento",
"Sunder": "Hendir",
"Suppression": "Supresión",
"Taunt": "Provocación (Taunt)",
"Tenacity": "Tenacidad",
"Torment": "Tormento",
"Tracking": "Rastreo",
"Tranquilize": "Tranquilizar",
"Trauma": "Trauma",
"True Accuracy": "Precisión verdadera",
"True Damage": "Daño verdadero",
"True Focus": "Enfoque verdadero",
"True Sense": "Sentido verdadero",
"True Strike": "Golpe verdadero",
"Unblockable": "Imbloqueable",
"Uncanny": "Asombroso",
"Undermine": "Socavar",
"Undying": "Imperecedero",
"Unstoppable": "Imparable",
"Untouchable": "Intocable",
"Vicious": "Despiadado",
"Vigilance": "Vigilancia",
"Vitality": "Vitalidad",
"Vulnerability": "Vulnerabilidad",
"Weakness": "Debilidad",
"Whiplash": "Latigazo",
"Wither": "Marchitar",
"Zeal": "Celo",
"Robotics": "Robótica",
"Sand": "Arena",
"Austere": "Austero",
"Aspect of Evolution": "Aspecto de la evolución",
"Orgnaic Magnetism": "Magnetismo orgánico",
"Oshtur’s Refusal": "Rechazo de Oshtur",
"Inverted Controls": "Controles invertidos",
"Reverse Control": "Controles invertidos",
"Reversed Controls": "Controles invertidos",
"Mind Control": "Control mental",
"Ability Accuracy Reduction": "Reducción de precisión de habilidades",
"Buff": "Inmunidad a mejoras (no puede recibir buffs)",
"Debuffs": "todos los perjuicios (debuffs)",
"Knockback": "Retroceso",
"Disable Willpower": "Desactivar voluntad",
"Daunt": "Intimidado",
"Special Attacks": "Ataques especiales",
"Contact Attacks": "Ataques de contacto",
"Unsteady": "Inestable",
"Power Burns": "Quema de poder"
}
```

### Títulos de tarjetas del kit
```json
{
"PASSIVE": "Pasiva",
"ALWAYS ACTIVE": "Siempre activa",
"SPECIAL ATTACK 1": "Ataque especial 1",
"SPECIAL ATTACK 2": "Ataque especial 2",
"SPECIAL ATTACK 3": "Ataque especial 3",
"SPECIAL ATTACKS": "Ataques especiales",
"HEAVY ATTACKS": "Ataques fuertes",
"HEAVY ATTACK": "Ataque fuerte",
"ALL ATTACKS": "Todos los ataques",
"CRITICAL HITS": "Golpes críticos",
"BLOCKING": "Al bloquear",
"BASIC ATTACKS": "Ataques básicos",
"WHEN CHARGING A HEAVY ATTACK": "Al cargar ataque fuerte",
"WHEN ATTACKED": "Al ser atacado",
"MEDIUM ATTACKS": "Ataques medios",
"LIGHT ATTACKS": "Ataques ligeros",
"FIGHT START": "Inicio del combate",
"PRE-FIGHT ABILITY": "Habilidad precombate",
"WELL TIMED BLOCKS": "Bloqueos a tiempo",
"WHEN KNOCKED DOWN": "Al ser derribado",
"OPPONENT ACTIVATES A SPECIAL ATTACK": "Cuando el rival activa un especial",
"SPECIAL ATTACKS 2 AND 3": "Ataques especiales 2 y 3",
"ON FOURTH LIGHT ATTACK": "Cuarto ataque ligero",
"PERSISTENT CHARGE": "Carga persistente"
}
```

## 7. Aviso legal para un repo público en GitHub
- MCOC, los nombres de campeones y sus retratos son marcas/copyright de **Kabam Games** y **Marvel**. Añadir un aviso de “fan project, no afiliado a Kabam ni Marvel” en el README y en la app.
- El repo **Shadey1/Mcoc-help** **no tiene licencia** (todos los derechos reservados por defecto) y los textos de los kits vienen de MCOCHUB. **No subir copias de sus JSON ni el texto completo de los kits** a un repo público sin permiso; enlazar a la ficha de MCOCHUB en su lugar. El Excel solo contiene resúmenes derivados (listas de efectos, inmunidades, tiers) + enlaces: aun así, conviene citar las fuentes en el README.
- No usar retratos oficiales sin comprobar los términos; mejor iconos propios o solo texto.
- Los tiers son opiniones de sus autores: citarlos (mcoc.app / Vega, Lagacy, MetalSonicDude, Seatin; MCOC Guide Blog / CaptainRaghu).

## 8. Ideas para la app (a decidir con el usuario)
- Buscador de reliquias: «qué reliquia pongo a X» y «a quién le pongo esta reliquia».
- Buscador de campeones + filtros por clase, rareza (7★/6★), tier, inmunidad (“inmune a Veneno”), efecto que provoca (“provoca Anular”).
- Ficha de campeón: tiers de ambas fuentes, inmunidades (permanentes/condicionales/no confirmadas), mapa de habilidades, sinergias, enlace a MCOCHUB.
- Vista “contadores”: dado un nodo/efecto, qué campeones son inmunes.
- Stack sugerido (simple y gratis en GitHub Pages): Vite + React + TypeScript, datos en `public/data/champions.json`, despliegue con GitHub Actions a GitHub Pages. Preguntar al usuario antes de elegir.
- Script reproducible `scripts/excel-to-json` para regenerar el JSON cuando se actualice el Excel.
- Las tier lists cambian cada mes: diseñar el JSON para poder actualizar tiers sin tocar el resto.

## 9. Actualización del 12/11/2026
- **Nuevo:** Star-Lord (Kitty Pryde) — Tecnología, 7★ (salió el 09/11/2026), tier A en mcoc.app (solo la valora MetalSonicDude), sin tier en Guide Blog. No está aún en `abilities.json` ni en MCOC Scout: su fila se rellenó a mano a partir del **Champion Spotlight oficial** (WP-JSON `playcontestofchampions.com/wp-json/wp/v2/posts?slug=champion-spotlight-star-lord-kitty-pryde`), que también da las reliquias (Battlecast Vision, Statcast Stark Spidey). Cuando MCOCHUB/Mcoc-help la incluyan, regenerar `kit-effects.js`.
- **Excluido a propósito:** Summoned Symbioid (clase Universal, 1★ de evento: no existe en 6★/7★).
- **Tiers mcoc.app cambiados:** Beast F→A, Colossus (AoA) A→B, Ghost Rider (Robbie Reyes) A→B, Hit-Monkey A→S, Mephisto F→D, Vision (Aarkus) C→B. Guide Blog actualizado el 05/11/2026 sin cambios respecto al Excel.
- MCOCHUB cambió el slug de Captain America (Infinity War): `caiw` → `captainamerica-movie` (enlace corregido).
- Kits de MCOCHUB (Mcoc-help, importados 09/11/2026) y reliquias de MCOC Scout: sin cambios.
- Excel se edita con **Excel COM desde PowerShell** (en este PC no hay Python ni Node). Fandom (`static.wikia.nocookie.net`) bloquea descargas con Cloudflare; los retratos nuevos de MCOCHUB vienen en `.webp`, que `System.Drawing` no lee: se convierten con WPF (`PresentationCore`, `BitmapDecoder`).
- mcoc.app usa nombres cortos: `Carina` = Carina Tivan, `Scarlet Witch` = Scarlet Witch (Sigil), `Spider-Man` = Spider-Man (Classic).

## 10. Tier list general (12/11/2026)
- Hoja nueva **«Tier list general»** (todas las clases juntas), generada por `scripts/update-tiers.ps1` (Excel COM; también actualiza las columnas de tier de Campeones y Reliquias y la Leyenda). Ejecutarlo cada mes y luego `excel-to-json.ps1` + subir `VERSION` en `sw.js`.
- Opiniones independientes (hasta 5, todas de ATAQUE): Vega, Lagacy, MetalSonicDude, Seatin (nota 0-100 de cada uno, leída de `mcoc.app/champions/<slug>`) + MCOC Guide Blog. Letras → centro de la franja de mcoc.app (S+≥96, S≥86, A≥72, B≥60, C≥46, D≥30). Media → tier general. Acuerdo por rango (≤20 Alto, ≤40 Medio, >40 Bajo).
- **MCOC Scout copia el Guide Blog** (300/300 iguales; llama «A++» al S+): solo se usa cuando el Guide Blog no valora al campeón.
- Las hojas de Google de los 4 youtubers se pueden exportar (`/export?format=xlsx`) pero tienen formatos libres (tiers en columnas, emojis): es más fiable leer mcoc.app.
- **Descartadas:** gacha.gg (duplicados con notas distintas, Doctor Strange en D, Stark Spidey en C), PocketGamer (Doctor Strange en el último tier, Jack O'Lantern como Místico), MMO Culture (solo ~5 por tier), findingdulcinea/ofzenandcomputing/rosenberryrooms (contenido reciclado, 403).
- App: pestaña Ranking → «Todas» + «General» (por defecto); en la ficha, tier general con puesto y desplegable «Comparar tier lists».
- PowerShell + Excel COM: tras asignar texto a `.Value2`, asignar un `object[,]` a `.Value2` falla («no se puede convertir a String»); usar `.Formula = $arr`. Los elementos del array deben ser tipos .NET simples (no PSObject).

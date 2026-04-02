# 🎵 Spotify Global Top Songs 2026 : Analytics dashboard

> Projet d'analyse et de visualisation des données Spotify Global Charts 2026, décliné en deux formats : **dashboard web interactif** (HTML/CSS/JS) et **classeur excel**.

---

## Structure du projet

```
spotify-dashboard-2026/
│
├── spotify-dashboard-2026.html     # Dashboard web interactif
├── spotify-dashboard-2026.xlsx     # Classeur excel 
└── README.md                       # Ce fichier
```

---

## Source des données

| Champ | Détail |
|---|---|
| **Dataset** | Spotify Global Top Songs 2026 |
| **Source** | [Kaggle — mkaur1141](https://www.kaggle.com/datasets/mkaur1141/spotify-global-top-songs-2026) |
| **Période** | 2026 (snapshot global) |
| **Volume** | 200 tracks, 13 variables |

### Variables disponibles

| Colonne | Type | Description |
|---|---|---|
| `track_name` | string | Titre du morceau |
| `artist_name` | string | Artiste principal |
| `streams` | int | Streams quotidiens |
| `stream_change` | int | Variation vs jour précédent (positif = hausse) |
| `7day` | int | Total streams sur 7 jours glissants |
| `genre` | string | Genre musical |
| `country` | string | Pays d'origine de l'artiste |
| `pos` | int | Position dans le classement global |
| `days` | int | Nombre de jours consécutifs dans le chart |
| `viral_score` | int | Score viral combiné (streams + engagement) |
| `trend` | string | `Rising` / `Falling` |
| `popularity_category` | string | `Trending` / `Average` |
| `longevity` | string | `Evergreen` / `Stable Hit` / `New` |

---

## Méthodologie & formules

### 1. KPIs principaux

Les indicateurs clés sont calculés dynamiquement dans l'onglet **Dashboard** via des formules Excel référençant la feuille `Raw_Data`.

#### Total Streams
```excel
=SUM(Raw_Data!C2:C201)
```
Somme brute de tous les streams quotidiens des 200 tracks.

#### Moyenne de streams par track
```excel
=AVERAGE(Raw_Data!C2:C201)
```

#### Nombre de tracks en hausse
```excel
=COUNTIF(Raw_Data!K2:K201, "Rising")
```
Compte les tracks dont la colonne `trend` vaut `"Rising"`.

#### Nombre d'artistes uniques
```excel
=SUMPRODUCT(1/COUNTIF(Raw_Data!B2:B201, Raw_Data!B2:B201))
```
On divise 1 par le nombre d'occurrences de chaque artiste, puis on somme. Chaque artiste ne compte alors que pour 1 au total, quelle que soit sa fréquence.

#### Âge moyen sur le chart
```excel
=AVERAGE(Raw_Data!I2:I201)
```
Moyenne de la colonne `days`. Indique depuis combien de jours les tracks sont en moyenne présentes dans le classement.

#### Total 7 jours
```excel
=SUM(Raw_Data!E2:E201)
```

#### Peak streams (record single track)
```excel
=MAX(Raw_Data!C2:C201)
```

#### Tracks Trending
```excel
=COUNTIF(Raw_Data!L2:L201, "Trending")
```

#### Tracks Evergreen
```excel
=COUNTIF(Raw_Data!M2:M201, "Evergreen")
```

#### Tracks New
```excel
=COUNTIF(Raw_Data!M2:M201, "New")
```

---

### 2. Analyse par genre

Calculée dans l'onglet **Genre Analysis**, pour chaque genre :

#### Nombre de tracks par genre
```excel
=COUNTIF(Raw_Data!F$2:F$201, A3)
```
où `A3` contient le nom du genre.

#### Total streams par genre
```excel
=SUMIF(Raw_Data!F$2:F$201, A3, Raw_Data!C$2:C$201)
```

#### Streams moyens par genre
```excel
=IFERROR(SUMIF(Raw_Data!F$2:F$201, A3, Raw_Data!C$2:C$201) / COUNTIF(Raw_Data!F$2:F$201, A3), 0)
```
Le `IFERROR` évite les erreurs `#DIV/0!` si un genre n'a qu'une seule occurrence ou est absent.

#### Part de marché streaming (%)
```excel
=SUMIF(Raw_Data!F$2:F$201, A3, Raw_Data!C$2:C$201) / SUM(Raw_Data!C$2:C$201)
```
Formaté en `0.00%` dans Excel.

---

### 3. Analyse des tendances (Trend Analysis)

#### Variation de streams (stream change)
La colonne `stream_change` représente la différence absolue de streams entre J et J-1 :

```
stream_change = streams(J) - streams(J-1)
```

- Valeur **positive** → track en croissance (`Rising`)
- Valeur **négative** → track en déclin (`Falling`)

#### Classement des tracks en hausse
```excel
=LARGE(IF(Raw_Data!D$2:D$201>0, Raw_Data!D$2:D$201), ROW()-3)
```
Formule matricielle (Ctrl+Shift+Entrée) pour extraire les N plus fortes hausses.

#### Score viral
Le `viral_score` fourni dans le dataset est un indicateur composite :
```
viral_score ≈ streams + (7day_streams × 0.5)
```
Il combine la performance immédiate et la dynamique hebdomadaire.

---

### 4. Analyse géographique (Country Analysis)

#### Tracks par pays
```excel
=COUNTIF(Raw_Data!G$2:G$201, A3)
```

#### Streams totaux par pays
```excel
=SUMIF(Raw_Data!G$2:G$201, A3, Raw_Data!C$2:C$201)
```

#### Artistes uniques par pays
```excel
=SUMPRODUCT((Raw_Data!G$2:G$201=A3) / COUNTIFS(Raw_Data!G$2:G$201, Raw_Data!G$2:G$201, Raw_Data!B$2:B$201, Raw_Data!B$2:B$201))
```
Variante du `SUMPRODUCT` dédoublonnant, filtré par pays.

#### Part de marché pays
```excel
=SUMIF(Raw_Data!G$2:G$201, A3, Raw_Data!C$2:C$201) / SUM(Raw_Data!C$2:C$201)
```

---

### 5. Mise en forme conditionnelle

| Règle | Colonne | Couleur |
|---|---|---|
| `stream_change > 0` | Stream Change | 🟢 Vert `#1DB954` |
| `stream_change < 0` | Stream Change | 🔴 Rouge `#FF4D6D` |
| `trend = "Rising"` | Trend | 🟢 Vert |
| `trend = "Falling"` | Trend | 🔴 Rouge |
| Color Scale min→max | Streams, Total Streams | Dégradé sombre → cyan |
| `longevity = "Evergreen"` | Longevity | 🔵 Cyan `#00D4FF` |
| `longevity = "Stable Hit"` | Longevity | 🟦 Teal `#00B8A0` |
| `longevity = "New"` | Longevity | 🟡 Amber `#F5A623` |

---

## 🎨 Design & Charte graphique

Le dashboard reprend ma palette personnalisée **dark mode** inspirée de mon propre portfolio :

```css
--bg:     #080B0F   /* Fond principal */
--cyan:   #00D4FF   /* Accent primaire :streams, KPIs clés */
--teal:   #00B8A0   /* Accent secondaire : 7 jours, pays */
--amber:  #F5A623   /* Alerte / New tracks */
--blue:   #5B9CF6   /* Artistes, données secondaires */
--green:  #1DB954   /* Rising / hausse */
--red:    #FF4D6D   /* Falling / baisse */
--ink:    #E4EDF5   /* Texte principal */
--mid:    #7A8FA0   /* Texte secondaire */
--muted:  #435060   /* Labels discrets */
```

**Typographies (version web) :**
- `Playfair Display` pour les titres (serif élégant)
- `JetBrains Mono` pour les valeurs, KPIs, labels techniques
- `DM Sans` pour le corps de texte

---

##  Stack technique

### Version Web (HTML)
| Outil | Usage |
|---|---|
| HTML5 / CSS3 | Structure & mise en page |
| Chart.js 4.4 | Graphiques (doughnut, bar horizontale) |
| Vanilla JS | Logique de filtres, calculs dynamiques |
| CSS Variables | Theming cohérent |
| IntersectionObserver | Animations au scroll |

### Version Excel
| Outil | Usage |
|---|---|
| Formules Excel | KPIs dynamiques, agrégations |
| Mise en forme conditionnelle | Color scales, règles visuelles |
| Filtres automatiques | Navigation dans Raw_Data |

---

## Insights clés

- **BTS domine** le top 10 avec 9 tracks issues d'une même sortie d'album, cumulant plus de **50M de streams** en deux jours, signal d'un pic de lancement typique des groupes K-Pop.
- **Les tracks Evergreen** (ex. *Mr. Brightside*, *Creep*, *Sweater Weather*) maintiennent 1.5M–2.5M streams/jour après plusieurs années, leur `days` dépasse parfois **2000 jours**.
- **La Pop reste le genre dominant** en volume de tracks, mais le **Reggaeton** (Bad Bunny) et le **Regional Mexicano** (Peso Pluma, Fuerza Regida) concentrent une part disproportionnée des streams latam.
- **Les USA et le Royaume-Uni** représentent à eux deux plus de 50% des tracks en chart, mais la Corée du Sud (`KR`) monopolise le top 12 sur cette période.
- Les tracks en mode **Rising** enregistrent en moyenne **+75K streams/jour** de gain, contre **-280K/jour** pour les tracks en chute, asymétrie classique de la dynamique de chart.

---

## Lancer le projet

```bash
# Cloner le repo
git clone https://github.com/ton-username/spotify-dashboard-2026.git
cd spotify-dashboard-2026

# Ouvrir le dashboard web
open spotify-dashboard-2026.html
# ou double-clic sur le fichier

# Ouvrir le classeur Excel
open spotify_dashboard_2026.xlsx
```

Aucune dépendance à installer pour la version web, tout est auto-contenu.

---

## À venir (idées d'évolution)

- [ ] Connecter à l'API Spotify officielle pour des données temps réel
- [ ] Ajouter un onglet **Corrélation** (viral score vs longevity)
- [ ] Intégrer une timeline d'évolution des positions
- [ ] Version Power BI `.pbix` avec slicers interactifs
- [ ] Analyse NLP des titres de tracks (longueur, langue, sentiment)

---

*Dataset : Kaggle · [mkaur1141/spotify-global-top-songs-2026](https://www.kaggle.com/datasets/mkaur1141/spotify-global-top-songs-2026)*

# 📦 Panel MRP · Enaex S.A.

Aplicación en **Streamlit** para la planificación de materiales (MRP): histórico
y pronóstico de demanda, y estado de abastecimiento (solped, OC, disponibilidad
y TAT). Reemplaza los reportes de Power BI.

Toda la aplicación está en **un solo archivo**: `app.py`.

---

## 🗂️ Contenido de la carpeta

```
mrp-panel-enaex/
├── app.py                 ← TODA la aplicación (cálculos + páginas)
├── requirements.txt       ← librerías que instala Streamlit
├── README.md              ← este archivo
├── .gitignore
├── .streamlit/
│   ├── config.toml            ← tema visual
│   └── secrets.toml.example   ← plantilla de contraseña y token
├── assets/
│   └── Tipos_de_demanda.png   ← imagen explicativa
└── data/                  ← aquí van los Excel de SAP
    ├── MB51/   (movimientos)          ← ⚠️ pon aquí tu MB51 real
    ├── MB5B/   (stock por mes)        ← ⚠️ pon aquí tus MB5B reales
    ├── MRP/    (Planificacion_Simpl)  ✅ incluido
    ├── MM60/   (maestro materiales)   ✅ incluido
    ├── ME5A/   (solped)               ✅ incluido
    ├── ME2M/   (órdenes de compra)    ✅ incluido
    └── TAT/    (Vista Ejecutiva)      ✅ incluido
```

> Las carpetas **MB51** y **MB5B** vienen vacías: debes poner ahí tus archivos
> reales (los de ejemplo se eliminaron para que no ensucien los cálculos).

---

## 🚀 Subir a GitHub y publicar

1. Sube esta carpeta completa a tu repositorio.
2. En [share.streamlit.io](https://share.streamlit.io) → **Create app**:
   - **Repository:** tu repositorio
   - **Branch:** `main`
   - **Main file path:** `mrp-panel-enaex/app.py`
     *(ajusta la ruta si la dejas dentro de otra carpeta)*
3. **Deploy**. La primera vez tarda unos minutos.

Para correrlo en tu computador:

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 🧭 Las páginas

| Sección | Página | Qué muestra |
|---|---|---|
| Inicio | 🏠 Inicio | Estado de los Excel cargados y resumen de KPIs |
| 01 Visualizaciones | 🎯 **Control de Materiales** | **Vista integrada**: stock y criticidad, estado del stock, demanda proyectada y tiempo hasta la demanda, si el stock alcanza o no, solped/OC con sus días de gestión, TAT y comentario |
| 01 Visualizaciones | 📈 Demanda y Pronóstico | Histórico por material (ingresos, egresos, stock), clasificación y pronóstico |
| 01 Visualizaciones | 💰 **Costos** | Valor del stock total, por condición, por grupo de compra y valor en tránsito |
| 01 Visualizaciones | 🚚 MRP E002 | Solped, OC, días de gestión, nacional/internacional, disponibilidad y TAT |
| 02 Datos | 📥 Cargar archivos | Subir las descargas de SAP |
| 02 Datos | 📖 Cómo usar | Guía completa |

### 🎯 Control de Materiales — qué responde

Para cada material, en una sola tabla:

1. **Material y stock:** código, descripción, centro, área, criticidad, stock,
   stock de seguridad y **estado del stock** (Sobre Stock · Stock OK · Bajo Stock ·
   Quiebre Stock).
2. **Demanda:** tipo de demanda, **demanda proyectada**, **tiempo hasta la próxima
   demanda**, **stock que queda tras la demanda** y el veredicto:

   | Veredicto | Significa |
   |---|---|
   | **Cumple** | El stock cubre la demanda y conserva el stock de seguridad |
   | **Queda en bajo stock** | Alcanza, pero queda bajo el stock de seguridad |
   | **Consume todo el stock** | La demanda deja el stock en cero: hay que reponer |
   | **No cumple** | El stock **no alcanza** para la próxima demanda |

3. **Gestión (semana más actualizada del MRP):** estado (Con Solped · Con OC ·
   **Solped bloqueada** · **Validación** · Sin gestión), número de solped y OC, y
   los **días de gestión** de cada una.
4. **TAT:** promedio, mínimo, máximo y recurrencia de compra.
5. **Comentario** del material (columna Observación del MRP).

---

## 🧪 Datos de ejemplo (bórralos)

Las carpetas `data/MB51` y `data/MB5B` traen archivos llamados
**`EJEMPLO_BORRAR_MB51.xlsx`** y **`EJEMPLO_BORRAR_MB5B.xlsx`**. Son datos
inventados (con códigos de material reales) solo para que el panel arranque
mostrando todo funcionando.

**Bórralos y sube tus descargas reales de MB51 y MB5B.** Mientras estén, los
números de demanda no son reales.

---

## 📥 Qué Excel se suben y cada cuánto

| Excel | Origen / layout | Cada cuánto | Al subirlo |
|---|---|---|---|
| **MB51** | Transacción MB51, layout **`/CALCDEMANDA`** | Semanal | Reemplaza |
| **MB5B** | Transacción MB5B (columnas: Material · Descripción · De fecha · A fecha · Stock inicial · Total ctd.entrada mcía. · Total cantidades salida · Stock de cierre · Unidad medida base · Stock especial) | Mensual | **Se agrega** |
| **MRP semanal** | `Planificacion_Simpl…xlsx`, hoja `data` | Semanal | Reemplaza |
| **MM60** | Transacción MM60 | Mensual | Reemplaza |
| **ME5A** | Transacción ME5A | Semanal | Reemplaza |
| **ME2M** | Transacción ME2M | Semanal | Reemplaza |
| **TAT** | Vista Ejecutiva de Materiales, hoja `Dias_TAT` | Mensual | Reemplaza |

Se suben desde la página **📥 Cargar archivos** o dejándolos en las carpetas de
`data/`.

---

## 🔐 Guardar en GitHub desde la app (opcional)

Para que los archivos que subes desde la app queden **permanentes** y protegidos
con contraseña:

1. **Crea un token en GitHub:** foto de perfil → *Settings* → *Developer settings*
   → *Personal access tokens* → *Fine-grained tokens* → *Generate new token*.
   En **Repository access** elige tu repositorio y en **Permissions → Contents**
   marca **Read and write**. Copia el token (`github_pat_…`).
2. **Configura los secrets:** en tu app de Streamlit → *Settings* → *Secrets*,
   pega (con tus valores):

   ```toml
   APP_PASSWORD       = "una-clave-secreta"
   GITHUB_TOKEN       = "github_pat_xxxxxxxx"
   GITHUB_REPO        = "victorhoraci/ENAEX"
   GITHUB_BRANCH      = "main"
   GITHUB_DATA_PREFIX = "mrp-panel-enaex/data"
   ```

Sin esta configuración la app funciona igual, pero los archivos que subas desde
el navegador serán temporales (se pierden al reiniciar).

---

## 🧮 Cómo se calcula

**Clasificación de demanda (ADI / CV², Syntetos & Boylan)** — ADI mide cada
cuánto hay demanda; CV² cuánto varía su tamaño. Cortes: ADI = 1,32 · CV² = 0,49.

| Tipo | Método | Tiempo hasta la próxima demanda |
|---|---|---|
| **Constante** | SES | Días hasta el próximo mes |
| **Errática** | COMBINADO (SES + media móvil 3 + media móvil 6) | Días hasta el próximo mes |
| **Intermitente / Irregular** con **< 4** demandas | SBA | **Indeterminado** |
| **Intermitente / Irregular** con **≥ 4** demandas | PR (Proceso de Renovación) | **Días estimados** |
| **Sin demanda** | — | — |

**Uniones del panel MRP E002** (todas *left join*, así **ningún material se pierde**):

| Fuente | Se une por |
|---|---|
| MRP semanal | tabla base |
| MM60 | Material + Centro |
| ME5A | Solped + Material |
| ME2M | OC + Material |
| TAT | Material (todos los centros) |
| Demanda | Material |

Los parámetros (α = 0,3, cortes ADI/CV², umbral de PR = 4, días por mes = 30)
están al inicio de `app.py`, en la sección **CONFIGURACIÓN DEL MODELO**.

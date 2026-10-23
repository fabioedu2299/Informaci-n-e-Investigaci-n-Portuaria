# Rentabilidad de los puertos concesionados del Perú, 2021 – 2025

Análisis financiero y operativo de los siete terminales portuarios concesionados bajo supervisión del Ositrán, construido sobre los Informes de Desempeño que publica el regulador.

**Elaborado por Fabio Edú García Rojas.** Análisis personal sobre fuentes públicas

---

## Qué hay aquí

| Ruta | Qué es |
|---|---|
| `docs/index.html` | La carta interactiva: panorama del sector, comparativo, economía unitaria, modelo de caja, ficha por terminal, lectura operativa y lectura financiera. Un solo archivo; se abre con doble clic. |
| `Modelo_3Estados_Puertos.xlsx` | Modelo integrado de tres estados con proyección 2026–2030, ficha por terminal, análisis DuPont y factor X de las revisiones tarifarias. 20 hojas, fórmulas vivas. |
| `datos/` | La fuente, en CSV. Es lo único que hay que editar para incorporar un ejercicio nuevo. |
| `scripts/` | El código que convierte los CSV en los dos entregables anteriores. |

## Los terminales

| Código | Terminal | Operador |
|---|---|---|
| `APMT` | Terminal Norte Multipropósito — Callao | APM Terminals Callao |
| `DPW` | Terminal de Contenedores Muelle Sur — Callao | DP World Callao |
| `TCSA` | Terminal de Embarque de Concentrados — Callao | Transportadora Callao |
| `TPP` | Terminal Portuario de Paita | Terminales Portuarios Euroandinos |
| `TISUR` | Terminal Portuario de Matarani | TISUR |
| `STI` | Terminal Portuario de Salaverry | Salaverry Terminal Internacional |
| `TPGSM` | Terminal Portuario General San Martín — Pisco | Terminal Portuario Paracas |

## Los datos

Todas las cifras monetarias están **en miles de dólares**, tal como las publica el Ositrán.

| Archivo | Contenido |
|---|---|
| `terminales.csv` | Identificación de cada concesión: operador, tipo de carga, modalidad, inicio y plazo, compromiso de inversión y avance. |
| `estado_resultados.csv` | Estado de resultados auditado, una fila por terminal y ejercicio. |
| `estado_situacion.csv` | Estado de situación financiera auditado, misma estructura. |
| `trafico.csv` | Naves atendidas, carga en toneladas, TEU y contenedores. Fuente: Declaración Estadística del Ositrán. |
| `margen_ebitda_oficial.csv` | Margen EBITDA publicado por el regulador. Solo existe para 2024 y 2025, salvo en TISUR, que tiene la serie completa. De ahí se deriva la depreciación y amortización del resto de terminales. |
| `depreciacion_amortizacion.csv` | D&A desagregada, donde el informe la publica directamente (DP World). |
| `ratios_oficiales.csv` | Ratios publicados por el Ositrán. No alimentan ningún cálculo: sirven para contrastar lo que calcula `kpis.py`. Una celda vacía es un ratio que el regulador no publicó, distinto de un cero. |

### Incorporar un ejercicio nuevo

1. Añadir las filas del año a `estado_resultados.csv`, `estado_situacion.csv` y `trafico.csv`.
2. Añadir el margen EBITDA y los ratios publicados, si el informe los trae.
3. Ejecutar `python scripts/build.py`.

No hay que tocar código ni editar el HTML a mano. El diff de git deja constancia de qué
cifra cambió y en qué ejercicio.

## Regenerar los entregables

```bash
pip install openpyxl
python scripts/build.py
```

El proceso corre en este orden:

1. **`kpis.py`** calcula los ratios por terminal y ejercicio, y los contrasta contra los que
   publica el Ositrán. Si alguna desviación supera un punto porcentual se detiene: eso
   significa que un estado financiero se cargó mal.
2. **`modelo.py`** construye el puente del EBIT al flujo de caja libre por el método
   indirecto, y las métricas unitarias por segmento de carga.
3. **`build.py`** sustituye el bloque de datos incrustado en `docs/index.html`.
4. **`m3s.py`** reconstruye el libro de tres estados.

La desviación máxima actual contra los ratios publicados es de **0,40 puntos porcentuales**,
atribuible al redondeo con que el regulador los difunde.

## Criterios de cálculo

Decisiones que afectan la comparabilidad y conviene conocer antes de leer las cifras:

- **Deuda financiera.** Incluye las cuentas por pagar a relacionadas de largo plazo. En
  Matarani y Salaverry ese saldo devenga interés y es financiamiento intragrupo, no crédito
  comercial; excluirlo inflaba el ROIC de Matarani de 32,7 % a 79,7 %.
- **EBITDA.** Utilidad operativa más depreciación y amortización. Para DP World se usa la D&A
  que desagrega el informe; para el resto se deriva del margen EBITDA que publica el
  regulador, disponible solo desde 2024. Donde no hay dato, la página muestra un guion en vez
  de una estimación.
- **Ingresos por construcción (CINIIF 12).** Se muestran separados de los ingresos por
  servicios portuarios: no generan margen y distorsionan la comparación entre terminales.
- **Unidad de trabajo.** El costo por tonelada no aplica a todos los terminales, porque unos
  mueven contenedores y otros granel. Los unitarios se comparan solo dentro de cada segmento;
  entre segmentos se usan EBITDA por recalada y EBITDA sobre activo total.
- **Flujo de caja libre.** Método indirecto, a partir del estado de resultados y de la
  variación del balance, porque el regulador no publica estado de flujos. Requiere el balance
  del ejercicio anterior, de modo que 2021 no admite cálculo.
- **ROE.** No es comparable entre terminales: en las concesiones con reparto agresivo de
  dividendos el patrimonio contable se ha erosionado, lo que infla el ratio sin que medie
  mejora operativa. Para juzgar creación de valor se priorizan el ROA, el ROIC y el EBITDA
  por tonelada.

## Limitaciones

- **Paita (TPP)** reconoce parte de su inversión como activo financiero de la concesión
  (modelo bifurcado CINIIF 12) y no como intangible, de modo que el capex derivado del
  balance subestima su inversión real.
- La proyección 2026–2030 del libro de Excel es un modelo de supuestos editables, no un
  pronóstico del regulador ni de los concesionarios.
- **Chancay** no es una concesión bajo el ámbito del Ositrán y por eso no aparece en el
  tablero, pese a competir directamente por el transbordo del Callao.

## Fuente

Informes de Desempeño 2025 del Ositrán para cada concesión portuaria. Los estados financieros
que reproduce el regulador provienen de los EEFF auditados de cada concesionario; el tráfico
procede de la Declaración Estadística del Ositrán. Documentos públicos, disponibles en
<https://www.ositran.gob.pe>.

## Publicar la página

`docs/index.html` es autocontenido salvo Chart.js, que carga desde cdnjs. Para servirlo con
GitHub Pages: *Settings → Pages → Source: deploy from a branch → rama `main`, carpeta `/docs`*.

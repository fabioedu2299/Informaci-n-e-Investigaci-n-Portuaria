# Scripts

Orden de ejecución; `build.py` los encadena y no hay que invocarlos a mano.

| Archivo | Qué hace | Qué escribe |
|---|---|---|
| `dataset.py` | Lee los CSV de `datos/` y arma las estructuras que usan los demás. No calcula nada. | — |
| `kpis.py` | Ratios por terminal y ejercicio, y validación contra los que publica el regulador. | `datos.json` |
| `modelo.py` | Puente del EBIT al flujo de caja libre y métricas unitarias por segmento. | `modelo.json` |
| `m3s.py` | Libro de tres estados con proyección, ficha por terminal, DuPont y factor X. | `Modelo_3Estados_Puertos.xlsx` |
| `build.py` | Encadena todo y actualiza `docs/index.html` y el libro en la raíz. | — |

`datos.json` y `modelo.json` son intermedios: se regeneran en cada ejecución y están
versionados solo para poder diffear el resultado del cálculo, no la fuente.

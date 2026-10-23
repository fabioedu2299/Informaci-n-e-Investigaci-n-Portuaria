# -*- coding: utf-8 -*-
"""
Regenera todos los entregables a partir de los CSV de `datos/`.

    python scripts/build.py

Qué hace, en orden:

  1. kpis.py      calcula ratios por terminal y año, y los contrasta contra los que publica
                  el Ositrán. Si alguna desviación pasa de 1 punto porcentual, aborta: eso
                  significa que un estado financiero se cargó mal.
  2. modelo.py    construye el puente del EBIT al flujo de caja libre por el método indirecto.
  3. payload      funde ambos en el bloque de datos que lleva incrustado docs/index.html y lo
                  sustituye en el archivo, sin tocar el resto de la página.
  4. m3s.py       reconstruye Modelo_3Estados_Puertos.xlsx.

Para incorporar un ejercicio nuevo basta con añadir sus filas a los CSV y volver a ejecutar
esto: no hay que editar código ni el HTML a mano.
"""
import json
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SCRIPTS = RAIZ / "scripts"
HTML = RAIZ / "docs" / "index.html"
XLSX = RAIZ / "Modelo_3Estados_Puertos.xlsx"

sys.path.insert(0, str(SCRIPTS))


def paso(titulo):
    print(f"\n── {titulo}")


def main():
    import dataset

    paso(f"Datos: {len(dataset.ORDEN)} terminales, ejercicios "
         f"{dataset.YEARS[0]}–{dataset.YEARS[-1]}")

    # 1 y 2 ── ratios y modelo de caja
    for modulo in ("kpis", "modelo"):
        paso(f"Ejecutando {modulo}.py")
        r = subprocess.run([sys.executable, str(SCRIPTS / f"{modulo}.py")],
                           cwd=SCRIPTS, capture_output=True, text=True)
        print(r.stdout.strip()[-900:])
        if r.returncode:
            print(r.stderr.strip()[-1500:])
            sys.exit(f"{modulo}.py terminó con error; no se escribió nada.")

    # 3 ── bloque de datos de la página
    paso("Reconstruyendo el bloque de datos de docs/index.html")
    datos = json.load(open(SCRIPTS / "datos.json", encoding="utf-8"))
    mod = json.load(open(SCRIPTS / "modelo.json", encoding="utf-8"))
    anios = dataset.YEARS

    carga = {"orden": dataset.ORDEN, "anios": anios, "wacc": [0.08, 0.10, 0.12], "p": {}}
    ult = str(anios[-1])
    for k in dataset.ORDEN:
        meta = dict(datos[k]["meta"])
        # La página necesita además el segmento de operación y la participación
        # contenedorizada del tonelaje, que calcula modelo.py sobre el último ejercicio.
        meta.pop("tm_contenedorizada_2025", None)
        meta["segmento"] = mod[k]["segmento"]
        meta["pct_contenedorizada"] = mod[k]["unitarios"][ult]["pct_contenedorizada"]
        carga["p"][k] = {"meta": meta,
                         "s": datos[k]["serie"],
                         "m": {str(x["anio"]): x for x in mod[k]["modelo"]},
                         "u": mod[k]["unitarios"]}

    html = HTML.read_text(encoding="utf-8")
    i = html.index("const D={")
    j = html.index("\n", i)
    cola = html[i:j]
    nuevo = "const D=" + json.dumps(carga, ensure_ascii=False, separators=(",", ":")) \
            + cola[cola.rindex("}") + 1:]
    HTML.write_text(html[:i] + nuevo + html[j:], encoding="utf-8")
    print(f"   bloque de datos: {len(nuevo):,} caracteres")

    # 4 ── libro de tres estados
    paso("Reconstruyendo el modelo de tres estados")
    r = subprocess.run([sys.executable, str(SCRIPTS / "m3s.py")],
                       cwd=SCRIPTS, capture_output=True, text=True)
    print(r.stdout.strip()[-600:])
    if r.returncode:
        print(r.stderr.strip()[-1500:])
        sys.exit("m3s.py terminó con error.")
    generado = SCRIPTS / "Modelo_3Estados_Puertos.xlsx"
    if generado.exists():
        XLSX.write_bytes(generado.read_bytes())
        generado.unlink()

    paso("Listo")
    print(f"   {HTML.relative_to(RAIZ)}")
    print(f"   {XLSX.relative_to(RAIZ)}")


if __name__ == "__main__":
    main()

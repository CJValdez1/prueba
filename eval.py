#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
abrir_revisiones.py

Deja abierto un pull request de "Revisión 2" por estudiante, en el repositorio de su
equipo. Cada pull request lleva un commit tuyo con "Co-authored-by" del estudiante, la
rama revision/USUARIO, el archivo revision/USUARIO.md y las instrucciones de cuatro pasos
para que la persona lo llene, lo apruebe y lo fusione.

Requisitos: Python 3.8 o superior y GitHub CLI (gh) con la sesión iniciada (gh auth login).
El script no pide ni guarda contraseñas ni tokens: usa la sesión de gh.

Uso (en Windows escribe py en lugar de python3):
  python3 abrir_revisiones.py estudiantes.csv                          # simulación, no crea nada
  python3 abrir_revisiones.py estudiantes.csv --equipo E5              # simulación de un equipo
  python3 abrir_revisiones.py estudiantes.csv --equipo E5 --ejecutar   # crea de verdad

Si lo vuelves a correr, se salta lo que ya está abierto: no duplica pull requests.
"""

import argparse
import base64
import csv
import json
import re
import shutil
import subprocess
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except AttributeError:
    pass

PLANTILLA_ARCHIVO = """# Revisión 2: {usuario}

Esta persona llena este archivo antes de la revisión en vivo, criterio por criterio:

1. Cambia [ ] por [x] en las condiciones que cumple.
2. Pega los enlaces donde se ven (la historia de Jira, el pull request, el commit o el archivo).
3. Escribe el nivel que se pone: E si cumple las 3 condiciones, S si cumple 2, I si cumple 1 o ninguna.

El nivel que se pone no es la calificación: el docente la define al revisar la evidencia en vivo.

## 1. Su historia terminada se ve en Jira

- [ ] 1a. Tiene a su nombre una historia en Done al cierre del sprint.
- [ ] 1b. El panel Desarrollo de esa historia muestra su pull request fusionado y un commit suyo.
- [ ] 1c. Dejó al menos dos comentarios de avance.

Enlaces a la evidencia:

Nivel que se pone (E, S o I):

## 2. Su pull request y su prueba

- [ ] 2a. Su pull request tiene la plantilla llena.
- [ ] 2b. Trae al menos una prueba automática que comprueba un criterio de aceptación.
- [ ] 2c. Esa prueba pasa en Actions.

Enlaces a la evidencia:

Nivel que se pone (E, S o I):

## 3. Revisiones y pareja

- [ ] 3a. Recibió al menos un comentario en una línea de su pull request y respondió a cada uno.
- [ ] 3b. Revisó con un comentario en línea el pull request de un compañero.
- [ ] 3c. Tiene un commit con Co-authored-by de su pareja.

Enlaces a la evidencia:

Nivel que se pone (E, S o I):

## 4. Su retrospectiva en main

- [ ] 4a. Su archivo en retrospectiva/, con su usuario de GitHub, llegó a tiempo a main en un pull request que revisó otra persona.
- [ ] 4b. Tiene las tres secciones y cita un work item por su clave.
- [ ] 4c. Cierra con una acción concreta.

Enlaces a la evidencia:

Nivel que se pone (E, S o I):
"""

# La línea Co-authored-by va en un bloque de código: sin él, GitHub oculta o convierte en
# enlace lo que está entre < y >, y la persona copiaría la línea incompleta.
PLANTILLA_PR = """## Revisión 2 de {usuario}

Esta persona cierra su revisión en cuatro pasos, unos 5 minutos.

**1.** Abrir `revision/{usuario}.md` en Files changed (menú ··· > Edit file), llenar los cuatro criterios y escribir el nivel que se pone en cada uno.

**2.** Guardar el cambio en esta misma rama (Commit directly to the revision/{usuario} branch). En la descripción del commit, dejar una línea en blanco y pegar esta línea:

```
Co-authored-by: {docente_nombre} <{docente_correo}>
```

**3.** Aprobar el pull request con Review changes > Approve.

**4.** Fusionarlo con Create a merge commit.

Si GitHub no deja aprobar porque el último commit es de esta persona, que lo apruebe otro integrante del equipo.
"""


class ErrorGh(Exception):
    def __init__(self, estado, mensaje):
        super().__init__(mensaje)
        self.estado = estado


def gh_api(ruta, metodo="GET", datos=None):
    """Llama a la API de GitHub con gh y devuelve el JSON (o None si no hay cuerpo)."""
    cmd = ["gh", "api", "-X", metodo, ruta]
    entrada = None
    if datos is not None:
        cmd += ["--input", "-", "-H", "Content-Type: application/json"]
        entrada = json.dumps(datos).encode("utf-8")
    p = subprocess.run(cmd, input=entrada, capture_output=True)
    salida = p.stdout.decode("utf-8", errors="replace")
    error = p.stderr.decode("utf-8", errors="replace")
    if p.returncode != 0:
        m = re.search(r"HTTP (\d{3})", error + salida)
        raise ErrorGh(int(m.group(1)) if m else None, (error or salida).strip())
    return json.loads(salida) if salida.strip() else None


def limpiar(texto):
    """Quita < > y saltos de línea para que el nombre no rompa la línea Co-authored-by."""
    return re.sub(r"[<>\r\n]", "", texto).strip()


def leer_filas(ruta):
    with open(ruta, "rb") as f:
        crudo = f.read()
    try:
        texto = crudo.decode("utf-8-sig")
    except UnicodeDecodeError:  # Excel en español puede guardar el CSV en cp1252
        texto = crudo.decode("cp1252", errors="replace")
    lineas = texto.splitlines()
    if not lineas:
        sys.exit("El CSV está vacío.")
    delim = ";" if lineas[0].count(";") > lineas[0].count(",") else ","
    lector = csv.DictReader(lineas, delimiter=delim)
    campos = {c.strip().lower(): c for c in (lector.fieldnames or [])}
    for obligatorio in ("repo", "usuario"):
        if obligatorio not in campos:
            sys.exit("Falta la columna '%s' en el CSV. Columnas esperadas: equipo, repo, usuario." % obligatorio)
    filas, vistos = [], set()
    for fila in lector:
        valores = [v.strip() for v in fila.values() if isinstance(v, str) and v.strip()]
        if not valores or valores[0].startswith("#"):
            continue  # fila vacía o comentario
        repo = (fila.get(campos["repo"]) or "").strip()
        usuario = (fila.get(campos["usuario"]) or "").strip().lstrip("@")
        equipo = (fila.get(campos["equipo"]) or "").strip() if "equipo" in campos else ""
        if (repo.lower(), usuario.lower()) in vistos:
            continue
        vistos.add((repo.lower(), usuario.lower()))
        filas.append({"equipo": equipo, "repo": repo, "usuario": usuario})
    return filas


def preparar(fila, docente, args):
    """Revisa todo sin crear nada. Devuelve {'omitir': motivo} o {'plan': {...}}."""
    repo, usuario = fila["repo"], fila["usuario"]
    if repo.count("/") != 1 or not usuario:
        return {"omitir": "fila incompleta: repo debe ser dueño/repositorio y usuario no puede ir vacío"}
    dueno = repo.split("/")[0]
    try:
        info = gh_api("repos/" + repo)
    except ErrorGh as e:
        return {"omitir": "no puedo abrir el repositorio (HTTP %s): revisa el nombre" % e.estado}
    if info.get("archived"):
        return {"omitir": "el repositorio está archivado"}
    if not (info.get("permissions") or {}).get("push"):
        return {"omitir": "no tienes permiso de escritura; GitHub haría un fork y no cuenta para las insignias"}
    base = info["default_branch"]
    try:
        est = gh_api("users/" + usuario)
    except ErrorGh as e:
        return {"omitir": "el usuario %s no existe en GitHub (HTTP %s): revisa el CSV" % (usuario, e.estado)}
    if est.get("type") != "User":
        return {"omitir": "%s no es una cuenta de persona" % usuario}
    login = est["login"]
    avisos = []
    if not args.sin_verificar_colaborador:
        try:
            gh_api("repos/%s/collaborators/%s" % (repo, login))
        except ErrorGh as e:
            if e.estado == 404:
                return {"omitir": "%s no figura como colaborador de %s: revisa que sea el repositorio correcto" % (login, repo)}
            avisos.append("no pude verificar si %s es colaborador (HTTP %s)" % (login, e.estado))
    rama = "revision/" + login
    prs = gh_api("repos/%s/pulls?state=all&head=%s:%s" % (repo, dueno, rama)) or []
    if prs:
        return {"omitir": "ya existe el pull request " + prs[0]["html_url"], "url": prs[0]["html_url"], "estado": "ya existía"}
    try:
        gh_api("repos/%s/git/ref/heads/%s" % (repo, rama))
        return {"omitir": "la rama %s ya existe y no tiene pull request: revísala a mano" % rama}
    except ErrorGh as e:
        if e.estado != 404:
            return {"omitir": "no pude comprobar la rama %s (HTTP %s)" % (rama, e.estado)}
    nombre_est = limpiar(est.get("name") or login)
    correo_est = "%s+%s@users.noreply.github.com" % (est["id"], login)
    mensaje = "Revisión 2: %s\n\nCo-authored-by: %s <%s>" % (login, nombre_est, correo_est)
    return {
        "plan": {
            "repo": repo,
            "login": login,
            "base": base,
            "rama": rama,
            "ruta": "revision/%s.md" % login,
            "mensaje": mensaje,
            "coautor": "%s <%s>" % (nombre_est, correo_est),
            "archivo": PLANTILLA_ARCHIVO.format(usuario=login),
            "titulo": "Revisión 2: " + login,
            "cuerpo": PLANTILLA_PR.format(
                usuario=login, docente_nombre=docente["nombre"], docente_correo=docente["correo"]
            )
            + (("\n**Fecha límite:** %s.\n" % args.fecha_limite) if args.fecha_limite else ""),
            "avisos": avisos,
        }
    }


def ejecutar(plan):
    """Crea la rama, el archivo con el commit coautorado, el pull request y la solicitud de revisión."""
    repo, rama, base, login = plan["repo"], plan["rama"], plan["base"], plan["login"]
    sha = gh_api("repos/%s/git/ref/heads/%s" % (repo, base))["object"]["sha"]
    gh_api("repos/%s/git/refs" % repo, "POST", {"ref": "refs/heads/" + rama, "sha": sha})
    gh_api(
        "repos/%s/contents/%s" % (repo, plan["ruta"]),
        "PUT",
        {
            "message": plan["mensaje"],
            "content": base64.b64encode(plan["archivo"].encode("utf-8")).decode("ascii"),
            "branch": rama,
        },
    )
    pr = gh_api(
        "repos/%s/pulls" % repo,
        "POST",
        {"title": plan["titulo"], "head": rama, "base": base, "body": plan["cuerpo"]},
    )
    nota = ""
    try:
        gh_api("repos/%s/pulls/%s/requested_reviewers" % (repo, pr["number"]), "POST", {"reviewers": [login]})
    except ErrorGh as e:
        nota = "no pude pedirle la revisión a %s: %s" % (login, str(e)[:120])
    return pr["html_url"], nota


def main():
    ap = argparse.ArgumentParser(description="Abre un pull request de revisión por estudiante, con coautoría.")
    ap.add_argument("csv", help="CSV con las columnas equipo, repo (dueño/repositorio) y usuario")
    ap.add_argument("--ejecutar", action="store_true", help="crea de verdad; sin esto solo simula")
    ap.add_argument("--equipo", help="procesa solo las filas de este equipo (columna equipo)")
    ap.add_argument("--sin-preguntar", action="store_true", help="no pide confirmación antes de crear")
    ap.add_argument("--nombre-docente", help="tu nombre para la línea Co-authored-by (por defecto, el de tu perfil)")
    ap.add_argument("--correo-docente", help="tu correo no-reply (por defecto ID+usuario@users.noreply.github.com)")
    ap.add_argument("--sin-verificar-colaborador", action="store_true", help="no comprueba que la persona sea colaboradora del repositorio")
    ap.add_argument("--fecha-limite", help='texto que se agrega al pull request, por ejemplo "viernes 5 de diciembre, 23:45"')
    ap.add_argument("--pausa", type=float, default=2.0, help="segundos de espera entre pull requests (por defecto 2)")
    ap.add_argument("--registro", default="revisiones_abiertas.csv", help="CSV donde se anota qué pasó con cada fila")
    args = ap.parse_args()

    if not shutil.which("gh"):
        sys.exit("No encuentro GitHub CLI (gh). Instálalo desde https://cli.github.com y luego ejecuta: gh auth login")
    if subprocess.run(["gh", "auth", "status"], capture_output=True).returncode != 0:
        sys.exit("gh no tiene la sesión iniciada. Ejecuta: gh auth login")

    yo = gh_api("user")
    docente = {
        "nombre": limpiar(args.nombre_docente or yo.get("name") or yo["login"]),
        "correo": args.correo_docente or "%s+%s@users.noreply.github.com" % (yo["id"], yo["login"]),
    }
    print("Docente: %s <%s>" % (docente["nombre"], docente["correo"]))
    print("Compara ese correo con el que aparece en github.com/settings/emails; si es otro, usa --correo-docente.\n")

    filas = leer_filas(args.csv)
    if args.equipo:
        filas = [f for f in filas if f["equipo"].lower() == args.equipo.lower()]
    if not filas:
        sys.exit("No hay filas para procesar (revisa el CSV y el valor de --equipo).")

    listos, registro = [], []
    for fila in filas:
        etiqueta = "[%s] %s / %s" % (fila["equipo"] or "-", fila["repo"], fila["usuario"])
        try:
            r = preparar(fila, docente, args)
        except ErrorGh as e:
            r = {"omitir": "error de GitHub (HTTP %s): %s" % (e.estado, str(e)[:120])}
        if "plan" in r:
            p = r["plan"]
            print("LISTO    %s: rama %s, archivo %s, coautor %s" % (etiqueta, p["rama"], p["ruta"], p["coautor"]))
            for a in p["avisos"]:
                print("         aviso: " + a)
            listos.append((fila, p))
        else:
            print("OMITIDO  %s: %s" % (etiqueta, r["omitir"]))
            registro.append([fila["equipo"], fila["repo"], fila["usuario"], r.get("estado", "omitido"), r.get("url", ""), r["omitir"]])

    print("\n%d listos, %d omitidos." % (len(listos), len(registro)))

    if not args.ejecutar:
        print("Esto fue una simulación: no se creó nada. Para crear de verdad, agrega --ejecutar.")
    elif listos:
        if not args.sin_preguntar:
            if input("Se van a abrir %d pull requests. Escribe SI para continuar: " % len(listos)).strip().upper() != "SI":
                sys.exit("Cancelado: no se creó nada.")
        for i, (fila, p) in enumerate(listos):
            etiqueta = "[%s] %s / %s" % (fila["equipo"] or "-", fila["repo"], fila["usuario"])
            try:
                url, nota = ejecutar(p)
                print("ABIERTO  %s: %s%s" % (etiqueta, url, ("  (" + nota + ")") if nota else ""))
                registro.append([fila["equipo"], fila["repo"], fila["usuario"], "abierto", url, nota])
            except ErrorGh as e:
                print("ERROR    %s: %s" % (etiqueta, str(e)[:200]))
                registro.append([fila["equipo"], fila["repo"], fila["usuario"], "error", "", str(e)[:200]])
            if i < len(listos) - 1:
                time.sleep(args.pausa)
        with open(args.registro, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["equipo", "repo", "usuario", "estado", "pull_request", "nota"])
            w.writerows(registro)
        print("\nRegistro guardado en %s." % args.registro)
        print("Siguiente paso: abre el primer pull request, entra a Commits y confirma que aparece la foto de la persona junto a la tuya.")
        print("Si GitHub responde con un límite de uso, espera unos minutos y vuelve a correr el mismo comando: se salta lo ya abierto.")


if __name__ == "__main__":
    main()

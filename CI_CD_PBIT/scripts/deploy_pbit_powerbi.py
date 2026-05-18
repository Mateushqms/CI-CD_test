#!/usr/bin/env python3
"""Publica arquivos .pbit no workspace do Power BI via Service Principal."""

import os
import sys
import time
import requests

TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
PBI_BASE = "https://api.powerbi.com/v1.0/myorg"


def get_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    resp = requests.post(
        TOKEN_URL.format(tenant_id=tenant_id),
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": PBI_SCOPE,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def publish(token: str, workspace_id: str, pbit_path: str) -> None:
    name = os.path.splitext(os.path.basename(pbit_path))[0]
    url = (
        f"{PBI_BASE}/groups/{workspace_id}/imports"
        f"?datasetDisplayName={name}&nameConflict=CreateOrOverwrite"
    )

    with open(pbit_path, "rb") as f:
        resp = requests.post(
            url,
            headers={"Authorization": f"Bearer {token}"},
            files={"file": (os.path.basename(pbit_path), f, "application/octet-stream")},
            timeout=300,
        )

    resp.raise_for_status()
    import_id = resp.json()["id"]
    print(f"[{name}] Import iniciado (id: {import_id})")
    wait_for_import(token, workspace_id, import_id, name)


def wait_for_import(token: str, workspace_id: str, import_id: str, name: str) -> None:
    url = f"{PBI_BASE}/groups/{workspace_id}/imports/{import_id}"
    headers = {"Authorization": f"Bearer {token}"}

    for attempt in range(24):
        time.sleep(10)
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        state = resp.json().get("importState", "")
        print(f"[{name}] Estado: {state} (tentativa {attempt + 1}/24)")
        if state == "Succeeded":
            print(f"[{name}] Publicado com sucesso.")
            return
        if state == "Failed":
            raise RuntimeError(f"A API retornou falha: {resp.json()}")

    raise TimeoutError(f"[{name}] Timeout aguardando conclusão do import.")


def validate_env() -> tuple[str, str, str, str]:
    missing = [v for v in ("PBI_TENANT_ID", "PBI_CLIENT_ID", "PBI_CLIENT_SECRET", "PBI_WORKSPACE_ID") if not os.environ.get(v)]
    if missing:
        print(f"ERRO: variáveis de ambiente não configuradas: {', '.join(missing)}")
        print("Configure os secrets no GitHub: Settings → Secrets and variables → Actions")
        sys.exit(1)
    return (
        os.environ["PBI_TENANT_ID"],
        os.environ["PBI_CLIENT_ID"],
        os.environ["PBI_CLIENT_SECRET"],
        os.environ["PBI_WORKSPACE_ID"],
    )


def main(files: list[str]) -> None:
    pbit_files = [f.strip() for f in files if f.strip().endswith(".pbit")]

    if not pbit_files:
        print("Nenhum arquivo .pbit informado. Nada a publicar.")
        return

    tenant_id, client_id, client_secret, workspace_id = validate_env()

    print("Obtendo token de acesso...")
    token = get_token(tenant_id, client_id, client_secret)
    print("Token obtido com sucesso.")

    errors = []
    for path in pbit_files:
        if not os.path.isfile(path):
            print(f"Arquivo não encontrado, pulando: {path}")
            continue
        try:
            publish(token, workspace_id, path)
        except Exception as exc:
            print(f"ERRO ao publicar {path}: {exc}", file=sys.stderr)
            errors.append(path)

    if errors:
        print(f"\nFalha em {len(errors)} arquivo(s): {errors}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv[1:])

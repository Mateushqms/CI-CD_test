#!/usr/bin/env python3
"""Publica arquivos .pbix no workspace do Power BI do Orgão2 via REST API."""

import os
import sys
import time
import requests

AUTHORITY_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
PBI_BASE_URL = "https://api.powerbi.com/v1.0/myorg"
PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"


def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = AUTHORITY_URL.format(tenant_id=tenant_id)
    resp = requests.post(
        url,
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


def publish_pbix(token: str, workspace_id: str, pbix_path: str) -> None:
    dataset_name = os.path.splitext(os.path.basename(pbix_path))[0]
    url = (
        f"{PBI_BASE_URL}/groups/{workspace_id}/imports"
        f"?datasetDisplayName={dataset_name}&nameConflict=CreateOrOverwrite"
    )
    headers = {"Authorization": f"Bearer {token}"}

    with open(pbix_path, "rb") as f:
        resp = requests.post(
            url,
            headers=headers,
            files={"file": (os.path.basename(pbix_path), f, "application/octet-stream")},
            timeout=300,
        )

    resp.raise_for_status()
    import_id = resp.json().get("id")
    print(f"[{dataset_name}] Import iniciado — id: {import_id}")
    _wait_import(token, workspace_id, import_id, dataset_name)


def _wait_import(token: str, workspace_id: str, import_id: str, name: str) -> None:
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/imports/{import_id}"
    headers = {"Authorization": f"Bearer {token}"}

    for attempt in range(20):
        time.sleep(10)
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        state = resp.json().get("importState", "")
        print(f"[{name}] Estado: {state} (tentativa {attempt + 1})")
        if state == "Succeeded":
            print(f"[{name}] Publicado com sucesso.")
            return
        if state == "Failed":
            raise RuntimeError(f"[{name}] Falha na publicação: {resp.json()}")

    raise TimeoutError(f"[{name}] Timeout aguardando conclusão do import.")


def main(files: list[str]) -> None:
    tenant_id = os.environ["PBI_TENANT_ID"]
    client_id = os.environ["PBI_CLIENT_ID"]
    client_secret = os.environ["PBI_CLIENT_SECRET"]
    workspace_id = os.environ["PBI_WORKSPACE_ID"]

    if not files:
        print("Nenhum arquivo .pbix informado. Nada a publicar.")
        return

    print("Obtendo token de acesso...")
    token = get_access_token(tenant_id, client_id, client_secret)

    errors = []
    for path in files:
        path = path.strip()
        if not path:
            continue
        if not path.endswith(".pbix"):
            print(f"Ignorando (não é .pbix): {path}")
            continue
        if not os.path.isfile(path):
            print(f"Arquivo não encontrado, pulando: {path}")
            continue
        try:
            publish_pbix(token, workspace_id, path)
        except Exception as exc:
            print(f"ERRO ao publicar {path}: {exc}", file=sys.stderr)
            errors.append(path)

    if errors:
        print(f"\nFalha em {len(errors)} arquivo(s): {errors}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv[1:])

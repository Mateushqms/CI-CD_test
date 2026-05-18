#!/usr/bin/env python3
"""Publica arquivos .pbit no SharePoint via Microsoft Graph API."""

import os
import sys
import requests

GRAPH_TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

CHUNK_SIZE = 4 * 1024 * 1024  # 4 MB — limite para upload simples vs. sessão


def get_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    resp = requests.post(
        GRAPH_TOKEN_URL.format(tenant_id=tenant_id),
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": GRAPH_SCOPE,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def _simple_upload(headers: dict, url: str, file_path: str) -> None:
    with open(file_path, "rb") as f:
        resp = requests.put(
            url,
            headers={**headers, "Content-Type": "application/octet-stream"},
            data=f,
            timeout=120,
        )
    resp.raise_for_status()


def _session_upload(headers: dict, session_url: str, file_path: str) -> None:
    resp = requests.post(
        session_url,
        headers=headers,
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    resp.raise_for_status()
    upload_url = resp.json()["uploadUrl"]

    file_size = os.path.getsize(file_path)
    with open(file_path, "rb") as f:
        start = 0
        while True:
            chunk = f.read(CHUNK_SIZE)
            if not chunk:
                break
            end = start + len(chunk) - 1
            resp = requests.put(
                upload_url,
                headers={
                    "Content-Range": f"bytes {start}-{end}/{file_size}",
                    "Content-Length": str(len(chunk)),
                },
                data=chunk,
                timeout=120,
            )
            if resp.status_code not in (200, 201, 202):
                resp.raise_for_status()
            start += len(chunk)


def upload_file(
    token: str,
    site_id: str,
    drive_id: str,
    folder_path: str,
    file_path: str,
) -> None:
    filename = os.path.basename(file_path)
    destination = f"{folder_path.strip('/')}/{filename}"
    headers = {"Authorization": f"Bearer {token}"}

    base_url = f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/root:/{destination}"

    if os.path.getsize(file_path) <= CHUNK_SIZE:
        _simple_upload(headers, f"{base_url}:/content", file_path)
    else:
        _session_upload(headers, f"{base_url}:/createUploadSession", file_path)

    print(f"[{filename}] Publicado com sucesso em '{folder_path}'.")


def validate_env() -> tuple[str, str, str, str, str, str]:
    required = (
        "SP_TENANT_ID",
        "SP_CLIENT_ID",
        "SP_CLIENT_SECRET",
        "SP_SITE_ID",
        "SP_DRIVE_ID",
        "SP_FOLDER_PATH",
    )
    missing = [v for v in required if not os.environ.get(v)]
    if missing:
        print(f"ERRO: variáveis de ambiente não configuradas: {', '.join(missing)}")
        print("Configure os secrets no GitHub: Settings → Secrets and variables → Actions")
        sys.exit(1)
    return tuple(os.environ[v] for v in required)  # type: ignore[return-value]


def main(files: list[str]) -> None:
    pbit_files = [f.strip() for f in files if f.strip().endswith(".pbit")]

    if not pbit_files:
        print("Nenhum arquivo .pbit informado. Nada a publicar.")
        return

    tenant_id, client_id, client_secret, site_id, drive_id, folder_path = validate_env()

    print("Obtendo token de acesso...")
    token = get_token(tenant_id, client_id, client_secret)
    print("Token obtido com sucesso.")

    errors = []
    for path in pbit_files:
        if not os.path.isfile(path):
            print(f"Arquivo não encontrado, pulando: {path}")
            continue
        try:
            print(f"[{os.path.basename(path)}] Enviando para SharePoint...")
            upload_file(token, site_id, drive_id, folder_path, path)
        except Exception as exc:
            print(f"ERRO ao publicar {path}: {exc}", file=sys.stderr)
            errors.append(path)

    if errors:
        print(f"\nFalha em {len(errors)} arquivo(s): {errors}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv[1:])

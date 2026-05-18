#!/usr/bin/env python3
"""Publica arquivos .pbit no OneDrive pessoal via Microsoft Graph API."""

import os
import sys
import requests

SCOPES = "https://graph.microsoft.com/Files.ReadWrite offline_access"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

CHUNK_SIZE = 4 * 1024 * 1024  # 4 MB — limite upload simples vs. sessão


TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"


def get_token(client_id: str, refresh_token: str) -> str:
    resp = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "client_id": client_id,
            "refresh_token": refresh_token,
            "scope": SCOPES,
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


def upload_file(token: str, folder_path: str, file_path: str) -> None:
    filename = os.path.basename(file_path)
    destination = f"{folder_path.strip('/')}/{filename}"
    headers = {"Authorization": f"Bearer {token}"}

    base_url = f"{GRAPH_BASE}/me/drive/root:/{destination}"

    if os.path.getsize(file_path) <= CHUNK_SIZE:
        _simple_upload(headers, f"{base_url}:/content", file_path)
    else:
        _session_upload(headers, f"{base_url}:/createUploadSession", file_path)

    print(f"[{filename}] Publicado com sucesso em '{folder_path}'.")


def validate_env() -> tuple[str, str, str]:
    required = ("ONEDRIVE_CLIENT_ID", "ONEDRIVE_REFRESH_TOKEN", "ONEDRIVE_FOLDER_PATH")
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

    client_id, refresh_token, folder_path = validate_env()

    print("Obtendo token de acesso...")
    token = get_token(client_id, refresh_token)
    print("Token obtido com sucesso.")

    errors = []
    for path in pbit_files:
        if not os.path.isfile(path):
            print(f"Arquivo não encontrado, pulando: {path}")
            continue
        try:
            print(f"[{os.path.basename(path)}] Enviando para OneDrive...")
            upload_file(token, folder_path, path)
        except Exception as exc:
            print(f"ERRO ao publicar {path}: {exc}", file=sys.stderr)
            errors.append(path)

    if errors:
        print(f"\nFalha em {len(errors)} arquivo(s): {errors}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv[1:])

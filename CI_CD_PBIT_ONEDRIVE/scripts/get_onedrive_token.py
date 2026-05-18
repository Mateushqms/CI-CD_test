#!/usr/bin/env python3
"""
Obtém refresh_token do OneDrive pessoal via device code flow.

Execute UMA VEZ localmente:
    python get_onedrive_token.py

Siga as instruções, copie o refresh_token gerado e salve como
secret ONEDRIVE_REFRESH_TOKEN no GitHub.
"""

import time
import requests

SCOPES = "https://graph.microsoft.com/Files.ReadWrite offline_access"


DEVICE_CODE_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode"
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"


def get_device_code(client_id: str) -> dict:
    resp = requests.post(
        DEVICE_CODE_URL,
        data={"client_id": client_id, "scope": SCOPES},
        timeout=30,
    )
    if not resp.ok:
        raise RuntimeError(f"Erro ao solicitar device code: {resp.status_code} — {resp.json()}")
    return resp.json()


def poll_for_token(client_id: str, device_code: str, interval: int) -> dict:
    while True:
        time.sleep(interval)
        resp = requests.post(
            TOKEN_URL,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": client_id,
                "device_code": device_code,
            },
            timeout=30,
        )
        data = resp.json()

        if "access_token" in data:
            return data

        error = data.get("error", "")
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue

        raise RuntimeError(f"Erro ao obter token: {data}")


def main() -> None:
    print("=== Gerador de Refresh Token — OneDrive Pessoal ===\n")
    client_id = input("Cole o Client ID do seu App Registration: ").strip()

    print("\nSolicitando código de dispositivo...")
    device_data = get_device_code(client_id)

    print(f"\n{'='*60}")
    print(f"1. Acesse: {device_data['verification_uri']}")
    print(f"2. Digite o código: {device_data['user_code']}")
    print(f"{'='*60}\n")
    print("Aguardando autenticação no browser...")

    token_data = poll_for_token(client_id, device_data["device_code"], device_data["interval"])

    print("\n✓ Autenticado com sucesso!\n")
    print("=" * 60)
    print("Salve o valor abaixo como secret ONEDRIVE_REFRESH_TOKEN no GitHub:")
    print("=" * 60)
    print(token_data["refresh_token"])
    print("=" * 60)
    print("\nAtenção: o refresh_token expira se ficar sem uso por 90 dias.")


if __name__ == "__main__":
    main()

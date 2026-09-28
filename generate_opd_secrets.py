"""Run locally once: python generate_opd_secrets.py

Prints a new Fernet encryption key using only Python's standard library.
Does not contact a server, save a file, or generate an app password.
Keep the key in Streamlit Secrets and a secure backup. Do not regenerate it
when updating an app with an existing archive.
"""
import base64
import secrets

if __name__ == "__main__":
    key = base64.urlsafe_b64encode(secrets.token_bytes(32)).decode("ascii")
    print("# PRIVATE KEY: copy to Streamlit Secrets and a secure backup; NOT GitHub.")
    print(f'encryption_key = "{key}"')
    print("# Keep your existing encryption key when updating an existing archive.")

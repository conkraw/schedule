"""Run locally to print a new SECTION password. Does not create an encryption key."""
import secrets

if __name__ == "__main__":
    print("# Add to Streamlit Secrets; do not commit the printed value to GitHub.")
    print("# Keep your existing [opd_archive] settings and encryption_key unchanged.")
    print("[evaluation_access]")
    print(f'password = "{secrets.token_urlsafe(32)}"')

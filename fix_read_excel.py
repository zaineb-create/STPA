import re

with open("generate_dashboard_devicelogin.py", "r") as f:
    content = f.read()

old = '''def read_excel(token: str) -> pd.DataFrame:
    headers = {"Authorization": f"Bearer {token}"}

    # Récupérer l'ID du site
    site_url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{SP_SITE}:{SP_SITE_PATH}"
    )
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]
    print(f"  Site ID trouvé : {site_id.split(',')[1]}")

    # Télécharger le fichier Excel
    file_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{EXCEL_PATH}:/content"
    )
    file_resp = requests.get(file_url, headers=headers)
    file_resp.raise_for_status()

    df = pd.read_excel(io.BytesIO(file_resp.content),
                       sheet_name=SHEET_NAME, header=0)
    print(f"  {len(df)} lignes lues — feuille '{SHEET_NAME}'")
    return df'''

new = '''def read_excel(token: str) -> pd.DataFrame:
    print("  Téléchargement direct via SharePoint...")
    download_url = (
        "https://roseblanchetn.sharepoint.com/sites/SDAHSESTPA"
        "/_layouts/15/download.aspx"
        "?UniqueId=0761FA65-3D84-4B10-B009-8CA5BF050C98"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "*/*",
        "User-Agent": "Mozilla/5.0"
    }
    resp = requests.get(download_url, headers=headers, timeout=60, allow_redirects=True)
    if resp.status_code != 200 or b"<!DOCTYPE" in resp.content[:200]:
        raise Exception(f"Erreur {resp.status_code} — supprime .token_cache.json et relance")
    df = pd.read_excel(io.BytesIO(resp.content), sheet_name=SHEET_NAME, header=0)
    print(f"  {len(df)} lignes lues — feuille '{SHEET_NAME}'")
    return df'''

content = content.replace(old, new)

with open("generate_dashboard_devicelogin.py", "w") as f:
    f.write(content)

print("OK — fonction read_excel remplacée")

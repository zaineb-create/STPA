"""
Lancer ce script UNE SEULE FOIS sur ton PC pour obtenir le token Microsoft.
Il va afficher un JSON à copier dans les variables d'environnement Render.

    pip install msal
    python get_token.py
"""

import json, os, webbrowser
import msal

CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES    = ["https://graph.microsoft.com/Sites.Read.All",
             "https://graph.microsoft.com/Files.ReadWrite.All"]

cache = msal.SerializableTokenCache()

app = msal.PublicClientApplication(
    client_id=CLIENT_ID,
    authority=AUTHORITY,
    token_cache=cache
)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    raise Exception(f"Erreur : {flow}")

print("\n" + "="*55)
print("  CONNEXION MICROSOFT REQUISE")
print("="*55)
print(f"\n  1. Ouvre : https://microsoft.com/devicelogin")
print(f"  2. Entre ce code : {flow['user_code']}")
print(f"  3. Connecte-toi avec ton compte entreprise")
print("\n  En attente...\n")

try:
    webbrowser.open("https://microsoft.com/devicelogin")
except:
    pass

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise Exception(f"Échec : {result.get('error_description', result)}")

print("  ✅ Connecté avec succès !")
print("\n" + "="*55)
print("  COPIE CE TOKEN DANS RENDER → Environment Variables")
print("="*55)
print(f"\n  Clé   : TOKEN_CACHE")
print(f"  Valeur: {cache.serialize()}")
print("\n" + "="*55)

# Sauvegarde aussi en fichier local
with open("token_cache_export.json", "w") as f:
    f.write(cache.serialize())
print("\n  Aussi sauvegardé dans : token_cache_export.json")

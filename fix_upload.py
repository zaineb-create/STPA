with open("generate_dashboard_devicelogin.py", "r") as f:
    content = f.read()

old = '''    print("\\n[5/5] Upload sur SharePoint...")
    url = upload_html(token, html)

    print("\\n" + "="*55)
    print("  TERMINÉ !")
    print("="*55)
    print(f"\\n  Anomalies : {payload['total_anomalies']}")
    print(f"  Analyses  : {payload['total_analyses']}")
    print(f"  Généré le : {payload['generated_at']}")
    print(f"\\n  URL Power Apps → Web viewer :")
    print(f\'  "{url}"\')
    print()'''

new = '''    print("\\n[5/5] Sauvegarde locale + GitHub Pages...")
    with open("dashboard_ssse.html", "w", encoding="utf-8") as f:
        f.write(html)
    import subprocess
    subprocess.run(["git", "add", "dashboard_ssse.html"])
    subprocess.run(["git", "commit", "-m", "Update dashboard SSSE"])
    subprocess.run(["git", "push"])

    print("\\n" + "="*55)
    print("  TERMINÉ !")
    print("="*55)
    print(f"\\n  Anomalies : {payload[\'total_anomalies\']}")
    print(f"  Analyses  : {payload[\'total_analyses\']}")
    print(f"  Généré le : {payload[\'generated_at\']}")
    print(f"\\n  URL Power Apps → Web viewer :")
    print(f\'  "https://zaineb-create.github.io/STPA/dashboard_ssse.html"\')
    print()'''

content = content.replace(old, new)

with open("generate_dashboard_devicelogin.py", "w") as f:
    f.write(content)

print("OK — upload remplacé par GitHub Pages")

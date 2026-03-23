with open("generate_dashboard_devicelogin.py", "r") as f:
    content = f.read()

old_imports = "import requests\nimport pandas as pd\nimport plotly.graph_objects as go\nimport plotly.io as pio\nimport msal"
new_imports = """import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
import msal"""
content = content.replace(old_imports, new_imports)
content = content.replace('OUTPUT_HTML  = "dashboard_ssse.html"', 'OUTPUT_EXCEL = "anomalies_ssse.xlsx"')

with open("generate_dashboard_devicelogin.py", "w") as f:
    f.write(content)
print("OK")

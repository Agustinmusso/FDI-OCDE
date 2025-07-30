# FDI-OCDE
import requests
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import StringIO
from statsmodels.tsa.seasonal import seasonal_decompose
from openpyxl.drawing.image import Image
from openpyxl import Workbook

# ===============================
# CONFIGURACIÓN
# ===============================
COUNTRIES = ["BRA", "ARG", "MEX", "CHN", "DEU", "USA"]
START = "2010-Q1"

BASE = "https://sdmx.oecd.org/public/rest/data/OECD.DAF.INV,DSD_FDI@DF_FDI_AGGR_SUMM"
KEY = "all"  # traemos todo y filtramos en pandas

PARAMS = {
    "startPeriod": START,
    "dimensionAtObservation": "AllDimensions",
    "format": "csvfilewithlabels"
}

output_file = r"C:/Users/HP/Desktop/Consultora/Consultorías/Global Ports/Bases/Gráficos/FDI_Graficos.xlsx"

# ===============================
# DESCARGA DE DATOS
# ===============================
local_file = r"C:/Users/HP/Desktop/Consultora/Consultorías/Global Ports/Bases/OCDE/FDI_data.csv"
raw = pd.read_csv(local_file, low_memory=False)

# ===============================
# FILTRADO
# ===============================
print("Filtrando datos...")
df = raw[
    (raw["REF_AREA"].isin(COUNTRIES))
    & (raw["FDI_COMP"] == "T_FA_F")
    & (raw["MEASURE"] == "USD_EXC_RC")
    & (raw["FREQ"] == "Q")
]

df = df.rename(columns={"OBS_VALUE": "FDI_inflows"})
df["TIME_PERIOD"] = pd.PeriodIndex(df["TIME_PERIOD"], freq="Q")
df = df.set_index(["TIME_PERIOD", "REF_AREA"])
df = df["FDI_inflows"].unstack().sort_index()  # Esto aplana la MultiIndex

# ===============================
# CRECIMIENTOS
# ===============================
yoy = df.pct_change(periods=4) * 100

# Ajuste estacional
print("Aplicando ajuste estacional (seasonal_decompose)...")
df_sa = pd.DataFrame(index=df.index)
for col in df.columns:
    ts = df[col].dropna().to_timestamp()
    if len(ts) >= 8:
        result = seasonal_decompose(ts, model="additive", period=4, extrapolate_trend="freq")
        sa = ts - result.seasonal
        df_sa[col] = sa.to_period("Q")
    else:
        df_sa[col] = df[col]

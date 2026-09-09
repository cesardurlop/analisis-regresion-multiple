"""
Anexo E - Regresión lineal múltiple para precios de metales
Casa de Moneda de México

Periodo histórico esperado: 2015-01 a 2024-12 (mensual)

Dependientes:
    aluminio, cobre, niquel, zinc

Explicativas:
    tipo_cambio, petroleo, industria

Modelo principal:
    ΔMetal_t = b0 + b1 ΔTC_t + b2 ΔPetroleo_t + b3 ΔIndustria_t + error_t

El script:
- lee los Excel de la carpeta actual;
- unifica las series;
- ejecuta diagnóstico y gráficos;
- valida fuera de muestra 2023-2024;
- ajusta el modelo final 2015-2024;
- genera un escenario 2025 para las X o usa escenario_2025.xlsx;
- reconstruye el nivel mensual del precio de cada metal en 2025.
"""

from __future__ import annotations

import math
import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.stattools import durbin_watson
from statsmodels.tsa.stattools import adfuller
from statsmodels.tsa.holtwinters import Holt

from scipy.stats import jarque_bera
from sklearn.metrics import mean_absolute_error, mean_squared_error


# ============================================================
# CONFIGURACIÓN
# ============================================================

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "resultados"
PLOTS = OUT / "graficas"

OUT.mkdir(exist_ok=True)
PLOTS.mkdir(exist_ok=True)

START = "2015-01-01"
END = "2024-12-01"
TEST_START = "2023-01-01"

METALS = ["aluminio", "cobre", "niquel", "zinc"]
X_COLS = ["tipo_cambio", "petroleo", "industria"]


FILE_PATTERNS = {
    "aluminio": [r"^aluminio.*\.xlsx$"],
    "cobre": [r"^cobre.*\.xlsx$"],
    "niquel": [r"^niquel.*\.xlsx$", r"^níquel.*\.xlsx$"],
    "zinc": [r"^zinc.*\.xlsx$"],
    "industria": [r"^industria.*\.xlsx$"],
    "tipo_cambio": [
        r"^tipo.*cambio.*\.xlsx$",
        r"^tc.*\.xlsx$",
    ],
    "petroleo": [
        r"^petroleo.*\.xlsx$",
        r"^petróleo.*\.xlsx$",
        r"^wti.*\.xlsx$",
    ],
}


# ============================================================
# UTILIDADES
# ============================================================

def normalize_text(text: str) -> str:
    text = str(text).strip().lower()
    text = "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )
    text = re.sub(r"\s+", " ", text)
    return text


def locate_file(series_name: str) -> Path:
    files = list(ROOT.glob("*.xlsx"))

    for pattern in FILE_PATTERNS[series_name]:
        rx = re.compile(pattern, flags=re.IGNORECASE)
        for f in files:
            if f.name.lower() == "escenario_2025.xlsx":
                continue
            if rx.search(normalize_text(f.name)):
                return f

    raise FileNotFoundError(
        f"\nNo encontré el Excel para '{series_name}'.\n"
        f"Archivos .xlsx disponibles: {[f.name for f in files]}\n"
        f"Revisa FILE_PATTERNS en main.py si utilizaste otro nombre."
    )


def detect_date_column(df: pd.DataFrame) -> str:
    candidates = []
    for c in df.columns:
        name = normalize_text(c)
        if any(k in name for k in ["fecha", "date", "periodo", "mes"]):
            candidates.append(c)

    if candidates:
        return candidates[0]

    # Como respaldo, busca la primera columna que convierta bien a fecha.
    for c in df.columns:
        parsed = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
        if parsed.notna().mean() >= 0.80:
            return c

    raise ValueError(
        "No fue posible detectar la columna de fecha. "
        f"Columnas encontradas: {list(df.columns)}"
    )


def detect_value_column(df: pd.DataFrame, date_col: str) -> str:
    preferred_words = [
        "precio", "valor", "indice", "índice", "total",
        "tipo de cambio", "petroleo", "petróleo", "imai"
    ]

    # 1) Nombres preferidos.
    for word in preferred_words:
        for c in df.columns:
            if c == date_col:
                continue
            if word in normalize_text(c):
                numeric = pd.to_numeric(df[c], errors="coerce")
                if numeric.notna().mean() >= 0.70:
                    return c

    # 2) Primera columna mayormente numérica.
    for c in df.columns:
        if c == date_col:
            continue
        numeric = pd.to_numeric(df[c], errors="coerce")
        if numeric.notna().mean() >= 0.70:
            return c

    raise ValueError(
        "No fue posible detectar una columna numérica de valor. "
        f"Columnas encontradas: {list(df.columns)}"
    )


def read_monthly_series(path: Path, output_name: str) -> pd.Series:
    """
    Lee exclusivamente la primera hoja.
    Detecta fecha y valor automáticamente.
    """
    df = pd.read_excel(path, sheet_name=0)

    # Elimina columnas completamente vacías.
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    date_col = detect_date_column(df)
    value_col = detect_value_column(df, date_col)

    out = df[[date_col, value_col]].copy()
    out.columns = ["fecha", output_name]

    out["fecha"] = pd.to_datetime(out["fecha"], errors="coerce", dayfirst=True)
    out[output_name] = pd.to_numeric(out[output_name], errors="coerce")
    out = out.dropna(subset=["fecha", output_name])

    # Lleva todas las fechas al primer día del mes.
    out["fecha"] = out["fecha"].dt.to_period("M").dt.to_timestamp()

    if out["fecha"].duplicated().any():
        dup = out.loc[out["fecha"].duplicated(keep=False), "fecha"].tolist()
        raise ValueError(
            f"{path.name}: existen meses duplicados. Ejemplos: {dup[:5]}"
        )

    out = out.set_index("fecha").sort_index()[output_name]
    out.name = output_name

    return out


def load_all_series() -> pd.DataFrame:
    series = {}

    print("\nARCHIVOS DETECTADOS")
    print("=" * 70)

    for name in METALS + X_COLS:
        path = locate_file(name)
        print(f"{name:15s} -> {path.name}")
        series[name] = read_monthly_series(path, name)

    base = pd.concat(series.values(), axis=1).sort_index()
    base = base.loc[START:END].copy()

    expected = pd.date_range(START, END, freq="MS")
    base = base.reindex(expected)
    base.index.name = "fecha"

    return base


def validate_base(df: pd.DataFrame) -> None:
    print("\nVALIDACIÓN DE LA BASE")
    print("=" * 70)
    print(f"Periodo: {df.index.min().date()} a {df.index.max().date()}")
    print(f"Filas:   {len(df)}")
    print("\nFaltantes por columna:")
    print(df.isna().sum())

    if len(df) != 120:
        print(
            f"\nADVERTENCIA: se esperaban 120 meses, pero hay {len(df)}."
        )

    missing = df.isna().sum()
    if (missing > 0).any():
        cols = missing[missing > 0].to_dict()
        raise ValueError(
            "\nLa base unificada contiene datos faltantes. "
            f"Faltantes: {cols}\n"
            "Corrige los Excel antes de continuar para evitar "
            "comparaciones con distintos periodos."
        )


def adf_result(s: pd.Series) -> dict:
    s = s.dropna()
    stat, pvalue, lags, nobs, crit, _ = adfuller(s, autolag="AIC")
    return {
        "ADF": stat,
        "p_value": pvalue,
        "lags": lags,
        "nobs": nobs,
        "stationary_5pct": pvalue < 0.05,
    }


def mape(y_true, y_pred) -> float:
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    mask = np.abs(y_true) > 1e-12
    if mask.sum() == 0:
        return np.nan
    return np.mean(
        np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])
    ) * 100


def metrics(y_true, y_pred) -> dict:
    return {
        "MAE": mean_absolute_error(y_true, y_pred),
        "RMSE": math.sqrt(mean_squared_error(y_true, y_pred)),
        "MAPE_pct": mape(y_true, y_pred),
    }


def calculate_vif(X: pd.DataFrame) -> pd.DataFrame:
    Xc = sm.add_constant(X)
    rows = []

    for i, col in enumerate(Xc.columns):
        if col == "const":
            continue
        rows.append({
            "variable": col,
            "VIF": variance_inflation_factor(Xc.values, i)
        })

    return pd.DataFrame(rows)


# ============================================================
# GRÁFICAS
# ============================================================

def save_series_plots(base: pd.DataFrame):
    for col in base.columns:
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(base.index, base[col])
        ax.set_title(f"Serie mensual: {col}")
        ax.set_xlabel("Fecha")
        ax.set_ylabel(col)
        ax.grid(alpha=0.25)
        fig.tight_layout()
        fig.savefig(PLOTS / f"serie_{col}.png", dpi=160)
        plt.close(fig)


def save_correlation_plot(df_diff: pd.DataFrame):
    corr = df_diff.corr()

    fig, ax = plt.subplots(figsize=(9, 7))
    im = ax.imshow(corr.values, aspect="auto")

    ax.set_xticks(range(len(corr.columns)))
    ax.set_xticklabels(corr.columns, rotation=45, ha="right")
    ax.set_yticks(range(len(corr.columns)))
    ax.set_yticklabels(corr.columns)

    for i in range(len(corr.columns)):
        for j in range(len(corr.columns)):
            ax.text(
                j, i, f"{corr.iloc[i, j]:.2f}",
                ha="center", va="center", fontsize=8
            )

    ax.set_title("Correlaciones de primeras diferencias")
    fig.colorbar(im, ax=ax)
    fig.tight_layout()
    fig.savefig(PLOTS / "correlaciones_diferencias.png", dpi=160)
    plt.close(fig)


def save_model_plots(
    metal: str,
    y_test_level: pd.Series,
    pred_test_level: pd.Series,
    model
):
    # Real vs pronosticado
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(y_test_level.index, y_test_level.values, label="Real")
    ax.plot(pred_test_level.index, pred_test_level.values, label="Regresión")
    ax.set_title(f"{metal.capitalize()}: validación 2023-2024")
    ax.set_xlabel("Fecha")
    ax.set_ylabel("Precio")
    ax.legend()
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(PLOTS / f"{metal}_validacion.png", dpi=160)
    plt.close(fig)

    # Residuos
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(model.resid.index, model.resid.values)
    ax.axhline(0, linewidth=1)
    ax.set_title(f"{metal.capitalize()}: residuos del modelo")
    ax.set_xlabel("Fecha")
    ax.set_ylabel("Residuo")
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(PLOTS / f"{metal}_residuos.png", dpi=160)
    plt.close(fig)

    # QQ plot
    fig = plt.figure(figsize=(6, 6))
    ax = fig.add_subplot(111)
    sm.qqplot(model.resid, line="45", ax=ax, fit=True)
    ax.set_title(f"{metal.capitalize()}: QQ plot de residuos")
    fig.tight_layout()
    fig.savefig(PLOTS / f"{metal}_qqplot.png", dpi=160)
    plt.close(fig)


# ============================================================
# MODELOS
# ============================================================

def build_differences(base: pd.DataFrame) -> pd.DataFrame:
    """
    Primera diferencia simple.

    Se utiliza para mantener consistencia con los anexos,
    donde las series de precios alcanzan estacionariedad
    después de diferenciar.
    """
    return base.diff().dropna()


def fit_regression(train_diff: pd.DataFrame, metal: str):
    y = train_diff[metal]
    X = sm.add_constant(train_diff[X_COLS])
    return sm.OLS(y, X).fit()


def reconstruct_level_from_differences(
    last_level: float,
    predicted_diffs: pd.Series
) -> pd.Series:
    values = []
    current = float(last_level)

    for d in predicted_diffs:
        current += float(d)
        values.append(current)

    return pd.Series(
        values,
        index=predicted_diffs.index,
        name=predicted_diffs.name
    )


def backtest_metal(base: pd.DataFrame, df_diff: pd.DataFrame, metal: str):
    train_diff = df_diff.loc[df_diff.index < TEST_START].copy()
    test_diff = df_diff.loc[df_diff.index >= TEST_START].copy()

    model = fit_regression(train_diff, metal)

    X_test = sm.add_constant(
        test_diff[X_COLS],
        has_constant="add"
    )

    pred_diff = pd.Series(
        model.predict(X_test),
        index=test_diff.index,
        name=f"{metal}_pred_diff"
    )

    # Nivel inmediatamente anterior al periodo de prueba.
    last_train_level = base.loc[
        base.index < pd.Timestamp(TEST_START), metal
    ].iloc[-1]

    pred_level = reconstruct_level_from_differences(
        last_train_level,
        pred_diff
    )

    y_test_level = base.loc[pred_level.index, metal]

    m = metrics(y_test_level, pred_level)

    # Diagnósticos del modelo de entrenamiento.
    vif = calculate_vif(train_diff[X_COLS])

    jb = jarque_bera(model.resid)
    bp = het_breuschpagan(
        model.resid,
        model.model.exog
    )

    diagnostics = {
        "R2_train": model.rsquared,
        "R2_adj_train": model.rsquared_adj,
        "F_pvalue": model.f_pvalue,
        "Durbin_Watson": durbin_watson(model.resid),
        "Jarque_Bera_pvalue": jb.pvalue,
        "Breusch_Pagan_pvalue": bp[1],
    }

    save_model_plots(
        metal,
        y_test_level,
        pred_level,
        model
    )

    return model, m, diagnostics, vif, pred_level


def forecast_x_holt(base: pd.DataFrame, periods: int = 12) -> pd.DataFrame:
    """
    Escenario automático 2025 para X.

    Holt amortiguado se usa aquí como generador de escenario,
    no como sustituto de la regresión del metal.
    """
    future_index = pd.date_range(
        base.index.max() + pd.offsets.MonthBegin(1),
        periods=periods,
        freq="MS"
    )

    forecasts = {}

    for col in X_COLS:
        s = base[col].dropna()
        fit = Holt(
            s,
            damped_trend=True,
            initialization_method="estimated"
        ).fit(optimized=True)

        fc = fit.forecast(periods)
        fc.index = future_index
        forecasts[col] = fc

    return pd.DataFrame(forecasts, index=future_index)


def load_or_create_2025_scenario(base: pd.DataFrame) -> pd.DataFrame:
    scenario_file = ROOT / "escenario_2025.xlsx"

    if not scenario_file.exists():
        print(
            "\nNo existe escenario_2025.xlsx. "
            "Se proyectarán las variables explicativas con Holt amortiguado."
        )
        scenario = forecast_x_holt(base, 12)
        scenario.to_excel(
            OUT / "escenario_2025_generado.xlsx",
            index_label="Fecha"
        )
        return scenario

    print("\nSe utilizará escenario_2025.xlsx.")

    raw = pd.read_excel(scenario_file, sheet_name=0)
    date_col = detect_date_column(raw)

    data = raw.copy()
    data[date_col] = pd.to_datetime(
        data[date_col],
        errors="coerce",
        dayfirst=True
    )
    data[date_col] = (
        data[date_col]
        .dt.to_period("M")
        .dt.to_timestamp()
    )
    data = data.set_index(date_col).sort_index()

    # Intento de mapeo flexible por nombre.
    rename = {}
    for c in data.columns:
        n = normalize_text(c)
        if "cambio" in n or n in {"tc", "usd/mxn", "usdmxn"}:
            rename[c] = "tipo_cambio"
        elif "petrole" in n or "wti" in n:
            rename[c] = "petroleo"
        elif (
            "industr" in n or
            "imai" in n or
            n == "total"
        ):
            rename[c] = "industria"

    data = data.rename(columns=rename)

    missing_cols = [c for c in X_COLS if c not in data.columns]
    if missing_cols:
        raise ValueError(
            "escenario_2025.xlsx no contiene las columnas requeridas: "
            f"{missing_cols}"
        )

    scenario = data[X_COLS].apply(
        pd.to_numeric,
        errors="coerce"
    )

    expected = pd.date_range(
        "2025-01-01",
        "2025-12-01",
        freq="MS"
    )
    scenario = scenario.reindex(expected)

    if scenario.isna().any().any():
        raise ValueError(
            "escenario_2025.xlsx debe contener 12 meses completos "
            "de enero a diciembre de 2025."
        )

    return scenario


def fit_final_and_forecast(
    base: pd.DataFrame,
    df_diff: pd.DataFrame,
    scenario_2025: pd.DataFrame,
    metal: str
):
    model = fit_regression(df_diff, metal)

    # Para utilizar el modelo en diferencias necesitamos las diferencias
    # de las X entre dic-2024 y cada mes de 2025.
    x_full = pd.concat(
        [
            base[X_COLS].iloc[[-1]],
            scenario_2025[X_COLS]
        ]
    )
    x_diff_2025 = x_full.diff().iloc[1:]

    X_future = sm.add_constant(
        x_diff_2025,
        has_constant="add"
    )

    pred_diff = pd.Series(
        model.predict(X_future),
        index=scenario_2025.index,
        name=f"{metal}_pred_diff"
    )

    last_level = base[metal].iloc[-1]

    pred_level = reconstruct_level_from_differences(
        last_level,
        pred_diff
    )
    pred_level.name = metal

    return model, pred_level, pred_diff


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():
    print("\nANEXO E - REGRESIÓN MÚLTIPLE DE METALES")
    print("=" * 70)

    base = load_all_series()
    validate_base(base)

    base.to_csv(
        OUT / "base_unificada.csv",
        encoding="utf-8-sig"
    )

    print("\nESTADÍSTICAS DESCRIPTIVAS")
    print("=" * 70)
    print(base.describe().T)

    save_series_plots(base)

    # ---------------------------
    # ADF
    # ---------------------------
    adf_rows = []

    for col in base.columns:
        level = adf_result(base[col])
        diff = adf_result(base[col].diff().dropna())

        adf_rows.append({
            "variable": col,
            "ADF_nivel": level["ADF"],
            "p_nivel": level["p_value"],
            "estacionaria_nivel_5pct": level["stationary_5pct"],
            "ADF_diff1": diff["ADF"],
            "p_diff1": diff["p_value"],
            "estacionaria_diff1_5pct": diff["stationary_5pct"],
        })

    adf_df = pd.DataFrame(adf_rows)
    adf_df.to_csv(
        OUT / "pruebas_adf.csv",
        index=False,
        encoding="utf-8-sig"
    )

    print("\nPRUEBAS ADF")
    print("=" * 70)
    print(adf_df.to_string(index=False))

    # ---------------------------
    # Diferencias y correlaciones
    # ---------------------------
    df_diff = build_differences(base)
    df_diff.to_csv(
        OUT / "base_diferencias.csv",
        encoding="utf-8-sig"
    )

    save_correlation_plot(df_diff)

    print("\nCORRELACIONES EN PRIMERAS DIFERENCIAS")
    print("=" * 70)
    print(df_diff.corr().round(3))

    # ---------------------------
    # Backtesting
    # ---------------------------
    metric_rows = []
    summary_blocks = []
    all_vif = []

    for metal in METALS:
        print(f"\n{'=' * 70}")
        print(f"MODELO: {metal.upper()}")
        print("=" * 70)

        model, met, diag, vif, pred = backtest_metal(
            base, df_diff, metal
        )

        print(model.summary())
        print("\nMétricas 2023-2024:")
        for k, v in met.items():
            print(f"  {k}: {v:.6f}")

        print("\nDiagnósticos:")
        for k, v in diag.items():
            print(f"  {k}: {v:.6f}")

        print("\nVIF:")
        print(vif.to_string(index=False))

        row = {
            "metal": metal,
            **met,
            **diag
        }
        metric_rows.append(row)

        vif2 = vif.copy()
        vif2.insert(0, "metal", metal)
        all_vif.append(vif2)

        summary_blocks.append(
            f"\n\n{'=' * 80}\n"
            f"{metal.upper()}\n"
            f"{'=' * 80}\n"
            f"{model.summary().as_text()}\n"
            f"\nMétricas fuera de muestra 2023-2024:\n"
            f"{met}\n"
            f"\nDiagnósticos:\n"
            f"{diag}\n"
            f"\nVIF:\n{vif.to_string(index=False)}\n"
        )

    metrics_df = pd.DataFrame(metric_rows)
    metrics_df.to_csv(
        OUT / "metricas_modelos.csv",
        index=False,
        encoding="utf-8-sig"
    )

    pd.concat(all_vif, ignore_index=True).to_csv(
        OUT / "vif_modelos.csv",
        index=False,
        encoding="utf-8-sig"
    )

    (OUT / "resumen_modelos.txt").write_text(
        "".join(summary_blocks),
        encoding="utf-8"
    )

    # ---------------------------
    # Escenario X 2025
    # ---------------------------
    scenario = load_or_create_2025_scenario(base)
    scenario.to_csv(
        OUT / "escenario_2025_utilizado.csv",
        encoding="utf-8-sig"
    )

    # ---------------------------
    # Pronóstico final 2025
    # ---------------------------
    forecast_levels = {}
    forecast_diffs = {}

    for metal in METALS:
        final_model, pred_level, pred_diff = fit_final_and_forecast(
            base,
            df_diff,
            scenario,
            metal
        )
        forecast_levels[metal] = pred_level
        forecast_diffs[metal] = pred_diff

    forecast_2025 = pd.DataFrame(forecast_levels)
    forecast_2025.index.name = "Fecha"

    forecast_2025.to_csv(
        OUT / "pronosticos_2025.csv",
        encoding="utf-8-sig"
    )

    pd.DataFrame(forecast_diffs).to_csv(
        OUT / "pronosticos_diferencias_2025.csv",
        encoding="utf-8-sig"
    )

    # Gráficas 2025
    for metal in METALS:
        fig, ax = plt.subplots(figsize=(10, 5))

        hist = base[metal].loc["2022-01-01":]
        ax.plot(hist.index, hist.values, label="Histórico")
        ax.plot(
            forecast_2025.index,
            forecast_2025[metal].values,
            label="Regresión múltiple 2025"
        )

        ax.set_title(
            f"{metal.capitalize()}: histórico y pronóstico 2025"
        )
        ax.set_xlabel("Fecha")
        ax.set_ylabel("Precio")
        ax.legend()
        ax.grid(alpha=0.25)
        fig.tight_layout()
        fig.savefig(
            PLOTS / f"{metal}_pronostico_2025.png",
            dpi=160
        )
        plt.close(fig)

    print("\n" + "=" * 70)
    print("RESULTADOS FUERA DE MUESTRA")
    print("=" * 70)
    print(
        metrics_df[
            ["metal", "MAE", "RMSE", "MAPE_pct", "R2_adj_train"]
        ].to_string(index=False)
    )

    print("\n" + "=" * 70)
    print("PRONÓSTICOS DE REGRESIÓN 2025")
    print("=" * 70)
    print(forecast_2025.round(6).to_string())

    print("\nArchivos generados en:")
    print(OUT.resolve())
    print(
        "\nSiguiente paso metodológico: comparar el RMSE anterior "
        "contra el RMSE ARIMA/Holt de los Anexos A-D."
    )


if __name__ == "__main__":
    main()

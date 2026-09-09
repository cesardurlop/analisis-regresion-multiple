Regresión múltiple para precios de metales — Casa de Moneda de México

Este repositorio contiene el análisis de regresión lineal múltiple que complementa los modelos ARIMA/Holt de los Anexos A–D del trabajo de titulación.

Objetivo

Evaluar si variables macroeconómicas mensuales ayudan a explicar y pronosticar los movimientos de los precios de:

Aluminio

Cobre

Níquel

Zinc

Variables explicativas:

Tipo de cambio USD/MXN

Precio del petróleo (WTI o la serie elegida)

Índice de actividad industrial de México (IMAI, serie desestacionalizada TOTAL)

El periodo histórico esperado es enero de 2015 a diciembre de 2024, con 120 observaciones mensuales por serie.

Estructura esperada del repositorio

.
├── main.py
├── README.md
├── aluminio.xlsx
├── cobre.xlsx
├── niquel.xlsx
├── zinc.xlsx
├── Industria 2015-2024.xlsx
├── Tipo de cambio 2025-2024.xlsx
└── Petroleo 2015-2024.xlsx

El nombre Tipo de cambio 2025-2024.xlsx se conserva porque aparece así en el repositorio mostrado.
Si lo renombras a Tipo de cambio 2015-2024.xlsx, main.py también intentará detectarlo.

Cada Excel debe tener en su primera hoja al menos dos columnas:

Fecha | Valor

En los archivos de metales puede llamarse:

Fecha | Precio

El script intenta detectar automáticamente la columna de fecha y la columna numérica principal.

Qué hace main.py

El script ejecuta el análisis completo:

Busca y lee automáticamente los siete archivos Excel.

Normaliza las fechas a frecuencia mensual.

Une todas las series por fecha.

Verifica:

rango de fechas;

observaciones faltantes;

duplicados;

valores no numéricos.

Genera estadísticas descriptivas.

Grafica las series históricas.

Calcula una matriz de correlaciones.

Ejecuta pruebas ADF de estacionariedad.

Construye variables transformadas mediante primera diferencia.

Ajusta, para cada metal, una regresión múltiple:

Δ Metal_t = β0
          + β1 Δ USD/MXN_t
          + β2 Δ Petróleo_t
          + β3 Δ IMAI_t
          + ε_t

Reporta:

coeficientes;

errores estándar;

valores p;

R²;

R² ajustado;

F-statistic;

Durbin-Watson;

VIF;

prueba de normalidad de residuos;

prueba de heterocedasticidad de Breusch-Pagan.

Realiza validación fuera de muestra:

entrenamiento: 2015–2022;

prueba: 2023–2024.

Calcula:

MAE;

RMSE;

MAPE.

Genera gráficas de:

real vs estimado;

residuos;

QQ plot;

correlaciones;

validación fuera de muestra.

Ajusta nuevamente el modelo con 2015–2024.

Proyecta las variables explicativas para 2025 con Holt amortiguado solo para poder producir un escenario base de 12 meses.

Con esas variables proyectadas obtiene un pronóstico de regresión para cada metal en 2025.

Exporta resultados a:

resultados/base_unificada.csv

resultados/metricas_modelos.csv

resultados/pronosticos_2025.csv

resultados/resumen_modelos.txt

carpeta resultados/graficas/

Importante sobre el pronóstico 2025

La regresión múltiple necesita valores 2025 de las variables explicativas.

El script ofrece dos opciones:

Opción A — Escenario automático

Si NO existe escenario_2025.xlsx, el programa proyecta:

USD/MXN

petróleo

IMAI

con Holt amortiguado y utiliza esos valores como escenario base.

Esto permite generar una cifra final reproducible, pero debe describirse como:

“escenario estadístico base condicionado a la proyección de las variables explicativas”.

Opción B — Escenario proporcionado

La opción metodológicamente más limpia para el presupuesto es crear:

escenario_2025.xlsx

con 12 filas:

Fecha | tipo_cambio | petroleo | industria

de enero a diciembre de 2025.

Si ese archivo existe, main.py lo utilizará en lugar de proyectar automáticamente las variables explicativas.

Esto permite usar los supuestos presupuestarios que decidas justificar.

Instalación

Se recomienda Python 3.11 o superior.

python -m venv .venv

Windows

.venv\Scripts\activate

macOS / Linux

source .venv/bin/activate

Instala dependencias:

pip install pandas numpy matplotlib openpyxl statsmodels scikit-learn scipy

Ejecución

python main.py

Al terminar, revisa la carpeta:

resultados/

Interpretación para el trabajo de titulación

La regresión múltiple no sustituye automáticamente a ARIMA/Holt.

Para cada metal se debe comparar el error fuera de muestra de:

modelo ARIMA/Holt del Anexo correspondiente;

modelo de regresión múltiple del Anexo E.

La regla sugerida es:

Seleccionar para fines presupuestales el modelo con menor RMSE fuera de muestra, siempre que el modelo presente diagnósticos estadísticos razonables y que su interpretación sea coherente.

La tabla final del Anexo E puede tener esta estructura:

Metal

Modelo univariado

RMSE univariado

RMSE regresión

Modelo seleccionado

Cobre

ARIMA(1,1,0)

...

...

...

Aluminio

ARIMA(0,1,0)

...

...

...

Níquel

Holt amortiguado

...

...

...

Zinc

ARIMA(1,1,1)

...

...

...

Nota metodológica

Los análisis previos encontraron que las series de precios en niveles son no estacionarias y que la primera diferencia mejora la estacionariedad. Por ello, el modelo principal de este repositorio trabaja con cambios mensuales (primeras diferencias) y no con una regresión simple en niveles.

Esto reduce el riesgo de obtener una regresión espuria ocasionada únicamente por tendencias comunes.

El script también ejecuta ADF sobre las variables explicativas para documentar esta decisión.

Resultado final

El archivo:

resultados/pronosticos_2025.csv

contendrá el precio mensual estimado de cada metal para 2025 mediante regresión múltiple.

Después estos precios pueden compararse contra los pronósticos ARIMA/Holt de los Anexos A–D y utilizar el modelo ganador como entrada del presupuesto determinístico de Casa de Moneda de México.

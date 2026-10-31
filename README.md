# 💰 Flujo de Caja y Control Financiero Personal (Multi-Anual 2026 - 2027)

Plantilla automatizada en Excel diseñada específicamente para el **control exhaustivo diario y semanal**, prevención de fugas por gastos hormiga, proyección de liquidez y **visión multi-anual continua (Octubre 2026 a Diciembre 2027)**.

---

## 📁 Archivos del Proyecto

* 📊 **[Flujo_de_Caja_Personal.xlsx](file:///C:/Users/Pite/Documents/CashFlowHogar/Flujo%20de%20caja%20Personal/Flujo_de_Caja_Personal.xlsx)**: Planilla principal de Excel con selector dinámico de mes, consolidado multi-anual y gráficos.
* 🐍 **[generate_workbook.py](file:///C:/Users/Pite/Documents/CashFlowHogar/Flujo%20de%20caja%20Personal/generate_workbook.py)**: Script en Python con `openpyxl` que estructura el libro, fórmulas y gráficos.
* 📖 **[README.md](file:///C:/Users/Pite/Documents/CashFlowHogar/Flujo%20de%20caja%20Personal/README.md)**: Guía y documentación del modelo financiero.

---

## 🚀 Novedades: Extensión para 2026 y Todo el Año 2027

1. **Selector Dinámico de Mes en el Tablero (`Tablero Mensual`)**:
   * En la celda **`F5`** dispones de un menú desplegable con los **15 meses** (desde *Octubre 2026* hasta *Diciembre 2027*).
   * Al elegir cualquier mes, **todo el tablero se recalcula al instante**: tarjetas KPI, la tabla de 5 semanas, los gráficos de barras/circular y el control de presupuesto por categoría.
   * **Saldo Inicial Automático y Continuo**: La celda `C5` calcula automáticamente el saldo inicial de cada mes arrastrando la liquidez final del mes anterior. ¡No tienes que copiar saldos a mano!

2. **Nueva Hoja Maestra: `Consolidado Anual`**:
   * Matriz panorámica con los **15 meses en una sola pantalla** (Oct-26 a Dic-27) con totales para **Q4 2026** y para **TODO 2027**.
   * **Resumen Ejecutivo de Cash Flow**: Saldo inicial, ingresos, egresos fijos/variables/hormiga, ahorro, flujo neto, saldo final y tasa de ahorro.
   * **Evolución por Categoría Mes a Mes**: Permite comparar meses, anticipar estacionalidad (vacaciones, aguinaldos, inicios de clases) y ver el acumulado anual de cada rubro.

3. **Libro Diario Continuo para 1.500 movimientos (`Registro Diario`)**:
   * Un único registro cronológico para todo el período 2026-2027.
   * Se mantuvieron intactos todos tus registros reales (Guillermo Mattioli, Starbucks, Ballet, Meli+, etc.).
   * Fórmulas preconfiguradas para numeración automática de fila (`#`) y asignación semanal (`Semana 1` a `Semana 5`).

4. **Planificador Semanal Flexible (`Proyeccion Semanal`)**:
   * Incluye selector de mes para proyectar semanas y anticipar vencimientos (alquiler, expensas, deudas, tarjetas) antes de cobrar.

5. **Calendario y Categorías en `Configuracion`**:
   * Tabla de calendario con fechas de inicio y fin exactas para cada uno de los 15 períodos, garantizando filtros independientes de idioma o formato de fecha.

---

## 📑 Estructura de las Pestañas

| Pestaña | Propósito y Funcionalidad |
| :--- | :--- |
| **`Tablero Mensual`** | **Dashboard Ejecutivo Dinámico**: Selecciona cualquier mes en la celda `F5` para ver sus KPIs, el flujo de sus 5 semanas, gráficos comparativos y el control de presupuesto vs. real por categoría. |
| **`Consolidado Anual`** | **Matriz Multi-Anual (15 Meses)**: Vista comparativa horizontal desde Octubre 2026 hasta Diciembre 2027, con totales acumulados de 2026 y 2027. |
| **`Registro Diario`** | **Motor Cotidiano**: Registras la fecha y el monto; el sistema numera la fila y calcula la semana automáticamente. Menús desplegables para Tipo, Categoría y Medio de Pago. |
| **`Proyeccion Semanal`** | **Control de Liquidez**: Planifica tus cobros y pagos semana a semana para evitar descubiertos bancarios. |
| **`Configuracion`** | **Metas y Calendario**: Ajusta presupuestos base por categoría y consulta el calendario de 15 meses. |
| **`Guia y Consejos`** | **Metodología**: Cómo realizar el cierre mensual, la regla de los 3 minutos diarios y cómo planificar el 2027. |

---

## 💡 Cómo empezar a usar el nuevo período (Noviembre 2026)

1. Abre **[Flujo_de_Caja_Personal.xlsx](file:///C:/Users/Pite/Documents/CashFlowHogar/Flujo%20de%20caja%20Personal/Flujo_de_Caja_Personal.xlsx)**.
2. En la pestaña **`Tablero Mensual`**, la celda `F5` ya está preseleccionada en **"Noviembre 2026"**. Verás que arranca automáticamente con el saldo acumulado que dejaste en Octubre (\$320.710,00).
3. En la pestaña **`Registro Diario`**, continúa agregando tus gastos a partir de la fila 12 ingresando las fechas de noviembre (`01/11/2026`, `02/11/2026`, etc.).
4. Si quieres consultar el cierre de octubre, solo cambia `F5` a **"Octubre 2026"** en el tablero o visita **`Consolidado Anual`** para ver ambos meses en paralelo.

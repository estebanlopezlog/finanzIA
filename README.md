# 💰 FinanzIA — Gestor Inteligente de Finanzas Personales

**Proyecto Final · Prompt Engineering para Programadores · Diplomatura en Data Science · CoderHouse**

---

## Descripción

**FinanzIA** es una aplicación web de finanzas personales potenciada por Inteligencia Artificial.
Permite cargar un archivo Excel con ingresos y gastos, organizarlos por categorías y obtener
análisis y recomendaciones personalizadas a través de GPT-3.5-turbo de OpenAI.

## Funcionalidades

| Módulo | Descripción |
|---|---|
| 📊 Dashboard | Visualización de ingresos vs gastos por mes, distribución por concepto y evolución temporal |
| 📋 Maestro de Gastos | Creación de familias/categorías y asignación de conceptos de gasto |
| 🤖 Análisis IA | Análisis financiero general y chat interactivo con GPT-3.5-turbo |
| ℹ️ Cómo Funciona | Guía de uso, formato del Excel y detalles técnicos |

## Tecnologías

- **Streamlit** — Framework web
- **Pandas** — Procesamiento de datos
- **Plotly** — Visualizaciones interactivas
- **OpenAI GPT-3.5-turbo** — Análisis con IA

## Instalación local

```bash
git clone https://github.com/tu-usuario/finanzIA.git
cd finanzIA
pip install -r requirements.txt
streamlit run app.py
```

## Formato del Excel

El archivo debe tener dos hojas:

**Gastos:** `Fecha | Mes | Year | Concepto | Monto | Medio de Pago | Moneda`

**Sueldos:** `Fecha | Mes | Año | Sueldo en Pesos | Sueldo en U$S`

## Uso de la IA — Prompt Inicial

```
Eres FinanzIA, un asesor experto en finanzas personales en Argentina.
Tu misión es ayudar al usuario a entender su situación financiera
y mejorar sus hábitos de gasto...
```

Solo se envían resúmenes estadísticos a la API (nunca datos crudos).
Costo estimado: ~$0.001 USD por análisis (gpt-3.5-turbo, max 450 tokens).

## Deploy en Streamlit Cloud

1. Subí el repo a GitHub
2. Ingresá a [share.streamlit.io](https://share.streamlit.io)
3. Conectá el repo y seleccioná `app.py`
4. ¡Listo!

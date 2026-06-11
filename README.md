# Scraper UTM - Indicadores Previsionales

## Descripción

Sistema de scraping que extrae indicadores previsionales desde **Previred** y **SII**, y los despliega en un dashboard web interactivo con gráficos, tablas ordenables, filtros, historial de ejecuciones y descarga de Excel con gráficos nativos.

Ideal para contadores, analistas de RR.HH. y cualquier profesional que necesite consultar valores UTM, UTA, IPC, tasas AFP, AFC, rentas mínimas, asignación familiar y otros indicadores previsionales actualizados.

---

## Stack Tecnológico

| Herramienta | Versión | Uso |
|---|---|---|
| **Python** | 3.14+ | Lenguaje principal |
| **requests** | 2.34+ | Cliente HTTP para scraping |
| **BeautifulSoup4** | 4.15+ | Parseo y extracción de HTML |
| **lxml** | 6.1+ | Parser HTML rápido |
| **openpyxl** | 3.1+ | Generación de Excel con gráficos |
| **http.server** | (stdlib) | Servidor web integrado |
| **Chart.js** | 4.4+ | Gráficos interactivos en dashboard |
| **PowerShell** | 5.1+ | Scripts de automatización |

---

## Estructura del Proyecto

```
scraperUTM/
├── server.py                        # Servidor HTTP (API + archivos estáticos)
├── scraper.py                       # Script principal de scraping (Python)
├── index.html                       # Dashboard web con Chart.js
├── dashboard_data.json              # Datos JSON para el dashboard
├── indicadores_previsionales.xlsx   # Excel descargable con 7 hojas + gráficos
├── scraper_log.txt                  # Historial persistente de ejecuciones
├── run_scraper.ps1                  # Script PowerShell para scraper manual
├── setup_scheduler.ps1              # Programar ejecución automática
└── README.md                        # Documentación
```

---

## API Endpoints

| Endpoint | Método | Descripción |
|---|---|---|
| `/` | GET | Dashboard web interactivo |
| `/api/data` | GET | Datos JSON actuales |
| `/api/refresh` | GET | Ejecuta el scraper y actualiza datos |
| `/api/logs` | GET | Últimas 100 líneas del historial de ejecuciones |
| `/api/download/excel` | GET | Descarga Excel con indicadores y gráficos |
| `/api/status` | GET | Health check del servidor |

---

## Funcionalidades del Dashboard

- **Gráficos interactivos** UTM/UTA y variación IPC con Chart.js
- **Filtro por rango de meses** en la tabla SII
- **Tablas ordenables** haciendo clic en los encabezados
- **Botón "Refrescar Datos"** que ejecuta el scraper desde el dashboard
- **Panel de logs** con historial persistente de ejecuciones
- **Descarga de Excel** con 7 hojas y gráficos nativos editables
- **Notificaciones toast** visuales al actualizar
- **Diseño responsive** adaptable a móviles

---

## Hojas del Excel (`indicadores_previsionales.xlsx`)

1. **Resumen** — Principales indicadores (UF, UTM, topes, rentas mínimas, APV)
2. **AFP** — Tasas de cotización por AFP (Trabajador, Empleador, Total, Independiente)
3. **AFC** — Seguro de Cesantía por tipo de contrato
4. **AsignacionFamiliar** — Montos y requisitos por tramo
5. **TrabajosPesados** — Tasas para trabajo pesado y menos pesado
6. **UTM_UTA_IPC** — Tabla mensual completa SII
7. **Graficos** — Gráficos de líneas UTM/UTA y variación IPC (nativos de Excel)

---

## Uso

### 1. Servidor web (recomendado)

```bash
python server.py
```

Esto inicia el servidor en `http://localhost:8080` y abre automáticamente el dashboard en el navegador.

### 2. Solo scraper

```powershell
.\run_scraper.ps1
```

O abrir `index.html` directamente en el navegador (sin botón de refrescar).

---

## Licencia

Proyecto académico — Bootcamp Python.

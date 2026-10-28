# Registro de Instalación de Candados Inteligentes - Luz del Sur

Frontend web responsive (HTML5 + CSS3 + JavaScript Puro) diseñado para el registro en campo de la instalación de candados inteligentes en Subestaciones de Distribución (SED) de Luz del Sur.

## Estructura del Proyecto
```
candados-inteligentes/
├── assets/
│   ├── logo-luz-del-sur.png   <-- Coloca aquí el logo PNG proporcionado
│   └── logo-luz-del-sur.svg   <-- Logo vectorial institucional (fallback automático)
├── index.html                 <-- Interfaz semántica y accesible
├── styles.css                 <-- Estilos corporativos y responsive (Mobile-First)
├── app.js                     <-- Lógica pura: catálogo SED, validaciones, carga de fotos y mock SharePoint
└── README.md                  <-- Guía de uso y despliegue
```

## Características Principales
1. **Buscador predictivo de SED**: Filtro en tiempo real entre más de 580 subestaciones con soporte táctil y de teclado (↑, ↓, Enter).
2. **Selección múltiple táctil**: Chips seleccionables para MULTILOCK y YARDENI.
3. **Código de candado inteligente**: Conversión automática a mayúsculas y validación de formato mediante expresión regular configurable.
4. **Carga y previsualización de evidencias**:
   - Soporte para tomar foto con la cámara del celular o seleccionar de la galería.
   - Zona de arrastrar y soltar (drag & drop) para escritorio.
   - Límite controlado de 1 a 10 fotos con miniaturas y botón de eliminación individual.
5. **Listo para SharePoint**: Función asíncrona aislada `enviarRegistro(datos)` documentada paso a paso para conectar con Microsoft Graph / SharePoint REST API o Power Automate.
6. **Validación en tiempo real y accesibilidad**: Mensajes en español, auto-scroll al primer campo con error y soporte para lectores de pantalla.

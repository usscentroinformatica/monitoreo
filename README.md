# Sistema de Monitoreo USS

Aplicación web React + Vite para el monitoreo y gestión de sesiones educativas en la Universidad Señor de Sipán (USS).

## Características

- **Gestión de archivos Excel**: Carga y procesamiento de archivos de asistencia
- **Integración con Zoom**: Correlación automática con reportes CSV de reuniones Zoom
- **Sistema de pestañas**: Manejo simultáneo de múltiples archivos
- **Autocompletado inteligente**: Creación automática de filas basada en datos de Zoom
- **Exportación profesional**: Formato USS con estilos corporativos

## Tecnologías

- React 19.1.1 con Vite 7.1.7
- ExcelJS para manipulación de archivos Excel
- PapaParse para procesamiento de CSV
- Tailwind CSS para estilos
- date-fns para manejo de fechas

## Comandos

```bash
# Desarrollo
npm run dev

# Build de producción
npm run build

# Previsualizar build
npm run preview

# Lint
npm run lint
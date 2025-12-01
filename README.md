# Sistema de Producción MI-EELO

Sistema automatizado de gestión de producción para Google Sheets con Google Apps Script.

## 📋 Descripción

Sistema completo para gestionar tareas de producción, asignar trabajo, monitorear eficiencia y automatizar notificaciones por correo electrónico.

## ✨ Características Principales

### 🎯 Gestión de Tareas
- Asignación automática de tareas
- Seguimiento de estado en tiempo real
- Reasignación inteligente de trabajo
- Detección automática de pendientes

### 📧 Notificaciones Automatizadas
- Alertas de tareas pendientes
- Notificaciones de nuevas asignaciones
- Reportes diarios automáticos
- Alertas de eficiencia baja

### 📊 Monitoreo y Reportes
- Panel de control en tiempo real
- Análisis de cargas de trabajo
- Indicadores de desempeño
- Historial completo de cambios

### 🔄 Reasignaciones
- Manual o automática
- Clasificación por tipo (reparación/sobrante)
- Registro en historial
- Notificaciones inmediatas

## 📁 Estructura del Proyecto

```
mi-eelo/
├── Code.gs              # Código principal y configuración
├── Funciones.gs         # Funciones de asignación y reasignación
├── Herramientas.gs      # Herramientas de análisis y búsqueda
├── Reportes.gs          # Generación de reportes y alertas
├── appsscript.json      # Configuración del proyecto
└── README.md            # Este archivo
```

## 🚀 Instalación

### 1. Crear una Nueva Hoja de Cálculo de Google

1. Ve a [Google Sheets](https://sheets.google.com)
2. Crea una nueva hoja de cálculo
3. Nómbrala "Sistema de Producción - MI-EELO"

### 2. Abrir el Editor de Apps Script

1. En la hoja de cálculo, ve a **Extensiones** → **Apps Script**
2. Esto abrirá el editor de Google Apps Script

### 3. Copiar el Código

1. **Elimina** el archivo `Code.gs` que viene por defecto
2. Crea los siguientes archivos nuevos haciendo clic en el **+** junto a "Archivos":

#### Archivo 1: `Code.gs`
- Copia todo el contenido del archivo `Code.gs` de este repositorio

#### Archivo 2: `Funciones.gs`
- Copia todo el contenido del archivo `Funciones.gs` de este repositorio

#### Archivo 3: `Herramientas.gs`
- Copia todo el contenido del archivo `Herramientas.gs` de este repositorio

#### Archivo 4: `Reportes.gs`
- Copia todo el contenido del archivo `Reportes.gs` de este repositorio

### 4. Configurar Correos Electrónicos

En el archivo `Code.gs`, actualiza la sección de configuración de correos (líneas 7-13):

```javascript
const CORREOS_CONFIG = {
  "Encargada de Taller 1": "stephany@mi-eelo.co",
  "Encargada de Taller 2": "encargada2@empresa.co",  // ← Cambiar por email real
  "Gerente de Producción": "adrian@creamosguatemala.org",
  "Admininstrador Sistema": "adrian@creamosguatemala.org"
};

const EMAIL_ADMIN = "adrian@creamosguatemala.org";
```

### 5. Guardar y Autorizar

1. Haz clic en el **icono de guardar** (💾)
2. Pon un nombre al proyecto: "Sistema Producción MI-EELO"
3. Cierra el editor de Apps Script

### 6. Ejecutar el Setup

1. **Refresca** la hoja de cálculo (F5 o Ctrl+R)
2. Aparecerá un nuevo menú **"🔄 REASIGNACIONES"** en la barra superior
3. Haz clic en **🔄 REASIGNACIONES** → **🔧 Setup Completo**
4. **Autoriza** el script cuando se te solicite:
   - Haz clic en "Revisar permisos"
   - Selecciona tu cuenta de Google
   - Haz clic en "Avanzado"
   - Haz clic en "Ir a Sistema Producción MI-EELO (no seguro)"
   - Haz clic en "Permitir"
5. Ejecuta nuevamente **🔄 REASIGNACIONES** → **🔧 Setup Completo**
6. Espera a que aparezca el mensaje: **"✅ SISTEMA INSTALADO CORRECTAMENTE"**

## 📊 Hojas del Sistema

El setup crea automáticamente 11 hojas:

1. **🎯 Panel Control** - Dashboard principal con métricas
2. **👥 Personas** - Base de datos de empleados
3. **📋 Tareas** - Registro de todas las tareas
4. **⏱️ Tiempos** - Control de tiempos de ejecución
5. **🔧 Reparaciones** - Registro de reparaciones
6. **📝 Historial** - Log de todos los cambios
7. **🔄 Reasignaciones** - Centro de reasignación
8. **📊 Reportes** - Reportes generados
9. **📈 Indicadores** - KPIs y métricas
10. **📖 Instrucciones** - Manual de uso
11. **⚙️ Configuración** - Configuración del sistema

## 🎮 Uso del Sistema

### Menú Principal: 🔄 REASIGNACIONES

#### ➕ Asignar Nueva Tarea
1. Clic en el menú
2. Ingresa descripción, cantidad, área
3. Selecciona persona y prioridad
4. Confirma la asignación

#### 🔍 Detectar Pendientes
- Muestra todas las tareas pendientes
- Opción de reasignar directamente

#### 🔄 Reasignar Manual
1. Ingresa el ID de la tarea
2. Selecciona tipo: Reparación o Sobrante
3. Elige nueva persona o mantener la misma
4. Confirma la reasignación

#### ⚡ Reasignación Automática
- Reasigna todas las pendientes automáticamente
- Mantiene a las mismas personas

#### 📝 Ver Historial
- Muestra los últimos 10 cambios del sistema

#### 📊 Actualizar Panel Control
- Refresca las métricas del dashboard

### Menú Secundario: ⚙️ HERRAMIENTAS

#### 🎯 Generar Reporte Manual
- Genera y envía reporte inmediatamente

#### 📧 Enviar Notificación
- Envía alerta de pendientes manualmente

#### 🔍 Buscar Tarea
- Busca tarea por ID

#### 👥 Análisis de Cargas
- Muestra carga de trabajo por persona

#### 📈 Ver Eficiencia
- Muestra eficiencia de cada persona

#### ℹ️ Ayuda
- Muestra ayuda rápida

## 🤖 Automatizaciones Activas

El sistema ejecuta automáticamente:

### 1. Detección de Pendientes
- **Frecuencia**: Cada 1 hora
- **Acción**: Detecta tareas pendientes y envía notificaciones

### 2. Reporte Diario
- **Frecuencia**: Todos los días a las 4:00 PM
- **Acción**: Genera y envía reporte completo

### 3. Verificación de Eficiencia
- **Frecuencia**: Cada 30 minutos
- **Acción**: Detecta eficiencia baja y envía alertas

## 📧 Configuración de Correos

### Destinatarios de Notificaciones

| Tipo de Notificación | Destinatarios |
|----------------------|---------------|
| Tareas Pendientes | Encargadas de Taller 1 y 2 |
| Nueva Tarea Asignada | Gerente de Producción |
| Reasignación Exitosa | Gerente de Producción |
| Reporte Diario | Gerente de Producción + Admin |
| Alerta de Eficiencia Baja | Administrador del Sistema |

### Formato de Correos

Todos los correos incluyen:
- HTML formateado profesionalmente
- Tablas con datos relevantes
- Recomendaciones de acción
- Fecha y hora de generación

## 🔒 Permisos Necesarios

El script requiere los siguientes permisos:

- **spreadsheets**: Leer y modificar la hoja de cálculo
- **script.container.ui**: Mostrar diálogos y menús
- **gmail.send**: Enviar correos electrónicos
- **script.scriptapp**: Crear triggers automáticos

## 📝 Estructura de Datos

### Personas
```
ID | Nombre | Turno | Área | Teléfono | Email | Eficiencia | Carga | Estado | Rol
```

### Tareas
```
Tarea ID | Persona | Área | Descripción | Cantidad | Fecha | Estado | Prioridad | Reasignable | Notas
```

### Historial
```
Fecha | Hora | Tipo | Tarea ID | Tipo Evento | De | Para | Detalle | Usuario | Estado
```

## 🐛 Solución de Problemas

### El menú no aparece
1. Refresca la página (F5)
2. Espera 10 segundos
3. Si no aparece, ejecuta manualmente `onOpen()` desde el editor

### Error al enviar correos
1. Verifica que los correos en `CORREOS_CONFIG` sean válidos
2. Asegúrate de haber autorizado el permiso de Gmail
3. Revisa los logs en Apps Script: **Ver** → **Registros de ejecución**

### Triggers no funcionan
1. Ve al editor de Apps Script
2. Clic en el icono del **reloj** (⏰) en la barra lateral
3. Verifica que existan los triggers
4. Si no existen, ejecuta manualmente `crearTriggers()`

## 📞 Soporte

Para soporte técnico, contacta a:
- **Email**: adrian@creamosguatemala.org
- **Rol**: Administrador del Sistema

## 📄 Licencia

Sistema desarrollado para MI-EELO.

---

## 🎯 Próximos Pasos Después de la Instalación

1. ✅ Verifica que todas las hojas se crearon correctamente
2. ✅ Actualiza los correos en la configuración
3. ✅ Revisa los datos de ejemplo en cada hoja
4. ✅ Modifica los datos de "Personas" con tu equipo real
5. ✅ Prueba asignar una tarea nueva
6. ✅ Prueba detectar pendientes
7. ✅ Prueba reasignar una tarea
8. ✅ Verifica que lleguen los correos
9. ✅ Revisa el historial de cambios
10. ✅ Capacita a tu equipo en el uso del sistema

## 🚀 ¡Listo para Producción!

El sistema está diseñado para uso inmediato. Todas las funcionalidades están activas desde el primer momento.

**¡Disfruta de tu nuevo sistema de producción automatizado!** 🎉

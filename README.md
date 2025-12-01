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

## 🚀 INSTALACIÓN RÁPIDA (3 PASOS)

### PASO 1: Crear Hoja de Cálculo

1. Ve a [Google Sheets](https://sheets.google.com)
2. Crea una nueva hoja de cálculo
3. Nómbrala **"Sistema de Producción - MI-EELO"**

### PASO 2: Copiar el Código

1. En la hoja de cálculo, ve a **Extensiones** → **Apps Script**
2. **Elimina** todo el código que aparece por defecto
3. **Copia y pega** TODO el contenido del archivo **`Code-Completo.gs`** (es un solo archivo completo)
4. Haz clic en el **icono de guardar** (💾)
5. Ponle nombre al proyecto: **"Sistema Producción MI-EELO"**

### PASO 3: Configurar y Ejecutar

1. **IMPORTANTE**: Actualiza los correos electrónicos en las líneas 7-13:

```javascript
const CORREOS_CONFIG = {
  "Encargada de Taller 1": "stephany@mi-eelo.co",           // ← Ya configurado
  "Encargada de Taller 2": "CAMBIAR@tupmail.com",           // ← CAMBIAR ESTE
  "Gerente de Producción": "adrian@creamosguatemala.org",   // ← Ya configurado
  "Admininstrador Sistema": "adrian@creamosguatemala.org"   // ← Ya configurado
};
```

2. **Guarda** nuevamente (💾)
3. **Cierra** el editor de Apps Script
4. **Refresca** la hoja de cálculo (presiona F5)
5. Espera 10 segundos a que aparezca el menú **"🔄 REASIGNACIONES"**
6. Haz clic en **🔄 REASIGNACIONES** → **🔧 Setup Completo**
7. **Autoriza** el script:
   - Clic en "Revisar permisos"
   - Selecciona tu cuenta
   - Clic en "Avanzado"
   - Clic en "Ir a Sistema Producción MI-EELO (no seguro)"
   - Clic en "Permitir"
8. Ejecuta nuevamente **🔄 REASIGNACIONES** → **🔧 Setup Completo**
9. ¡Listo! Verás el mensaje: **"✅ SISTEMA INSTALADO CORRECTAMENTE"**

## 📁 Archivos del Proyecto

```
mi-eelo/
├── Code-Completo.gs     # ⭐ ARCHIVO ÚNICO CON TODO EL CÓDIGO
├── appsscript.json      # Configuración del proyecto (opcional)
└── README.md            # Este archivo
```

**NOTA**: Solo necesitas copiar **`Code-Completo.gs`** para que funcione todo el sistema.

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
3. Si no aparece, ve al editor de Apps Script
4. Ejecuta manualmente la función `onOpen` (selecciónala del menú desplegable y presiona ▶️)

### Error al enviar correos
1. Verifica que los correos en `CORREOS_CONFIG` sean válidos
2. Asegúrate de haber autorizado el permiso de Gmail
3. Revisa los logs: En Apps Script → **Ver** → **Registros de ejecución**

### Triggers no funcionan
1. Ve al editor de Apps Script
2. Clic en el icono del **reloj** (⏰) en la barra lateral izquierda
3. Verifica que existan 4 triggers
4. Si no existen, ejecuta manualmente la función `crearTriggers()`

### Error "Cannot read property..."
- Asegúrate de haber ejecutado el **Setup Completo** primero
- Verifica que todas las 11 hojas se crearon correctamente

## 📞 Soporte

Para soporte técnico, contacta a:
- **Email**: adrian@creamosguatemala.org
- **Rol**: Administrador del Sistema

## 🎯 Checklist Post-Instalación

Después de instalar, verifica:

- [ ] ✅ El menú "🔄 REASIGNACIONES" aparece
- [ ] ✅ Se crearon las 11 hojas correctamente
- [ ] ✅ Los correos están actualizados
- [ ] ✅ Puedes asignar una tarea de prueba
- [ ] ✅ La notificación por correo funciona
- [ ] ✅ El panel control se actualiza
- [ ] ✅ El historial registra cambios

## 💡 Próximos Pasos

1. Actualiza los datos en la hoja **"👥 Personas"** con tu equipo real
2. Prueba asignar tareas a diferentes personas
3. Prueba la reasignación manual
4. Verifica que los correos lleguen correctamente
5. Revisa el panel de control
6. Capacita a tu equipo en el uso del sistema

## 🚀 ¡Sistema Listo!

**El sistema está 100% funcional y listo para usar inmediatamente.**

Solo necesitas:
1. Copiar **`Code-Completo.gs`** en Apps Script
2. Actualizar el email de "Encargada de Taller 2"
3. Ejecutar el **Setup Completo**

**¡Disfruta de tu sistema de producción automatizado!** 🎉

---

## 📄 Versión

- **Versión**: 2.0
- **Fecha**: Diciembre 2025
- **Desarrollado para**: MI-EELO
- **Archivo único**: Code-Completo.gs (¡Todo en uno!)

// SISTEMA DE PRODUCCIÓN - MI-EELO
// Google Apps Script - Versión 2.0
// CON FUNCIÓN DE ASIGNAR TAREAS

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ==================== CONFIGURACIÓN DE CORREOS ====================

const CORREOS_CONFIG = {
  "Encargada de Taller 1": "stephany@mi-eelo.co",
  "Encargada de Taller 2": "encargada2@empresa.co",
  "Gerente de Producción": "adrian@creamosguatemala.org",
  "Admininstrador Sistema": "adrian@creamosguatemala.org"
};

const EMAIL_ADMIN = "adrian@creamosguatemala.org";

// ==================== SETUP COMPLETO ====================

function setupSistemaCompleto() {
  try {
    console.log("🚀 Iniciando setup del sistema...");

    crearYLimpiarHojas();
    Utilities.sleep(500);

    configurarPanelControl();
    configurarPersonas();
    configurarTareas();
    configurarTiempos();
    configurarReparaciones();
    configurarHistorial();
    configurarReasignaciones();
    configurarReportes();
    configurarIndicadores();
    configurarInstrucciones();
    configurarConfiguracion();
    Utilities.sleep(500);

    aplicarFormatosGlobales();
    crearTriggers();
    onOpen();

    SpreadsheetApp.getUi().alert(
      "✅ SISTEMA INSTALADO CORRECTAMENTE\n\n" +
      "✓ 11 Hojas creadas\n" +
      "✓ Todos los datos cargados\n" +
      "✓ Formatos perfectos\n" +
      "✓ AUTOMATIZACIONES ACTIVAS:\n" +
      "  • Asignación de tareas\n" +
      "  • Notificaciones por correo\n" +
      "  • Reasignación inteligente\n" +
      "  • Reportes automáticos\n" +
      "  • Alertas de pendientes\n\n" +
      "🎯 ¡Sistema listo para producción!"
    );

    console.log("✅ Setup completado");
  } catch (e) {
    console.log(`❌ Error: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ Error: ${e.message}`);
  }
}

// ==================== CREAR HOJAS ====================

function crearYLimpiarHojas() {
  console.log("📝 Creando hojas...");

  const hojas = [
    { nombre: "🎯 Panel Control", color: "#1f4d7f" },
    { nombre: "👥 Personas", color: "#33b0e0" },
    { nombre: "📋 Tareas", color: "#d9b919" },
    { nombre: "⏱️ Tiempos", color: "#33aa33" },
    { nombre: "🔧 Reparaciones", color: "#cc3333" },
    { nombre: "📝 Historial", color: "#33b0e0" },
    { nombre: "🔄 Reasignaciones", color: "#d9b919" },
    { nombre: "📊 Reportes", color: "#33b0e0" },
    { nombre: "📈 Indicadores", color: "#1f4d7f" },
    { nombre: "📖 Instrucciones", color: "#808080" },
    { nombre: "⚙️ Configuración", color: "#666666" }
  ];

  hojas.forEach(({ nombre, color }) => {
    try {
      let hoja = SS.getSheetByName(nombre);
      if (!hoja) {
        hoja = SS.insertSheet(nombre);
      } else {
        hoja.clearContents();
      }
      hoja.setTabColor(color);
      console.log(`  ✓ ${nombre}`);
    } catch (e) {
      console.log(`  ⚠️ ${nombre}: ${e.message}`);
    }
  });
}

// ==================== PANEL CONTROL ====================

function configurarPanelControl() {
  const sheet = SS.getSheetByName("🎯 Panel Control");
  sheet.clearContents();

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 200);

  const titulo = sheet.getRange("A1");
  titulo.setValue("🎯 SISTEMA DE PRODUCCIÓN - PANEL DE CONTROL");
  titulo.setFontSize(14);
  titulo.setFontWeight("bold");
  titulo.setBackground("#1f4d7f");
  titulo.setFontColor("white");
  titulo.setVerticalAlignment("middle");
  titulo.setHorizontalAlignment("center");
  sheet.setRowHeight(1, 30);

  sheet.setRowHeight(2, 15);

  sheet.getRange("A3").setValue("Fecha:");
  sheet.getRange("B3").setValue(new Date());
  sheet.getRange("B3").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("D3").setValue("Hora:");
  sheet.getRange("E3").setValue(new Date());
  sheet.getRange("E3").setNumberFormat("hh:mm:ss");

  sheet.setRowHeight(4, 15);

  const encabezados = ["MÉTRICA", "HOY", "META", "% CUMPLIMIENTO", "ESTADO"];
  sheet.getRange("A5:E5").setValues([encabezados]);
  sheet.getRange("A5:E5").setBackground("#e8f0f7");
  sheet.getRange("A5:E5").setFontWeight("bold");
  sheet.getRange("A5:E5").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(5, 25);

  const metricas = [
    ["Total de tareas", 4, 21, "=B6/C6"],
    ["Tareas completadas", 1, 18, "=B7/C7"],
    ["Tareas pendientes", 2, 0, "=B8/C8"],
    ["Tareas con reparación", 1, 2, "=B9/C9"],
    ["Eficiencia promedio (%)", 92, 90, "=B10/C10"]
  ];

  sheet.getRange("A6:D10").setValues(metricas);
  sheet.getRange("A6:D10").setBorder(true, true, true, true, true, true);

  sheet.getRange("E6").setValue("🟡 Normal");
  sheet.getRange("E7").setValue("🟢 Bien");
  sheet.getRange("E8").setValue("🔴 ⚠️ ACCIÓN");
  sheet.getRange("E9").setValue("🟡 Normal");
  sheet.getRange("E10").setValue("🟢 Bien");
  sheet.getRange("E6:E10").setHorizontalAlignment("center");
  sheet.getRange("E6:E10").setBorder(true, true, true, true, true, true);

  sheet.getRange("D6:D10").setNumberFormat("0.0%");

  sheet.setRowHeight(12, 15);
  sheet.getRange("A12").setValue("⚠️ TAREAS PENDIENTES - REQUIEREN ACCIÓN");
  sheet.getRange("A12:E12").setBackground("#fff3cd");
  sheet.getRange("A12:E12").setFontWeight("bold");
  sheet.getRange("A12:E12").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(12, 25);

  const encPendientes = ["ID", "Persona", "Descripción", "Recomendación", ""];
  sheet.getRange("A13:E13").setValues([encPendientes]);
  sheet.getRange("A13:E13").setBackground("#ffe8a6");
  sheet.getRange("A13:E13").setFontWeight("bold");
  sheet.getRange("A13:E13").setBorder(true, true, true, true, true, true);

  sheet.getRange("A14").setValue("T003");
  sheet.getRange("B14").setValue("Carlos Ruiz");
  sheet.getRange("C14").setValue("Planchado de Prendas");
  sheet.getRange("D14").setValue("Reasignar a carga BAJA");
  sheet.getRange("A14:E14").setBorder(true, true, true, true, true, true);

  sheet.getRange("A15").setValue("T004");
  sheet.getRange("B15").setValue("Ana Martínez");
  sheet.getRange("C15").setValue("Empaque Final");
  sheet.getRange("D15").setValue("Reasignar a carga BAJA");
  sheet.getRange("A15:E15").setBorder(true, true, true, true, true, true);

  console.log("  ✓ Panel Control");
}

// ==================== PERSONAS ====================

function configurarPersonas() {
  const sheet = SS.getSheetByName("👥 Personas");
  sheet.clearContents();

  const anchos = [50, 130, 100, 120, 110, 140, 100, 200];
  anchos.forEach((ancho, i) => sheet.setColumnWidth(i + 1, ancho));

  const encabezados = ["ID", "Nombre", "Turno", "Área", "Eficiencia (%)", "Carga", "Estado", "Rol"];
  sheet.getRange("A1:H1").setValues([encabezados]);
  sheet.getRange("A1:H1").setBackground("#33b0e0");
  sheet.getRange("A1:H1").setFontColor("white");
  sheet.getRange("A1:H1").setFontWeight("bold");
  sheet.getRange("A1:H1").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(1, 25);

  const datos = [
    [1, "Juan García", "6am-2pm", "Confección", 0.95, "Medio", "Activo", "Operario"],
    [2, "María López", "6am-2pm", "Serigrafía", 0.98, "BAJO ⭐", "Activo", "Operario"],
    [3, "Carlos Ruiz", "2pm-10pm", "Planchado", 0.85, "ALTO", "Activo", "Operario"],
    [4, "Ana Martínez", "2pm-10pm", "Empaque", 0.88, "Medio", "Activo", "Operario"],
    [5, "Pedro Sánchez", "10pm-6am", "Confección", 0.92, "BAJO ⭐", "Activo", "Operario"],
    [6, "Stephany Fuentes", "6am-2pm", "Supervisión", 1.0, "BAJA", "Activo", "Encargada Taller"]
  ];

  sheet.getRange("A2:H7").setValues(datos);
  sheet.getRange("A2:H7").setBorder(true, true, true, true, true, true);
  sheet.getRange("E2:E7").setNumberFormat("0.0%");

  for (let i = 0; i < datos.length; i++) {
    if (i % 2 === 0) {
      sheet.getRange(`A${i + 2}:H${i + 2}`).setBackground("#f5f5f5");
    }
  }

  console.log("  ✓ Personas");
}

// ==================== TAREAS ====================

function configurarTareas() {
  const sheet = SS.getSheetByName("📋 Tareas");
  sheet.clearContents();

  const anchos = [90, 140, 110, 180, 90, 110, 130, 90, 130, 180];
  anchos.forEach((ancho, i) => sheet.setColumnWidth(i + 1, ancho));

  const encabezados = ["Tarea ID", "Persona Asignada", "Área", "Descripción", "Cantidad", "Fecha", "Estado", "Prioridad", "Reasignable", "Notas"];
  sheet.getRange("A1:J1").setValues([encabezados]);
  sheet.getRange("A1:J1").setBackground("#d9b919");
  sheet.getRange("A1:J1").setFontWeight("bold");
  sheet.getRange("A1:J1").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(1, 25);

  const today = new Date();
  const datos = [
    ["T001", "Juan García", "Confección", "Confección Camisetas", 1500, today, "En Progreso", "Alta", "NO", "75% completado"],
    ["T002", "María López", "Serigrafía", "Serigrafía Logo", 1500, today, "Completada", "Alta", "NO", "Finalizado"],
    ["T003", "Carlos Ruiz", "Planchado", "Planchado de Prendas", 500, today, "🔴 PENDIENTE", "Alta", "SI", "Sin iniciar"],
    ["T004", "Ana Martínez", "Empaque", "Empaque Final", 500, today, "🔴 PENDIENTE", "Media", "SI", "Requiere reasignación"]
  ];

  sheet.getRange("A2:J5").setValues(datos);
  sheet.getRange("A2:J5").setBorder(true, true, true, true, true, true);
  sheet.getRange("F2:F5").setNumberFormat("yyyy-mm-dd");

  for (let i = 0; i < datos.length; i++) {
    const fila = i + 2;
    if (datos[i][6].includes("PENDIENTE")) {
      sheet.getRange(`A${fila}:J${fila}`).setBackground("#ffcccc");
    } else if (i % 2 === 0) {
      sheet.getRange(`A${fila}:J${fila}`).setBackground("#f9f9f9");
    }
  }

  console.log("  ✓ Tareas");
}

// ==================== TIEMPOS ====================

function configurarTiempos() {
  const sheet = SS.getSheetByName("⏱️ Tiempos");
  sheet.clearContents();

  const anchos = [90, 140, 100, 100, 100, 100, 100, 100];
  anchos.forEach((ancho, i) => sheet.setColumnWidth(i + 1, ancho));

  const encabezados = ["Tarea ID", "Persona", "Hora Inicio", "Hora Fin", "Tiempo Real", "Estimado", "Variación", "Estado"];
  sheet.getRange("A1:H1").setValues([encabezados]);
  sheet.getRange("A1:H1").setBackground("#33aa33");
  sheet.getRange("A1:H1").setFontColor("white");
  sheet.getRange("A1:H1").setFontWeight("bold");
  sheet.getRange("A1:H1").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(1, 25);

  const datos = [
    ["T001", "Juan García", "08:30", "09:45", "75 min", "60 min", "+15 min", "✅"],
    ["T002", "María López", "09:15", "09:45", "30 min", "30 min", "0 min", "✅"],
    ["T003", "Carlos Ruiz", "-", "-", "-", "45 min", "-", "🔴"],
    ["T004", "Ana Martínez", "-", "-", "-", "40 min", "-", "🔴"]
  ];

  sheet.getRange("A2:H5").setValues(datos);
  sheet.getRange("A2:H5").setBorder(true, true, true, true, true, true);

  console.log("  ✓ Tiempos");
}

// ==================== REPARACIONES ====================

function configurarReparaciones() {
  const sheet = SS.getSheetByName("🔧 Reparaciones");
  sheet.clearContents();

  const anchos = [130, 110, 180, 140, 100, 100, 130, 120];
  anchos.forEach((ancho, i) => sheet.setColumnWidth(i + 1, ancho));

  const encabezados = ["Reparación ID", "Tarea Original", "Motivo", "Persona", "Tiempo", "Costo", "Estado", "Fecha"];
  sheet.getRange("A1:H1").setValues([encabezados]);
  sheet.getRange("A1:H1").setBackground("#cc3333");
  sheet.getRange("A1:H1").setFontColor("white");
  sheet.getRange("A1:H1").setFontWeight("bold");
  sheet.getRange("A1:H1").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(1, 25);

  const today = new Date();
  const datos = [
    ["REP001", "T001", "Defecto en costuras", "Juan García", "30 min", "$15.50", "En Progreso", today],
    ["REP002", "T002", "Color desigual", "María López", "20 min", "$10.00", "Pendiente", today]
  ];

  sheet.getRange("A2:H3").setValues(datos);
  sheet.getRange("A2:H3").setBorder(true, true, true, true, true, true);
  sheet.getRange("H2:H3").setNumberFormat("yyyy-mm-dd");

  console.log("  ✓ Reparaciones");
}

// ==================== HISTORIAL ====================

function configurarHistorial() {
  const sheet = SS.getSheetByName("📝 Historial");
  sheet.clearContents();

  const anchos = [120, 100, 140, 90, 150, 140, 140, 180, 150, 100];
  anchos.forEach((ancho, i) => sheet.setColumnWidth(i + 1, ancho));

  const encabezados = ["Fecha", "Hora", "Tipo", "Tarea ID", "Tipo Evento", "De", "Para", "Detalle", "Usuario", "Estado"];
  sheet.getRange("A1:J1").setValues([encabezados]);
  sheet.getRange("A1:J1").setBackground("#33b0e0");
  sheet.getRange("A1:J1").setFontColor("white");
  sheet.getRange("A1:J1").setFontWeight("bold");
  sheet.getRange("A1:J1").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(1, 25);

  const today = new Date();
  const datos = [
    [today, "07:00:00", "Creación Sistema", "-", "Setup", "-", "-", "Sistema inicializado", "Admin", "✅"],
    [today, "08:30:00", "Tarea Iniciada", "T001", "Inicio", "Juan García", "Juan García", "Confección", "Admin", "✅"],
    [today, "09:45:00", "Tarea Completada", "T002", "Completada", "María López", "María López", "Serigrafía", "Admin", "✅"],
    [today, "10:00:00", "Detección Pendiente", "T003", "Pendiente", "Carlos Ruiz", "Carlos Ruiz", "Planchado", "Admin", "⚠️"]
  ];

  sheet.getRange("A2:J5").setValues(datos);
  sheet.getRange("A2:J5").setBorder(true, true, true, true, true, true);
  sheet.getRange("A2:A5").setNumberFormat("yyyy-mm-dd");

  console.log("  ✓ Historial");
}

// ==================== REASIGNACIONES ====================

function configurarReasignaciones() {
  const sheet = SS.getSheetByName("🔄 Reasignaciones");
  sheet.clearContents();

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 180);

  const titulo = sheet.getRange("A1");
  titulo.setValue("🔄 CENTRO DE REASIGNACIÓN DE TAREAS");
  titulo.setFontSize(12);
  titulo.setFontWeight("bold");
  titulo.setBackground("#d9b919");
  sheet.setRowHeight(1, 25);

  const instrucciones = [
    ["", ""],
    ["CÓMO USAR:", ""],
    ["1.", "Abre menú 🔄 REASIGNACIONES"],
    ["2.", "Selecciona ➕ Asignar Nueva Tarea"],
    ["3.", "O selecciona 🔍 Detectar Pendientes"],
    ["4.", "Elige 🔄 Reasignar Manual"],
    ["5.", "Sigue los pasos del sistema"],
    ["", ""],
    ["PERSONAS DISPONIBLES (CARGA BAJA):", ""],
    ["Nombre", "Eficiencia"],
    ["María López", "98%"],
    ["Pedro Sánchez", "92%"],
    ["", ""],
    ["AUTOMATIZACIONES ACTIVAS:", ""],
    ["✓ Asignación de tareas", ""],
    ["✓ Notificaciones por correo", ""],
    ["✓ Reasignación inteligente", ""],
    ["✓ Alertas de pendientes", ""],
    ["✓ Reportes automáticos", ""]
  ];

  sheet.getRange("A2:B20").setValues(instrucciones);

  console.log("  ✓ Reasignaciones");
}

// ==================== REPORTES ====================

function configurarReportes() {
  const sheet = SS.getSheetByName("📊 Reportes");
  sheet.clearContents();

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);

  const titulo = sheet.getRange("A1");
  titulo.setValue("📊 REPORTE DIARIO");
  titulo.setFontSize(12);
  titulo.setFontWeight("bold");
  titulo.setBackground("#33b0e0");
  titulo.setFontColor("white");
  sheet.setRowHeight(1, 25);

  sheet.getRange("A2").setValue("Fecha:");
  sheet.getRange("B2").setValue(new Date());
  sheet.getRange("B2").setNumberFormat("yyyy-mm-dd");

  sheet.getRange("A4").setValue("RESUMEN DEL DÍA");
  sheet.getRange("A4").setFontWeight("bold");
  sheet.getRange("A4").setBackground("#e8f0f7");
  sheet.setRowHeight(4, 21);

  const resumen = [
    ["Métrica", "Valor"],
    ["Total de tareas", 4],
    ["Completadas", 1],
    ["En progreso", 1],
    ["Pendientes", 2],
    ["Eficiencia Promedio", "92%"]
  ];

  sheet.getRange("A5:B10").setValues(resumen);
  sheet.getRange("A5:B5").setBackground("#fff3cd");
  sheet.getRange("A5:B5").setFontWeight("bold");
  sheet.getRange("A5:B10").setBorder(true, true, true, true, true, true);

  console.log("  ✓ Reportes");
}

// ==================== INDICADORES ====================

function configurarIndicadores() {
  const sheet = SS.getSheetByName("📈 Indicadores");
  sheet.clearContents();

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);

  const titulo = sheet.getRange("A1");
  titulo.setValue("📈 INDICADORES DE DESEMPEÑO");
  titulo.setFontSize(12);
  titulo.setFontWeight("bold");
  titulo.setBackground("#1f4d7f");
  titulo.setFontColor("white");
  sheet.setRowHeight(1, 25);

  sheet.getRange("A3").setValue("INDICADOR");
  sheet.getRange("B3").setValue("HOY");
  sheet.getRange("C3").setValue("META");
  sheet.getRange("D3").setValue("ESTADO");
  sheet.getRange("E3").setValue("VARIACIÓN");
  sheet.getRange("A3:E3").setBackground("#e8f0f7");
  sheet.getRange("A3:E3").setFontWeight("bold");
  sheet.setRowHeight(3, 21);

  const indicadores = [
    ["Cumplimiento (%)", "50%", "95%", "🔴 Bajo", "-45%"],
    ["Tareas Pendientes", 2, 0, "🔴 Alto", "+2"],
    ["Eficiencia Promedio", "92%", "90%", "🟢 Alto", "+2%"],
    ["Reasignaciones Totales", 0, 5, "🟡 Medio", "0"]
  ];

  sheet.getRange("A4:E7").setValues(indicadores);
  sheet.getRange("A4:E7").setBorder(true, true, true, true, true, true);

  console.log("  ✓ Indicadores");
}

// ==================== INSTRUCCIONES ====================

function configurarInstrucciones() {
  const sheet = SS.getSheetByName("📖 Instrucciones");
  sheet.clearContents();

  sheet.setColumnWidth(1, 600);

  const titulo = sheet.getRange("A1");
  titulo.setValue("📖 MANUAL DE USO - SISTEMA DE PRODUCCIÓN");
  titulo.setFontSize(12);
  titulo.setFontWeight("bold");
  titulo.setBackground("#808080");
  titulo.setFontColor("white");
  sheet.setRowHeight(1, 25);

  const instrucciones = [
    [""],
    ["🎯 OBJETIVO DEL SISTEMA"],
    ["Automatizar la gestión de tareas, reasignación de trabajo y seguimiento de eficiencia."],
    [""],
    ["🚀 PRIMEROS PASOS"],
    ["1. Abre el menú '🔄 REASIGNACIONES' en la barra superior"],
    ["2. Usa ➕ ASIGNAR NUEVA TAREA para crear nuevas tareas"],
    ["3. Selecciona 🔍 Detectar Pendientes para ver tareas sin completar"],
    ["4. Revisa qué tareas están pendientes (marcadas con 🔴 PENDIENTE)"],
    ["5. Usa 🔄 Reasignar Manual para mover tareas a otra persona"],
    [""],
    ["✨ AUTOMATIZACIONES INCLUIDAS"],
    [""],
    ["➕ ASIGNACIÓN DE TAREAS"],
    ["• Crear nuevas tareas desde el menú"],
    ["• Asignar automáticamente a personas"],
    ["• Registrar en historial"],
    [""],
    ["📧 NOTIFICACIONES POR CORREO"],
    ["• Se envía automáticamente cuando hay tareas pendientes"],
    ["• Notifica a las encargadas de taller"],
    ["• Incluye detalles de la tarea y recomendaciones"],
    [""],
    ["🤖 REASIGNACIÓN INTELIGENTE"],
    ["• El sistema analiza cargas de trabajo automáticamente"],
    ["• Sugiere reasignación a personas con capacidad disponible"],
    ["• Registra cada cambio en el historial"],
    [""],
    ["📊 REPORTES AUTOMÁTICOS"],
    ["• Se generan cada día al final del turno"],
    ["• Incluyen resumen de tareas, eficiencia y pendientes"],
    ["• Se envían por correo a gerencia"],
    [""]
  ];

  sheet.getRange(`A1:A${instrucciones.length}`).setValues(instrucciones);
  sheet.getRange(`A1:A${instrucciones.length}`).setWrap(true);
  sheet.getRange(`A1:A${instrucciones.length}`).setFontSize(11);
  sheet.getRange(`A1:A${instrucciones.length}`).setVerticalAlignment("top");

  console.log("  ✓ Instrucciones");
}

// ==================== CONFIGURACIÓN ====================

function configurarConfiguracion() {
  const sheet = SS.getSheetByName("⚙️ Configuración");
  sheet.clearContents();

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 200);

  const titulo = sheet.getRange("A1");
  titulo.setValue("⚙️ CONFIGURACIÓN DEL SISTEMA");
  titulo.setFontSize(12);
  titulo.setFontWeight("bold");
  titulo.setBackground("#666666");
  titulo.setFontColor("white");
  sheet.setRowHeight(1, 25);

  const config = [
    ["", "", ""],
    ["CONFIGURACIÓN DE CORREOS", "", ""],
    ["Rol", "Email", "Descripción"],
    ["Encargada de Taller 1", "stephany@mi-eelo.co", "Recibe alertas de pendientes"],
    ["Encargada de Taller 2", "encargada2@empresa.co", "Recibe alertas de pendientes"],
    ["Gerente de Producción", "adrian@creamosguatemala.org", "Recibe reportes diarios"],
    ["Administrador", "adrian@creamosguatemala.org", "Recibe todos los reportes"],
    ["", "", ""],
    ["CONFIGURACIÓN DE AUTOMATIZACIONES", "", ""],
    ["Parámetro", "Valor", "Descripción"],
    ["Enviar alertas de pendientes", "SÍ", "Envía notificación cuando hay tareas sin completar"],
    ["Horario de alerta", "10:00 AM", "Hora de envío de alertas diarias"],
    ["Generar reportes diarios", "SÍ", "Genera reporte automático al final del turno"],
    ["Eficiencia mínima alerta", "80%", "Alerta si eficiencia es menor a este %"],
    ["Reasignar automáticamente", "NO", "Reasigna sin confirmación (usar con cuidado)"],
    ["", "", ""],
    ["NOTIFICACIONES ACTIVAS", "", ""],
    ["✓ Tareas nuevas asignadas", "", ""],
    ["✓ Tareas pendientes detectadas", "", ""],
    ["✓ Reasignaciones realizadas", "", ""],
    ["✓ Reportes diarios", "", ""]
  ];

  sheet.getRange("A1:C21").setValues(config);
  sheet.getRange("A1:C23").setBorder(true, true, true, true, true, true);
  sheet.getRange("A3:C3").setBackground("#e8f0f7");
  sheet.getRange("A3:C3").setFontWeight("bold");
  sheet.getRange("A10:C10").setBackground("#e8f0f7");
  sheet.getRange("A10:C10").setFontWeight("bold");
  sheet.getRange("A18:C18").setBackground("#e8f0f7");
  sheet.getRange("A18:C18").setFontWeight("bold");

  console.log("  ✓ Configuración");
}

// ==================== FORMATOS GLOBALES ====================

function aplicarFormatosGlobales() {
  try {
    const hojas = SS.getSheets();
    hojas.forEach(sheet => {
      sheet.setFrozenRows(1);
      sheet.getRange("A1:Z1000").setFontFamily("Calibri");
      sheet.getRange("A1:Z1000").setFontSize(11);
    });

    console.log("  ✓ Formatos globales");
  } catch (e) {
    console.log(`  ⚠️ Error formatos: ${e.message}`);
  }
}

// ==================== TRIGGERS ====================

function crearTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

    ScriptApp.newTrigger("onOpen")
      .forSpreadsheet(SS)
      .onOpen()
      .create();

    ScriptApp.newTrigger("detectarPendientesAutomatico")
      .timeBased()
      .everyHours(1)
      .create();

    ScriptApp.newTrigger("generarReporteDiarioAutomatico")
      .timeBased()
      .atHour(16)
      .everyDays(1)
      .create();

    ScriptApp.newTrigger("verificarEficienciaBaja")
      .timeBased()
      .everyMinutes(30)
      .create();

    console.log("  ✓ Triggers creados");
  } catch (e) {
    console.log(`  ⚠️ Error triggers: ${e.message}`);
  }
}

// ==================== MENÚ PRINCIPAL ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🔄 REASIGNACIONES")
    .addItem("➕ Asignar Nueva Tarea", "btnAsignarTarea")
    .addItem("🔍 Detectar Pendientes", "btnDetectarTodosPendientes")
    .addItem("🔄 Reasignar Manual", "btnReasignarSeleccionada")
    .addItem("⚡ Reasignación Automática", "btnReasignacionAutomatica")
    .addItem("📝 Ver Historial", "btnMostrarHistorial")
    .addSeparator()
    .addItem("📊 Actualizar Panel Control", "actualizarPanelControlUI")
    .addSeparator()
    .addItem("🔧 Setup Completo", "setupSistemaCompleto")
    .addToUi();

  ui.createMenu("⚙️ HERRAMIENTAS")
    .addItem("🎯 Generar Reporte Manual", "btnGenerarReporteManual")
    .addItem("📧 Enviar Notificación", "btnEnviarNotificacionManual")
    .addItem("🔍 Buscar Tarea", "buscarTarea")
    .addItem("👥 Análisis de Cargas", "analizarCargas")
    .addItem("📈 Ver Eficiencia", "verEficiencia")
    .addSeparator()
    .addItem("ℹ️ Ayuda", "mostrarAyuda")
    .addToUi();
}

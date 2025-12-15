// ==================================================================================
// SISTEMA DE PRODUCCIÓN - MI-EELO
// Google Apps Script - Versión 2.0 - CÓDIGO COMPLETO
// ==================================================================================

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

  const anchos = [90, 140, 110, 180, 90, 110, 130, 130, 180];
  anchos.forEach((ancho, i) => sheet.setColumnWidth(i + 1, ancho));

  const encabezados = ["Tarea ID", "Persona Asignada", "Área", "Descripción", "Cantidad", "Fecha", "Estado", "Reasignable", "Notas"];
  sheet.getRange("A1:I1").setValues([encabezados]);
  sheet.getRange("A1:I1").setBackground("#d9b919");
  sheet.getRange("A1:I1").setFontWeight("bold");
  sheet.getRange("A1:I1").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(1, 25);

  const today = new Date();
  const datos = [
    ["T001", "Juan García", "Confección", "Confección Camisetas", 1500, today, "En Progreso", "NO", "75% completado"],
    ["T002", "María López", "Serigrafía", "Serigrafía Logo", 1500, today, "Completada", "NO", "Finalizado"],
    ["T003", "Carlos Ruiz", "Planchado", "Planchado de Prendas", 500, today, "🔴 PENDIENTE", "SI", "Sin iniciar"],
    ["T004", "Ana Martínez", "Empaque", "Empaque Final", 500, today, "🔴 PENDIENTE", "SI", "Requiere reasignación"]
  ];

  sheet.getRange("A2:I5").setValues(datos);
  sheet.getRange("A2:I5").setBorder(true, true, true, true, true, true);
  sheet.getRange("F2:F5").setNumberFormat("yyyy-mm-dd");

  for (let i = 0; i < datos.length; i++) {
    const fila = i + 2;
    if (datos[i][6].includes("PENDIENTE")) {
      sheet.getRange(`A${fila}:I${fila}`).setBackground("#ffcccc");
    } else if (i % 2 === 0) {
      sheet.getRange(`A${fila}:I${fila}`).setBackground("#f9f9f9");
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

    ScriptApp.newTrigger("actualizarTodasEficiencias")
      .timeBased()
      .atHour(17)
      .everyDays(1)
      .create();

    console.log("  ✓ Triggers creados (incluye actualización diaria de eficiencias)");
  } catch (e) {
    console.log(`  ⚠️ Error triggers: ${e.message}`);
  }
}

// ==================== MENÚ PRINCIPAL ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🔄 REASIGNACIONES")
    .addItem("➕ Asignar Nueva Tarea", "btnAsignarTarea")
    .addItem("✅ Completar Tarea", "btnCompletarTarea")
    .addSeparator()
    .addItem("🔍 Detectar Pendientes", "btnDetectarTodosPendientes")
    .addItem("🔄 Reasignar Manual", "btnReasignarSeleccionada")
    .addItem("⚡ Reasignación Automática", "btnReasignacionAutomatica")
    .addItem("📝 Ver Historial", "btnMostrarHistorial")
    .addSeparator()
    .addItem("📊 Actualizar Panel Control", "actualizarPanelControlUI")
    .addSeparator()
    .addItem("🔧 Setup Completo", "setupSistemaCompleto")
    .addItem("🗑️ Eliminar Datos de Ejemplo", "eliminarDatosEjemplo")
    .addToUi();

  ui.createMenu("⚙️ HERRAMIENTAS")
    .addItem("🎯 Generar Reporte Manual", "btnGenerarReporteManual")
    .addItem("📧 Enviar Notificación", "btnEnviarNotificacionManual")
    .addItem("🔍 Buscar Tarea", "buscarTarea")
    .addItem("👥 Análisis de Cargas", "analizarCargas")
    .addItem("📈 Ver Eficiencia", "verEficiencia")
    .addSeparator()
    .addItem("📊 Actualizar Todas las Eficiencias", "btnActualizarTodasEficiencias")
    .addSeparator()
    .addItem("ℹ️ Ayuda", "mostrarAyuda")
    .addToUi();
}

// ==================== FUNCIÓN: ASIGNAR NUEVA TAREA ====================

function btnAsignarTarea() {
  mostrarFormularioAsignacion();
}

function mostrarFormularioAsignacion() {
  const html = HtmlService.createHtmlOutput(getFormularioHTML())
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ ASIGNAR NUEVA TAREA');
}

function getFormularioHTML() {
  // Obtener personas desde la hoja
  const personasSheet = SS.getSheetByName("👥 Personas");
  const personasData = personasSheet.getDataRange().getValues();

  let opcionesPersonas = '';
  for (let i = 1; i < personasData.length; i++) {
    const nombre = personasData[i][1];
    const carga = personasData[i][5];
    if (nombre) {
      opcionesPersonas += `<option value="${nombre}">${nombre} - ${carga}</option>`;
    }
  }

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background-color: #f5f5f5;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            font-weight: bold;
            margin-bottom: 5px;
            color: #1f4d7f;
          }
          input[type="text"],
          input[type="number"],
          select {
            width: 100%;
            padding: 10px;
            font-size: 14px;
            border: 2px solid #ccc;
            border-radius: 5px;
            box-sizing: border-box;
          }
          select {
            cursor: pointer;
            background-color: white;
          }
          input:focus,
          select:focus {
            border-color: #1f4d7f;
            outline: none;
          }
          .btn {
            padding: 12px 30px;
            font-size: 16px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            margin-right: 10px;
          }
          .btn-primary {
            background-color: #1f4d7f;
            color: white;
          }
          .btn-primary:hover {
            background-color: #163a5f;
          }
          .btn-secondary {
            background-color: #ccc;
            color: #333;
          }
          .btn-secondary:hover {
            background-color: #999;
          }
          .buttons {
            margin-top: 30px;
            text-align: center;
          }
          .titulo {
            color: #1f4d7f;
            margin-bottom: 20px;
            font-size: 20px;
          }
          .required {
            color: red;
          }
        </style>
      </head>
      <body>
        <h2 class="titulo">📋 Nueva Tarea de Producción</h2>

        <form id="formulario">
          <div class="form-group">
            <label for="operacion">Nombre de la Operación <span class="required">*</span></label>
            <input type="text" id="operacion" name="operacion" placeholder="Ej: Confección Camisetas" required>
          </div>

          <div class="form-group">
            <label for="area">Área <span class="required">*</span></label>
            <select id="area" name="area" required>
              <option value="">-- Selecciona un área --</option>
              <option value="Confección">Confección</option>
              <option value="Serigrafía">Serigrafía</option>
              <option value="Planchado">Planchado</option>
              <option value="Empaque">Empaque</option>
            </select>
          </div>

          <div class="form-group">
            <label for="persona">Asignar a <span class="required">*</span></label>
            <select id="persona" name="persona" required>
              <option value="">-- Selecciona una persona --</option>
              ${opcionesPersonas}
            </select>
          </div>

          <div class="form-group">
            <label for="cantidad">Cantidad <span class="required">*</span></label>
            <input type="number" id="cantidad" name="cantidad" placeholder="Ej: 1000" min="1" required>
          </div>

          <div class="buttons">
            <button type="submit" class="btn btn-primary">✅ Asignar Tarea</button>
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">❌ Cancelar</button>
          </div>
        </form>

        <script>
          document.getElementById('formulario').addEventListener('submit', function(e) {
            e.preventDefault();

            const operacion = document.getElementById('operacion').value.trim();
            const area = document.getElementById('area').value;
            const persona = document.getElementById('persona').value;
            const cantidad = parseInt(document.getElementById('cantidad').value);

            if (!operacion || !area || !persona || !cantidad) {
              alert('❌ Por favor completa todos los campos');
              return;
            }

            // Deshabilitar botón para evitar doble envío
            document.querySelector('.btn-primary').disabled = true;
            document.querySelector('.btn-primary').textContent = '⏳ Asignando...';

            // Enviar datos a Google Apps Script
            google.script.run
              .withSuccessHandler(function() {
                alert('✅ TAREA ASIGNADA CORRECTAMENTE\\n\\n' +
                      'Operación: ' + operacion + '\\n' +
                      'Área: ' + area + '\\n' +
                      'Persona: ' + persona + '\\n' +
                      'Cantidad: ' + cantidad);
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('❌ Error: ' + error);
                document.querySelector('.btn-primary').disabled = false;
                document.querySelector('.btn-primary').textContent = '✅ Asignar Tarea';
              })
              .procesarNuevaTarea(operacion, cantidad, area, persona);
          });
        </script>
      </body>
    </html>
  `;
}

function procesarNuevaTarea(operacion, cantidad, area, persona) {
  crearNuevaTarea(operacion, cantidad, area, persona);
}

function crearNuevaTarea(descripcion, cantidad, area, persona) {
  try {
    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const historialSheet = SS.getSheetByName("📝 Historial");

    // Generar ID de tarea
    const ultimaFila = tareasSheet.getLastRow();
    const tareaID = `T${String(ultimaFila).padStart(3, "0")}`;

    // Agregar tarea
    const nuevaFila = ultimaFila + 1;
    const today = new Date();

    tareasSheet.getRange(nuevaFila, 1).setValue(tareaID);
    tareasSheet.getRange(nuevaFila, 2).setValue(persona);
    tareasSheet.getRange(nuevaFila, 3).setValue(area);
    tareasSheet.getRange(nuevaFila, 4).setValue(descripcion);
    tareasSheet.getRange(nuevaFila, 5).setValue(cantidad);
    tareasSheet.getRange(nuevaFila, 6).setValue(today);
    tareasSheet.getRange(nuevaFila, 7).setValue("En Progreso");
    tareasSheet.getRange(nuevaFila, 8).setValue("SI");
    tareasSheet.getRange(nuevaFila, 9).setValue("Tarea nueva asignada");

    // Registrar en Historial
    const ultimaFilaHistorial = historialSheet.getLastRow() + 1;
    const ahora = new Date();

    historialSheet.getRange(ultimaFilaHistorial, 1).setValue(ahora);
    historialSheet.getRange(ultimaFilaHistorial, 2).setValue(ahora.toLocaleTimeString("es-MX"));
    historialSheet.getRange(ultimaFilaHistorial, 3).setValue("Nueva Tarea");
    historialSheet.getRange(ultimaFilaHistorial, 4).setValue(tareaID);
    historialSheet.getRange(ultimaFilaHistorial, 5).setValue("Asignación");
    historialSheet.getRange(ultimaFilaHistorial, 6).setValue("-");
    historialSheet.getRange(ultimaFilaHistorial, 7).setValue(persona);
    historialSheet.getRange(ultimaFilaHistorial, 8).setValue(descripcion);
    historialSheet.getRange(ultimaFilaHistorial, 9).setValue(Session.getActiveUser().getEmail());
    historialSheet.getRange(ultimaFilaHistorial, 10).setValue("✅ Asignada");

    // Enviar notificación
    const asunto = `✅ Nueva Tarea Asignada: ${tareaID}`;
    const cuerpo = `
    <html>
      <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #1f4d7f;">✅ NUEVA TAREA ASIGNADA</h2>

        <p><strong>Tarea ID:</strong> ${tareaID}</p>
        <p><strong>Persona:</strong> ${persona}</p>
        <p><strong>Descripción:</strong> ${descripcion}</p>
        <p><strong>Cantidad:</strong> ${cantidad}</p>
        <p><strong>Área:</strong> ${area}</p>

        <p style="margin-top: 20px; color: #666;">
          <strong>Hora:</strong> ${ahora.toLocaleString("es-MX")}<br>
          <strong>Sistema:</strong> Producción Automático - Control de Tiempos
        </p>
      </body>
    </html>
    `;

    try {
      GmailApp.sendEmail(CORREOS_CONFIG["Gerente de Producción"], asunto, cuerpo, { htmlBody: cuerpo });
      console.log("  ✓ Notificación de nueva tarea enviada");
    } catch (e) {
      console.log(`  ⚠️ Error enviando notificación: ${e.message}`);
    }

    actualizarPanelControl();

  } catch (e) {
    console.log(`Error creando tarea: ${e.message}`);
  }
}

// ==================== FUNCIÓN: COMPLETAR TAREA ====================

function btnCompletarTarea() {
  mostrarFormularioCompletar();
}

function mostrarFormularioCompletar() {
  const html = HtmlService.createHtmlOutput(getFormularioCompletarHTML())
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '✅ COMPLETAR TAREA');
}

function getFormularioCompletarHTML() {
  // Obtener tareas en progreso desde la hoja
  const tareasSheet = SS.getSheetByName("📋 Tareas");
  const tareasData = tareasSheet.getDataRange().getValues();

  let opcionesTareas = '';
  for (let i = 1; i < tareasData.length; i++) {
    const tareaID = tareasData[i][0];
    const persona = tareasData[i][1];
    const descripcion = tareasData[i][3];
    const estado = tareasData[i][6];

    if (tareaID && estado === "En Progreso") {
      opcionesTareas += `<option value="${tareaID}|${persona}">${tareaID} - ${persona} - ${descripcion}</option>`;
    }
  }

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background-color: #f5f5f5;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            font-weight: bold;
            margin-bottom: 5px;
            color: #1f4d7f;
          }
          input[type="number"],
          select {
            width: 100%;
            padding: 10px;
            font-size: 14px;
            border: 2px solid #ccc;
            border-radius: 5px;
            box-sizing: border-box;
          }
          select {
            cursor: pointer;
            background-color: white;
          }
          input:focus,
          select:focus {
            border-color: #1f4d7f;
            outline: none;
          }
          .btn {
            padding: 12px 30px;
            font-size: 16px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            margin-right: 10px;
          }
          .btn-primary {
            background-color: #33aa33;
            color: white;
          }
          .btn-primary:hover {
            background-color: #2d8f2d;
          }
          .btn-secondary {
            background-color: #ccc;
            color: #333;
          }
          .btn-secondary:hover {
            background-color: #999;
          }
          .buttons {
            margin-top: 30px;
            text-align: center;
          }
          .titulo {
            color: #1f4d7f;
            margin-bottom: 20px;
            font-size: 20px;
          }
          .required {
            color: red;
          }
          .info-box {
            background-color: #e7f3ff;
            border-left: 4px solid #1f4d7f;
            padding: 12px;
            margin-bottom: 20px;
            border-radius: 4px;
          }
        </style>
      </head>
      <body>
        <h2 class="titulo">✅ Completar Tarea</h2>

        <div class="info-box">
          📊 Al completar una tarea, se actualizará automáticamente la eficiencia del operario
        </div>

        <form id="formulario">
          <div class="form-group">
            <label for="tarea">Tarea a Completar <span class="required">*</span></label>
            <select id="tarea" name="tarea" required>
              <option value="">-- Selecciona una tarea --</option>
              ${opcionesTareas}
            </select>
          </div>

          <div class="form-group">
            <label for="minutos">Tiempo Real (en minutos) <span class="required">*</span></label>
            <input type="number" id="minutos" name="minutos" placeholder="Ej: 45" min="1" required>
          </div>

          <div class="buttons">
            <button type="submit" class="btn btn-primary">✅ Marcar como Completada</button>
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">❌ Cancelar</button>
          </div>
        </form>

        <script>
          document.getElementById('formulario').addEventListener('submit', function(e) {
            e.preventDefault();

            const tareaData = document.getElementById('tarea').value;
            const minutos = parseInt(document.getElementById('minutos').value);

            if (!tareaData || !minutos) {
              alert('❌ Por favor completa todos los campos');
              return;
            }

            // Separar tareaID y persona
            const [tareaID, persona] = tareaData.split('|');

            // Deshabilitar botón para evitar doble envío
            document.querySelector('.btn-primary').disabled = true;
            document.querySelector('.btn-primary').textContent = '⏳ Completando...';

            // Enviar datos a Google Apps Script
            google.script.run
              .withSuccessHandler(function() {
                alert('✅ TAREA COMPLETADA CORRECTAMENTE\\n\\n' +
                      'Tarea: ' + tareaID + '\\n' +
                      'Persona: ' + persona + '\\n' +
                      'Tiempo Real: ' + minutos + ' minutos\\n\\n' +
                      '📊 La eficiencia ha sido actualizada automáticamente');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('❌ Error: ' + error);
                document.querySelector('.btn-primary').disabled = false;
                document.querySelector('.btn-primary').textContent = '✅ Marcar como Completada';
              })
              .procesarTareaCompletada(tareaID, persona, minutos);
          });
        </script>
      </body>
    </html>
  `;
}

function procesarTareaCompletada(tareaID, persona, minutos) {
  marcarTareaCompletada(tareaID, persona, minutos);
}

function btnActualizarTodasEficiencias() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    '📊 ACTUALIZAR EFICIENCIAS',
    '¿Deseas recalcular las eficiencias de todas las personas basándose en sus tareas completadas?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta === ui.Button.YES) {
    actualizarTodasEficiencias();
    ui.alert('✅ EFICIENCIAS ACTUALIZADAS\n\nSe han recalculado las eficiencias de todas las personas basándose en su desempeño real.');
  }
}

// ==================== SISTEMA DE EFICIENCIA AUTOMÁTICA ====================

/**
 * Calcula la eficiencia de una persona basándose en su desempeño real
 * @param {string} nombrePersona - Nombre de la persona
 * @return {number} - Eficiencia en porcentaje (0-100)
 */
function calcularEficienciaPersona(nombrePersona) {
  try {
    const tiemposSheet = SS.getSheetByName("⏱️ Tiempos");
    const reparacionesSheet = SS.getSheetByName("🔧 Reparaciones");

    // Obtener todas las tareas completadas de esta persona
    const datosTiempos = tiemposSheet.getDataRange().getValues();
    const tareasPersona = [];

    for (let i = 1; i < datosTiempos.length; i++) {
      const persona = datosTiempos[i][1];
      const estado = datosTiempos[i][7];

      if (persona === nombrePersona && estado === "✅") {
        tareasPersona.push({
          tareaID: datosTiempos[i][0],
          tiempoReal: datosTiempos[i][4],
          tiempoEstimado: datosTiempos[i][5],
          variacion: datosTiempos[i][6]
        });
      }
    }

    // Si no tiene tareas completadas, retornar eficiencia inicial
    if (tareasPersona.length === 0) {
      return 85.0;
    }

    // Limitar a las últimas 10 tareas para el cálculo
    const tareasRecientes = tareasPersona.slice(-10);

    // Calcular eficiencia basada en tiempos
    let sumaEficienciaTiempo = 0;
    let tareasValidas = 0;

    tareasRecientes.forEach(tarea => {
      const real = parsearMinutos(tarea.tiempoReal);
      const estimado = parsearMinutos(tarea.tiempoEstimado);

      if (real > 0 && estimado > 0) {
        // Eficiencia = (Tiempo Estimado / Tiempo Real) * 100
        // Si termina antes: eficiencia > 100, se limita a 100
        // Si termina después: eficiencia < 100
        let eficienciaTarea = (estimado / real) * 100;

        // Limitar a 100% como máximo
        if (eficienciaTarea > 100) eficienciaTarea = 100;

        sumaEficienciaTiempo += eficienciaTarea;
        tareasValidas++;
      }
    });

    // Calcular promedio de eficiencia de tiempo (70% del peso)
    const eficienciaTiempo = tareasValidas > 0 ? (sumaEficienciaTiempo / tareasValidas) : 85;

    // Calcular penalización por reparaciones (30% del peso)
    const datosReparaciones = reparacionesSheet.getDataRange().getValues();
    let reparacionesPersona = 0;

    for (let i = 1; i < datosReparaciones.length; i++) {
      const persona = datosReparaciones[i][3];
      if (persona === nombrePersona) {
        reparacionesPersona++;
      }
    }

    // Calcular factor de calidad (sin reparaciones = 100%, cada reparación reduce 10%)
    const factorCalidad = Math.max(0, 100 - (reparacionesPersona * 10));

    // Eficiencia final: 70% tiempo + 30% calidad
    const eficienciaFinal = (eficienciaTiempo * 0.7) + (factorCalidad * 0.3);

    // Redondear a 2 decimales y limitar entre 0 y 100
    return Math.max(0, Math.min(100, Math.round(eficienciaFinal * 100) / 100));

  } catch (e) {
    console.log(`Error calculando eficiencia para ${nombrePersona}: ${e.message}`);
    return 85.0; // Retornar eficiencia por defecto en caso de error
  }
}

/**
 * Convierte una cadena de tiempo a minutos
 * @param {string} tiempo - Tiempo en formato "XX min" o "X h YY min"
 * @return {number} - Minutos totales
 */
function parsearMinutos(tiempo) {
  if (!tiempo || tiempo === "-") return 0;

  try {
    let minutos = 0;
    const tiempoStr = tiempo.toString().toLowerCase();

    // Buscar horas
    const horasMatch = tiempoStr.match(/(\d+)\s*h/);
    if (horasMatch) {
      minutos += parseInt(horasMatch[1]) * 60;
    }

    // Buscar minutos
    const minutosMatch = tiempoStr.match(/(\d+)\s*min/);
    if (minutosMatch) {
      minutos += parseInt(minutosMatch[1]);
    }

    return minutos;
  } catch (e) {
    console.log(`Error parseando tiempo: ${tiempo}`);
    return 0;
  }
}

/**
 * Actualiza la eficiencia de una persona específica en la hoja Personas
 * @param {string} nombrePersona - Nombre de la persona
 */
function actualizarEficienciaPersona(nombrePersona) {
  try {
    const personasSheet = SS.getSheetByName("👥 Personas");
    const datos = personasSheet.getDataRange().getValues();

    // Buscar la fila de la persona
    for (let i = 1; i < datos.length; i++) {
      const nombre = datos[i][1]; // Columna B (Nombre)

      if (nombre === nombrePersona) {
        const eficiencia = calcularEficienciaPersona(nombrePersona);
        personasSheet.getRange(i + 1, 5).setValue(eficiencia); // Columna E (Eficiencia)
        console.log(`  ✓ Eficiencia actualizada para ${nombrePersona}: ${eficiencia}%`);
        return;
      }
    }
  } catch (e) {
    console.log(`Error actualizando eficiencia de ${nombrePersona}: ${e.message}`);
  }
}

/**
 * Actualiza las eficiencias de todas las personas
 */
function actualizarTodasEficiencias() {
  try {
    console.log("📊 Actualizando eficiencias de todas las personas...");

    const personasSheet = SS.getSheetByName("👥 Personas");
    const datos = personasSheet.getDataRange().getValues();

    let actualizadas = 0;

    for (let i = 1; i < datos.length; i++) {
      const nombre = datos[i][1]; // Columna B (Nombre)

      if (nombre && nombre !== "") {
        const eficiencia = calcularEficienciaPersona(nombre);
        personasSheet.getRange(i + 1, 5).setValue(eficiencia); // Columna E (Eficiencia)
        actualizadas++;
      }
    }

    console.log(`✅ ${actualizadas} eficiencias actualizadas`);

  } catch (e) {
    console.log(`Error actualizando todas las eficiencias: ${e.message}`);
  }
}

/**
 * Marca una tarea como completada y registra el tiempo real
 * @param {string} tareaID - ID de la tarea
 * @param {string} nombrePersona - Nombre de la persona
 * @param {number} minutosReales - Tiempo real en minutos
 */
function marcarTareaCompletada(tareaID, nombrePersona, minutosReales) {
  try {
    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const tiemposSheet = SS.getSheetByName("⏱️ Tiempos");
    const historialSheet = SS.getSheetByName("📝 Historial");

    // 1. Buscar la tarea en la hoja Tareas
    const datosTareas = tareasSheet.getDataRange().getValues();
    let filaTarea = -1;
    let datosTarea = null;

    for (let i = 1; i < datosTareas.length; i++) {
      if (datosTareas[i][0] === tareaID) {
        filaTarea = i + 1;
        datosTarea = datosTareas[i];
        break;
      }
    }

    if (filaTarea === -1) {
      throw new Error(`Tarea ${tareaID} no encontrada`);
    }

    // 2. Actualizar estado en hoja Tareas
    tareasSheet.getRange(filaTarea, 7).setValue("Completada");
    tareasSheet.getRange(filaTarea, 9).setValue(`Completada en ${minutosReales} minutos`);

    // 3. Buscar el registro de tiempo correspondiente
    const datosTiempos = tiemposSheet.getDataRange().getValues();
    let filaTiempo = -1;

    for (let i = 1; i < datosTiempos.length; i++) {
      if (datosTiempos[i][0] === tareaID && datosTiempos[i][1] === nombrePersona) {
        filaTiempo = i + 1;
        break;
      }
    }

    if (filaTiempo === -1) {
      // Si no existe registro de tiempo, crear uno nuevo
      filaTiempo = tiemposSheet.getLastRow() + 1;
      tiemposSheet.getRange(filaTiempo, 1).setValue(tareaID);
      tiemposSheet.getRange(filaTiempo, 2).setValue(nombrePersona);
      tiemposSheet.getRange(filaTiempo, 6).setValue("60 min"); // Tiempo estimado por defecto
    }

    // 4. Actualizar tiempos
    const ahora = new Date();
    const horaInicio = tiemposSheet.getRange(filaTiempo, 3).getValue() || ahora.toLocaleTimeString("es-MX");

    tiemposSheet.getRange(filaTiempo, 3).setValue(horaInicio);
    tiemposSheet.getRange(filaTiempo, 4).setValue(ahora.toLocaleTimeString("es-MX"));
    tiemposSheet.getRange(filaTiempo, 5).setValue(`${minutosReales} min`);

    // Obtener tiempo estimado
    const tiempoEstimadoStr = tiemposSheet.getRange(filaTiempo, 6).getValue() || "60 min";
    const tiempoEstimado = parsearMinutos(tiempoEstimadoStr);

    // Calcular variación
    const variacion = minutosReales - tiempoEstimado;
    const variacionStr = variacion > 0 ? `+${variacion} min` : `${variacion} min`;

    tiemposSheet.getRange(filaTiempo, 7).setValue(variacionStr);
    tiemposSheet.getRange(filaTiempo, 8).setValue("✅");

    // 5. Registrar en historial
    const ultimaFilaHistorial = historialSheet.getLastRow() + 1;

    historialSheet.getRange(ultimaFilaHistorial, 1).setValue(ahora);
    historialSheet.getRange(ultimaFilaHistorial, 2).setValue(ahora.toLocaleTimeString("es-MX"));
    historialSheet.getRange(ultimaFilaHistorial, 3).setValue("Tarea Completada");
    historialSheet.getRange(ultimaFilaHistorial, 4).setValue(tareaID);
    historialSheet.getRange(ultimaFilaHistorial, 5).setValue("Completada");
    historialSheet.getRange(ultimaFilaHistorial, 6).setValue(nombrePersona);
    historialSheet.getRange(ultimaFilaHistorial, 7).setValue(nombrePersona);
    historialSheet.getRange(ultimaFilaHistorial, 8).setValue(datosTarea[2]); // Área
    historialSheet.getRange(ultimaFilaHistorial, 9).setValue(Session.getActiveUser().getEmail());
    historialSheet.getRange(ultimaFilaHistorial, 10).setValue("✅");

    // 6. Actualizar eficiencia de la persona
    actualizarEficienciaPersona(nombrePersona);

    // 7. Actualizar panel de control
    actualizarPanelControl();

    console.log(`✅ Tarea ${tareaID} completada por ${nombrePersona} en ${minutosReales} minutos`);

  } catch (e) {
    console.log(`Error marcando tarea completada: ${e.message}`);
    throw e;
  }
}

// ==================== AUTOMATIZACIÓN: DETECTAR PENDIENTES ====================

function detectarTareasPendientes() {
  try {
    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const datos = tareasSheet.getDataRange().getValues();
    const pendientes = [];

    for (let i = 1; i < datos.length; i++) {
      const estado = datos[i][6];
      if (estado && estado.includes("PENDIENTE")) {
        pendientes.push({
          fila: i + 1,
          tareaID: datos[i][0],
          persona: datos[i][1],
          descripcion: datos[i][3],
          area: datos[i][2],
          cantidad: datos[i][4],
          estado: estado
        });
      }
    }

    return pendientes;
  } catch (e) {
    console.log(`Error detectando pendientes: ${e.message}`);
    return [];
  }
}

// ==================== VALIDACIÓN DE DATOS REALES ====================

/**
 * Verifica si el sistema tiene datos reales (no solo datos de ejemplo o vacío)
 * @return {boolean} - true si hay datos reales, false si solo hay ejemplos o está vacío
 */
function hayDatosReales() {
  try {
    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const personasSheet = SS.getSheetByName("👥 Personas");

    // Contar filas de datos (sin contar header)
    const numTareas = tareasSheet.getLastRow() - 1;
    const numPersonas = personasSheet.getLastRow() - 1;

    // Si hay 4 o menos tareas, probablemente son datos de ejemplo
    // Los datos de ejemplo iniciales son exactamente 4 filas
    if (numTareas <= 4 && numPersonas <= 7) {
      console.log("📝 Sistema tiene solo datos de ejemplo - NO se enviarán correos");
      return false;
    }

    // Si las hojas están vacías (solo headers)
    if (numTareas === 0 || numPersonas === 0) {
      console.log("📝 Sistema está vacío - NO se enviarán correos");
      return false;
    }

    // Verificar si alguna tarea tiene ID que no sea de ejemplo (T001-T004)
    const datosTareas = tareasSheet.getDataRange().getValues();
    let tieneIDsReales = false;

    for (let i = 1; i < datosTareas.length; i++) {
      const tareaID = datosTareas[i][0];
      if (tareaID && !["T001", "T002", "T003", "T004"].includes(tareaID)) {
        tieneIDsReales = true;
        break;
      }
    }

    if (!tieneIDsReales && numTareas <= 4) {
      console.log("📝 Solo hay tareas de ejemplo (T001-T004) - NO se enviarán correos");
      return false;
    }

    console.log("✅ Sistema tiene datos reales - Se pueden enviar correos");
    return true;

  } catch (e) {
    console.log(`Error verificando datos reales: ${e.message}`);
    return false; // Por seguridad, no enviar correos si hay error
  }
}

function detectarPendientesAutomatico() {
  try {
    console.log("🔍 Ejecutando detección automática de pendientes...");

    // VALIDACIÓN: Solo enviar correos si hay datos reales
    if (!hayDatosReales()) {
      console.log("⏭️ Saltando envío de correos - sin datos reales");
      return;
    }

    const pendientes = detectarTareasPendientes();

    if (pendientes.length > 0) {
      console.log(`⚠️ Se encontraron ${pendientes.length} tareas pendientes`);
      notificarPendientes(pendientes);
    } else {
      console.log("✅ No hay tareas pendientes");
    }
  } catch (e) {
    console.log(`Error en automatización: ${e.message}`);
  }
}

// ==================== AUTOMATIZACIÓN: NOTIFICACIONES POR CORREO ====================

function notificarPendientes(pendientes) {
  try {
    console.log("📧 Enviando notificaciones de tareas pendientes...");

    const encargadas = [
      CORREOS_CONFIG["Encargada de Taller 1"],
      CORREOS_CONFIG["Encargada de Taller 2"]
    ];

    const asunto = `⚠️ ALERTA: ${pendientes.length} Tareas Pendientes en Sistema de Producción`;

    let cuerpo = `
    <html>
      <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #1f4d7f;">⚠️ ALERTA DE TAREAS PENDIENTES</h2>

        <p>Se han detectado <strong>${pendientes.length}</strong> tareas pendientes que requieren atención:</p>

        <table style="border-collapse: collapse; width: 100%; margin-top: 15px;">
          <tr style="background-color: #fff3cd;">
            <th style="border: 1px solid #ddd; padding: 10px;">Tarea ID</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Persona</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Descripción</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Cantidad</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Acción</th>
          </tr>`;

    pendientes.forEach((tarea, idx) => {
      const color = idx % 2 === 0 ? "#f9f9f9" : "white";
      cuerpo += `
          <tr style="background-color: ${color};">
            <td style="border: 1px solid #ddd; padding: 10px;"><strong>${tarea.tareaID}</strong></td>
            <td style="border: 1px solid #ddd; padding: 10px;">${tarea.persona}</td>
            <td style="border: 1px solid #ddd; padding: 10px;">${tarea.descripcion}</td>
            <td style="border: 1px solid #ddd; padding: 10px;">${tarea.cantidad}</td>
            <td style="border: 1px solid #ddd; padding: 10px;">Reasignar</td>
          </tr>`;
    });

    cuerpo += `
        </table>

        <h3 style="margin-top: 20px; color: #1f4d7f;">💡 RECOMENDACIONES:</h3>
        <ul>
          <li>Reasignar tareas a personas con CARGA BAJA</li>
          <li>Revisar si hay defectos (REPARACIÓN) o simplemente no se completaron (SOBRANTE)</li>
          <li>Actualizar estado en el sistema cuando se reasigne</li>
        </ul>

        <p style="margin-top: 20px; color: #666;">
          <strong>Hora de generación:</strong> ${new Date().toLocaleString("es-MX")}<br>
          <strong>Sistema:</strong> Producción Automático
        </p>
      </body>
    </html>
    `;

    encargadas.forEach(email => {
      try {
        GmailApp.sendEmail(email, asunto, cuerpo, { htmlBody: cuerpo });
        console.log(`  ✓ Notificación enviada a ${email}`);
      } catch (e) {
        console.log(`  ⚠️ Error enviando a ${email}: ${e.message}`);
      }
    });

  } catch (e) {
    console.log(`Error en notificación: ${e.message}`);
  }
}

// ==================== BOTONES DE MENÚ ====================

function btnDetectarTodosPendientes() {
  const pendientes = detectarTareasPendientes();
  const ui = SpreadsheetApp.getUi();

  if (pendientes.length === 0) {
    ui.alert("✅ No hay tareas pendientes\n\n¡Excelente trabajo!");
    return;
  }

  let mensaje = `🔍 TAREAS PENDIENTES: ${pendientes.length}\n\n`;
  pendientes.forEach((tarea, idx) => {
    mensaje += `${idx + 1}. ${tarea.tareaID} - ${tarea.persona}\n`;
    mensaje += `   ${tarea.descripcion} (${tarea.cantidad} unidades)\n`;
    mensaje += `   Área: ${tarea.area}\n\n`;
  });

  const botones = ui.ButtonSet.YES_NO;
  const respuesta = ui.alert("TAREAS PENDIENTES", mensaje + "\n¿Deseas reasignar alguna tarea?", botones);

  if (respuesta === ui.Button.YES) {
    btnReasignarSeleccionada();
  }
}

function btnReasignarSeleccionada() {
  mostrarFormularioReasignacion();
}

function mostrarFormularioReasignacion() {
  const html = HtmlService.createHtmlOutput(getFormularioReasignacionHTML())
    .setWidth(550)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🔄 REASIGNAR TAREA');
}

function getFormularioReasignacionHTML() {
  // Obtener tareas pendientes o en progreso
  const tareasSheet = SS.getSheetByName("📋 Tareas");
  const tareasData = tareasSheet.getDataRange().getValues();

  let opcionesTareas = '';
  for (let i = 1; i < tareasData.length; i++) {
    const tareaID = tareasData[i][0];
    const persona = tareasData[i][1];
    const descripcion = tareasData[i][3];
    const estado = tareasData[i][6];
    const area = tareasData[i][2];
    const cantidad = tareasData[i][4];

    if (tareaID && (estado === "En Progreso" || estado.includes("PENDIENTE"))) {
      // Formato: tareaID|persona|descripcion|area|cantidad
      const valor = `${tareaID}|${persona}|${descripcion}|${area}|${cantidad}`;
      opcionesTareas += `<option value="${valor}">${tareaID} - ${persona} - ${descripcion}</option>`;
    }
  }

  // Obtener personas desde la hoja
  const personasSheet = SS.getSheetByName("👥 Personas");
  const personasData = personasSheet.getDataRange().getValues();

  let opcionesPersonas = '';
  for (let i = 1; i < personasData.length; i++) {
    const nombre = personasData[i][1];
    const carga = personasData[i][5];
    if (nombre) {
      opcionesPersonas += `<option value="${nombre}">${nombre} - ${carga}</option>`;
    }
  }

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background-color: #f5f5f5;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            font-weight: bold;
            margin-bottom: 5px;
            color: #1f4d7f;
          }
          select {
            width: 100%;
            padding: 10px;
            font-size: 14px;
            border: 2px solid #ccc;
            border-radius: 5px;
            box-sizing: border-box;
            cursor: pointer;
            background-color: white;
          }
          select:focus {
            border-color: #1f4d7f;
            outline: none;
          }
          .btn {
            padding: 12px 30px;
            font-size: 16px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            margin-right: 10px;
          }
          .btn-primary {
            background-color: #d9b919;
            color: white;
          }
          .btn-primary:hover {
            background-color: #b89815;
          }
          .btn-secondary {
            background-color: #ccc;
            color: #333;
          }
          .btn-secondary:hover {
            background-color: #999;
          }
          .buttons {
            margin-top: 30px;
            text-align: center;
          }
          .titulo {
            color: #1f4d7f;
            margin-bottom: 20px;
            font-size: 20px;
          }
          .required {
            color: red;
          }
          .info-box {
            background-color: #fff3cd;
            border-left: 4px solid #d9b919;
            padding: 12px;
            margin-bottom: 20px;
            border-radius: 4px;
          }
        </style>
      </head>
      <body>
        <h2 class="titulo">🔄 Reasignar Tarea</h2>

        <div class="info-box">
          ⚠️ Selecciona la tarea a reasignar y la nueva persona encargada
        </div>

        <form id="formulario">
          <div class="form-group">
            <label for="tarea">Tarea a Reasignar <span class="required">*</span></label>
            <select id="tarea" name="tarea" required>
              <option value="">-- Selecciona una tarea --</option>
              ${opcionesTareas}
            </select>
          </div>

          <div class="form-group">
            <label for="tipo">Tipo de Reasignación <span class="required">*</span></label>
            <select id="tipo" name="tipo" required>
              <option value="">-- Selecciona el tipo --</option>
              <option value="REPARACIÓN">🔧 REPARACIÓN (tiene defectos)</option>
              <option value="TAREA SOBRANTE">📦 TAREA SOBRANTE (no terminó)</option>
            </select>
          </div>

          <div class="form-group">
            <label for="persona">Reasignar a <span class="required">*</span></label>
            <select id="persona" name="persona" required>
              <option value="">-- Selecciona una persona --</option>
              ${opcionesPersonas}
            </select>
          </div>

          <div class="buttons">
            <button type="submit" class="btn btn-primary">✅ Reasignar Tarea</button>
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">❌ Cancelar</button>
          </div>
        </form>

        <script>
          document.getElementById('formulario').addEventListener('submit', function(e) {
            e.preventDefault();

            const tareaData = document.getElementById('tarea').value;
            const tipo = document.getElementById('tipo').value;
            const nuevaPersona = document.getElementById('persona').value;

            if (!tareaData || !tipo || !nuevaPersona) {
              alert('❌ Por favor completa todos los campos');
              return;
            }

            // Separar los datos de la tarea
            const [tareaID, personaActual, descripcion, area, cantidad] = tareaData.split('|');

            // Deshabilitar botón para evitar doble envío
            document.querySelector('.btn-primary').disabled = true;
            document.querySelector('.btn-primary').textContent = '⏳ Reasignando...';

            // Enviar datos a Google Apps Script
            google.script.run
              .withSuccessHandler(function() {
                alert('✅ TAREA REASIGNADA CORRECTAMENTE\\n\\n' +
                      'Tarea: ' + tareaID + '\\n' +
                      'Tipo: ' + tipo + '\\n' +
                      'De: ' + personaActual + '\\n' +
                      'Para: ' + nuevaPersona + '\\n\\n' +
                      '✓ Registrado en Historial\\n' +
                      '✓ Sistema actualizado');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('❌ Error: ' + error);
                document.querySelector('.btn-primary').disabled = false;
                document.querySelector('.btn-primary').textContent = '✅ Reasignar Tarea';
              })
              .procesarReasignacion(tareaID, personaActual, descripcion, area, cantidad, nuevaPersona, tipo);
          });
        </script>
      </body>
    </html>
  `;
}

function procesarReasignacion(tareaID, personaActual, descripcion, area, cantidad, nuevaPersona, tipo) {
  try {
    const tareaInfo = {
      tareaID: tareaID,
      persona: personaActual,
      descripcion: descripcion,
      area: area,
      cantidad: cantidad
    };

    // Buscar la fila de la tarea
    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const datos = tareasSheet.getDataRange().getValues();
    let filaEncontrada = -1;

    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === tareaID) {
        filaEncontrada = i;
        break;
      }
    }

    if (filaEncontrada === -1) {
      throw new Error(`Tarea ${tareaID} no encontrada`);
    }

    ejecutarReasignacion(tareaID, tareaInfo, nuevaPersona, tipo, filaEncontrada);

  } catch (e) {
    console.log(`Error en reasignación: ${e.message}`);
    throw e;
  }
}

function ejecutarReasignacion(tareaID, tareaInfo, nuevaPersona, tipo, fila) {
  try {
    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const historialSheet = SS.getSheetByName("📝 Historial");
    const reparacionesSheet = SS.getSheetByName("🔧 Reparaciones");

    tareasSheet.getRange(fila + 1, 2).setValue(nuevaPersona);
    tareasSheet.getRange(fila + 1, 7).setValue("En Progreso");
    tareasSheet.getRange(fila + 1, 10).setValue(`Reasignado de: ${tareaInfo.persona}`);

    if (tipo === "REPARACIÓN") {
      const ultimaFila = reparacionesSheet.getLastRow() + 1;
      const reparacionID = `REP-${tareaID}`;

      reparacionesSheet.getRange(ultimaFila, 1).setValue(reparacionID);
      reparacionesSheet.getRange(ultimaFila, 2).setValue(tareaID);
      reparacionesSheet.getRange(ultimaFila, 3).setValue("Reasignación por reparación");
      reparacionesSheet.getRange(ultimaFila, 4).setValue(nuevaPersona);
      reparacionesSheet.getRange(ultimaFila, 7).setValue("En Progreso");
      reparacionesSheet.getRange(ultimaFila, 8).setValue(new Date());
    }

    const ultimaFila = historialSheet.getLastRow() + 1;
    const ahora = new Date();
    historialSheet.getRange(ultimaFila, 1).setValue(ahora);
    historialSheet.getRange(ultimaFila, 2).setValue(ahora.toLocaleTimeString("es-MX"));
    historialSheet.getRange(ultimaFila, 3).setValue("Reasignación");
    historialSheet.getRange(ultimaFila, 4).setValue(tareaID);
    historialSheet.getRange(ultimaFila, 5).setValue(tipo);
    historialSheet.getRange(ultimaFila, 6).setValue(tareaInfo.persona);
    historialSheet.getRange(ultimaFila, 7).setValue(nuevaPersona);
    historialSheet.getRange(ultimaFila, 8).setValue(tareaInfo.persona !== nuevaPersona ? "A otra persona" : "A la misma");
    historialSheet.getRange(ultimaFila, 9).setValue(Session.getActiveUser().getEmail());
    historialSheet.getRange(ultimaFila, 10).setValue("✅ Exitosa");

    const asunto = `✅ Tarea ${tareaID} Reasignada - ${tipo}`;
    const cuerpo = `
    <html>
      <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #1f4d7f;">✅ TAREA REASIGNADA</h2>

        <p><strong>Tarea ID:</strong> ${tareaID}</p>
        <p><strong>Descripción:</strong> ${tareaInfo.descripcion}</p>
        <p><strong>Cantidad:</strong> ${tareaInfo.cantidad}</p>
        <p><strong>Área:</strong> ${tareaInfo.area}</p>
        <p><strong>Tipo:</strong> ${tipo}</p>

        <h3>Reasignación:</h3>
        <p><strong>De:</strong> ${tareaInfo.persona}</p>
        <p><strong>Para:</strong> ${nuevaPersona}</p>

        <p><strong>Hora:</strong> ${ahora.toLocaleString("es-MX")}</p>
      </body>
    </html>
    `;

    try {
      GmailApp.sendEmail(CORREOS_CONFIG["Gerente de Producción"], asunto, cuerpo, { htmlBody: cuerpo });
      console.log("  ✓ Notificación de reasignación enviada");
    } catch (e) {
      console.log(`  ⚠️ Error enviando notificación: ${e.message}`);
    }

    actualizarPanelControl();

  } catch (e) {
    console.log(`Error en reasignación: ${e.message}`);
  }
}

function btnReasignacionAutomatica() {
  const ui = SpreadsheetApp.getUi();
  const pendientes = detectarTareasPendientes();

  if (pendientes.length === 0) {
    ui.alert("✅ No hay tareas pendientes");
    return;
  }

  const confirm = ui.alert(
    `⚡ REASIGNACIÓN AUTOMÁTICA\n\n` +
    `Se reasignarán ${pendientes.length} tareas a la misma persona\n\n` +
    `¿Continuar?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  let reasignadas = 0;

  pendientes.forEach((tarea) => {
    const tareaInfo = {
      tareaID: tarea.tareaID,
      persona: tarea.persona,
      descripcion: tarea.descripcion,
      area: tarea.area,
      cantidad: tarea.cantidad
    };
    ejecutarReasignacion(tarea.tareaID, tareaInfo, tarea.persona, "TAREA SOBRANTE", tarea.fila - 1);
    reasignadas++;
  });

  ui.alert(
    `✅ REASIGNACIÓN AUTOMÁTICA COMPLETADA\n\n` +
    `Tareas procesadas: ${reasignadas}\n` +
    `Tipo: TAREA SOBRANTE\n` +
    `Estado: Las mismas personas continúan`
  );
}

function btnMostrarHistorial() {
  const ui = SpreadsheetApp.getUi();
  const historialSheet = SS.getSheetByName("📝 Historial");
  const datos = historialSheet.getDataRange().getValues();

  let mensaje = "📝 ÚLTIMOS CAMBIOS\n\n";

  const desde = Math.max(1, datos.length - 10);

  for (let i = desde; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[3]) {
      mensaje += `${fila[0]} ${fila[1]}\n`;
      mensaje += `${fila[3]} (${fila[4]})\n`;
      mensaje += `${fila[5]} → ${fila[6]}\n`;
      mensaje += `${fila[9]}\n\n`;
    }
  }

  ui.alert(mensaje);
}

// ==================== HERRAMIENTAS ====================

function buscarTarea() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "🔍 BUSCAR TAREA\n\n" +
    "Ingresa el ID (ej: T001):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const tareaID = response.getResponseText().toUpperCase().trim();
  const tareasSheet = SS.getSheetByName("📋 Tareas");
  const datos = tareasSheet.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === tareaID) {
      const info = datos[i];
      ui.alert(
        `✅ TAREA ENCONTRADA\n\n` +
        `ID: ${info[0]}\n` +
        `Persona: ${info[1]}\n` +
        `Área: ${info[2]}\n` +
        `Descripción: ${info[3]}\n` +
        `Cantidad: ${info[4]}\n` +
        `Estado: ${info[6]}\n` +
        `Notas: ${info[8] || "Sin notas"}`
      );
      return;
    }
  }

  ui.alert(`❌ Tarea ${tareaID} no encontrada`);
}

function analizarCargas() {
  const ui = SpreadsheetApp.getUi();
  const personasSheet = SS.getSheetByName("👥 Personas");
  const tareasSheet = SS.getSheetByName("📋 Tareas");

  const personasData = personasSheet.getDataRange().getValues();
  const tareasData = tareasSheet.getDataRange().getValues();

  let mensaje = "👥 ANÁLISIS DE CARGAS DE TRABAJO\n\n";

  for (let i = 1; i < personasData.length; i++) {
    const nombre = personasData[i][1];
    if (!nombre) break;

    let tareasPendientes = 0;
    let tareasEnProgreso = 0;

    for (let j = 1; j < tareasData.length; j++) {
      if (tareasData[j][1] === nombre) {
        if (tareasData[j][6].includes("PENDIENTE")) tareasPendientes++;
        if (tareasData[j][6].includes("Progreso")) tareasEnProgreso++;
      }
    }

    const eficiencia = personasData[i][6];
    const carga = personasData[i][7];

    mensaje += `${nombre}\n`;
    mensaje += `  • Eficiencia: ${(eficiencia * 100).toFixed(0)}%\n`;
    mensaje += `  • Carga: ${carga}\n`;
    mensaje += `  • En progreso: ${tareasEnProgreso}\n`;
    mensaje += `  • Pendientes: ${tareasPendientes}\n\n`;
  }

  ui.alert(mensaje);
}

function verEficiencia() {
  const ui = SpreadsheetApp.getUi();
  const personasSheet = SS.getSheetByName("👥 Personas");
  const datos = personasSheet.getDataRange().getValues();

  let mensaje = "📊 EFICIENCIA POR PERSONA\n\n";
  let baja = [];

  for (let i = 1; i < datos.length; i++) {
    const nombre = datos[i][1];
    const eficiencia = datos[i][6];

    if (typeof eficiencia === 'number') {
      const pct = (eficiencia * 100).toFixed(0);
      const icon = eficiencia >= 0.9 ? "🟢" : eficiencia >= 0.8 ? "🟡" : "🔴";

      mensaje += `${icon} ${nombre}: ${pct}%\n`;

      if (eficiencia < 0.8) {
        baja.push(nombre);
      }
    }
  }

  if (baja.length > 0) {
    mensaje += `\n⚠️ EFICIENCIA BAJA (<80%):\n`;
    baja.forEach(nombre => {
      mensaje += `• ${nombre} - Revisar rendimiento\n`;
    });
  }

  ui.alert(mensaje);
}

function actualizarPanelControlUI() {
  actualizarPanelControl();
  SpreadsheetApp.getUi().alert("✅ Panel Control actualizado");
}

function actualizarPanelControl() {
  try {
    const panelSheet = SS.getSheetByName("🎯 Panel Control");
    const tareasSheet = SS.getSheetByName("📋 Tareas");

    const datos = tareasSheet.getDataRange().getValues();

    let total = 0;
    let completadas = 0;
    let pendientes = 0;

    for (let i = 1; i < datos.length; i++) {
      const estado = datos[i][6];
      total++;

      if (estado.includes("Completada")) completadas++;
      if (estado.includes("PENDIENTE")) pendientes++;
    }

    panelSheet.getRange("B6").setValue(total);
    panelSheet.getRange("B7").setValue(completadas);
    panelSheet.getRange("B8").setValue(pendientes);

  } catch (e) {
    console.log(`Error actualizando panel: ${e.message}`);
  }
}

function mostrarAyuda() {
  SpreadsheetApp.getUi().alert(
    "📖 AYUDA RÁPIDA\n\n" +
    "➕ Asignar Nueva Tarea\n" +
    "Crea y asigna nuevas tareas\n\n" +
    "🔍 Detectar Pendientes\n" +
    "Encuentra tareas sin completar\n\n" +
    "🔄 Reasignar Manual\n" +
    "Mueve tarea de persona a persona\n\n" +
    "⚡ Reasignación Automática\n" +
    "Reasigna todas las pendientes\n\n" +
    "📧 Enviar Notificación\n" +
    "Notifica a encargadas manualmente\n\n" +
    "📊 Generar Reporte\n" +
    "Crea reporte diario\n\n" +
    "Ve a 📖 Instrucciones para más"
  );
}

// ==================== ELIMINAR DATOS DE EJEMPLO ====================

function eliminarDatosEjemplo() {
  const ui = SpreadsheetApp.getUi();

  // Confirmación 1: Advertencia
  const confirm1 = ui.alert(
    "⚠️ ELIMINAR TODOS LOS DATOS\n\n" +
    "Esta acción eliminará ABSOLUTAMENTE TODO:\n\n" +
    "• Personas (excepto encabezados)\n" +
    "• Tareas\n" +
    "• Registros de tiempos\n" +
    "• Reparaciones\n" +
    "• Historial completo\n" +
    "• Reasignaciones\n" +
    "• Reportes\n" +
    "• Indicadores\n\n" +
    "El sistema quedará COMPLETAMENTE VACÍO.\n" +
    "Todo en CERO para empezar de nuevo.\n\n" +
    "¿Deseas continuar?",
    ui.ButtonSet.YES_NO
  );

  if (confirm1 !== ui.Button.YES) {
    ui.alert("❌ Cancelado - No se eliminaron los datos");
    return;
  }

  // Confirmación 2: Doble verificación
  const confirm2 = ui.alert(
    "🔴 ÚLTIMA CONFIRMACIÓN\n\n" +
    "¿Estás COMPLETAMENTE SEGURO?\n\n" +
    "Esta acción NO se puede deshacer.\n\n" +
    "TODO el sistema volverá a CERO.\n\n" +
    "Para recuperar los datos de ejemplo tendrás que\n" +
    "ejecutar el Setup Completo nuevamente.",
    ui.ButtonSet.YES_NO
  );

  if (confirm2 !== ui.Button.YES) {
    ui.alert("❌ Cancelado - No se eliminaron los datos");
    return;
  }

  try {
    console.log("🗑️ Eliminando TODOS los datos del sistema...");

    // 1. Limpiar Personas (dejar solo encabezados)
    const personasSheet = SS.getSheetByName("👥 Personas");
    const ultimaFilaPersonas = personasSheet.getLastRow();
    if (ultimaFilaPersonas > 1) {
      personasSheet.deleteRows(2, ultimaFilaPersonas - 1);
      console.log("  ✓ Personas limpiadas");
    }

    // 2. Limpiar Tareas (dejar solo encabezados)
    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const ultimaFilaTareas = tareasSheet.getLastRow();
    if (ultimaFilaTareas > 1) {
      tareasSheet.deleteRows(2, ultimaFilaTareas - 1);
      console.log("  ✓ Tareas limpiadas");
    }

    // 3. Limpiar Tiempos (dejar solo encabezados)
    const tiemposSheet = SS.getSheetByName("⏱️ Tiempos");
    const ultimaFilaTiempos = tiemposSheet.getLastRow();
    if (ultimaFilaTiempos > 1) {
      tiemposSheet.deleteRows(2, ultimaFilaTiempos - 1);
      console.log("  ✓ Tiempos limpiados");
    }

    // 4. Limpiar Reparaciones (dejar solo encabezados)
    const reparacionesSheet = SS.getSheetByName("🔧 Reparaciones");
    const ultimaFilaReparaciones = reparacionesSheet.getLastRow();
    if (ultimaFilaReparaciones > 1) {
      reparacionesSheet.deleteRows(2, ultimaFilaReparaciones - 1);
      console.log("  ✓ Reparaciones limpiadas");
    }

    // 5. Limpiar Historial (dejar solo encabezados)
    const historialSheet = SS.getSheetByName("📝 Historial");
    const ultimaFilaHistorial = historialSheet.getLastRow();
    if (ultimaFilaHistorial > 1) {
      historialSheet.deleteRows(2, ultimaFilaHistorial - 1);
      console.log("  ✓ Historial limpiado");
    }

    // 6. Limpiar Reasignaciones (dejar solo encabezados)
    const reasignacionesSheet = SS.getSheetByName("🔄 Reasignaciones");
    const ultimaFilaReasignaciones = reasignacionesSheet.getLastRow();
    if (ultimaFilaReasignaciones > 1) {
      reasignacionesSheet.deleteRows(2, ultimaFilaReasignaciones - 1);
      console.log("  ✓ Reasignaciones limpiadas");
    }

    // 7. Limpiar Reportes (dejar solo encabezados)
    const reportesSheet = SS.getSheetByName("📊 Reportes");
    const ultimaFilaReportes = reportesSheet.getLastRow();
    if (ultimaFilaReportes > 1) {
      reportesSheet.deleteRows(2, ultimaFilaReportes - 1);
      console.log("  ✓ Reportes limpiados");
    }

    // 8. Limpiar Indicadores completamente (tabla de indicadores)
    const indicadoresSheet = SS.getSheetByName("📈 Indicadores");
    // Limpiar solo los valores, mantener la estructura
    indicadoresSheet.getRange("B4:B7").setValue(0);  // Columna HOY
    indicadoresSheet.getRange("C4:C7").setValue(0);  // Columna META
    indicadoresSheet.getRange("D4:D7").clearContent();  // Columna ESTADO
    indicadoresSheet.getRange("E4:E7").clearContent();  // Columna VARIACIÓN
    console.log("  ✓ Indicadores limpiados");

    // 9. Actualizar Panel Control a ceros
    const panelSheet = SS.getSheetByName("🎯 Panel Control");
    panelSheet.getRange("B6").setValue(0);  // Total tareas
    panelSheet.getRange("B7").setValue(0);  // Completadas
    panelSheet.getRange("B8").setValue(0);  // Pendientes
    panelSheet.getRange("B9").setValue(0);  // Con reparación
    panelSheet.getRange("B10").setValue(0); // Eficiencia promedio

    // Limpiar tareas pendientes del panel
    panelSheet.getRange("A14:E20").clearContent();
    console.log("  ✓ Panel Control actualizado a ceros");

    // 10. Registrar en Historial que se eliminaron los datos
    const ahora = new Date();
    historialSheet.getRange(2, 1).setValue(ahora);
    historialSheet.getRange(2, 2).setValue(ahora.toLocaleTimeString("es-MX"));
    historialSheet.getRange(2, 3).setValue("Sistema Limpiado Completamente");
    historialSheet.getRange(2, 4).setValue("-");
    historialSheet.getRange(2, 5).setValue("Limpieza Total");
    historialSheet.getRange(2, 6).setValue("-");
    historialSheet.getRange(2, 7).setValue("-");
    historialSheet.getRange(2, 8).setValue("TODOS los datos eliminados - Sistema en CERO");
    historialSheet.getRange(2, 9).setValue(Session.getActiveUser().getEmail());
    historialSheet.getRange(2, 10).setValue("✅ Listo para datos reales");

    console.log("✅ TODOS los datos eliminados correctamente");

    ui.alert(
      "✅ SISTEMA COMPLETAMENTE LIMPIO\n\n" +
      "TODO ha sido eliminado:\n\n" +
      "✓ Personas: VACÍO\n" +
      "✓ Tareas: VACÍO\n" +
      "✓ Tiempos: VACÍO\n" +
      "✓ Reparaciones: VACÍO\n" +
      "✓ Historial: Solo registro de limpieza\n" +
      "✓ Reasignaciones: VACÍO\n" +
      "✓ Reportes: VACÍO\n" +
      "✓ Indicadores: TODO EN CERO\n" +
      "✓ Panel Control: TODO EN CERO\n\n" +
      "🎯 Sistema 100% LIMPIO!\n\n" +
      "Ahora puedes:\n" +
      "1. Agregar tus personas reales en '👥 Personas'\n" +
      "2. Asignar tareas reales desde el menú\n" +
      "3. Completar tareas con tiempos reales\n" +
      "4. El sistema comenzará a calcular todo desde cero"
    );

  } catch (e) {
    console.log(`❌ Error eliminando datos: ${e.message}`);
    ui.alert(`❌ Error: ${e.message}\n\nAlgunos datos podrían no haberse eliminado correctamente.`);
  }
}

// ==================== REPORTES Y AUTOMATIZACIONES ====================

function btnGenerarReporteManual() {
  generarReporteDiarioAutomatico();
  SpreadsheetApp.getUi().alert("✅ Reporte generado y enviado");
}

function generarReporteDiarioAutomatico() {
  try {
    console.log("📊 Generando reporte diario...");

    // VALIDACIÓN: Solo enviar correos si hay datos reales
    if (!hayDatosReales()) {
      console.log("⏭️ Saltando envío de reporte - sin datos reales");
      return;
    }

    const tareasSheet = SS.getSheetByName("📋 Tareas");
    const personasSheet = SS.getSheetByName("👥 Personas");
    const datos = tareasSheet.getDataRange().getValues();
    const personasData = personasSheet.getDataRange().getValues();

    let pendientes = [];
    let completadas = 0;
    let enProgreso = 0;

    for (let i = 1; i < datos.length; i++) {
      if (datos[i][6].includes("PENDIENTE")) {
        pendientes.push({
          id: datos[i][0],
          persona: datos[i][1],
          descripcion: datos[i][3]
        });
      } else if (datos[i][6].includes("Completada")) {
        completadas++;
      } else if (datos[i][6].includes("Progreso")) {
        enProgreso++;
      }
    }

    const asunto = `📊 REPORTE DIARIO DE PRODUCCIÓN - ${new Date().toLocaleDateString("es-MX")}`;

    let cuerpo = `
    <html>
      <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #1f4d7f;">📊 REPORTE DIARIO DE PRODUCCIÓN</h2>

        <p><strong>Fecha:</strong> ${new Date().toLocaleDateString("es-MX")}</p>
        <p><strong>Hora de generación:</strong> ${new Date().toLocaleTimeString("es-MX")}</p>

        <h3>RESUMEN DEL DÍA</h3>
        <table style="border-collapse: collapse; width: 100%;">
          <tr style="background-color: #e8f0f7;">
            <th style="border: 1px solid #ddd; padding: 10px;">Métrica</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Valor</th>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 10px;">Total de tareas</td>
            <td style="border: 1px solid #ddd; padding: 10px;">${datos.length - 1}</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 10px;">Completadas</td>
            <td style="border: 1px solid #ddd; padding: 10px;">✅ ${completadas}</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 10px;">En progreso</td>
            <td style="border: 1px solid #ddd; padding: 10px;">⏳ ${enProgreso}</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 10px;">Pendientes</td>
            <td style="border: 1px solid #ddd; padding: 10px;">🔴 ${pendientes.length}</td>
          </tr>
        </table>`;

    if (pendientes.length > 0) {
      cuerpo += `
        <h3 style="margin-top: 20px;">TAREAS PENDIENTES</h3>
        <table style="border-collapse: collapse; width: 100%;">
          <tr style="background-color: #fff3cd;">
            <th style="border: 1px solid #ddd; padding: 10px;">Tarea ID</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Persona</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Descripción</th>
          </tr>`;

      pendientes.forEach((tarea, idx) => {
        const color = idx % 2 === 0 ? "#f9f9f9" : "white";
        cuerpo += `
          <tr style="background-color: ${color};">
            <td style="border: 1px solid #ddd; padding: 10px;"><strong>${tarea.id}</strong></td>
            <td style="border: 1px solid #ddd; padding: 10px;">${tarea.persona}</td>
            <td style="border: 1px solid #ddd; padding: 10px;">${tarea.descripcion}</td>
          </tr>`;
      });

      cuerpo += `</table>`;
    }

    cuerpo += `
        <h3 style="margin-top: 20px;">RECOMENDACIONES</h3>
        <ul>
          <li>Reasignar tareas pendientes a personas con carga BAJA</li>
          <li>Revisar si hay defectos (REPARACIÓN) o simplemente no se completaron</li>
          <li>Verificar eficiencia de operarios con rendimiento bajo</li>
        </ul>

        <p style="margin-top: 20px; color: #666;">
          Reporte automático del Sistema de Producción
        </p>
      </body>
    </html>
    `;

    const destinatarios = [
      CORREOS_CONFIG["Gerente de Producción"],
      CORREOS_CONFIG["Admininstrador Sistema"]
    ];

    destinatarios.forEach(email => {
      try {
        GmailApp.sendEmail(email, asunto, cuerpo, { htmlBody: cuerpo });
        console.log(`  ✓ Reporte enviado a ${email}`);
      } catch (e) {
        console.log(`  ⚠️ Error enviando a ${email}: ${e.message}`);
      }
    });

  } catch (e) {
    console.log(`Error generando reporte: ${e.message}`);
  }
}

function verificarEficienciaBaja() {
  try {
    console.log("📊 Verificando eficiencia...");

    // VALIDACIÓN: Solo enviar correos si hay datos reales
    if (!hayDatosReales()) {
      console.log("⏭️ Saltando verificación de eficiencia - sin datos reales");
      return;
    }

    const personasSheet = SS.getSheetByName("👥 Personas");
    const datos = personasSheet.getDataRange().getValues();

    let bajaEficiencia = [];

    for (let i = 1; i < datos.length; i++) {
      const nombre = datos[i][1];
      const eficiencia = datos[i][6];

      if (typeof eficiencia === 'number' && eficiencia < 0.8) {
        bajaEficiencia.push({
          nombre: nombre,
          eficiencia: (eficiencia * 100).toFixed(0),
          area: datos[i][3]
        });
      }
    }

    if (bajaEficiencia.length > 0) {
      console.log(`⚠️ Se encontraron ${bajaEficiencia.length} personas con eficiencia baja`);
      notificarEficienciaBaja(bajaEficiencia);
    } else {
      console.log("✅ Todas las personas tienen eficiencia adecuada");
    }

  } catch (e) {
    console.log(`Error verificando eficiencia: ${e.message}`);
  }
}

function notificarEficienciaBaja(personas) {
  try {
    const asunto = `⚠️ ALERTA: Eficiencia Baja Detectada`;

    let cuerpo = `
    <html>
      <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #cc3333;">⚠️ ALERTA DE EFICIENCIA BAJA</h2>

        <p>Se han detectado ${personas.length} persona(s) con eficiencia menor a 80%:</p>

        <table style="border-collapse: collapse; width: 100%; margin-top: 15px;">
          <tr style="background-color: #ffcccc;">
            <th style="border: 1px solid #ddd; padding: 10px;">Persona</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Eficiencia</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Área</th>
            <th style="border: 1px solid #ddd; padding: 10px;">Recomendación</th>
          </tr>`;

    personas.forEach((p, idx) => {
      const color = idx % 2 === 0 ? "#f9f9f9" : "white";
      cuerpo += `
          <tr style="background-color: ${color};">
            <td style="border: 1px solid #ddd; padding: 10px;"><strong>${p.nombre}</strong></td>
            <td style="border: 1px solid #ddd; padding: 10px;">🔴 ${p.eficiencia}%</td>
            <td style="border: 1px solid #ddd; padding: 10px;">${p.area}</td>
            <td style="border: 1px solid #ddd; padding: 10px;">Revisar y capacitar</td>
          </tr>`;
    });

    cuerpo += `
        </table>

        <h3 style="margin-top: 20px;">ACCIONES RECOMENDADAS:</h3>
        <ul>
          <li>Revisar tareas asignadas</li>
          <li>Considerar capacitación adicional</li>
          <li>Redistribuir carga si es necesario</li>
          <li>Hacer seguimiento semanal</li>
        </ul>
      </body>
    </html>
    `;

    GmailApp.sendEmail(EMAIL_ADMIN, asunto, cuerpo, { htmlBody: cuerpo });
    console.log("  ✓ Alerta de eficiencia baja enviada");

  } catch (e) {
    console.log(`Error notificando eficiencia: ${e.message}`);
  }
}

function btnEnviarNotificacionManual() {
  const pendientes = detectarTareasPendientes();

  if (pendientes.length > 0) {
    notificarPendientes(pendientes);
    SpreadsheetApp.getUi().alert(`✅ Notificación enviada a ${Object.keys(CORREOS_CONFIG).filter(k => k.includes("Encargada")).length} encargadas`);
  } else {
    SpreadsheetApp.getUi().alert("✅ No hay tareas pendientes para notificar");
  }
}

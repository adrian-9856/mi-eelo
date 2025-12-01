// ==================== HERRAMIENTAS ADICIONALES ====================

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
        `Prioridad: ${info[7]}\n` +
        `Notas: ${info[9] || "Sin notas"}`
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

// ==================== FUNCIÓN: ASIGNAR NUEVA TAREA ====================

function btnAsignarTarea() {
  const ui = SpreadsheetApp.getUi();

  // PASO 1: Descripción
  const response1 = ui.prompt(
    "➕ ASIGNAR NUEVA TAREA\n\n" +
    "Descripción de la tarea (ej: Confección Camisetas):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response1.getSelectedButton() !== ui.Button.OK) return;
  const descripcion = response1.getResponseText().trim();

  if (!descripcion) {
    ui.alert("❌ Debes ingresar una descripción");
    return;
  }

  // PASO 2: Cantidad
  const response2 = ui.prompt(
    "➕ ASIGNAR NUEVA TAREA\n\n" +
    "Cantidad (ej: 1000):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response2.getSelectedButton() !== ui.Button.OK) return;
  const cantidad = parseInt(response2.getResponseText().trim());

  if (isNaN(cantidad) || cantidad <= 0) {
    ui.alert("❌ Debes ingresar una cantidad válida");
    return;
  }

  // PASO 3: Seleccionar Área
  const response3 = ui.prompt(
    "➕ ASIGNAR NUEVA TAREA\n\n" +
    "Área (Confección / Serigrafía / Planchado / Empaque):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response3.getSelectedButton() !== ui.Button.OK) return;
  const area = response3.getResponseText().trim();

  // PASO 4: Seleccionar Persona
  const personasSheet = SS.getSheetByName("👥 Personas");
  const personasData = personasSheet.getDataRange().getValues();

  let personas = [];
  for (let i = 1; i < personasData.length; i++) {
    const nombre = personasData[i][1];
    const carga = personasData[i][7];
    if (nombre) {
      personas.push(`${nombre} (${carga})`);
    }
  }

  const response4 = ui.prompt(
    `➕ ASIGNAR NUEVA TAREA\n\n` +
    `Selecciona persona:\n\n${personas.join("\n")}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response4.getSelectedButton() !== ui.Button.OK) return;
  const persona = response4.getResponseText().split("(")[0].trim();

  // PASO 5: Prioridad
  const response5 = ui.alert(
    "➕ ASIGNAR NUEVA TAREA\n\n" +
    "¿Prioridad ALTA?\n\n" +
    "SÍ = Alta\nNO = Media",
    ui.ButtonSet.YES_NO
  );

  const prioridad = response5 === ui.Button.YES ? "Alta" : "Media";

  // CONFIRMAR
  const confirm = ui.alert(
    `✅ RESUMEN DE NUEVA TAREA\n\n` +
    `Descripción: ${descripcion}\n` +
    `Cantidad: ${cantidad}\n` +
    `Área: ${area}\n` +
    `Persona: ${persona}\n` +
    `Prioridad: ${prioridad}\n\n` +
    `¿Confirmar asignación?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ui.alert("❌ Cancelado");
    return;
  }

  // CREAR TAREA
  crearNuevaTarea(descripcion, cantidad, area, persona, prioridad);

  ui.alert(
    `✅ TAREA ASIGNADA CORRECTAMENTE\n\n` +
    `Descripción: ${descripcion}\n` +
    `Persona: ${persona}\n` +
    `Cantidad: ${cantidad}\n\n` +
    `✓ Registrada en el sistema\n` +
    `✓ Añadida al Historial`
  );
}

function crearNuevaTarea(descripcion, cantidad, area, persona, prioridad) {
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
    tareasSheet.getRange(nuevaFila, 8).setValue(prioridad);
    tareasSheet.getRange(nuevaFila, 9).setValue("SI");
    tareasSheet.getRange(nuevaFila, 10).setValue("Tarea nueva asignada");

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
        <p><strong>Prioridad:</strong> ${prioridad}</p>

        <p style="margin-top: 20px; color: #666;">
          <strong>Hora:</strong> ${ahora.toLocaleString("es-MX")}<br>
          <strong>Sistema:</strong> Producción Automático
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

function detectarPendientesAutomatico() {
  try {
    console.log("🔍 Ejecutando detección automática de pendientes...");

    const pendientes = detectarTareasPendientes();

    if (pendientes.length > 0) {
      console.log(`⚠️ Se encontraron ${pendientes.length} tareas pendientes`);
      notificarPendientes(pendientes);
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

// ==================== RESTO DE FUNCIONES ====================

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
  const ui = SpreadsheetApp.getUi();
  const tareasSheet = SS.getSheetByName("📋 Tareas");

  const response = ui.prompt(
    "🔄 REASIGNACIÓN MANUAL\n\n" +
    "Ingresa el ID de la tarea (ej: T003):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const tareaID = response.getResponseText().toUpperCase().trim();
  const datos = tareasSheet.getDataRange().getValues();

  let tareaInfo = null;
  let filaEncontrada = -1;

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === tareaID) {
      tareaInfo = {
        tareaID: datos[i][0],
        persona: datos[i][1],
        descripcion: datos[i][3],
        area: datos[i][2],
        cantidad: datos[i][4]
      };
      filaEncontrada = i;
      break;
    }
  }

  if (!tareaInfo) {
    ui.alert(`❌ Tarea ${tareaID} no encontrada`);
    return;
  }

  const tipoRespuesta = ui.alert(
    `📌 TAREA: ${tareaID}\n\n` +
    `Persona: ${tareaInfo.persona}\n` +
    `Descripción: ${tareaInfo.descripcion}\n` +
    `Cantidad: ${tareaInfo.cantidad}\n\n` +
    `¿Qué tipo de reasignación?\n\n` +
    `SÍ = REPARACIÓN (tiene defectos)\n` +
    `NO = TAREA SOBRANTE (no terminó)`,
    ui.ButtonSet.YES_NO
  );

  const tipo = tipoRespuesta === ui.Button.YES ? "REPARACIÓN" : "TAREA SOBRANTE";

  const personasSheet = SS.getSheetByName("👥 Personas");
  const personasData = personasSheet.getDataRange().getValues();

  let personasDisponibles = [];
  for (let i = 1; i < personasData.length; i++) {
    const nombre = personasData[i][1];
    const carga = personasData[i][7];
    if (nombre && carga && carga.includes("BAJO")) {
      personasDisponibles.push(`${nombre} (${carga} ⭐)`);
    }
  }

  const personaRespuesta = ui.alert(
    `¿A quién deseas reasignar?\n\n` +
    `SÍ = A la misma persona (${tareaInfo.persona})\n` +
    `NO = A otra persona (disponibles: ${personasDisponibles.join(", ")})`,
    ui.ButtonSet.YES_NO
  );

  let nuevaPersona = tareaInfo.persona;

  if (personaRespuesta === ui.Button.NO) {
    const respuesta2 = ui.prompt(
      `Selecciona nueva persona:\n\n` +
      `${personasDisponibles.join("\n")}\n\n` +
      `O escribe otro nombre:`,
      ui.ButtonSet.OK_CANCEL
    );

    if (respuesta2.getSelectedButton() === ui.Button.OK) {
      nuevaPersona = respuesta2.getResponseText().split("(")[0].trim();
    } else {
      ui.alert("❌ Cancelado");
      return;
    }
  }

  const confirmRespuesta = ui.alert(
    `✅ RESUMEN DE REASIGNACIÓN\n\n` +
    `Tarea: ${tareaID}\n` +
    `Tipo: ${tipo}\n` +
    `De: ${tareaInfo.persona}\n` +
    `Para: ${nuevaPersona}\n\n` +
    `¿Confirmar?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmRespuesta !== ui.Button.YES) {
    ui.alert("❌ Cancelado");
    return;
  }

  ejecutarReasignacion(tareaID, tareaInfo, nuevaPersona, tipo, filaEncontrada);

  ui.alert(
    `✅ REASIGNACIÓN EXITOSA\n\n` +
    `Tarea: ${tareaID}\n` +
    `Tipo: ${tipo}\n` +
    `De: ${tareaInfo.persona}\n` +
    `Para: ${nuevaPersona}\n\n` +
    `✓ Registrado en Historial\n` +
    `✓ Notificaciones enviadas`
  );
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

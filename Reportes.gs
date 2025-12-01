// ==================== REPORTES Y AUTOMATIZACIONES ====================

function btnGenerarReporteManual() {
  generarReporteDiarioAutomatico();
  SpreadsheetApp.getUi().alert("✅ Reporte generado y enviado");
}

function generarReporteDiarioAutomatico() {
  try {
    console.log("📊 Generando reporte diario...");

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

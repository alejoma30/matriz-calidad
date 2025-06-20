// app.js

document.addEventListener('DOMContentLoaded', function () {
  const asesoresPorProceso = {
    "PQRSD": ["Avila, Marcela Paola", "Beltran, Nicole Juliana", "Castro, Diana Angelica", "Fuentes, Sandra Milena", "Lopez, Angela Marcela", "Roa, Cindy Paola", "Sanchez, Hellen Viviana", "Nadya Eliscet Bernal Escobar"],
    "NOTIFICACIONES - PQRSD": ["Insuasty, Daniel Ismael"],
    "NOTIFICACIONES": ["Gomez, Natalia", "Gutierrez, Valentina", "Alvarez, Carlos William", "Garavito, Gabriela Alexandra", "Mahecha, Diego Andres", "Peña, Jairo Esteban", "Rincon, Nathaly Dayana", "Sandoval, Diego Mauricio", "Santamaria, Edinson Yesid", "Hernandez, Diego Andres", "John Edwar Olarte"],
    "LEGALIZACIONES": ["Castiblanco, Jonathan Javier", "Saavedra, Jenny Alexandra", "Ojeda, Maria Alejandra", "Rodriguez, Andrés Eduardo", "Ruiz, Daissy Katerine"],
    "ANTENCION PRESENCIAL": ["Alvarez, Katherine"],
  };

  const lideresCalidad = ["Rene Alejandro Mayorga", "Andrea Guzman Botache"];

  const evaluadorSelect = document.getElementById('evaluador');
  lideresCalidad.forEach(nombre => {
    const option = document.createElement('option');
    option.value = nombre;
    option.textContent = nombre;
    evaluadorSelect.appendChild(option);
  });

  const procesoSelect = document.getElementById('proceso');
  const asesorSelect = document.getElementById('asesor');

  procesoSelect.addEventListener('change', () => {
    const proceso = procesoSelect.value;
    const asesores = asesoresPorProceso[proceso] || [];
    asesorSelect.innerHTML = '<option value="">-- Selecciona --</option>';
    asesores.forEach(nombre => {
      const option = document.createElement('option');
      option.value = nombre;
      option.textContent = nombre;
      asesorSelect.appendChild(option);
    });
  });

  const cumplimientoSelects = document.querySelectorAll('.cumplimiento');
  const notaSpan = document.getElementById('nota');
  const pesos = [50, 30, 20];

  const calcularNota = () => {
    let nota = 100;
    for (let i = 0; i < cumplimientoSelects.length; i++) {
      const val = cumplimientoSelects[i].value;
      if (val === 'NO') {
        if (i < 3) nota -= pesos[i];
        else return 0;
      }
    }
    return nota;
  };

  const updateNota = () => {
    const nota = calcularNota();
    notaSpan.textContent = `${nota}%`;
  };

  cumplimientoSelects.forEach(select => {
    select.addEventListener('change', updateNota);
  });

  // Fecha auditoría por defecto
  document.getElementById('fecha-auditoria').valueAsDate = new Date();

  document.getElementById('formulario').addEventListener('submit', function (e) {
    e.preventDefault();
    const nota = calcularNota();
    const semaforo = nota >= 90 ? '🟢 Excelente' : nota >= 80 ? '🟡 Aceptable' : '🔴 Debe mejorar';
    const evaluador = document.getElementById('evaluador').value;

    alert(`${evaluador}, la auditoría se ha guardado con éxito.\nNota: ${nota}%\n${semaforo}`);

    const data = {
      fechaAuditoria: document.getElementById('fecha-auditoria').value,
      fechaGestion: document.getElementById('fecha-gestion').value,
      proceso: document.getElementById('proceso').value,
      asesor: document.getElementById('asesor').value,
      evaluador: evaluador,
      radicado: document.getElementById('radicado').value,
      "Uso de plantillas": document.getElementById('c1').value,
      "Claridad del lenguaje": document.getElementById('c2').value,
      "Redacción – puntuación": document.getElementById('c3').value,
      "Redacción – ortografía": document.getElementById('c4').value,
      "Interpretación de la solicitud": document.getElementById('c5').value,
      "Oportunidad en la respuesta": document.getElementById('c6').value,
      "Pertinencia de la respuesta": document.getElementById('c7').value,
      "Desempeño": document.getElementById('c8').value,
      observaciones: document.getElementById('observaciones').value,
      retroalimentacion: document.getElementById('feedback').value,
      nota: nota,
      semaforo: semaforo
    };

    fetch("https://script.google.com/macros/s/AKfycbz875DEEnTkNiYEjmdJI15MR0gvrW07GQNtt_JSG0KVOHmx58zQN3GVHgl1XZq3f_Y9/exec", {
      method: 'POST',
      body: JSON.stringify(data),
      headers: { 'Content-Type': 'application/json' }
    }).then(res => res.text())
      .then(resp => console.log("Enviado a Sheets:", resp))
      .catch(err => console.error("Error al enviar a Sheets:", err));

    this.reset();
    notaSpan.textContent = '100%';
    document.getElementById('fecha-auditoria').valueAsDate = new Date();
  });

  document.getElementById('btnExportarExcel').addEventListener('click', function () {
    const headers = [
      "Fecha Auditoría", "Fecha Gestión", "Proceso", "Asesor", "Evaluador", "Radicado",
      "Uso de plantillas", "Claridad del lenguaje", "Redacción – puntuación", "Redacción – ortografía",
      "Interpretación de la solicitud", "Oportunidad en la respuesta", "Pertinencia de la respuesta", "Desempeño",
      "Observaciones", "Retroalimentación", "Nota", "Semáforo"
    ];
    const fila = [
      document.getElementById('fecha-auditoria').value,
      document.getElementById('fecha-gestion').value,
      document.getElementById('proceso').value,
      document.getElementById('asesor').value,
      document.getElementById('evaluador').value,
      document.getElementById('radicado').value,
      document.getElementById('c1').value,
      document.getElementById('c2').value,
      document.getElementById('c3').value,
      document.getElementById('c4').value,
      document.getElementById('c5').value,
      document.getElementById('c6').value,
      document.getElementById('c7').value,
      document.getElementById('c8').value,
      document.getElementById('observaciones').value,
      document.getElementById('feedback').value,
      document.getElementById('nota').textContent,
      calcularNota() >= 90 ? '🟢 Excelente' : calcularNota() >= 80 ? '🟡 Aceptable' : '🔴 Debe mejorar'
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, fila]);
    XLSX.utils.book_append_sheet(wb, ws, "Monitoreo");
    XLSX.writeFile(wb, "monitoreo.xlsx");
  });
});

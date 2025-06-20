document.addEventListener('DOMContentLoaded', function () {
  // ... (todo tu código previo intacto)

  // Enviar a Google Sheets
  fetch("https://script.google.com/macros/s/AKfycbwiUeBzoOXJWlzEbZ2CxNZ0pEn-UliN7D-hwLnFz3ObFPbY4oO7rjulTwxd45PvtjkCKQ/exec", {
    method: 'POST',
    body: JSON.stringify(data),
    headers: { 'Content-Type': 'application/json' }
  }).then(res => res.text())
    .then(resp => console.log("Enviado a Sheets:", resp))
    .catch(err => console.error("Error al enviar a Sheets:", err));
});

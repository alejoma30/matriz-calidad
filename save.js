import { google } from 'googleapis';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).send({ message: 'Sólo se permiten solicitudes POST' });
  }

  try {
    const serviceAccount = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const body = req.body;

    const values = [[
      body.fechaAuditoria,
      body.fechaGestion,
      body.proceso,
      body.asesor,
      body.evaluador,
      body.radicado,
      body.c1,
      body.c2,
      body.c3,
      body.c4,
      body.c5,
      body.c6,
      body.c7,
      body.c8,
      body.observaciones,
      body.retroalimentacion,
      body.nota,
      body.semaforo
    ]];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'A1',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values }
    });

    res.status(200).json({ message: 'Datos guardados con éxito en Google Sheets' });
  } catch (error) {
    console.error('Error al guardar en Google Sheets:', error);
    res.status(500).json({ message: 'Error al guardar en Google Sheets', error: error.message });
  }
}
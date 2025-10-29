import express from "express";
import multer from "multer";
import cors from "cors";
import ExcelJS from "exceljs";
import pkg from "pg";
import sharp from "sharp";
const { Pool } = pkg;

const app = express();
const PORT = 3000;

app.use(express.static("public"));
// --- Configura PostgreSQL ---
const pool = new Pool({
  connectionString: process.env.DATABASE_URL || "postgresql://usuario:avlUBxKlYBmWlbNkHcpljYHJgow7x8Py@dpg-d4175l63jp1c73cm7ktg-a.oregon-postgres.render.com/basedatos_cxga",
  ssl: { rejectUnauthorized: false },
});

// Crear tablas si no existen
const initDB = async () => {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS respuestas (
      id SERIAL PRIMARY KEY,
      nombreColaborador TEXT,
      cedula TEXT,
      nombreNino TEXT,
      parentesco TEXT,
      edad TEXT,
      dibujo TEXT,
      valor TEXT,
      categoria TEXT,
      fecha TIMESTAMP DEFAULT NOW()
    );
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS imagenes (
      id SERIAL PRIMARY KEY,
      respuesta_id INT REFERENCES respuestas(id) ON DELETE CASCADE,
      imagen BYTEA
    );
  `);
};
initDB();

// --- Configuración de Multer (límite 5MB por imagen) ---
const upload = multer({
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB
  storage: multer.memoryStorage(),
});

app.use(cors());
app.use(express.json());

// --- Guardar formulario ---
app.post("/api/form", upload.array("fotos", 4), async (req, res) => {
  try {
    if (!req.files || req.files.length !== 4)
      return res.status(400).json({ ok: false, msg: "Debes subir exactamente 4 fotos." });

    const { nombreColaborador, cedula, nombreNino, parentesco, edad, dibujo, valor, categoria } = req.body;

    const result = await pool.query(
      `INSERT INTO respuestas (nombreColaborador, cedula, nombreNino, parentesco, edad, dibujo, valor, categoria)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8) RETURNING id`,
      [nombreColaborador, cedula, nombreNino, parentesco, edad, dibujo, valor, categoria]
    );

    const respuesta_id = result.rows[0].id;

    for (const f of req.files) {
      // Redimensionar si es necesario
      const imgSharp = sharp(f.buffer);
      const metadata = await imgSharp.metadata();

      let resizedBuffer = f.buffer;

      if (metadata.width > 768 || metadata.height > 768) {
        resizedBuffer = await imgSharp
          .resize({
            width: metadata.width > metadata.height ? 768 : null,
            height: metadata.height >= metadata.width ? 768 : null,
            fit: "inside",
          })
          .jpeg({ quality: 80 }) // opcional: comprimir un poco
          .toBuffer();
      }

      await pool.query("INSERT INTO imagenes (respuesta_id, imagen) VALUES ($1, $2)", [
        respuesta_id,
        resizedBuffer,
      ]);
    }

    res.json({ ok: true, msg: "Datos guardados correctamente" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, msg: "Error al guardar los datos" });
  }
});

// --- Obtener todas las respuestas ---
app.get("/api/data", async (req, res) => {
  const respuestas = await pool.query("SELECT * FROM respuestas ORDER BY id DESC");
  for (const r of respuestas.rows) {
    const imgs = await pool.query("SELECT id FROM imagenes WHERE respuesta_id=$1", [r.id]);
    r.fotos = imgs.rows.map(i => `/api/img/${i.id}`);
  }
  res.json(respuestas.rows);
});

// --- Servir imágenes ---
app.get("/api/img/:id", async (req, res) => {
  const img = await pool.query("SELECT imagen FROM imagenes WHERE id=$1", [req.params.id]);
  if (img.rowCount === 0) return res.status(404).send("No encontrada");
  res.setHeader("Content-Type", "image/jpeg");
  res.send(img.rows[0].imagen);
});

// --- Panel admin ---
app.get("/admin", async (req, res) => {
  const respuestas = await pool.query("SELECT * FROM respuestas ORDER BY id DESC");

  let rows = "";
  for (const [i, r] of respuestas.rows.entries()) {
    const imgs = await pool.query("SELECT id FROM imagenes WHERE respuesta_id=$1", [r.id]);
    const links = imgs.rows.map(im => `<a href="/api/img/${im.id}" target="_blank">📷</a>`).join(" ");
    rows += `
      <tr>
        <td>${i + 1}</td>
        <td>${r.nombrecolaborador}</td>
        <td>${r.cedula}</td>
        <td>${r.nombrenino}</td>
        <td>${r.parentesco}</td>
        <td>${r.edad}</td>
        <td>${r.dibujo || ""}</td>
        <td>${r.valor || ""}</td>
        <td>${r.categoria || ""}</td>
        <td>${r.fecha.toISOString().split("T")[0]}</td>
        <td>${links}</td>
      </tr>`;
  }

  res.send(`
  <html><head><meta charset="utf-8"/>
  <title>Panel Admin</title>
  <style>
    body{font-family:sans-serif;padding:20px;}
    table{border-collapse:collapse;width:100%;}
    th,td{border:1px solid #ccc;padding:5px;font-size:14px;}
    th{background:#f4f4f4;}
    a.btn{display:inline-block;margin-bottom:10px;padding:6px 12px;background:#28a745;color:white;text-decoration:none;border-radius:4px;}
  </style></head>
  <body>
    <h2>Respuestas del Formulario</h2>
    <a href="/admin/export" class="btn">📥 Descargar Excel</a>
    <table>
      <tr>
        <th>#</th><th>Colaborador</th><th>Cédula</th><th>Niño(a)</th><th>Parentesco</th>
        <th>Edad</th><th>Dibujo</th><th>Valor</th><th>Categoría</th><th>Fecha</th><th>Fotos</th>
      </tr>${rows}
    </table>
  </body></html>`);
});

// --- Exportar Excel con 4 hipervínculos ---
// --- Exportar Excel con 4 hipervínculos funcionales ---
app.get("/admin/export", async (req, res) => {
  try {
    const respuestas = (await pool.query("SELECT * FROM respuestas ORDER BY id")).rows;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Respuestas");

    // Encabezados
    sheet.columns = [
      { header: "#", key: "num", width: 5 },
      { header: "Colaborador", key: "colaborador", width: 25 },
      { header: "Cédula", key: "cedula", width: 15 },
      { header: "Niño(a)", key: "nino", width: 25 },
      { header: "Parentesco", key: "parentesco", width: 15 },
      { header: "Edad", key: "edad", width: 10 },
      { header: "Dibujo", key: "dibujo", width: 15 },
      { header: "Valor", key: "valor", width: 15 },
      { header: "Categoría", key: "categoria", width: 15 },
      { header: "Fecha", key: "fecha", width: 15 },
      { header: "Foto 1", key: "foto1", width: 30 },
      { header: "Foto 2", key: "foto2", width: 30 },
      { header: "Foto 3", key: "foto3", width: 30 },
      { header: "Foto 4", key: "foto4", width: 30 },
    ];

    for (let i = 0; i < respuestas.length; i++) {
      const r = respuestas[i];
      const imgs = await pool.query("SELECT id FROM imagenes WHERE respuesta_id=$1 LIMIT 4", [r.id]);
      const baseUrl = `${req.protocol}://${req.get("host")}/api/img/`;

      // Crear fila
      const rowValues = {
        num: i + 1,
        colaborador: r.nombrecolaborador,
        cedula: r.cedula,
        nino: r.nombrenino,
        parentesco: r.parentesco,
        edad: r.edad,
        dibujo: r.dibujo,
        valor: r.valor,
        categoria: r.categoria,
        fecha: r.fecha.toISOString().split("T")[0],
      };

      // Asignar hipervínculos a cada columna de foto
      imgs.rows.forEach((img, idx) => {
        rowValues[`foto${idx + 1}`] = { text: `Foto ${idx + 1}`, hyperlink: baseUrl + img.id };
      });

      sheet.addRow(rowValues);
    }

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Disposition", "attachment; filename=respuestas.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send("Error al generar el Excel");
  }
});

app.get("/admin/erase/cuidado/87654321", async (req, res) => {
  try {
    await pool.query("DELETE FROM imagenes");
    await pool.query("DELETE FROM respuestas");
    res.json({ ok: true, msg: "Datos borrados correctamente" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, msg: "Error al borrar los datos" });
  }
});

app.listen(PORT, () => console.log(`Servidor en http://localhost:${PORT}`));

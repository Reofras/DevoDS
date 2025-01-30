const express = require('express');
const multer = require('multer');
const fs = require('fs');
const xml2js = require('xml2js');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.post('/upload', upload.single('xmlFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'Nenhum arquivo enviado.' });
    }

    const filePath = req.file.path;
    const fileContent = fs.readFileSync(filePath, 'utf-8');

    xml2js.parseString(fileContent, (err, result) => {
        if (err) {
            return res.status(400).json({ error: 'Erro ao processar o arquivo XML.' });
        }

        // Processar o XML conforme necessário
        res.json(result);
    });
});

app.listen(3000, () => {
    console.log('Servidor rodando na porta 3000');
});
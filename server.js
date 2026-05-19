require('dotenv').config();
const express = require('express');
const multer = require('multer');
const Groq = require('groq-sdk');
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } = require('docx');

const app = express();

const storage = multer.diskStorage({
    destination: 'uploads/',
    filename: (req, file, cb) => {
        const ext = path.extname(file.originalname);
        cb(null, Date.now() + ext);
    }
});
const upload = multer({ storage });

const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

app.use(express.static('public'));
app.use(express.json());

app.post('/transcribe', upload.single('audio'), async (req, res) => {
    try {
        const audioPath = req.file.path;
        const transcription = await groq.audio.transcriptions.create({
            file: fs.createReadStream(audioPath),
            model: "whisper-large-v3",
            language: "fr"
        });
        fs.unlinkSync(audioPath);
        res.json({ transcription: transcription.text });
    } catch (error) {
        console.error('Erreur:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.post('/refine', async (req, res) => {
    try {
        const { transcription, instruction } = req.body;
        
        const completion = await groq.chat.completions.create({
            messages: [{
                role: "user",
                content: `Voici une retranscription : "${transcription}"\n\nInstruction : ${instruction}\n\nDonne-moi la version améliorée.`
            }],
            model: "llama-3.3-70b-versatile",
            temperature: 0.5
        });
        
        const refined = completion.choices[0].message.content;
        res.json({ refined });
    } catch (error) {
        console.error('Erreur:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.post('/export-docx', async (req, res) => {
    try {
        const { text, patiente, date, numero } = req.body;
        if (!text) return res.status(400).json({ error: 'Texte manquant' });

        const children = [];

        // Titre du document
        const titleLine = patiente
            ? `Compte rendu — ${patiente}${numero ? ` (N°${numero})` : ''}`
            : 'Compte rendu';
        children.push(new Paragraph({
            text: titleLine,
            heading: HeadingLevel.HEADING_1,
            spacing: { after: 120 },
        }));

        if (date) {
            children.push(new Paragraph({
                children: [new TextRun({ text: `Séance du ${date}`, color: '666666', size: 20 })],
                spacing: { after: 320 },
            }));
        }

        // Parse les sections ## ...
        const sections = text.split(/(?=^## )/m);
        sections.forEach(section => {
            if (!section.trim()) return;
            const lines = section.trim().split('\n');
            const sectionTitle = lines[0].replace(/^##\s*/, '').trim();
            const body = lines.slice(1);

            children.push(new Paragraph({
                text: sectionTitle,
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 280, after: 100 },
            }));

            body.forEach(line => {
                line = line.trim();
                if (!line) return;
                const isBullet = line.startsWith('- ') || line.startsWith('• ');
                children.push(new Paragraph({
                    children: [new TextRun({ text: isBullet ? line.replace(/^[-•]\s*/, '') : line, size: 22 })],
                    bullet: isBullet ? { level: 0 } : undefined,
                    spacing: { after: 80 },
                }));
            });
        });

        const doc = new Document({ sections: [{ children }] });
        const buffer = await Packer.toBuffer(doc);

        const filename = patiente
            ? `CR_${patiente.replace(/\s+/g, '_')}_${new Date().toISOString().slice(0,10)}.docx`
            : `Compte_rendu_${new Date().toISOString().slice(0,10)}.docx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);
    } catch (error) {
        console.error('Erreur export docx:', error.message);
        res.status(500).json({ error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Application lancée sur http://localhost:${PORT}`);
});

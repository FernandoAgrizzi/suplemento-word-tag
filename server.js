const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.static(__dirname));

// Utiliza os certificados dev-cert do ecossistema se existirem ou porta local HTTPS
const PORT = 3000;

// Configuracao basica do servidor para servir taskpane.html e JS
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'taskpane.html'));
});

// Caso nao haja certs SSL configurados no node, o ecossistema usa npx http-server -S
console.log(`Servidor configurado para rodar na porta ${PORT}`);

// node-compat-test.js

const report = {
    passed: [],
    failed: [],
    missing: []
};

// Função auxiliar para isolar e registrar os testes
function testFeature(featureName, testFn) {
    try {
        testFn();
        report.passed.push(featureName);
        console.log(`[ OK ] ${featureName}`);
    } catch (error) {
        const errMsg = error.message || error.toString();
        // Classifica como missing se o erro indicar falta de módulo ou variável indefinida
        if (errMsg.includes('Cannot find module') || errMsg.includes('not defined') || errMsg.includes('is not a function')) {
            report.missing.push({ name: featureName, error: errMsg });
            console.log(`[FALTA] ${featureName} -> ${errMsg}`);
        } else {
            report.failed.push({ name: featureName, error: errMsg });
            console.log(`[ERRO] ${featureName} -> ${errMsg}`);
        }
    }
}

console.log("=========================================");
console.log(" Iniciando Testes de Compatibilidade Node");
console.log("=========================================\n");

// 1. Variáveis e Funções Globais
testFeature('Global: console (log, warn, error)', () => {
    if (typeof console.log !== 'function' || typeof console.error !== 'function') {
        throw new Error("console.log ou console.error ausentes");
    }
});

testFeature('Global: process (env, version, cwd)', () => {
    if (!process.env || !process.version || typeof process.cwd !== 'function') {
        throw new Error("Objeto 'process' incompleto");
    }
});

testFeature('Global: Buffer', () => {
    const b = Buffer.from('axon', 'utf8');
    if (b.toString('hex') !== '61786f6e') {
        throw new Error("Implementação do Buffer com falha na conversão");
    }
});

testFeature('Global: setTimeout e clearTimeout', () => {
    if (typeof setTimeout !== 'function' || typeof clearTimeout !== 'function') {
        throw new Error("setTimeout is not defined");
    }
});

testFeature('Global: __dirname e __filename', () => {
    if (typeof __dirname === 'undefined' || typeof __filename === 'undefined') {
        throw new Error("__dirname ou __filename not defined");
    }
});

// 2. Resolução de Módulos Core (CommonJS)
testFeature('Módulo: fs (File System)', () => {
    const fs = require('fs');
    if (typeof fs.readFileSync !== 'function') {
        throw new Error("fs.readFileSync is not a function");
    }
});

testFeature('Módulo: path', () => {
    const path = require('path');
    const joined = path.join('pasta', 'arquivo.js');
    if (!joined.includes('pasta') || !joined.includes('arquivo.js')) {
        throw new Error("Falha na lógica do path.join");
    }
});

testFeature('Módulo: os (Operating System)', () => {
    const os = require('os');
    if (typeof os.platform !== 'function' || typeof os.cpus !== 'function') {
        throw new Error("Funções básicas do módulo 'os' ausentes");
    }
});

testFeature('Módulo: crypto', () => {
    const crypto = require('crypto');
    const hash = crypto.createHash('md5').update('axon').digest('hex');
    if (!hash) throw new Error("Falha ao gerar hash com crypto");
});

testFeature('Módulo: events (EventEmitter)', () => {
    const EventEmitter = require('events');
    const emitter = new EventEmitter();
    let triggered = false;
    emitter.on('teste', () => { triggered = true; });
    emitter.emit('teste');
    if (!triggered) throw new Error("EventEmitter não disparou o evento");
});

testFeature('Módulo: http', () => {
    const http = require('http');
    if (typeof http.createServer !== 'function') {
        throw new Error("http.createServer is not a function");
    }
});

testFeature('Módulo: util', () => {
    const util = require('util');
    if (typeof util.promisify !== 'function') {
        throw new Error("util.promisify is not a function");
    }
});

testFeature('Módulo: child_process', () => {
    const cp = require('child_process');
    if (typeof cp.execSync !== 'function') {
        throw new Error("child_process.execSync is not a function");
    }
});

// 3. Suporte a ES6+ (caso o motor JavaScript seja moderno)
testFeature('Sintaxe: Promises', () => {
    if (typeof Promise === 'undefined') throw new Error("Promise is not defined");
    const p = new Promise((resolve) => resolve(true));
});

testFeature('Sintaxe: Async/Await', () => {
    // Avalia dinamicamente para evitar erro de parse se a engine for muito antiga (ex: JScript antigo)
    eval("async function teste() { await Promise.resolve(true); }");
});


// Relatório Final
console.log("\n=========================================");
console.log(" Resumo da Compatibilidade");
console.log("=========================================");
console.log(`Sucessos     : ${report.passed.length}`);
console.log(`Falhas Lógicas: ${report.failed.length}`);
console.log(`Não Implemen.: ${report.missing.length}`);
console.log("=========================================");

if (report.missing.length > 0) {
    console.log("\nRecursos que você precisa implementar no AxonASP:");
    report.missing.forEach(m => console.log(`- ${m.name}`));
}
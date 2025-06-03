let tiposPesos = [];
let pesosAtuais = {};
let tabelaSemana = [];
let sortColumn = 'usuario';
let sortAscending = true;
let usuariosComparacao = [];
let dadosComparacao = [];
let chartComparacao = null;
let chartInstance = null;
let isDarkTheme = true;
let excelData = null;
let processedData = {
    usuarios: [],
    totalMinutas: 0,
    mediaUsuario: 0,
    usuariosAtivos: 0,
    diaProdutivo: '--',
    topUsers: [],
    mesesAtivos: [],
    mesesDisponiveis: []
};

function initializeApp() {
    updateTheme();
    createChart();
    renderMonths();
    updateKPIs();
}

function showDashboard() {
    document.querySelectorAll('.page-content').forEach(page => page.style.display = 'none');
    document.getElementById('dashboard-page').style.display = 'block';
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelectorAll('.nav-item')[0].classList.add('active');
}

function showComparacaoPage() {
    document.querySelectorAll('.page-content').forEach(page => page.style.display = 'none');
    document.getElementById('comparar-page').style.display = 'block';
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelectorAll('.nav-item')[1].classList.add('active');
    gerarDadosComparacao();
}

function showSemanaPage() {
    document.querySelectorAll('.page-content').forEach(page => page.style.display = 'none');
    document.getElementById('semana-page').style.display = 'block';
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelectorAll('.nav-item')[2].classList.add('active');
    gerarTabelaSemana();
}

function gerarDadosComparacao() {
    if (!excelData || excelData.length === 0) return;

    const usuariosMap = new Map();
    let filteredData = excelData;

    if (processedData.mesesAtivos.length > 0 && processedData.mesesAtivos.length < processedData.mesesDisponiveis.length) {
        filteredData = excelData.filter(row => {
            if (!row['Data criação']) return false;
            
            try {
                let date = row['Data criação'];
                if (!(date instanceof Date)) {
                    date = new Date(date);
                }
                
                if (!isNaN(date.getTime())) {
                    const mesNomes = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                                     'jul', 'ago', 'set', 'out', 'nov', 'dez'];
                    const mesAno = `${mesNomes[date.getMonth()]}/${date.getFullYear()}`;
                    return processedData.mesesAtivos.includes(mesAno);
                }
                return false;
            } catch (e) {
                return false;
            }
        });
    }

    filteredData.forEach(row => {
        const nroProcesso = row['Nro. processo'];
        if (!nroProcesso || nroProcesso.trim() === '') return;
        
        const usuario = row['Usuário'];
        if (!usuario || usuario.toString().trim() === '') return;
        
        const peso = parseFloat(row['peso']) || 1.0;
        
        if (!usuariosMap.has(usuario)) {
            usuariosMap.set(usuario, 0);
        }
        usuariosMap.set(usuario, usuariosMap.get(usuario) + peso);
    });

    const usuarios = Array.from(usuariosMap.entries())
        .map(([nome, minutas]) => ({ nome, minutas, selecionado: false }))
        .sort((a, b) => b.minutas - a.minutas);

    usuariosComparacao = usuarios;
    renderizarListaUsuarios();
}

function renderizarListaUsuarios() {
    const container = document.getElementById('usuarios-list');
    if (!container) return;
    
    container.innerHTML = '';

    const filtro = document.getElementById('filtro-usuarios').value.toLowerCase();
    const usuariosFiltrados = usuariosComparacao.filter(usuario => 
        usuario.nome.toLowerCase().includes(filtro)
    );

    usuariosFiltrados.forEach((usuario, index) => {
        const usuarioDiv = document.createElement('div');
        usuarioDiv.className = `usuario-item ${usuario.selecionado ? 'selected' : ''}`;
        usuarioDiv.innerHTML = `
            <div class="usuario-checkbox">
                <input type="checkbox" 
                       id="user-${index}" 
                       ${usuario.selecionado ? 'checked' : ''}
                       onchange="toggleUsuarioComparacao('${usuario.nome}')">
                <label for="user-${index}"></label>
            </div>
            <div class="usuario-info">
                <span class="usuario-nome">${usuario.nome}</span>
                <span class="usuario-minutas">${Math.round(usuario.minutas)} minutas</span>
            </div>
        `;
        container.appendChild(usuarioDiv);
    });
}

function toggleUsuarioComparacao(nomeUsuario) {
    const usuario = usuariosComparacao.find(u => u.nome === nomeUsuario);
    if (usuario) {
        usuario.selecionado = !usuario.selecionado;
        renderizarListaUsuarios();
        atualizarGraficoComparacao();
    }
}

function selecionarTodosUsuarios() {
    usuariosComparacao.forEach(usuario => usuario.selecionado = true);
    renderizarListaUsuarios();
    atualizarGraficoComparacao();
}

function limparSelecaoUsuarios() {
    usuariosComparacao.forEach(usuario => usuario.selecionado = false);
    renderizarListaUsuarios();
    atualizarGraficoComparacao();
}

function atualizarGraficoComparacao() {
    const ctx = document.getElementById('comparacao-chart');
    if (!ctx) return;

    if (chartComparacao) {
        chartComparacao.destroy();
    }

    const usuariosSelecionados = usuariosComparacao.filter(u => u.selecionado);
    
    if (usuariosSelecionados.length === 0) {
        chartComparacao = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Nenhum usuário selecionado'],
                datasets: [{
                    label: 'Minutas',
                    data: [0],
                    backgroundColor: '#cccccc'
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Selecione usuários para comparar',
                        color: '#666666'
                    }
                }
            }
        });
        return;
    }

    const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
    const cores = [
        '#4a90e2', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6',
        '#1abc9c', '#34495e', '#e67e22', '#95a5a6', '#d35400'
    ];

    chartComparacao = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: usuariosSelecionados.map(u => u.nome),
            datasets: [{
                label: 'Minutas',
                data: usuariosSelecionados.map(u => u.minutas),
                backgroundColor: usuariosSelecionados.map((_, index) => cores[index % cores.length]),
                borderWidth: 0
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                title: {
                    display: true,
                    text: 'Comparação de Produtividade',
                    color: isDark ? '#ffffff' : '#232946',
                    font: {
                        size: 16,
                        weight: 'bold'
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: {
                        color: isDark ? '#404040' : '#e0e0e0'
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946'
                    }
                },
                y: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946'
                    }
                }
            }
        }
    });
}

function updateKPIs() {
    console.log('Atualizando KPIs:', processedData);
    
    const totalElement = document.getElementById('total-minutas');
    const diaElement = document.getElementById('dia-produtivo');
    const mediaElement = document.getElementById('media-usuario');
    const usuariosElement = document.getElementById('usuarios-ativos');
    
    if (totalElement) {
        totalElement.textContent = processedData.totalMinutas;
    }
    
    if (diaElement) {
        diaElement.textContent = processedData.diaProdutivo;
    }
    
    if (mediaElement) {
        mediaElement.textContent = processedData.mediaUsuario > 0 ? processedData.mediaUsuario.toFixed(1) : '--';
    }
    
    if (usuariosElement) {
        usuariosElement.textContent = processedData.usuariosAtivos;
    }
    
    const topUsersContainer = document.getElementById('top-users');
    if (topUsersContainer) {
        topUsersContainer.innerHTML = '';
        
        if (processedData.topUsers && processedData.topUsers.length > 0) {
            processedData.topUsers.forEach((user, index) => {
                const userDiv = document.createElement('div');
                userDiv.className = 'user-item';
                userDiv.textContent = `${index + 1}º ${user.nome}: ${Math.round(user.minutas)}`;
                userDiv.style.marginBottom = '4px';
                userDiv.style.fontSize = '12px';
                userDiv.style.lineHeight = '1.2';
                topUsersContainer.appendChild(userDiv);
            });
        } else {
            const userDiv = document.createElement('div');
            userDiv.className = 'user-item';
            userDiv.textContent = 'Nenhum dado disponível';
            topUsersContainer.appendChild(userDiv);
        }
    }
}

function renderMonths() {
    const monthsGrid = document.getElementById('months-grid');
    monthsGrid.innerHTML = '';
    
    processedData.mesesDisponiveis.forEach(month => {
        const isActive = processedData.mesesAtivos.includes(month);
        
        const monthChip = document.createElement('div');
        monthChip.className = `month-chip ${isActive ? 'active' : ''}`;
        monthChip.textContent = month;
        monthChip.onclick = () => toggleMonth(month);
        
        monthsGrid.appendChild(monthChip);
    });
}

function toggleMonth(month) {
    const index = processedData.mesesAtivos.indexOf(month);
    
    if (index > -1) {
        processedData.mesesAtivos.splice(index, 1);
    } else {
        processedData.mesesAtivos.push(month);
    }
    
    renderMonths();
    
    if (excelData && excelData.length > 0) {
        processExcelData();
    }
}

function createChart() {
    const ctx = document.getElementById('produtividade-chart');
    if (!ctx) return;
    
    const context = ctx.getContext('2d');
    
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
    const chartData = processedData.usuarios ? processedData.usuarios.slice(0, 20) : [];
    
    if (chartData.length === 0) {
        chartInstance = new Chart(context, {
            type: 'bar',
            data: {
                labels: ['Sem dados'],
                datasets: [{
                    label: 'Minutas',
                    data: [0],
                    backgroundColor: isDark ? '#4a90e2' : '#3cb3e6'
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Total de Minutas',
                        color: isDark ? '#ffffff' : '#232946'
                    }
                },
                scales: {
                    x: {
                        beginAtZero: true,
                        grid: {
                            color: isDark ? '#404040' : '#e0e0e0'
                        },
                        ticks: {
                            color: isDark ? '#ffffff' : '#232946'
                        }
                    },
                    y: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            color: isDark ? '#ffffff' : '#232946'
                        }
                    }
                }
            }
        });
        return;
    }
    
    chartInstance = new Chart(context, {
        type: 'bar',
        data: {
            labels: chartData.map(u => u.nome),
            datasets: [{
                label: 'Minutas',
                data: chartData.map(u => u.minutas),
                backgroundColor: isDark ? '#4a90e2' : '#3cb3e6',
                borderColor: isDark ? '#4a90e2' : '#3cb3e6',
                borderWidth: 0
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                title: {
                    display: true,
                    text: 'Minutas por Usuário',
                    color: isDark ? '#ffffff' : '#232946',
                    font: {
                        size: 16,
                        weight: 'bold'
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: {
                        color: isDark ? '#404040' : '#e0e0e0'
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946',
                        font: {
                            size: 12
                        }
                    }
                },
                y: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946',
                        font: {
                            size: 11
                        }
                    }
                }
            }
        },
        plugins: [{
            id: 'datalabels',
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const data = dataset.data[index];
                        ctx.fillStyle = isDark ? '#ffffff' : '#232946';
                        ctx.font = 'bold 11px Arial';
                        ctx.textAlign = 'left';
                        ctx.textBaseline = 'middle';
                        ctx.fillText(Math.round(data), bar.x + 5, bar.y);
                    });
                });
            }
        }]
    });
}

function reprocessarDados() {
    if (excelData && excelData.length > 0) {
        extractMonthsFromData();
        processExcelData();
    }
}

function updateChart() {
    if (chartInstance && processedData.usuarios.length > 0) {
        const chartData = processedData.usuarios.slice(0, 20);
        
        chartInstance.data.labels = chartData.map(u => u.nome);
        chartInstance.data.datasets[0].data = chartData.map(u => u.minutas);
        chartInstance.update();
    }
}

function importData() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.multiple = true;
    
    input.onchange = function(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;
        
        let processedFiles = 0;
        let totalRows = 0;
        let fileNames = [];
        excelData = [];
        
        files.forEach((file, index) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    const range = XLSX.utils.decode_range(worksheet['!ref']);
                    let currentFileRows = 0;
                    
                    for (let R = 1; R <= range.e.r; ++R) {
                        const row = {};
                        
                        const tipoCell = worksheet[XLSX.utils.encode_cell({r: R, c: 0})];
                        const codigoCell = worksheet[XLSX.utils.encode_cell({r: R, c: 1})];
                        const nroProcessoCell = worksheet[XLSX.utils.encode_cell({r: R, c: 2})];
                        const usuarioCell = worksheet[XLSX.utils.encode_cell({r: R, c: 3})];
                        const dataCell = worksheet[XLSX.utils.encode_cell({r: R, c: 4})];
                        const statusCell = worksheet[XLSX.utils.encode_cell({r: R, c: 5})];
                        const agendamentoCell = worksheet[XLSX.utils.encode_cell({r: R, c: 6})];
                        
                        if (!usuarioCell || !usuarioCell.v) continue;
                        
                        let usuario = usuarioCell.v.toString().trim();
                        if (!usuario || usuario === '' || 
                            usuario.toUpperCase() === "ANGELOBRASIL" || 
                            usuario.toUpperCase() === "SECAUTOLOC" ||
                            usuario.toUpperCase() === "USUÁRIO" ||
                            usuario.toUpperCase() === "USUARIO") {
                            continue;
                        }
                        
                        row['Tipo'] = tipoCell && tipoCell.v ? tipoCell.v.toString().trim() : '';
                        row['Código'] = codigoCell && codigoCell.v ? codigoCell.v.toString().trim() : '';
                        row['Nro. processo'] = nroProcessoCell && nroProcessoCell.v ? nroProcessoCell.v.toString().trim() : '';
                        row['Usuário'] = usuario;
                        row['Status'] = statusCell && statusCell.v ? statusCell.v.toString().trim() : '';
                        row['Agendamento'] = agendamentoCell && agendamentoCell.v ? agendamentoCell.v.toString().trim() : '';
                        
                        if (dataCell && dataCell.v) {
                            if (dataCell.t === 'n') {
                                row['Data criação'] = new Date((dataCell.v - 25569) * 86400 * 1000);
                            } else {
                                const dateStr = dataCell.v.toString();
                                if (dateStr.includes('/')) {
                                    const [datePart, timePart] = dateStr.split(' ');
                                    const [day, month, year] = datePart.split('/');
                                    row['Data criação'] = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                                } else {
                                    row['Data criação'] = new Date(dataCell.v);
                                }
                            }
                        }
                        
                        row['peso'] = 1.0;
                        excelData.push(row);
                        currentFileRows++;
                    }
                    
                    totalRows += currentFileRows;
                    fileNames.push(file.name);
                    processedFiles++;
                    
                    if (processedFiles === files.length) {
                        extractMonthsFromData();
                        processExcelData();
                        
                        const fileList = fileNames.map(name => `• ${name}`).join('\n');
                        alert(`Dados carregados com sucesso!\n\nArquivos processados:\n${fileList}\n\nTotal de registros: ${totalRows}`);
                    }
                } catch (error) {
                    console.error('Erro ao processar arquivo:', file.name, error);
                    processedFiles++;
                    
                    if (processedFiles === files.length && excelData.length > 0) {
                        extractMonthsFromData();
                        processExcelData();
                        
                        const fileList = fileNames.map(name => `• ${name}`).join('\n');
                        alert(`Alguns arquivos foram processados!\n\nArquivos processados:\n${fileList}\n\nTotal de registros: ${totalRows}`);
                    }
                }
            };
            reader.readAsArrayBuffer(file);
        });
    };
    
    input.click();
}


function extractMonthsFromData() {
    if (!excelData || excelData.length === 0) {
        processedData.mesesDisponiveis = [];
        processedData.mesesAtivos = [];
        renderMonths();
        return;
    }
    
    const meses = new Set();
    
    excelData.forEach(row => {
        if (row['Data criação']) {
            try {
                let date = row['Data criação'];
                if (!(date instanceof Date)) {
                    date = new Date(date);
                }
                
                if (!isNaN(date.getTime())) {
                    const mesNomes = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                                     'jul', 'ago', 'set', 'out', 'nov', 'dez'];
                    const mesAno = `${mesNomes[date.getMonth()]}/${date.getFullYear()}`;
                    meses.add(mesAno);
                }
            } catch (e) {
                console.log('Erro ao processar data:', row['Data criação']);
            }
        }
    });
    
    const mesesArray = Array.from(meses).sort((a, b) => {
        const [mesA, anoA] = a.split('/');
        const [mesB, anoB] = b.split('/');
        const mesesOrdem = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
        
        if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
        return mesesOrdem.indexOf(mesA) - mesesOrdem.indexOf(mesB);
    });
    
    processedData.mesesDisponiveis = mesesArray;
    processedData.mesesAtivos = [...mesesArray];
    
    renderMonths();
}

function mostrarPopupExportacao() {
    document.getElementById('popup-exportacao').style.display = 'flex';
}

function fecharPopupExportacao() {
    document.getElementById('popup-exportacao').style.display = 'none';
}

window.onclick = function(event) {
    const popup = document.getElementById('popup-exportacao');
    if (event.target === popup) {
        fecharPopupExportacao();
    }
}

function processExcelData() {
    console.log('Iniciando processamento dos dados...');
    console.log('Dados brutos:', excelData ? excelData.length : 0);
    
    if (!excelData || excelData.length === 0) {
        processedData = {
            usuarios: [],
            totalMinutas: 0,
            mediaUsuario: 0,
            usuariosAtivos: 0,
            diaProdutivo: '--',
            topUsers: [],
            mesesAtivos: processedData ? processedData.mesesAtivos : [],
            mesesDisponiveis: processedData ? processedData.mesesDisponiveis : []
        };
        updateKPIs();
        createChart();
        return;
    }
    
    let filteredData = excelData;
    
    if (processedData.mesesAtivos && processedData.mesesAtivos.length > 0 && 
        processedData.mesesDisponiveis && processedData.mesesAtivos.length < processedData.mesesDisponiveis.length) {
        filteredData = excelData.filter(row => {
            if (!row['Data criação']) return false;
            
            try {
                let date = row['Data criação'];
                if (!(date instanceof Date)) {
                    date = new Date(date);
                }
                
                if (!isNaN(date.getTime())) {
                    const mesNomes = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                                     'jul', 'ago', 'set', 'out', 'nov', 'dez'];
                    const mesAno = `${mesNomes[date.getMonth()]}/${date.getFullYear()}`;
                    return processedData.mesesAtivos.includes(mesAno);
                }
                return false;
            } catch (e) {
                return false;
            }
        });
    }
    
    console.log(`Dados filtrados: ${filteredData.length} registros`);
    
    const usuariosMap = new Map();
    const diasSemana = new Map();
    
    filteredData.forEach(row => {
        const nroProcesso = row['Nro. processo'];
        if (!nroProcesso || nroProcesso.trim() === '') return;
        
        const usuario = row['Usuário'];
        if (!usuario || usuario.toString().trim() === '') return;
        
        const peso = parseFloat(row['peso']) || 1.0;
        
        if (!usuariosMap.has(usuario)) {
            usuariosMap.set(usuario, 0);
        }
        usuariosMap.set(usuario, usuariosMap.get(usuario) + peso);
        
        if (row['Data criação']) {
            try {
                let date = row['Data criação'];
                if (!(date instanceof Date)) {
                    date = new Date(date);
                }
                
                if (!isNaN(date.getTime())) {
                    const diasNomes = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
                    const diaSemana = diasNomes[date.getDay()];
                    
                    if (!diasSemana.has(diaSemana)) {
                        diasSemana.set(diaSemana, 0);
                    }
                    diasSemana.set(diaSemana, diasSemana.get(diaSemana) + peso);
                }
            } catch (e) {
                console.log('Erro ao processar data para dia da semana:', row['Data criação']);
            }
        }
    });
    
    const usuarios = Array.from(usuariosMap.entries())
        .map(([nome, minutas]) => ({ nome, minutas }))
        .sort((a, b) => b.minutas - a.minutas);
    
    console.log(`Usuários processados: ${usuarios.length}`);
    console.log('Top 5 usuários:', usuarios.slice(0, 5));
    
    const totalMinutas = Math.round(usuarios.reduce((sum, u) => sum + u.minutas, 0));
    const usuariosAtivos = usuarios.length;
    const mediaUsuario = usuarios.length > 0 ? usuarios.reduce((sum, u) => sum + u.minutas, 0) / usuarios.length : 0;
    
    let diaProdutivo = '--';
    if (diasSemana.size > 0) {
        const diaTop = Array.from(diasSemana.entries())
            .sort((a, b) => b[1] - a[1])[0];
        diaProdutivo = `${diaTop[0]} (${Math.round(diaTop[1])})`;
    }
    
    const topUsers = usuarios.slice(0, 3);
    
    processedData = {
        usuarios,
        totalMinutas,
        mediaUsuario,
        usuariosAtivos,
        diaProdutivo,
        topUsers,
        mesesAtivos: processedData.mesesAtivos || [],
        mesesDisponiveis: processedData.mesesDisponiveis || []
    };
    
    console.log('Dados processados finais:', processedData);
    
    updateKPIs();
    createChart();
}

function calcularERenderizarRanking() {
    if (!tabelaSemana || tabelaSemana.length === 0) {
        const container = document.getElementById('ranking-dias');
        if (container) {
            container.innerHTML = '<div class="ranking-item">Nenhum dado disponível</div>';
        }
        return;
    }

    const diasTotais = {
        segunda: 0,
        terca: 0,
        quarta: 0,
        quinta: 0,
        sexta: 0,
        sabado: 0,
        domingo: 0
    };

    tabelaSemana.forEach(usuario => {
        diasTotais.segunda += usuario.segunda || 0;
        diasTotais.terca += usuario.terca || 0;
        diasTotais.quarta += usuario.quarta || 0;
        diasTotais.quinta += usuario.quinta || 0;
        diasTotais.sexta += usuario.sexta || 0;
        diasTotais.sabado += usuario.sabado || 0;
        diasTotais.domingo += usuario.domingo || 0;
    });

    const diasNomes = {
        segunda: 'Segunda-feira',
        terca: 'Terça-feira',
        quarta: 'Quarta-feira',
        quinta: 'Quinta-feira',
        sexta: 'Sexta-feira',
        sabado: 'Sábado',
        domingo: 'Domingo'
    };

    const rankingOrdenado = Object.entries(diasTotais)
        .map(([dia, total]) => ({
            dia: diasNomes[dia],
            chave: dia,
            total: Math.round(total * 10) / 10
        }))
        .sort((a, b) => b.total - a.total);

    renderizarRanking(rankingOrdenado);
}

function renderizarRanking(rankingData) {
    const container = document.getElementById('ranking-dias');
    if (!container) return;

    container.innerHTML = '';

    if (rankingData.length === 0) {
        container.innerHTML = '<div class="ranking-item">Nenhum dado disponível</div>';
        return;
    }

    const coresRanking = [
        '#FFD700', // Ouro - 1º lugar
        '#C0C0C0', // Prata - 2º lugar  
        '#CD7F32', // Bronze - 3º lugar
        '#4a90e2', // Azul padrão - outros
        '#4a90e2',
        '#4a90e2',
        '#4a90e2'
    ];

    const iconesRanking = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣'];

    rankingData.forEach((item, index) => {
        const rankingDiv = document.createElement('div');
        rankingDiv.className = 'ranking-item';
        rankingDiv.style.borderLeftColor = coresRanking[index];
        
        const porcentagem = rankingData[0].total > 0 ? 
            Math.round((item.total / rankingData[0].total) * 100) : 0;

        rankingDiv.innerHTML = `
            <div class="ranking-posicao">
                <span class="ranking-icone">${iconesRanking[index]}</span>
                <span class="ranking-numero">#${index + 1}</span>
            </div>
            <div class="ranking-dia">${item.dia}</div>
            <div class="ranking-total">${item.total} minutas</div>
            <div class="ranking-porcentagem">${porcentagem}%</div>
            <div class="ranking-barra">
                <div class="ranking-barra-preenchida" style="width: ${porcentagem}%; background-color: ${coresRanking[index]};"></div>
            </div>
        `;
        
        container.appendChild(rankingDiv);
    });
}

function gerarTabelaSemana() {
    if (!excelData || excelData.length === 0) return;

    let filteredData = excelData;
    
    if (processedData.mesesAtivos.length > 0 && processedData.mesesAtivos.length < processedData.mesesDisponiveis.length) {
        filteredData = excelData.filter(row => {
            if (!row['Data criação']) return false;
            
            try {
                let date = row['Data criação'];
                if (!(date instanceof Date)) {
                    date = new Date(date);
                }
                
                if (!isNaN(date.getTime())) {
                    const mesNomes = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                                     'jul', 'ago', 'set', 'out', 'nov', 'dez'];
                    const mesAno = `${mesNomes[date.getMonth()]}/${date.getFullYear()}`;
                    return processedData.mesesAtivos.includes(mesAno);
                }
                return false;
            } catch (e) {
                return false;
            }
        });
    }

    const usuariosMap = new Map();
    const diasSemana = ['domingo', 'segunda', 'terca', 'quarta', 'quinta', 'sexta', 'sabado'];
    
    filteredData.forEach(row => {
        const nroProcesso = row['Nro. processo'];
        if (!nroProcesso || nroProcesso.trim() === '') return;
        
        const usuario = row['Usuário'];
        if (!usuario || usuario.toString().trim() === '') return;
        
        const peso = parseFloat(row['peso']) || 1.0;
        
        if (!usuariosMap.has(usuario)) {
            usuariosMap.set(usuario, {
                usuario: usuario,
                domingo: 0,
                segunda: 0,
                terca: 0,
                quarta: 0,
                quinta: 0,
                sexta: 0,
                sabado: 0
            });
        }
        
        if (row['Data criação']) {
            try {
                let date = row['Data criação'];
                if (!(date instanceof Date)) {
                    date = new Date(date);
                }
                
                if (!isNaN(date.getTime())) {
                    const diaSemana = diasSemana[date.getDay()];
                    usuariosMap.get(usuario)[diaSemana] += peso;
                }
            } catch (e) {
                console.log('Erro ao processar data:', row['Data criação']);
            }
        }
    });

    tabelaSemana = Array.from(usuariosMap.values());
    renderizarTabelaSemana();
}

function renderizarTabelaSemana() {
    const container = document.getElementById('tabela-semana');
    if (!container) return;
    
    container.innerHTML = '';
    
    const filtro = document.getElementById('filtro-semana').value.toLowerCase();
    const usuariosFiltrados = tabelaSemana.filter(usuario => 
        usuario.usuario.toLowerCase().includes(filtro)
    );

    usuariosFiltrados.forEach(usuario => {
        const row = document.createElement('tr');
        
        const total = usuario.domingo + usuario.segunda + usuario.terca + 
                     usuario.quarta + usuario.quinta + usuario.sexta + usuario.sabado;
        
        row.innerHTML = `
            <td class="usuario-cell">${usuario.usuario}</td>
            <td class="dia-cell">${Math.round(usuario.segunda)}</td>
            <td class="dia-cell">${Math.round(usuario.terca)}</td>
            <td class="dia-cell">${Math.round(usuario.quarta)}</td>
            <td class="dia-cell">${Math.round(usuario.quinta)}</td>
            <td class="dia-cell">${Math.round(usuario.sexta)}</td>
            <td class="dia-cell">${Math.round(usuario.sabado)}</td>
            <td class="dia-cell">${Math.round(usuario.domingo)}</td>
            <td class="total-cell">${Math.round(total)}</td>
        `;
        
        container.appendChild(row);
    });

    calcularERenderizarRanking();
}

function exportarTabelaSemana(formato) {
    if (!tabelaSemana || tabelaSemana.length === 0) {
        alert('Não há dados para exportar!');
        return;
    }
    
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    
    if (formato === 'pdf') {
        exportarPDF(timestamp);
    } else if (formato === 'xlsx') {
        exportarExcel(timestamp);
    }
}

function exportarPDF(timestamp) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('landscape');
    
    doc.setFillColor(74, 144, 226);
    doc.rect(0, 0, doc.internal.pageSize.width, 25, 'F');
    
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('APROD - Produtividade por Dia da Semana', 15, 16);
    
    doc.setTextColor(74, 144, 226);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')}`, 15, 35);
    
    let mesesTexto = `Meses: ${processedData.mesesAtivos.join(', ')}`;
    if (processedData.mesesAtivos.length === 0) {
        mesesTexto = 'Nenhum mês selecionado';
    }
    doc.text(mesesTexto, 15, 42);
    
    const tableData = tabelaSemana.map(usuario => {
        const total = usuario.domingo + usuario.segunda + usuario.terca + 
                     usuario.quarta + usuario.quinta + usuario.sexta + usuario.sabado;
        return [
            usuario.usuario,
            Math.round(usuario.segunda),
            Math.round(usuario.terca),
            Math.round(usuario.quarta),
            Math.round(usuario.quinta),
            Math.round(usuario.sexta),
            Math.round(usuario.sabado),
            Math.round(usuario.domingo),
            Math.round(total)
        ];
    });
    
    doc.autoTable({
        head: [['Usuário', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo', 'Total']],
        body: tableData,
        startY: 52,
        styles: {
            fontSize: 8,
            cellPadding: 3,
            textColor: [35, 41, 70],
            lineColor: [74, 144, 226],
            lineWidth: 0.5
        },
        headStyles: {
            fillColor: [74, 144, 226],
            textColor: [255, 255, 255],
            fontStyle: 'bold',
            fontSize: 9
        },
        alternateRowStyles: {
            fillColor: [248, 249, 250]
        },
        columnStyles: {
            0: { cellWidth: 40, fontStyle: 'bold' },
            8: { fillColor: [227, 242, 253], fontStyle: 'bold' }
        },
        margin: { left: 15, right: 15 }
    });
    
    const totalGeral = tabelaSemana.reduce((acc, usuario) => {
        return acc + usuario.domingo + usuario.segunda + usuario.terca + 
               usuario.quarta + usuario.quinta + usuario.sexta + usuario.sabado;
    }, 0);
    
    const finalY = doc.previousAutoTable.finalY + 10;
    doc.setFillColor(227, 242, 253);
    doc.rect(15, finalY, doc.internal.pageSize.width - 30, 15, 'F');
    
    doc.setTextColor(74, 144, 226);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(`Total Geral: ${Math.round(totalGeral)} minutas`, 20, finalY + 10);
    
    doc.save(`tabela_semana_${timestamp}.pdf`);
}

function exportarExcel(timestamp) {
    let mesesTexto = `Meses: ${processedData.mesesAtivos.join(', ')}`;
    if (processedData.mesesAtivos.length === 0) {
        mesesTexto = 'Nenhum mês selecionado';
    }
    
    const wb = XLSX.utils.book_new();
    const ws = {};
    
    ws['A1'] = { 
        v: 'APROD - Produtividade por Dia da Semana', 
        t: 's',
        s: {
            font: { name: "Calibri", sz: 18, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "4A90E2" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
                top: { style: "thick", color: { rgb: "4A90E2" } },
                bottom: { style: "thick", color: { rgb: "4A90E2" } },
                left: { style: "thick", color: { rgb: "4A90E2" } },
                right: { style: "thick", color: { rgb: "4A90E2" } }
            }
        }
    };
    
    ws['A3'] = { 
        v: `Gerado em: ${new Date().toLocaleString('pt-BR')}`, 
        t: 's',
        s: {
            font: { name: "Calibri", sz: 11, color: { rgb: "4A90E2" } },
            alignment: { horizontal: "left", vertical: "center" }
        }
    };
    
    ws['A4'] = { 
        v: mesesTexto, 
        t: 's',
        s: {
            font: { name: "Calibri", sz: 11, color: { rgb: "4A90E2" } },
            alignment: { horizontal: "left", vertical: "center" }
        }
    };
    
    const headers = ['Usuário', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo', 'Total'];
    const headerCells = ['A6', 'B6', 'C6', 'D6', 'E6', 'F6', 'G6', 'H6', 'I6'];
    
    headers.forEach((header, index) => {
        ws[headerCells[index]] = {
            v: header,
            t: 's',
            s: {
                font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "4A90E2" } },
                alignment: { horizontal: "center", vertical: "center" },
                border: {
                    top: { style: "medium", color: { rgb: "4A90E2" } },
                    bottom: { style: "medium", color: { rgb: "4A90E2" } },
                    left: { style: "medium", color: { rgb: "4A90E2" } },
                    right: { style: "medium", color: { rgb: "4A90E2" } }
                }
            }
        };
    });
    
    let row = 7;
    tabelaSemana.forEach((usuario, index) => {
        const total = usuario.domingo + usuario.segunda + usuario.terca + 
                     usuario.quarta + usuario.quinta + usuario.sexta + usuario.sabado;
        
        const isEvenRow = index % 2 === 0;
        const rowData = [
            usuario.usuario,
            Math.round(usuario.segunda),
            Math.round(usuario.terca),
            Math.round(usuario.quarta),
            Math.round(usuario.quinta),
            Math.round(usuario.sexta),
            Math.round(usuario.sabado),
            Math.round(usuario.domingo),
            Math.round(total)
        ];
        
        const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
        
        rowData.forEach((value, colIndex) => {
            const cellAddress = `${columns[colIndex]}${row}`;
            
            if (colIndex === 0) {
                ws[cellAddress] = {
                    v: value,
                    t: 's',
                    s: {
                        font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "232946" } },
                        fill: { fgColor: { rgb: isEvenRow ? "FFFFFF" : "F8F9FA" } },
                        alignment: { horizontal: "left", vertical: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "D0D0D0" } },
                            bottom: { style: "thin", color: { rgb: "D0D0D0" } },
                            left: { style: "thin", color: { rgb: "D0D0D0" } },
                            right: { style: "thin", color: { rgb: "D0D0D0" } }
                        }
                    }
                };
            } else if (colIndex === 8) {
                ws[cellAddress] = {
                    v: value,
                    t: 'n',
                    s: {
                        font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "4A90E2" } },
                        fill: { fgColor: { rgb: "E3F2FD" } },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "4A90E2" } },
                            bottom: { style: "thin", color: { rgb: "4A90E2" } },
                            left: { style: "thin", color: { rgb: "4A90E2" } },
                            right: { style: "thin", color: { rgb: "4A90E2" } }
                        }
                    }
                };
            } else {
                ws[cellAddress] = {
                    v: value,
                    t: 'n',
                    s: {
                        font: { name: "Calibri", sz: 11, color: { rgb: "232946" } },
                        fill: { fgColor: { rgb: isEvenRow ? "FFFFFF" : "F8F9FA" } },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "D0D0D0" } },
                            bottom: { style: "thin", color: { rgb: "D0D0D0" } },
                            left: { style: "thin", color: { rgb: "D0D0D0" } },
                            right: { style: "thin", color: { rgb: "D0D0D0" } }
                        }
                    }
                };
            }
        });
        
        row++;
    });
    
    const totalGeral = tabelaSemana.reduce((acc, usuario) => {
        return acc + usuario.domingo + usuario.segunda + usuario.terca + 
               usuario.quarta + usuario.quinta + usuario.sexta + usuario.sabado;
    }, 0);
    
    row += 1;
    const totalColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
    const totalValues = ['Total Geral:', '', '', '', '', '', '', '', Math.round(totalGeral)];
    
    totalValues.forEach((value, colIndex) => {
        const cellAddress = `${totalColumns[colIndex]}${row}`;
        ws[cellAddress] = {
            v: value,
            t: colIndex === 0 || colIndex === 8 ? (typeof value === 'string' ? 's' : 'n') : 's',
            s: {
                font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "4A90E2" } },
                alignment: { horizontal: colIndex === 0 ? "left" : "center", vertical: "center" },
                border: {
                    top: { style: "thick", color: { rgb: "4A90E2" } },
                    bottom: { style: "thick", color: { rgb: "4A90E2" } },
                    left: { style: "thick", color: { rgb: "4A90E2" } },
                    right: { style: "thick", color: { rgb: "4A90E2" } }
                }
            }
        };
    });
    
    const lastRow = row;
    ws['!ref'] = `A1:I${lastRow}`;
    
    ws['!cols'] = [
        { wch: 25 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
        { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }
    ];
    
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }];
    
    ws['!rows'] = [
        { hpt: 30 },
        { hpt: 15 },
        { hpt: 18 },
        { hpt: 18 },
        { hpt: 15 },
        { hpt: 25 }
    ];
    
    for (let i = 6; i < lastRow; i++) {
        if (!ws['!rows'][i]) ws['!rows'][i] = {};
        ws['!rows'][i].hpt = i === lastRow - 1 ? 25 : 20;
    }
    
    XLSX.utils.book_append_sheet(wb, ws, "Produtividade Semanal");
    XLSX.writeFile(wb, `tabela_semana_${timestamp}.xlsx`);
}

function ordenarTabelaSemana(coluna) {
    if (!tabelaSemana) return;
    
    const isAscending = sortColumn === coluna ? !sortAscending : true;
    sortColumn = coluna;
    sortAscending = isAscending;
    
    tabelaSemana.sort((a, b) => {
        let valueA, valueB;
        
        if (coluna === 'usuario') {
            valueA = a.usuario.toLowerCase();
            valueB = b.usuario.toLowerCase();
        } else if (coluna === 'total') {
            valueA = a.domingo + a.segunda + a.terca + a.quarta + a.quinta + a.sexta + a.sabado;
            valueB = b.domingo + b.segunda + b.terca + b.quarta + b.quinta + b.sexta + b.sabado;
        } else {
            valueA = a[coluna];
            valueB = b[coluna];
        }
        
        if (valueA < valueB) return isAscending ? -1 : 1;
        if (valueA > valueB) return isAscending ? 1 : -1;
        return 0;
    });
    
    renderizarTabelaSemana();
}

function calcularRankingDias() {
    if (!tabelaSemana) return [];
    
    const diasTotais = {
        segunda: 0,
        terca: 0,
        quarta: 0,
        quinta: 0,
        sexta: 0,
        sabado: 0,
        domingo: 0
    };
    
    tabelaSemana.forEach(usuario => {
        diasTotais.segunda += usuario.segunda;
        diasTotais.terca += usuario.terca;
        diasTotais.quarta += usuario.quarta;
        diasTotais.quinta += usuario.quinta;
        diasTotais.sexta += usuario.sexta;
        diasTotais.sabado += usuario.sabado;
        diasTotais.domingo += usuario.domingo;
    });
    
    const diasNomes = {
        segunda: 'Segunda-feira',
        terca: 'Terça-feira',
        quarta: 'Quarta-feira',
        quinta: 'Quinta-feira',
        sexta: 'Sexta-feira',
        sabado: 'Sábado',
        domingo: 'Domingo'
    };
    
    return Object.entries(diasTotais)
        .map(([dia, total]) => ({
            dia: diasNomes[dia],
            total: Math.round(total)
        }))
        .sort((a, b) => b.total - a.total);
}

function showSemanaPage() {
    document.querySelectorAll('.page-content').forEach(page => page.style.display = 'none');
    document.getElementById('semana-page').style.display = 'block';
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelectorAll('.nav-item')[2].classList.add('active');
    gerarTabelaSemana();
    setTimeout(() => {
        calcularERenderizarRanking();
    }, 100);
}

function toggleTheme() {
    isDarkTheme = !isDarkTheme;
    updateTheme();
    createChart();
}

function updateTheme() {
    const theme = isDarkTheme ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', theme);
    document.getElementById('theme-icon').textContent = isDarkTheme ? '🌙' : '☀️';
    
    localStorage.setItem('theme', theme);
}

function showPesosPage() {
    document.querySelectorAll('.page-content').forEach(page => page.style.display = 'none');
    document.getElementById('pesos-page').style.display = 'block';
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelectorAll('.nav-item')[3].classList.add('active');
    carregarTiposPesos();
}

function carregarTiposPesos() {
    if (!excelData || excelData.length === 0) {
        const container = document.getElementById('tabela-pesos');
        container.innerHTML = '<tr><td colspan="3" style="text-align: center; color: var(--text-secondary);">Nenhum dado disponível. Importe arquivos Excel primeiro.</td></tr>';
        return;
    }

    const tipos = new Set();
    
    console.log('Total de registros para análise:', excelData.length);
    console.log('Primeiros 5 registros:', excelData.slice(0, 5));
    
    excelData.forEach((row, index) => {
        if (index < 10) {
            console.log(`Registro ${index}:`, row);
        }
        
        Object.keys(row).forEach(key => {
            if (key.toLowerCase().includes('tipo') || key.toLowerCase().includes('agendamento')) {
                const valor = row[key];
                if (valor !== undefined && valor !== null && valor !== '') {
                    const valorStr = valor.toString().trim();
                    if (valorStr !== '' && valorStr.toLowerCase() !== 'tipo' && valorStr.toLowerCase() !== 'agendamento') {
                        tipos.add(valorStr);
                        console.log(`Tipo encontrado na coluna "${key}": "${valorStr}"`);
                    }
                }
            }
        });
    });

    console.log('Tipos únicos encontrados:', Array.from(tipos));

    if (tipos.size === 0) {
        const container = document.getElementById('tabela-pesos');
        container.innerHTML = '<tr><td colspan="3" style="text-align: center; color: var(--text-secondary);">Nenhum tipo de agendamento válido encontrado nos dados importados. Verifique se a coluna existe e tem dados.</td></tr>';
        return;
    }

    tiposPesos = Array.from(tipos).sort();
    
    tiposPesos.forEach(tipo => {
        if (!(tipo in pesosAtuais)) {
            pesosAtuais[tipo] = 1.0;
        }
    });

    renderizarTabelaPesos();
}

function renderizarTabelaPesos() {
    const container = document.getElementById('tabela-pesos');
    container.innerHTML = '';

    if (tiposPesos.length === 0) {
        container.innerHTML = '<tr><td colspan="3" style="text-align: center; color: var(--text-secondary);">Nenhum tipo de agendamento encontrado.</td></tr>';
        return;
    }

    tiposPesos.forEach(tipo => {
        const row = document.createElement('tr');
        
        let tipoLimpo = tipo;
        if (tipo.includes('(') && tipo.includes(')')) {
            tipoLimpo = tipo.substring(0, tipo.indexOf('(')).trim();
        }
        
        row.innerHTML = `
            <td class="tipo-cell">${tipoLimpo}</td>
            <td>
                <input type="number" 
                       class="peso-input" 
                       value="${pesosAtuais[tipo]}" 
                       min="0" 
                       max="10" 
                       step="1"
                       onchange="atualizarPeso('${tipo}', this.value)">
            </td>
            <td>
                <div class="peso-controls-cell">
                    <button class="peso-btn decrement" onclick="decrementarPeso('${tipo}')">−</button>
                    <button class="peso-btn increment" onclick="incrementarPeso('${tipo}')">+</button>
                </div>
            </td>
        `;
        
        container.appendChild(row);
    });
}

function atualizarPeso(tipo, novoValor) {
    const valor = parseFloat(novoValor);
    if (!isNaN(valor) && valor >= 0) {
        pesosAtuais[tipo] = valor;
    }
}

function decrementarPeso(tipo) {
    if (pesosAtuais[tipo] > 0) {
        pesosAtuais[tipo] = Math.max(0, pesosAtuais[tipo] - 1);
        renderizarTabelaPesos();
    }
}

function incrementarPeso(tipo) {
    if (pesosAtuais[tipo] < 10) {
        pesosAtuais[tipo] = Math.min(10, pesosAtuais[tipo] + 1);
        renderizarTabelaPesos();
    }
}

function salvarPesos() {
    try {
        const pesosJson = JSON.stringify(pesosAtuais, null, 2);
        const blob = new Blob([pesosJson], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.href = url;
        link.download = `pesos_${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        URL.revokeObjectURL(url);
        
        alert('Pesos salvos com sucesso!');
    } catch (error) {
        console.error('Erro ao salvar pesos:', error);
        alert('Erro ao salvar pesos.');
    }
}

function aplicarPesos() {
    if (!excelData || excelData.length === 0) {
        alert('Nenhum dado disponível para aplicar pesos!');
        return;
    }

    excelData.forEach(row => {
        const tipo = row['Tipo'];
        if (tipo && pesosAtuais[tipo]) {
            row['peso'] = pesosAtuais[tipo];
        } else {
            row['peso'] = 1.0;
        }
    });

    processExcelData();
    
    if (document.getElementById('semana-page').style.display === 'block') {
        gerarTabelaSemana();
    }

    alert('Pesos aplicados com sucesso! Os gráficos e tabelas foram atualizados.');
}

function restaurarPesosPadrao() {
    if (confirm('Tem certeza que deseja restaurar todos os pesos para 1.0?')) {
        tiposPesos.forEach(tipo => {
            pesosAtuais[tipo] = 1.0;
        });
        
        renderizarTabelaPesos();
        alert('Pesos restaurados para o valor padrão (1.0).');
    }
}

document.addEventListener('DOMContentLoaded', function() {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme) {
        isDarkTheme = savedTheme === 'dark';
    }
    
    initializeApp();
    
    const filtroUsuarios = document.getElementById('filtro-usuarios');
    if (filtroUsuarios) {
        filtroUsuarios.addEventListener('input', function() {
            renderizarListaUsuarios();
        });
    }
    
    const filtroSemana = document.getElementById('filtro-semana');
    if (filtroSemana) {
        filtroSemana.addEventListener('input', function() {
            renderizarTabelaSemana();
        });
    }
});

window.addEventListener('resize', function() {
    if (chartInstance) {
        chartInstance.resize();
    }
});
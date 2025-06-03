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
    const ctx = document.getElementById('produtividade-chart').getContext('2d');
    
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
    const chartData = processedData.usuarios.slice(0, 20);
    
    if (chartData.length === 0) {
        chartInstance = new Chart(ctx, {
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
                    return;
                }
    
    chartInstance = new Chart(ctx, {
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
        if (files.length > 0) {
            files.forEach(file => processFile(file));
        }
    };
    
    input.click();
}

function processFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            processExcelFile(e.target.result, file.name);
        } catch (error) {
            alert('Erro ao processar arquivo: ' + error.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function processExcelFile(arrayBuffer, fileName) {
    try {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const data = [];
        
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
            if (usuario.toUpperCase() === "ANGELOBRASIL" || usuario.toUpperCase() === "SECAUTOLOC") {
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
            
            data.push(row);
        }
        
        console.log(`Dados processados: ${data.length} registros`);
        console.log('Amostra de dados processados:', data.slice(0, 3));
        
        if (!excelData) {
            excelData = [];
        }
        excelData = excelData.concat(data);
        extractMonthsFromData();
        processExcelData();
        
        alert(`Arquivo "${fileName}" carregado com sucesso!\n${data.length} registros processados.`);
    } catch (error) {
        console.error('Erro detalhado:', error);
        alert('Erro ao processar arquivo Excel: ' + error.message);
    }
}

function calcularPeso(agendamento) {
    if (!agendamento) return 1.0;
    
    const tipoLimpo = agendamento.replace(/\s*\(.*?\)/, '').trim().toUpperCase();
    return 1.0;
}

function extractMonthsFromData() {
    if (!excelData || excelData.length === 0) return;
    
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

function processExcelData() {
    console.log('Iniciando processamento dos dados...');
    
    if (!excelData || excelData.length === 0) {
        processedData = {
            usuarios: [],
            totalMinutas: 0,
            mediaUsuario: 0,
            usuariosAtivos: 0,
            diaProdutivo: '--',
            topUsers: [],
            mesesAtivos: processedData.mesesAtivos,
            mesesDisponiveis: processedData.mesesDisponiveis
        };
        updateKPIs();
        updateChart();
        return;
    }
    
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
                    const isIncluded = processedData.mesesAtivos.includes(mesAno);
                    return isIncluded;
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
        
        if (row['Data criação'] && nroProcesso && nroProcesso.trim() !== '') {
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
    
    const totalMinutas = filteredData.filter(row => row['Nro. processo'] && row['Nro. processo'].trim() !== '').length;
    const usuariosAtivos = usuarios.length;
    const mediaUsuario = usuarios.length > 0 ? usuarios.reduce((sum, u) => sum + u.minutas, 0) / usuarios.length : 0;
    
    let diaProdutivo = '--';
    if (diasSemana.size > 0) {
        const diaTop = Array.from(diasSemana.entries())
            .sort((a, b) => b[1] - a[1])[0];
        const valorFormatado = diaTop[1] >= 1000 ? `${Math.round(diaTop[1])}` : `${diaTop[1].toFixed(1)}`;
        diaProdutivo = `${diaTop[0]}\n(${valorFormatado})`;
    }
    
    const topUsers = usuarios.slice(0, 3);
    
    console.log('Top 3 usuários detalhado:', topUsers);
    console.log('KPIs calculados:', {
        totalMinutas,
        mediaUsuario: mediaUsuario.toFixed(1),
        usuariosAtivos,
        diaProdutivo,
        topUsers: topUsers.map(u => `${u.nome}: ${u.minutas.toFixed(1)}`)
    });
    
    processedData = {
        usuarios,
        totalMinutas,
        mediaUsuario,
        usuariosAtivos,
        diaProdutivo,
        topUsers,
        mesesAtivos: processedData.mesesAtivos,
        mesesDisponiveis: processedData.mesesDisponiveis
    };
    
    updateKPIs();
    updateChart();
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

document.addEventListener('DOMContentLoaded', function() {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme) {
        isDarkTheme = savedTheme === 'dark';
    }
    
    initializeApp();
});

window.addEventListener('resize', function() {
    if (chartInstance) {
        chartInstance.resize();
    }
});
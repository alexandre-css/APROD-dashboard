let sortColumnAssuntos = 'minutas';
let sortAscendingAssuntos = false;
let chartAssuntos = null;
let assuntosJSON = null;
let anosDisponiveis = [];
let anoSelecionado = new Date().getFullYear();
let tabelaMes = [];
let sortColumnMes = 'usuario';
let sortAscendingMes = true;
let tiposPesos = [];
let pesosAtuais = {};
let pesosAtivos = true;
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
let configuracoes = {
    nomeGabinete: '',
    incluirLogo: true,
    incluirData: true,
    formatoData: 'br',
    temaInicial: 'dark',
    itensGrafico: 20,
    pesoPadrao: 1.0,
    autoAplicarPesos: false,
    iaHabilitada: false,
    iaToken: '6f76bea40d73b6ff291fd29780ca51cc6e12d00c',
    iaModelo: 'finetuned-gpt-neox-20b'
};
let assuntosData = [];
let assuntosProcessados = {
    minutasPorAssunto: new Map(),
    processosPorAssunto: new Map(),
    codigosAssuntos: new Map(),
    totalMinutasAssuntos: 0,
    totalProcessosAssuntos: 0
};

function initializeApp() {
    carregarConfiguracoes();
    updateTheme();
    updatePesosButton();
    createChart();
    renderMonths();
    updateKPIs();
}

function showDashboard() {
    hideAllPages();
    document.getElementById('dashboard-page').style.display = 'block';
    setActiveNavItem(0);
}

function showComparacaoPage() {
    hideAllPages();
    document.getElementById('comparar-page').style.display = 'block';
    setActiveNavItem(1);
    gerarDadosComparacao();
}

function showMesPage() {
    hideAllPages();
    document.getElementById('mes-page').style.display = 'block';
    setActiveNavItem(2);
    gerarTabelaMes();
}

function showSemanaPage() {
    hideAllPages();
    document.getElementById('semana-page').style.display = 'block';
    setActiveNavItem(3);
    gerarTabelaSemana();
}

async function showAssuntosPage() {
    hideAllPages();
    document.getElementById('assuntos-page').style.display = 'block';
    setActiveNavItem(4);
    
    try {
        await carregarAssuntosJSON();
        await gerarDadosAssuntos();
        gerarTabelaAssuntos();
        renderizarResumoAssuntos();
        criarGraficoVariacaoAssuntos();
    } catch (error) {
        console.error('Erro ao carregar página de assuntos:', error);
        
        try {
            await gerarDadosAssuntos();
            gerarTabelaAssuntos();
            renderizarResumoAssuntos();
            criarGraficoVariacaoAssuntos();
        } catch (fallbackError) {
            console.error('Erro no fallback:', fallbackError);
            document.getElementById('tabela-assuntos').innerHTML = 
                '<tr><td colspan="5" style="text-align: center; padding: 40px; color: var(--text-secondary);">Erro ao carregar dados de assuntos. Verifique se há dados importados.</td></tr>';
        }
    }
}

function showPesosPage() {
    hideAllPages();
    document.getElementById('pesos-page').style.display = 'block';
    setActiveNavItem(5);
    carregarTiposPesos();
}

function showIAPage() {
    hideAllPages();
    document.getElementById('ia-page').style.display = 'block';
    setActiveNavItem(6);
    
    const modeloSelect = document.getElementById('ia-modelo-direto');
    if (modeloSelect) {
        for (let i = 0; i < modeloSelect.options.length; i++) {
            if (modeloSelect.options[i].value === configuracoes.iaModelo) {
                modeloSelect.selectedIndex = i;
                break;
            }
        }
    }
}

function showConfiguracoesPage() {
    hideAllPages();
    document.getElementById('configuracoes-page').style.display = 'block';
    setActiveNavItem(7);
    carregarConfiguracoes();
}

function hideAllPages() {
    document.querySelectorAll('.page-content').forEach(page => page.style.display = 'none');
}

function setActiveNavItem(index) {
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => item.classList.remove('active'));
    if (navItems[index]) {
        navItems[index].classList.add('active');
    }
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
        usuarioDiv.onclick = function() {
            toggleUsuarioComparacao(usuario.nome);
        };
        
        usuarioDiv.innerHTML = `
            <div class="usuario-checkbox">
                <input type="checkbox" 
                       id="user-${index}" 
                       ${usuario.selecionado ? 'checked' : ''}
                       onclick="event.stopPropagation();"
                       onchange="toggleUsuarioComparacao('${usuario.nome}')">
                <label for="user-${index}" onclick="event.stopPropagation();"></label>
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
        const chartContainer = ctx.parentElement;
        chartContainer.innerHTML = `
            <div class="grafico-header">
                <h3>Comparação de Produtividade</h3>
            </div>
            <div style="text-align: center; padding: 40px; color: var(--text-secondary);">
                <p>Selecione usuários para visualizar a comparação segmentada</p>
            </div>
            <canvas id="comparacao-chart"></canvas>
        `;
        
        const newCtx = document.getElementById('comparacao-chart');
        chartComparacao = new Chart(newCtx, {
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

    const detalhesUsuarios = gerarDetalhesTooltipUsuarios(usuariosSelecionados);
    
    const chartData = usuariosSelecionados.map(usuario => {
        const detalhes = detalhesUsuarios[usuario.nome] || [];
        
        let acordaoMinutas = 0;
        let despachoMinutas = 0;
        
        detalhes.forEach(tipo => {
            if (tipo.nome === 'Acórdão') {
                acordaoMinutas = tipo.minutas;
            } else {
                despachoMinutas += tipo.minutas;
            }
        });
        
        return {
            nome: usuario.nome,
            acordao: Math.round(acordaoMinutas * 10) / 10,
            despacho: Math.round(despachoMinutas * 10) / 10,
            total: Math.round((acordaoMinutas + despachoMinutas) * 10) / 10
        };
    });

    const isDark = document.documentElement.getAttribute('data-theme') !== 'light';

    const cores = [
        '#1565c0', '#43a047', '#e53935', '#8e24aa', '#fb8c00',
        '#00acc1', '#7cb342', '#d81b60', '#5e35b1', '#ff7043',
        '#26a69a', '#ab47bc', '#42a5f5', '#66bb6a', '#ef5350'
    ];

    const coresDespacho = [
        '#42a5f5', '#66bb6a', '#ef5350', '#ba68c8', '#ffb74d',
        '#4dd0e1', '#9ccc65', '#f06292', '#7986cb', '#ff8a65',
        '#4db6ac', '#ce93d8', '#64b5f6', '#81c784', '#e57373'
    ];

    const acordaoData = chartData.map((usuario, index) => usuario.acordao);
    const despachoData = chartData.map((usuario, index) => usuario.despacho);
    const backgroundColorsAcordao = chartData.map((_, index) => cores[index % cores.length]);
    const backgroundColorsDespacho = chartData.map((_, index) => coresDespacho[index % coresDespacho.length]);

    const chartContainer = ctx.parentElement;
    chartContainer.innerHTML = `
        <div class="grafico-header">
            <h3>Comparação de Produtividade por Tipo</h3>
        </div>
        <canvas id="comparacao-chart"></canvas>
    `;

    const newCtx = document.getElementById('comparacao-chart');

    chartComparacao = new Chart(newCtx, {
        type: 'bar',
        data: {
            labels: chartData.map(u => u.nome),
            datasets: [
                {
                    label: 'Acórdão',
                    data: acordaoData,
                    backgroundColor: backgroundColorsAcordao,
                    borderColor: backgroundColorsAcordao,
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                    stack: 'Stack 0'
                },
                {
                    label: 'Despacho/Decisão',
                    data: despachoData,
                    backgroundColor: backgroundColorsDespacho,
                    borderColor: backgroundColorsDespacho,
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                    stack: 'Stack 0'
                }
            ]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: {
                    right: 40,
                    left: 10,
                    top: 10,
                    bottom: 10
                }
            },
            scales: {
                x: {
                    stacked: true,
                    beginAtZero: true,
                    grid: {
                        color: isDark ? '#404040' : '#e0e0e0'
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946',
                        callback: function(value) {
                            return value;
                        },
                        maxTicksLimit: 10
                    }
                },
                y: {
                    stacked: true,
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946',
                        maxTicksLimit: Math.min(chartData.length, 15),
                        font: {
                            size: Math.max(10, Math.min(12, 400 / chartData.length))
                        }
                    }
                }
            },
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    enabled: false,
                    external: function(context) {
                        const chart = context.chart;
                        const tooltip = context.tooltip;
                        
                        let tooltipEl = chart.canvas.parentNode.querySelector('.chart-tooltip-custom');
                        
                        if (!tooltipEl) {
                            tooltipEl = document.createElement('div');
                            tooltipEl.className = 'chart-tooltip-custom';
                            tooltipEl.style.cssText = `
                                position: absolute;
                                pointer-events: none;
                                z-index: 1000;
                                background: var(--card-bg);
                                border: 2px solid var(--border-color);
                                border-radius: 8px;
                                padding: 12px;
                                font-size: 12px;
                                line-height: 1.4;
                                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
                                min-width: 200px;
                                transition: opacity 0.2s ease;
                            `;
                            chart.canvas.parentNode.appendChild(tooltipEl);
                        }
                        
                        if (tooltip.opacity === 0) {
                            tooltipEl.style.opacity = '0';
                            return;
                        }
                        
                        if (tooltip.body && tooltip.dataPoints && tooltip.dataPoints.length > 0) {
                            const activePoint = tooltip.dataPoints[0];
                            const dataIndex = activePoint.dataIndex;
                            const nomeUsuario = chartData[dataIndex].nome;
                            const detalhesComparacao = gerarDetalhesTooltipComparacao(usuariosSelecionados);
                            const dadosUsuario = detalhesComparacao[nomeUsuario];
                            
                            if (dadosUsuario) {
                                tooltipEl.innerHTML = `
                                    <div style="font-weight: bold; font-size: 14px; color: var(--text-primary); margin-bottom: 8px; border-bottom: 1px solid var(--border-color); padding-bottom: 4px; text-align: center;">
                                        ${nomeUsuario}
                                    </div>
                                    <div style="display: flex; align-items: center; gap: 8px; margin: 6px 0; padding: 2px 0;">
                                        <div style="width: 14px; height: 14px; border-radius: 3px; background-color: ${dadosUsuario.acordao.cor}; border: 1px solid rgba(255, 255, 255, 0.3); box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15); flex-shrink: 0;"></div>
                                        <span style="color: var(--text-primary); font-weight: 500; font-size: 12px;">Acórdão: ${dadosUsuario.acordao.minutas} minutas</span>
                                    </div>
                                    <div style="display: flex; align-items: center; gap: 8px; margin: 6px 0; padding: 2px 0;">
                                        <div style="width: 14px; height: 14px; border-radius: 3px; background-color: ${dadosUsuario.despacho.cor}; border: 1px solid rgba(255, 255, 255, 0.3); box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15); flex-shrink: 0;"></div>
                                        <span style="color: var(--text-primary); font-weight: 500; font-size: 12px;">Despacho/Decisão: ${dadosUsuario.despacho.minutas} minutas</span>
                                    </div>
                                    <div style="margin-top: 8px; padding-top: 8px; border-top: 1px solid var(--border-color); font-weight: bold; color: var(--accent); text-align: center; font-size: 13px;">
                                        Total: ${dadosUsuario.total} minutas
                                    </div>
                                `;
                            }
                        }
                        
                        const canvasRect = chart.canvas.getBoundingClientRect();
                        const containerRect = chart.canvas.parentNode.getBoundingClientRect();
                        
                        let left = tooltip.caretX + canvasRect.left - containerRect.left;
                        let top = tooltip.caretY + canvasRect.top - containerRect.top;
                        
                        const tooltipWidth = tooltipEl.offsetWidth;
                        const tooltipHeight = tooltipEl.offsetHeight;
                        const containerWidth = containerRect.width;
                        const containerHeight = containerRect.height;
                        
                        if (left + tooltipWidth > containerWidth - 20) {
                            left = containerWidth - tooltipWidth - 20;
                        }
                        if (left < 20) {
                            left = 20;
                        }
                        
                        if (top + tooltipHeight > containerHeight - 20) {
                            top = top - tooltipHeight - 20;
                        }
                        if (top < 20) {
                            top = 20;
                        }
                        
                        tooltipEl.style.left = left + 'px';
                        tooltipEl.style.top = top + 'px';
                        tooltipEl.style.opacity = '1';
                    }
                }
            },
            interaction: {
                mode: 'index',
                intersect: false,
                axis: 'y'
            },
            animation: {
                duration: 1000,
                easing: 'easeOutQuart'
            },
            onHover: function(event, activeElements) {
                event.native.target.style.cursor = activeElements.length > 0 ? 'pointer' : 'default';
            }
        },
        plugins: [{
            id: 'datalabels',
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
                
                chart.data.datasets.forEach((dataset, datasetIndex) => {
                    const meta = chart.getDatasetMeta(datasetIndex);
                    meta.data.forEach((bar, index) => {
                        if (datasetIndex === 1) {
                            const total = chartData[index].total;
                            ctx.fillStyle = isDark ? '#ffffff' : '#232946';
                            ctx.font = `bold ${Math.max(10, Math.min(12, 400 / chartData.length))}px Arial`;
                            ctx.textAlign = 'left';
                            ctx.textBaseline = 'middle';
                            ctx.fillText(Math.round(total), bar.x + 5, bar.y);
                        }
                    });
                });
            }
        }]
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
        
        const detalhesUsuario = document.querySelector('.detalhes-usuario');
        if (detalhesUsuario) {
            const titulo = detalhesUsuario.querySelector('h3');
            if (titulo) {
                const nomeUsuario = titulo.textContent.replace('Detalhamento - ', '');
                const chartData = processedData.usuarios ? processedData.usuarios.slice(0, 20) : [];
                const usuarioAtual = chartData.find(u => u.nome === nomeUsuario);
                
                if (usuarioAtual) {
                    const tiposDetalhados = gerarDetalhesTooltipUsuarios(chartData);
                    const detalhes = tiposDetalhados[nomeUsuario];
                    mostrarDetalhesUsuario(nomeUsuario, detalhes, usuarioAtual.minutas);
                }
            }
        }
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
    
    const tiposDetalhados = gerarDetalhesTooltipUsuarios(chartData);
    
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
            onClick: (event, activeElements) => {
                if (activeElements.length > 0) {
                    const index = activeElements[0].index;
                    const nomeUsuario = chartData[index].nome;
                    const detalhes = tiposDetalhados[nomeUsuario];
                    mostrarDetalhesUsuario(nomeUsuario, detalhes, chartData[index].minutas);
                }
            },
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
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            return context[0].label;
                        },
                        label: function(context) {
                            const nomeUsuario = context.label;
                            const detalhes = tiposDetalhados[nomeUsuario];
                            
                            if (!detalhes || detalhes.length === 0) {
                                return [`Total: ${Math.round(context.raw)} minutas`, 'Clique para mais detalhes'];
                            }
                            
                            const acordao = detalhes.find(d => d.nome === 'Acórdão')?.minutas || 0;
                            const outrosTotal = detalhes.filter(d => d.nome !== 'Acórdão')
                                                      .reduce((sum, item) => sum + item.minutas, 0);
                            
                            return [
                                `Total: ${Math.round(context.raw)} minutas`,
                                '',
                                `Acórdão: ${Math.round(acordao)} minutas`,
                                `Despacho/Decisão: ${Math.round(outrosTotal)} minutas`,
                                '',
                                'Clique para mais detalhes'
                            ];
                        }
                    },
                    backgroundColor: isDark ? '#2d2d2d' : '#ffffff',
                    titleColor: isDark ? '#ffffff' : '#232946',
                    bodyColor: isDark ? '#ffffff' : '#232946',
                    borderColor: isDark ? '#4a90e2' : '#3cb3e6',
                    borderWidth: 2,
                    cornerRadius: 8,
                    displayColors: false,
                    padding: 12,
                    titleFont: {
                        size: 14,
                        weight: 'bold'
                    },
                    bodyFont: {
                        size: 12
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

function mostrarDetalhesUsuario(nomeUsuario, detalhes, totalMinutas) {
    const chartContainer = document.querySelector('.chart-container');
    
    chartContainer.style.transition = 'all 0.3s ease';
    chartContainer.style.opacity = '0';
    chartContainer.style.transform = 'scale(0.95)';
    
    setTimeout(() => {
        const acordaoMinutas = detalhes.find(d => d.nome === 'Acórdão')?.minutas || 0;
        const despachoMinutas = totalMinutas - acordaoMinutas;
        
        const outrosTipos = detalhes.filter(d => d.nome !== 'Acórdão');
        
        chartContainer.innerHTML = `
            <div class="detalhes-usuario">
                <div class="detalhes-header">
                    <div class="detalhes-titulo">
                        <div class="detalhes-icone-titulo">
                            <svg viewBox="0 0 24 24" fill="none">
                                <rect x="3" y="4" width="18" height="12" rx="2" stroke="currentColor" stroke-width="2"/>
                                <path d="M7 8h10M7 12h7" stroke="currentColor" stroke-width="2"/>
                            </svg>
                        </div>
                        <div class="detalhes-texto">
                            <h3>Detalhamento - ${nomeUsuario}</h3>
                            <div class="detalhes-total">Total: ${Math.round(totalMinutas)} minutas</div>
                        </div>
                    </div>
                    <button class="btn-voltar" onclick="voltarGrafico()">
                        <svg viewBox="0 0 24 24" fill="none">
                            <path d="M19 12H5M12 19l-7-7 7-7" stroke="currentColor" stroke-width="2"/>
                        </svg>
                        <span>Voltar ao Gráfico</span>
                    </button>
                </div>
                
                <div class="detalhes-conteudo">
                    <div class="categoria-detalhes acordao-categoria">
                        <div class="categoria-header">
                            <div class="categoria-icone">
                                <canvas id="chart-acordao-${Date.now()}"></canvas>
                            </div>
                            <div class="categoria-info">
                                <h4>Acórdão</h4>
                                <div class="categoria-valor">${Math.round(acordaoMinutas)} minutas</div>
                                <div class="categoria-percentual">${Math.round((acordaoMinutas / totalMinutas) * 100)}% do total</div>
                            </div>
                        </div>
                        <div class="categoria-barra">
                            <div class="barra-preenchida acordao-barra" style="width: ${(acordaoMinutas / totalMinutas) * 100}%"></div>
                        </div>
                    </div>
                    
                    <div class="categoria-detalhes despacho-categoria">
                        <div class="categoria-header">
                            <div class="categoria-icone">
                                <canvas id="chart-despacho-${Date.now()}"></canvas>
                            </div>
                            <div class="categoria-info">
                                <h4>Despacho/Decisão</h4>
                                <div class="categoria-valor">${Math.round(despachoMinutas)} minutas</div>
                                <div class="categoria-percentual">${Math.round((despachoMinutas / totalMinutas) * 100)}% do total</div>
                            </div>
                        </div>
                        <div class="categoria-barra">
                            <div class="barra-preenchida despacho-barra" style="width: ${(despachoMinutas / totalMinutas) * 100}%"></div>
                        </div>
                        
                        ${outrosTipos.length > 0 ? `
                            <div class="subcategorias">
                                <div class="subcategorias-header">
                                    <svg viewBox="0 0 24 24" fill="none">
                                        <path d="M9 12l2 2 4-4M21 12a9 9 0 11-18 0 9 9 0 0118 0z" stroke="currentColor" stroke-width="2"/>
                                    </svg>
                                    <h5>Subcategorias</h5>
                                </div>
                                <div class="subcategorias-lista">
                                    ${outrosTipos.map(tipo => `
                                        <div class="subcategoria-item">
                                            <div class="subcategoria-indicador"></div>
                                            <span class="subcategoria-nome">${tipo.nome}</span>
                                            <span class="subcategoria-valor">${Math.round(tipo.minutas)} minutas</span>
                                            <div class="subcategoria-barra-mini">
                                                <div class="subcategoria-preenchida" style="width: ${(tipo.minutas / despachoMinutas) * 100}%"></div>
                                            </div>
                                        </div>
                                    `).join('')}
                                </div>
                            </div>
                        ` : ''}
                    </div>
                </div>
                
                <div class="detalhes-resumo">
                    <div class="resumo-estatisticas">
                        <div class="stat-item">
                            <span class="stat-label">Tipo Predominante</span>
                            <span class="stat-valor">${acordaoMinutas > despachoMinutas ? 'Acórdão' : 'Despacho/Decisão'}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">Variedade</span>
                            <span class="stat-valor">${detalhes.length} tipo${detalhes.length > 1 ? 's' : ''}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">Eficiência</span>
                            <span class="stat-valor">${totalMinutas > 50 ? 'Alta' : totalMinutas > 20 ? 'Média' : 'Baixa'}</span>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        chartContainer.style.opacity = '1';
        chartContainer.style.transform = 'scale(1)';
        
        setTimeout(() => {
            criarMiniGraficosPizza(acordaoMinutas, despachoMinutas, totalMinutas);
            
            setTimeout(() => {
                const acordaoIcone = document.querySelector('.acordao-categoria .categoria-icone');
                const despachoIcone = document.querySelector('.despacho-categoria .categoria-icone');
                
                if (acordaoIcone) {
                    acordaoIcone.classList.add('pulse');
                    setTimeout(() => acordaoIcone.classList.remove('pulse'), 4000);
                }
                
                if (despachoIcone) {
                    despachoIcone.classList.add('pulse');
                    setTimeout(() => despachoIcone.classList.remove('pulse'), 4000);
                }
            }, 800);
            
            const elementos = chartContainer.querySelectorAll('.categoria-detalhes, .subcategoria-item');
            elementos.forEach((el, index) => {
                el.style.opacity = '0';
                el.style.transform = 'translateY(20px)';
                setTimeout(() => {
                    el.style.transition = 'all 0.3s ease';
                    el.style.opacity = '1';
                    el.style.transform = 'translateY(0)';
                }, index * 100);
            });
        }, 100);
        
    }, 300);
}

function criarMiniGraficosPizza(acordaoMinutas, despachoMinutas, totalMinutas) {
    const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
    
    const acordaoCanvas = document.querySelector('[id^="chart-acordao-"]');
    const despachoCanvas = document.querySelector('[id^="chart-despacho-"]');
    
    if (acordaoCanvas) {
        const ctx = acordaoCanvas.getContext('2d');
        new Chart(ctx, {
            type: 'doughnut',
            data: {
                datasets: [{
                    data: [acordaoMinutas, despachoMinutas],
                    backgroundColor: ['#FFFFFF', '#1565c0'],
                    borderWidth: 2,
                    borderColor: ['#1565c0', '#1565c0'],
                    offset: [12, 0]
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                cutout: '65%',
                animation: {
                    animateRotate: true,
                    duration: 1500,
                    easing: 'easeOutCubic'
                },
                plugins: {
                    legend: { display: false },
                    tooltip: { enabled: false }
                }
            }
        });
    }
    
    if (despachoCanvas) {
        const ctx = despachoCanvas.getContext('2d');
        new Chart(ctx, {
            type: 'doughnut',
            data: {
                datasets: [{
                    data: [acordaoMinutas, despachoMinutas],
                    backgroundColor: ['#1565c0', '#FFFFFF'],
                    borderWidth: 2,
                    borderColor: ['#1565c0', '#4a90e2'],
                    offset: [0, 12]
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                cutout: '65%',
                animation: {
                    animateRotate: true,
                    duration: 1500,
                    easing: 'easeOutCubic'
                },
                plugins: {
                    legend: { display: false },
                    tooltip: { enabled: false }
                }
            }
        });
    }
}

function voltarGrafico() {
    const chartContainer = document.querySelector('.chart-container');
    
    chartContainer.style.transition = 'all 0.3s ease';
    chartContainer.style.opacity = '0';
    chartContainer.style.transform = 'scale(0.95)';
    
    setTimeout(() => {
        chartContainer.innerHTML = '<canvas id="produtividade-chart"></canvas>';
        chartContainer.style.opacity = '1';
        chartContainer.style.transform = 'scale(1)';
        
        setTimeout(() => {
            createChart();
        }, 100);
    }, 300);
}

function gerarDetalhesTooltipUsuarios(usuarios) {
    if (!excelData || excelData.length === 0) return {};
    
    const usuariosDetalhes = {};
    
    usuarios.forEach(usuario => {
        usuariosDetalhes[usuario.nome] = [];
    });
    
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
    
    filteredData.forEach(row => {
        const nroProcesso = row['Nro. processo'];
        if (!nroProcesso || nroProcesso.trim() === '') return;
        
        const usuario = row['Usuário'];
        if (!usuario || usuario.toString().trim() === '') return;
        
        if (!usuariosDetalhes[usuario]) return;
        
        const peso = parseFloat(row['peso']) || 1.0;
        const tipoValue = row['Tipo'];
        const agendamentoValue = row['Agendamento'];
        
        let tipoFinal = 'Outros';
        
        if (tipoValue && typeof tipoValue === 'string') {
            const tipoLimpo = tipoValue.toString().trim().toUpperCase();
            
            if (tipoLimpo === 'ACÓRDÃO' || tipoLimpo === 'VOTO-VISTA' || tipoLimpo === 'VOTO DIVERGENTE') {
                tipoFinal = 'Acórdão';
            }
            else if (tipoLimpo === 'DESPACHO/DECISÃO') {
                if (agendamentoValue && typeof agendamentoValue === 'string') {
                    let agendamentoLimpo = agendamentoValue.toString().trim();
                    
                    if (agendamentoLimpo.includes('(') && agendamentoLimpo.includes(')')) {
                        agendamentoLimpo = agendamentoLimpo.substring(0, agendamentoLimpo.indexOf('(')).trim();
                    }
                    
                    if (agendamentoLimpo && agendamentoLimpo !== '') {
                        tipoFinal = agendamentoLimpo;
                    }
                }
            }
            else if (tipoLimpo !== '' && tipoLimpo !== 'TIPO') {
                tipoFinal = tipoValue.toString().trim();
            }
        }
        
        const tipoExistente = usuariosDetalhes[usuario].find(t => t.nome === tipoFinal);
        if (tipoExistente) {
            tipoExistente.minutas += peso;
        } else {
            usuariosDetalhes[usuario].push({
                nome: tipoFinal,
                minutas: peso
            });
        }
    });
    
    Object.keys(usuariosDetalhes).forEach(usuario => {
        usuariosDetalhes[usuario].sort((a, b) => {
            if (a.nome === 'Acórdão') return -1;
            if (b.nome === 'Acórdão') return 1;
            return b.minutas - a.minutas;
        });
    });
    
    return usuariosDetalhes;
}

function gerarDetalhesTooltipComparacao(usuarios) {
    const cores = [
        '#1565c0', '#43a047', '#e53935', '#8e24aa', '#fb8c00',
        '#00acc1', '#7cb342', '#d81b60', '#5e35b1', '#ff7043',
        '#26a69a', '#ab47bc', '#42a5f5', '#66bb6a', '#ef5350'
    ];

    const coresDespacho = [
        '#42a5f5', '#66bb6a', '#ef5350', '#ba68c8', '#ffb74d',
        '#4dd0e1', '#9ccc65', '#f06292', '#7986cb', '#ff8a65',
        '#4db6ac', '#ce93d8', '#64b5f6', '#81c784', '#e57373'
    ];

    const resultado = {};
    
    if (!excelData || excelData.length === 0) return resultado;
    
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
    
    usuarios.forEach((usuario, index) => {
        let acordaoMinutas = 0;
        let despachoMinutas = 0;
        
        filteredData.forEach(row => {
            const nroProcesso = row['Nro. processo'];
            if (!nroProcesso || nroProcesso.trim() === '') return;
            
            const usuarioRow = row['Usuário'];
            if (!usuarioRow || usuarioRow.toString().trim() === '') return;
            
            if (usuarioRow !== usuario.nome) return;
            
            const peso = parseFloat(row['peso']) || 1.0;
            const tipoValue = row['Tipo'];
            const agendamentoValue = row['Agendamento'];
            
            if (tipoValue && typeof tipoValue === 'string') {
                const tipoLimpo = tipoValue.toString().trim().toUpperCase();
                
                if (tipoLimpo === 'ACÓRDÃO' || tipoLimpo === 'VOTO-VISTA' || tipoLimpo === 'VOTO DIVERGENTE') {
                    acordaoMinutas += peso;
                } else if (tipoLimpo === 'DESPACHO/DECISÃO') {
                    despachoMinutas += peso;
                } else {
                    despachoMinutas += peso;
                }
            } else {
                despachoMinutas += peso;
            }
        });
        
        resultado[usuario.nome] = {
            acordao: {
                minutas: Math.round(acordaoMinutas * 10) / 10,
                cor: cores[index % cores.length]
            },
            despacho: {
                minutas: Math.round(despachoMinutas * 10) / 10,
                cor: coresDespacho[index % coresDespacho.length]
            },
            total: Math.round((acordaoMinutas + despachoMinutas) * 10) / 10
        };
    });
    
    return resultado;
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
        
        logger.info(`Iniciando importação de ${files.length} arquivo(s)`, { arquivos: files.map(f => f.name) });
        
        let processedFiles = 0;
        let totalRows = 0;
        let fileNames = [];
        let allErrors = [];
        excelData = [];
        
        files.forEach((file, index) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    logger.info(`Processando arquivo: ${file.name}`);
                    
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    const range = XLSX.utils.decode_range(worksheet['!ref']);
                    
                    const colunaMapping = detectarColunas(worksheet, range);
                    if (colunaMapping.errors.length > 0) {
                        const errorMsg = `Colunas não encontradas: ${colunaMapping.errors.join(', ')}`;
                        logger.error(`Arquivo ${file.name}: ${errorMsg}`, { erros: colunaMapping.errors });
                        allErrors.push(`${file.name}: ${errorMsg}`);
                        processedFiles++;
                        checkProcessingComplete();
                        return;
                    }
                    
                    let currentFileRows = 0;
                    let invalidRows = 0;
                    let linhasExcluidas = [];
                    
                    for (let R = 2; R <= range.e.r; ++R) {
                        const numeroLinha = R + 1;
                        const row = {};
                        
                        const tipoCell = worksheet[XLSX.utils.encode_cell({r: R, c: colunaMapping.tipo})];
                        const codigoCell = worksheet[XLSX.utils.encode_cell({r: R, c: colunaMapping.codigo})];
                        const nroProcessoCell = worksheet[XLSX.utils.encode_cell({r: R, c: colunaMapping.processo})];
                        const usuarioCell = worksheet[XLSX.utils.encode_cell({r: R, c: colunaMapping.usuario})];
                        const dataCell = worksheet[XLSX.utils.encode_cell({r: R, c: colunaMapping.data})];
                        const statusCell = worksheet[XLSX.utils.encode_cell({r: R, c: colunaMapping.status})];
                        const agendamentoCell = worksheet[XLSX.utils.encode_cell({r: R, c: colunaMapping.agendamento})];
                        
                        if (!usuarioCell || !usuarioCell.v) {
                            invalidRows++;
                            linhasExcluidas.push({
                                linha: numeroLinha,
                                motivo: "Usuário vazio ou inexistente"
                            });
                            continue;
                        }
                        
                        let usuario = usuarioCell.v.toString().trim();
                        if (!usuario || usuario === '' || 
                            usuario.toUpperCase() === "ANGELOBRASIL" || 
                            usuario.toUpperCase() === "SECAUTOLOC" ||
                            usuario.toUpperCase() === "USUÁRIO" ||
                            usuario.toUpperCase() === "USUARIO") {
                            invalidRows++;
                            linhasExcluidas.push({
                                linha: numeroLinha,
                                motivo: `Usuário inválido: "${usuario}"`
                            });
                            continue;
                        }
                        
                        const tipo = tipoCell && tipoCell.v ? tipoCell.v.toString().trim() : '';
                        const status = statusCell && statusCell.v ? statusCell.v.toString().trim() : '';
                        const nroProcesso = nroProcessoCell && nroProcessoCell.v ? nroProcessoCell.v.toString().trim() : '';
                        
                        if (!isValidStatus(status)) {
                            invalidRows++;
                            linhasExcluidas.push({
                                linha: numeroLinha,
                                motivo: `Status inválido: "${status}"`
                            });
                            continue;
                        }
                        if (!isValidTipo(tipo)) {
                            invalidRows++;
                            linhasExcluidas.push({
                                linha: numeroLinha,
                                motivo: `Tipo inválido: "${tipo}"`
                            });
                            continue;
                        }
                        if (!nroProcesso || nroProcesso === '') {
                            invalidRows++;
                            linhasExcluidas.push({
                                linha: numeroLinha,
                                motivo: "Número do processo vazio"
                            });
                            continue;
                        }
                        
                        row['Tipo'] = tipo;
                        row['Cod. assunto'] = codigoCell && codigoCell.v ? codigoCell.v.toString().trim() : '';
                        row['Nro. processo'] = nroProcesso;
                        row['Usuário'] = usuario;
                        row['Status'] = status;
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
                    
                    if (linhasExcluidas.length > 0) {
                        const exemploPorMotivo = {};
                        const contagemPorMotivo = {};
                        
                        linhasExcluidas.forEach(item => {
                            const categoria = item.motivo.split(':')[0];
                            
                            if (!contagemPorMotivo[categoria]) {
                                contagemPorMotivo[categoria] = 0;
                                exemploPorMotivo[categoria] = [];
                            }
                            
                            contagemPorMotivo[categoria]++;
                            
                            if (exemploPorMotivo[categoria].length < 10) {
                                exemploPorMotivo[categoria].push(item);
                            }
                        });
                        
                        logger.warn(`Linhas excluídas no arquivo ${file.name}`, {
                            totalExcluidas: linhasExcluidas.length,
                            resumoDetalhado: Object.keys(contagemPorMotivo).map(categoria => ({
                                categoria: categoria,
                                total: contagemPorMotivo[categoria],
                                exemplos: exemploPorMotivo[categoria]
                            })),
                            resumoMotivos: contagemPorMotivo
                        });
                    }
                    
                    logger.info(`Arquivo ${file.name} processado`, { 
                        registrosValidos: currentFileRows, 
                        registrosInvalidos: invalidRows,
                        totalLinhas: range.e.r - 1
                    });
                    
                    totalRows += currentFileRows;
                    fileNames.push(file.name);
                    processedFiles++;
                    
                    checkProcessingComplete();
                    
                } catch (error) {
                    logger.error(`Erro ao processar arquivo ${file.name}`, error);
                    allErrors.push(`${file.name}: Erro na leitura do arquivo`);
                    processedFiles++;
                    checkProcessingComplete();
                }
            };
            reader.readAsArrayBuffer(file);
        });
        
        function checkProcessingComplete() {
            if (processedFiles === files.length) {
                logger.info(`Importação finalizada`, { 
                    totalArquivos: files.length,
                    arquivosProcessados: fileNames.length,
                    totalRegistros: totalRows,
                    erros: allErrors.length
                });
                
                if (allErrors.length > 0) {
                    alert(`Alguns arquivos não puderam ser processados:\n\n${allErrors.join('\n')}\n\nArquivos processados com sucesso:\n${fileNames.map(name => `• ${name}`).join('\n')}\n\nTotal de registros válidos: ${totalRows}`);
                } else {
                    alert(`Dados carregados com sucesso!\n\nArquivos processados:\n${fileNames.map(name => `• ${name}`).join('\n')}\n\nTotal de registros válidos: ${totalRows}`);
                }
                
                if (excelData.length > 0) {
                    extractMonthsFromData();
                    processExcelData();
                }
            }
        }
    };
    
    input.click();
}


function extrairMesAno(data) {
    const date = new Date(data);
    if (isNaN(date.getTime())) return null;
    
    const mesNomes = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                     'jul', 'ago', 'set', 'out', 'nov', 'dez'];
    return `${mesNomes[date.getMonth()]}/${date.getFullYear()}`;
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
    const popupDashboard = document.getElementById('popup-exportacao-dashboard');
    const popupMes = document.getElementById('popup-exportacao-mes');
    if (event.target === popup) {
        fecharPopupExportacao();
    }
    if (event.target === popupDashboard) {
        fecharPopupExportacaoDashboard();
    }
    if (event.target === popupMes) {
        fecharPopupExportacaoMes();
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
    
    const totalMinutas = Math.round(filteredData.length);
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

function detectarColunas(worksheet, range) {
    const headers = {};
    const errors = [];
    const colunaMapping = {};
    
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const headerCell = worksheet[XLSX.utils.encode_cell({r: 1, c: C})];
        if (headerCell && headerCell.v) {
            const headerValue = headerCell.v.toString().trim();
            headers[C] = headerValue;
        }
    }
    
    const colunasNecessarias = {
        'Tipo': ['Tipo'],
        'Cod. assunto': ['Código', 'Cod. assunto'],
        'Nro. processo': ['Nro. processo'],
        'Usuário': ['Usuário'],
        'Data criação': ['Data criação'],
        'Status': ['Status'],
        'Agendamento': ['Agendamento']
    };
    
    Object.keys(colunasNecessarias).forEach(coluna => {
        let encontrada = false;
        Object.keys(headers).forEach(indice => {
            const headerValue = headers[indice];
            if (colunasNecessarias[coluna].includes(headerValue)) {
                switch(coluna) {
                    case 'Tipo':
                        colunaMapping.tipo = parseInt(indice);
                        break;
                    case 'Cod. assunto':
                        colunaMapping.codigo = parseInt(indice);
                        break;
                    case 'Nro. processo':
                        colunaMapping.processo = parseInt(indice);
                        break;
                    case 'Usuário':
                        colunaMapping.usuario = parseInt(indice);
                        break;
                    case 'Data criação':
                        colunaMapping.data = parseInt(indice);
                        break;
                    case 'Status':
                        colunaMapping.status = parseInt(indice);
                        break;
                    case 'Agendamento':
                        colunaMapping.agendamento = parseInt(indice);
                        break;
                }
                encontrada = true;
            }
        });
        
        if (!encontrada) {
            errors.push(`Coluna "${coluna}" não encontrada`);
        }
    });
    
    return {
        ...colunaMapping,
        errors: errors
    };
}

let cacheProcessamento = {
    dadosOriginais: null,
    hashDados: null,
    resultadoProcessado: null
};

function gerarHashDados(dados) {
    return dados.length + '_' + dados.map(r => r['Nro. processo']).join('').slice(0, 50);
}

function validarDadosCompletos(row) {
    const camposObrigatorios = ['Nro. processo', 'Usuário', 'Tipo', 'Status', 'Data criação'];
    const erros = [];
    
    camposObrigatorios.forEach(campo => {
        if (!row[campo] || row[campo].toString().trim() === '') {
            erros.push(`Campo "${campo}" vazio`);
        }
    });
    
    if (row['Data criação'] && isNaN(new Date(row['Data criação']).getTime())) {
        erros.push('Data inválida');
    }
    
    return {
        valido: erros.length === 0,
        erros: erros
    };
}

function processarComValidacao(row, numeroLinha) {
    const validacao = validarDadosCompletos(row);
    
    if (!validacao.valido) {
        console.warn(`Linha ${numeroLinha}: ${validacao.erros.join(', ')}`);
        return false;
    }
    
    return true;
}

function processExcelDataComCache() {
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

    const hashAtual = gerarHashDados(excelData) + '_' + processedData.mesesAtivos.join(',');
    
    if (cacheProcessamento.hashDados === hashAtual && cacheProcessamento.resultadoProcessado) {
        processedData = { ...cacheProcessamento.resultadoProcessado };
        updateKPIs();
        createChart();
        return;
    }

    processExcelData();
    
    cacheProcessamento.hashDados = hashAtual;
    cacheProcessamento.resultadoProcessado = { ...processedData };
}

function isValidStatus(status) {
    if (!status || status === '') return false;
    
    const statusValidos = [
        'Anexada ao processo',
        'Assinada',
        'Conferida',
        'Enviada para jurisprudência',
        'Enviada para o diário eletrônico',
        'Para assinar',
        'Para conferir'
    ];
    
    return statusValidos.includes(status);
}

function isValidTipo(tipo) {
    if (!tipo || tipo === '') return false;
    
    const tiposValidos = [
        'ACÓRDÃO',
        'DESPACHO/DECISÃO',
        'VOTO DIVERGENTE',
        'VOTO-VISTA'
    ];
    
    return tiposValidos.includes(tipo);
}

async function processarDadosEmLotes(dados, tamanhoLote = 1000) {
    const resultados = [];
    
    for (let i = 0; i < dados.length; i += tamanhoLote) {
        const lote = dados.slice(i, i + tamanhoLote);
        const resultadoLote = await processarLote(lote);
        resultados.push(...resultadoLote);
        
        if (i % (tamanhoLote * 5) === 0) {
            await new Promise(resolve => setTimeout(resolve, 10));
        }
    }
    
    return resultados;
}

function gerarRelatorioQualidade(dados) {
    const relatorio = {
        totalRegistros: dados.length,
        registrosValidos: 0,
        registrosInvalidos: 0,
        camposVazios: {},
        tiposEncontrados: new Set(),
        statusEncontrados: new Set(),
        periodoCoberto: { inicio: null, fim: null },
        usuariosUnicos: new Set()
    };
    
    dados.forEach(row => {
        let registroValido = true;
        
        Object.keys(row).forEach(campo => {
            if (!row[campo] || row[campo].toString().trim() === '') {
                relatorio.camposVazios[campo] = (relatorio.camposVazios[campo] || 0) + 1;
                registroValido = false;
            }
        });
        
        if (row['Tipo']) relatorio.tiposEncontrados.add(row['Tipo']);
        if (row['Status']) relatorio.statusEncontrados.add(row['Status']);
        if (row['Usuário']) relatorio.usuariosUnicos.add(row['Usuário']);
        
        if (row['Data criação']) {
            const data = new Date(row['Data criação']);
            if (!isNaN(data.getTime())) {
                if (!relatorio.periodoCoberto.inicio || data < relatorio.periodoCoberto.inicio) {
                    relatorio.periodoCoberto.inicio = data;
                }
                if (!relatorio.periodoCoberto.fim || data > relatorio.periodoCoberto.fim) {
                    relatorio.periodoCoberto.fim = data;
                }
            }
        }
        
        if (registroValido) {
            relatorio.registrosValidos++;
        } else {
            relatorio.registrosInvalidos++;
        }
    });
    
    console.log('Relatório de Qualidade dos Dados:', relatorio);
    return relatorio;
}

const logger = {
    logs: [],
    
    info(mensagem, dados = null) {
        const log = {
            nivel: 'INFO',
            timestamp: new Date().toISOString(),
            mensagem,
            dados
        };
        this.logs.push(log);
        console.log(`[INFO] ${mensagem}`, dados);
    },
    
    warn(mensagem, dados = null) {
        const log = {
            nivel: 'WARN',
            timestamp: new Date().toISOString(),
            mensagem,
            dados
        };
        this.logs.push(log);
        console.warn(`[WARN] ${mensagem}`, dados);
    },
    
    error(mensagem, erro = null) {
        const log = {
            nivel: 'ERROR',
            timestamp: new Date().toISOString(),
            mensagem,
            erro: erro ? erro.message : null
        };
        this.logs.push(log);
        console.error(`[ERROR] ${mensagem}`, erro);
    },
    
    exportarLogs() {
        const blob = new Blob([JSON.stringify(this.logs, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `aprod_logs_${new Date().toISOString().slice(0, 10)}.json`;
        link.click();
        URL.revokeObjectURL(url);
    }
};

function mostrarLogsProcessamento() {
    const logsContainer = document.createElement('div');
    logsContainer.className = 'logs-container';
    logsContainer.innerHTML = `
        <div class="logs-overlay">
            <div class="logs-modal">
                <div class="logs-header">
                    <h3>Logs de Processamento</h3>
                    <button onclick="fecharLogsProcessamento()" class="btn-fechar">×</button>
                </div>
                <div class="logs-content">
                    <div class="logs-stats">
                        <div class="stat-item">
                            <span class="stat-label">Total de Logs:</span>
                            <span class="stat-valor">${logger.logs.length}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">Erros:</span>
                            <span class="stat-valor">${logger.logs.filter(l => l.nivel === 'ERROR').length}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">Avisos:</span>
                            <span class="stat-valor">${logger.logs.filter(l => l.nivel === 'WARN').length}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">Dados Importados:</span>
                            <span class="stat-valor">${excelData ? excelData.length : 0} registros</span>
                        </div>
                    </div>
                    <div class="logs-list">
                        ${logger.logs.length === 0 ? 
                            '<div class="log-item info"><div class="log-mensagem">Nenhum log de processamento disponível ainda. Os logs são gerados durante a importação de dados.</div></div>' :
                            logger.logs.map(log => `
                                <div class="log-item ${log.nivel.toLowerCase()}">
                                    <div class="log-header">
                                        <span class="log-nivel">${log.nivel}</span>
                                        <span class="log-timestamp">${new Date(log.timestamp).toLocaleString('pt-BR')}</span>
                                    </div>
                                    <div class="log-mensagem">${log.mensagem}</div>
                                    ${log.dados ? `<div class="log-dados">${JSON.stringify(log.dados)}</div>` : ''}
                                    ${log.erro ? `<div class="log-dados">Erro: ${log.erro}</div>` : ''}
                                </div>
                            `).join('')
                        }
                    </div>
                </div>
                <div class="logs-footer">
                    <button onclick="logger.exportarLogs()" class="btn-exportar" ${logger.logs.length === 0 ? 'disabled' : ''}>
                        Exportar Logs
                    </button>
                    <button onclick="limparLogs()" class="btn-limpar" ${logger.logs.length === 0 ? 'disabled' : ''}>
                        Limpar Logs
                    </button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(logsContainer);
}

function fecharLogsProcessamento() {
    const container = document.querySelector('.logs-container');
    if (container) {
        container.remove();
    }
}

function limparLogs() {
    if (confirm('Tem certeza que deseja limpar todos os logs?')) {
        logger.logs = [];
        fecharLogsProcessamento();
        alert('Logs limpos com sucesso!');
    }
}

function processarComFallback(dados) {
    try {
        return processExcelData();
    } catch (erro) {
        logger.error('Erro no processamento principal, tentando fallback', erro);
        
        try {
            return processarModoPadrao(dados);
        } catch (erroFallback) {
            logger.error('Erro também no fallback', erroFallback);
            return processarModoSeguro(dados);
        }
    }
}

function processarModoSeguro(dados) {
    const dadosBasicos = dados.filter(row => 
        row['Nro. processo'] && 
        row['Usuário'] && 
        row['Tipo'] && 
        row['Status']
    );
    
    logger.info(`Modo seguro: processando ${dadosBasicos.length} de ${dados.length} registros`);
    
    return {
        usuarios: [],
        totalMinutas: dadosBasicos.length,
        mediaUsuario: 0,
        usuariosAtivos: new Set(dadosBasicos.map(r => r['Usuário'])).size,
        diaProdutivo: '--',
        topUsers: [],
        mesesAtivos: [],
        mesesDisponiveis: []
    };
}

function processarLote(lote) {
    return new Promise(resolve => {
        const processados = lote.filter(row => {
            const nroProcesso = row['Nro. processo'];
            if (!nroProcesso || nroProcesso.trim() === '') return false;
            
            const usuario = row['Usuário'];
            if (!usuario || usuario.toString().trim() === '') return false;
            
            const tipo = row['Tipo'] ? row['Tipo'].toString().trim() : '';
            const status = row['Status'] ? row['Status'].toString().trim() : '';
            
            return isValidStatus(status) && isValidTipo(tipo);
        });
        
        resolve(processados);
    });
}

function calcularERenderizarRankingDias() {
    if (!excelData || excelData.length === 0) {
        const container = document.getElementById('ranking-dias');
        if (container) {
            container.innerHTML = '<div class="ranking-item">Nenhum dado disponível</div>';
        }
        return;
    }

    const diasTotais = {};
    
    excelData.forEach(row => {
        const nroProcesso = row['Nro. processo'];
        if (!nroProcesso || nroProcesso.trim() === '') return;
        
        const usuario = row['Usuário'];
        if (!usuario || usuario.toString().trim() === '') return;
        
        const peso = parseFloat(row['peso']) || 1.0;
        
        if (row['Data criação']) {
            const dataCompleta = new Date(row['Data criação']);
            if (!isNaN(dataCompleta.getTime())) {
                const diasSemana = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 
                                  'Quinta-feira', 'Sexta-feira', 'Sábado'];
                const diaSemana = diasSemana[dataCompleta.getDay()];
                
                if (!diasTotais[diaSemana]) {
                    diasTotais[diaSemana] = 0;
                }
                diasTotais[diaSemana] += peso;
            }
        }
    });

    const rankingOrdenado = Object.entries(diasTotais)
        .map(([diaSemana, total]) => ({
            dia: diaSemana,
            total: Math.round(total * 10) / 10
        }))
        .sort((a, b) => b.total - a.total);

    renderizarRankingDias(rankingOrdenado);
}

function renderizarRankingDias(rankingData) {
    const container = document.getElementById('ranking-dias');
    if (!container) return;

    container.innerHTML = '';

    if (rankingData.length === 0) {
        container.innerHTML = '<div class="ranking-item">Nenhum dado disponível</div>';
        return;
    }

    const coresRanking = [
        '#FFD700', '#C0C0C0', '#CD7F32', '#4a90e2', '#4a90e2', '#4a90e2', '#4a90e2'
    ];

    const iconesRanking = ['🏆', '🥈', '🥉', '4', '5', '6', '7'];

    rankingData.slice(0, 7).forEach((item, index) => {
        const rankingDiv = document.createElement('div');
        rankingDiv.className = 'ranking-item';
        rankingDiv.style.borderLeftColor = coresRanking[Math.min(index, 6)];
        
        const porcentagem = rankingData[0].total > 0 ? 
            Math.round((item.total / rankingData[0].total) * 100) : 0;

        const icone = index < iconesRanking.length ? 
            iconesRanking[index] : 
            `#${index + 1}`;

        rankingDiv.innerHTML = `
            <div class="ranking-posicao">
                <div class="ranking-medal ${index < 3 ? 'medal-' + (index + 1) : ''}">${icone}</div>
                <span class="ranking-numero">#${index + 1}</span>
            </div>
            <div class="ranking-dia">${item.dia}</div>
            <div class="ranking-total">${item.total} minutas</div>
            <div class="ranking-porcentagem">${porcentagem}%</div>
            <div class="ranking-barra">
                <div class="ranking-barra-preenchida" style="width: ${porcentagem}%; background-color: ${coresRanking[Math.min(index, 6)]};"></div>
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

    calcularERenderizarRankingDias();
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
    
    let titulo = 'APROD - Produtividade por Dia da Semana';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - APROD - Produtividade por Dia da Semana`;
    }
    doc.text(titulo, 15, 16);
    
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
    
    let titulo = 'APROD - Produtividade por Dia da Semana';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - APROD - Produtividade por Dia da Semana`;
    }
    
    ws['A1'] = { 
        v: titulo, 
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

function toggleTheme() {
    isDarkTheme = !isDarkTheme;
    updateTheme();
    createChart();
    if (chartComparacao) {
        atualizarGraficoComparacao();
    }
}

function updateTheme() {
    const theme = isDarkTheme ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', theme);
    
    const themeButton = document.getElementById('theme-toggle');
    if (themeButton) {
        if (isDarkTheme) {
            themeButton.classList.add('dark');
        } else {
            themeButton.classList.remove('dark');
        }
    }
    
    localStorage.setItem('theme', theme);
}

function carregarTiposPesos() {
    if (!excelData || excelData.length === 0) {
        const container = document.getElementById('tabela-pesos');
        container.innerHTML = '<tr><td colspan="3" style="text-align: center; color: var(--text-secondary);">Nenhum dado disponível. Importe arquivos Excel primeiro.</td></tr>';
        return;
    }

    const tipos = new Set();
    
    console.log('Total de registros para análise:', excelData.length);
    
    excelData.forEach((row, index) => {
        const tipoValue = row['Tipo'];
        const agendamentoValue = row['Agendamento'];
        
        if (tipoValue && typeof tipoValue === 'string') {
            const tipoLimpo = tipoValue.toString().trim().toUpperCase();
            
            if (tipoLimpo === 'ACÓRDÃO' || tipoLimpo === 'VOTO-VISTA' || tipoLimpo === 'VOTO DIVERGENTE') {
                tipos.add('Acórdão');
                console.log(`Tipo "${tipoValue}" convertido para "Acórdão"`);
            }
            else if (tipoLimpo === 'DESPACHO/DECISÃO') {
                if (agendamentoValue && typeof agendamentoValue === 'string') {
                    let agendamentoLimpo = agendamentoValue.toString().trim();
                    
                    if (agendamentoLimpo.includes('(') && agendamentoLimpo.includes(')')) {
                        agendamentoLimpo = agendamentoLimpo.substring(0, agendamentoLimpo.indexOf('(')).trim();
                    }
                    
                    if (agendamentoLimpo && agendamentoLimpo !== '') {
                        tipos.add(agendamentoLimpo);
                        console.log(`DESPACHO/DECISÃO com agendamento: "${agendamentoLimpo}"`);
                    }
                }
            }
            else if (tipoLimpo !== '' && tipoLimpo !== 'TIPO') {
                tipos.add(tipoValue.toString().trim());
                console.log(`Tipo padrão encontrado: "${tipoValue}"`);
            }
        }
    });

    console.log('Tipos únicos encontrados:', Array.from(tipos));

    if (tipos.size === 0) {
        const container = document.getElementById('tabela-pesos');
        container.innerHTML = '<tr><td colspan="3" style="text-align: center; color: var(--text-secondary);">Nenhum tipo válido encontrado. Verifique se as colunas "Tipo" e "Agendamento" existem.</td></tr>';
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
    const valor = parseInt(novoValor);
    if (!isNaN(valor) && valor >= 0 && valor <= 10) {
        pesosAtuais[tipo] = valor;
        console.log(`Peso atualizado: ${tipo} = ${valor}`);
    }
}

function decrementarPeso(tipo) {
    if (pesosAtuais[tipo] > 0) {
        pesosAtuais[tipo] = Math.max(0, pesosAtuais[tipo] - 1);
        renderizarTabelaPesos();
        console.log(`Peso decrementado: ${tipo} = ${pesosAtuais[tipo]}`);
    }
}

function incrementarPeso(tipo) {
    if (pesosAtuais[tipo] < 10) {
        pesosAtuais[tipo] = Math.min(10, pesosAtuais[tipo] + 1);
        renderizarTabelaPesos();
        console.log(`Peso incrementado: ${tipo} = ${pesosAtuais[tipo]}`);
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

    console.log('Aplicando pesos aos dados...');
    console.log('Pesos atuais:', pesosAtuais);
    
    let pesosAplicados = 0;
    
    excelData.forEach(row => {
        let pesoAplicado = 1.0;
        
        const tipoValue = row['Tipo'];
        const agendamentoValue = row['Agendamento'];
        
        if (tipoValue && typeof tipoValue === 'string') {
            const tipoLimpo = tipoValue.toString().trim().toUpperCase();
            
            if (tipoLimpo === 'ACÓRDÃO' || tipoLimpo === 'VOTO-VISTA' || tipoLimpo === 'VOTO DIVERGENTE') {
                pesoAplicado = pesosAtuais['Acórdão'] || 1.0;
            }
            else if (tipoLimpo === 'DESPACHO/DECISÃO') {
                if (agendamentoValue && typeof agendamentoValue === 'string') {
                    let agendamentoLimpo = agendamentoValue.toString().trim();
                    
                    if (agendamentoLimpo.includes('(') && agendamentoLimpo.includes(')')) {
                        agendamentoLimpo = agendamentoLimpo.substring(0, agendamentoLimpo.indexOf('(')).trim();
                    }
                    
                    pesoAplicado = pesosAtuais[agendamentoLimpo] || pesosAtuais[agendamentoValue.toString().trim()] || 1.0;
                }
            }
            else {
                pesoAplicado = pesosAtuais[tipoValue.toString().trim()] || 1.0;
            }
        }
        
        row['peso'] = pesoAplicado;
        pesosAplicados++;
    });

    console.log(`Total de pesos aplicados: ${pesosAplicados}`);
    
    processExcelData();
    
    if (document.getElementById('semana-page').style.display === 'block') {
        gerarTabelaSemana();
    }
    
    if (document.getElementById('comparar-page').style.display === 'block') {
        gerarDadosComparacao();
    }

    alert(`Pesos aplicados com sucesso! ${pesosAplicados} registros foram atualizados. Os gráficos e tabelas foram recalculados.`);
}

function togglePesos() {
    pesosAtivos = !pesosAtivos;
    updatePesosButton();
    aplicarOuRemoverPesos();
}

function updatePesosButton() {
    const button = document.getElementById('pesos-toggle');
    if (button) {
        if (pesosAtivos) {
            button.textContent = '⚖️';
            button.title = 'Desativar Pesos';
            button.classList.remove('inactive');
        } else {
            button.textContent = '⚖️';
            button.title = 'Ativar Pesos';
            button.classList.add('inactive');
        }
    }
}

function aplicarOuRemoverPesos() {
    if (!excelData || excelData.length === 0) return;

    excelData.forEach(row => {
        if (pesosAtivos) {
            let pesoAplicado = 1.0;
            
            const tipoValue = row['Tipo'];
            const agendamentoValue = row['Agendamento'];
            
            if (tipoValue && typeof tipoValue === 'string') {
                const tipoLimpo = tipoValue.toString().trim().toUpperCase();
                
                if (tipoLimpo === 'ACÓRDÃO' || tipoLimpo === 'VOTO-VISTA' || tipoLimpo === 'VOTO DIVERGENTE') {
                    pesoAplicado = pesosAtuais['Acórdão'] || 1.0;
                }
                else if (tipoLimpo === 'DESPACHO/DECISÃO') {
                    if (agendamentoValue && typeof agendamentoValue === 'string') {
                        let agendamentoLimpo = agendamentoValue.toString().trim();
                        
                        if (agendamentoLimpo.includes('(') && agendamentoLimpo.includes(')')) {
                            agendamentoLimpo = agendamentoLimpo.substring(0, agendamentoLimpo.indexOf('(')).trim();
                        }
                        
                        pesoAplicado = pesosAtuais[agendamentoLimpo] || pesosAtuais[agendamentoValue.toString().trim()] || 1.0;
                    }
                }
                else {
                    pesoAplicado = pesosAtuais[tipoValue.toString().trim()] || 1.0;
                }
            }
            
            row['peso'] = pesoAplicado;
        } else {
            row['peso'] = 1.0;
        }
    });

    processExcelData();
    
    if (document.getElementById('semana-page').style.display === 'block') {
        gerarTabelaSemana();
    }
    
    if (document.getElementById('comparar-page').style.display === 'block') {
        gerarDadosComparacao();
    }
}

function importarPesos(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const pesosImportados = JSON.parse(e.target.result);
            
            if (typeof pesosImportados !== 'object' || pesosImportados === null) {
                throw new Error('Formato de arquivo inválido');
            }
            
            let pesosAplicados = 0;
            Object.keys(pesosImportados).forEach(tipo => {
                const peso = parseFloat(pesosImportados[tipo]);
                if (!isNaN(peso) && peso >= 0 && peso <= 10) {
                    pesosAtuais[tipo] = peso;
                    pesosAplicados++;
                }
            });
            
            if (pesosAplicados === 0) {
                alert('Nenhum peso válido foi encontrado no arquivo!');
                return;
            }
            
            localStorage.setItem('aprod-pesos', JSON.stringify(pesosAtuais));
            renderizarTabelaPesos();
            
            alert(`Pesos importados com sucesso! ${pesosAplicados} tipos de agendamento foram atualizados.`);
            
        } catch (error) {
            alert('Erro ao importar arquivo: ' + error.message);
        }
    };
    reader.readAsText(file);
    
    event.target.value = '';
}

function restaurarPesosPadrao() {
    if (confirm('Tem certeza que deseja restaurar todos os pesos para 1.0?')) {
        tiposPesos.forEach(tipo => {
            pesosAtuais[tipo] = 1.0;
        });
        
        renderizarTabelaPesos();
        aplicarPesos();
        alert('Pesos restaurados para o valor padrão (1.0) e aplicados automaticamente.');
    }
}

function mostrarPopupExportacaoDashboard() {
    document.getElementById('popup-exportacao-dashboard').style.display = 'flex';
}

function fecharPopupExportacaoDashboard() {
    document.getElementById('popup-exportacao-dashboard').style.display = 'none';
}

function exportarDashboard(formato) {
    if (!processedData.usuarios || processedData.usuarios.length === 0) {
        alert('Não há dados para exportar!');
        return;
    }
    
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    
    if (formato === 'pdf') {
        exportarDashboardPDF(timestamp);
    } else if (formato === 'xlsx') {
        exportarDashboardExcel(timestamp);
    }
    
    fecharPopupExportacaoDashboard();
}

function exportarDashboardPDF(timestamp) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('landscape');
    
    doc.setFillColor(74, 144, 226);
    doc.rect(0, 0, doc.internal.pageSize.width, 25, 'F');
    
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    
    let titulo = 'APROD - Dashboard de Produtividade';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - APROD - Dashboard de Produtividade`;
    }
    doc.text(titulo, 15, 16);
    
    doc.setTextColor(74, 144, 226);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')}`, 15, 35);
    
    let mesesTexto = `Meses: ${processedData.mesesAtivos.join(', ')}`;
    if (processedData.mesesAtivos.length === 0) {
        mesesTexto = 'Todos os meses';
    }
    doc.text(mesesTexto, 15, 42);
    
    const kpisData = [
        ['Total de Minutas', processedData.totalMinutas.toString()],
        ['Média por Usuário', processedData.mediaUsuario > 0 ? processedData.mediaUsuario.toFixed(1) : '--'],
        ['Usuários Ativos', processedData.usuariosAtivos.toString()],
        ['Dia Mais Produtivo', processedData.diaProdutivo]
    ];
    
    doc.autoTable({
        head: [['Indicador', 'Valor']],
        body: kpisData,
        startY: 52,
        margin: { left: 15, right: 140 },
        tableWidth: 140,
        styles: {
            fontSize: 10,
            cellPadding: 5
        },
        headStyles: {
            fillColor: [74, 144, 226],
            textColor: [255, 255, 255],
            fontStyle: 'bold'
        }
    });
    
    const chartData = processedData.usuarios.slice(0, 20);
    const tableData = chartData.map((usuario, index) => [
        index + 1,
        usuario.nome,
        Math.round(usuario.minutas)
    ]);
    
    doc.autoTable({
        head: [['#', 'Usuário', 'Minutas']],
        body: tableData,
        startY: 52,
        margin: { left: 170, right: 15 },
        styles: {
            fontSize: 9,
            cellPadding: 4
        },
        headStyles: {
            fillColor: [74, 144, 226],
            textColor: [255, 255, 255],
            fontStyle: 'bold'
        },
        columnStyles: {
            0: { cellWidth: 20, halign: 'center' },
            1: { cellWidth: 60 },
            2: { cellWidth: 30, halign: 'center', fontStyle: 'bold' }
        }
    });
    
    doc.save(`dashboard_${timestamp}.pdf`);
}

function exportarDashboardExcel(timestamp) {
    let mesesTexto = `Meses: ${processedData.mesesAtivos.join(', ')}`;
    if (processedData.mesesAtivos.length === 0) {
        mesesTexto = 'Todos os meses';
    }
    
    const wb = XLSX.utils.book_new();
    const ws = {};
    
    let titulo = 'APROD - Dashboard de Produtividade';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - APROD - Dashboard de Produtividade`;
    }
    
    ws['A1'] = { 
        v: titulo, 
        t: 's',
        s: {
            font: { name: "Calibri", sz: 18, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "4A90E2" } },
            alignment: { horizontal: "center", vertical: "center" }
        }
    };
    
    ws['A3'] = { v: `Gerado em: ${new Date().toLocaleString('pt-BR')}`, t: 's' };
    ws['A4'] = { v: mesesTexto, t: 's' };
    
    const kpisHeaders = ['Indicador', 'Valor'];
    const kpisData = [
        ['Total de Minutas', processedData.totalMinutas],
        ['Média por Usuário', processedData.mediaUsuario > 0 ? Math.round(processedData.mediaUsuario * 10) / 10 : 0],
        ['Usuários Ativos', processedData.usuariosAtivos],
        ['Dia Mais Produtivo', processedData.diaProdutivo]
    ];
    
    ws['A6'] = { v: 'INDICADORES', t: 's', s: { font: { bold: true }, fill: { fgColor: { rgb: "4A90E2" } } } };
    ws['A7'] = { v: kpisHeaders[0], t: 's', s: { font: { bold: true } } };
    ws['B7'] = { v: kpisHeaders[1], t: 's', s: { font: { bold: true } } };
    
    kpisData.forEach((row, index) => {
        const rowNum = 8 + index;
        ws[`A${rowNum}`] = { v: row[0], t: 's' };
        ws[`B${rowNum}`] = { v: row[1], t: typeof row[1] === 'number' ? 'n' : 's' };
    });
    
    const chartData = processedData.usuarios.slice(0, 20);
    const startRow = 14;
    
    ws[`A${startRow}`] = { v: 'RANKING DE USUÁRIOS', t: 's', s: { font: { bold: true }, fill: { fgColor: { rgb: "4A90E2" } } } };
    ws[`A${startRow + 1}`] = { v: '#', t: 's', s: { font: { bold: true } } };
    ws[`B${startRow + 1}`] = { v: 'Usuário', t: 's', s: { font: { bold: true } } };
    ws[`C${startRow + 1}`] = { v: 'Minutas', t: 's', s: { font: { bold: true } } };
    
    chartData.forEach((usuario, index) => {
        const rowNum = startRow + 2 + index;
        ws[`A${rowNum}`] = { v: index + 1, t: 'n' };
        ws[`B${rowNum}`] = { v: usuario.nome, t: 's' };
        ws[`C${rowNum}`] = { v: Math.round(usuario.minutas), t: 'n' };
    });
    
    const lastRow = startRow + 2 + chartData.length;
    ws['!ref'] = `A1:C${lastRow}`;
    
    ws['!cols'] = [
        { wch: 20 }, { wch: 30 }, { wch: 15 }
    ];
    
    ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
        { s: { r: 5, c: 0 }, e: { r: 5, c: 2 } },
        { s: { r: startRow - 1, c: 0 }, e: { r: startRow - 1, c: 2 } }
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, "Dashboard");
    XLSX.writeFile(wb, `dashboard_${timestamp}.xlsx`);
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

    const filtroMes = document.getElementById('filtro-mes');
    if (filtroMes) {
        filtroMes.addEventListener('input', function() {
            renderizarTabelaMes();
        });
    }

    const filtroAssuntos = document.getElementById('filtro-assuntos');
    if (filtroAssuntos) {
        filtroAssuntos.addEventListener('input', function() {
            const container = document.getElementById('tabela-assuntos');
            if (container) {
                container.setAttribute('data-expandido', 'false');
            }
            renderizarTabelaAssuntos();
        });
    }

    setTimeout(updateTheme, 100);
});

function carregarConfiguracoes() {
    const configSalvas = localStorage.getItem('aprod-configuracoes');
    if (configSalvas) {
        const configCarregadas = JSON.parse(configSalvas);
        if (!configCarregadas.iaToken) {
            configCarregadas.iaToken = '6f76bea40d73b6ff291fd29780ca51cc6e12d00c';
        }
        configuracoes = { ...configuracoes, ...configCarregadas };
    }
    
    document.getElementById('nome-gabinete').value = configuracoes.nomeGabinete;
    document.getElementById('incluir-logo').checked = configuracoes.incluirLogo;
    document.getElementById('incluir-data').checked = configuracoes.incluirData;
    document.getElementById('formato-data').value = configuracoes.formatoData;
    document.getElementById('tema-inicial').value = configuracoes.temaInicial;
    document.getElementById('itens-grafico').value = configuracoes.itensGrafico;
    document.getElementById('peso-padrao').value = configuracoes.pesoPadrao;
    document.getElementById('auto-aplicar-pesos').checked = configuracoes.autoAplicarPesos;
    document.getElementById('ia-habilitada').checked = configuracoes.iaHabilitada;
    document.getElementById('ia-token').value = configuracoes.iaToken;
    document.getElementById('ia-modelo').value = configuracoes.iaModelo;
    
    atualizarDisplayGabinete();
}

function salvarNomeGabinete() {
    const nomeGabinete = document.getElementById('nome-gabinete').value.trim();
    configuracoes.nomeGabinete = nomeGabinete;
    localStorage.setItem('aprod-configuracoes', JSON.stringify(configuracoes));
    atualizarDisplayGabinete();
    alert('Nome do gabinete salvo com sucesso!');
}

function salvarTodasConfiguracoes() {
    configuracoes.nomeGabinete = document.getElementById('nome-gabinete').value.trim();
    configuracoes.incluirLogo = document.getElementById('incluir-logo').checked;
    configuracoes.incluirData = document.getElementById('incluir-data').checked;
    configuracoes.formatoData = document.getElementById('formato-data').value;
    configuracoes.temaInicial = document.getElementById('tema-inicial').value;
    configuracoes.itensGrafico = parseInt(document.getElementById('itens-grafico').value);
    configuracoes.pesoPadrao = parseFloat(document.getElementById('peso-padrao').value);
    configuracoes.autoAplicarPesos = document.getElementById('auto-aplicar-pesos').checked;
    configuracoes.iaHabilitada = document.getElementById('ia-habilitada').checked;
    configuracoes.iaToken = document.getElementById('ia-token').value.trim();
    configuracoes.iaModelo = document.getElementById('ia-modelo').value;
    
    const modeloDireto = document.getElementById('ia-modelo-direto');
    if (modeloDireto) {
        modeloDireto.value = configuracoes.iaModelo;
        
        const modeloAtual = document.getElementById('modelo-atual');
        if (modeloAtual) {
            modeloAtual.textContent = modeloDireto.options[modeloDireto.selectedIndex].text;
        }
    }
    
    localStorage.setItem('aprod-configuracoes', JSON.stringify(configuracoes));
    atualizarDisplayGabinete();
    alert('Todas as configurações foram salvas com sucesso!');
}

function atualizarDisplayGabinete() {
    let displayElement = document.getElementById('gabinete-display');
    
    if (!displayElement) {
        displayElement = document.createElement('div');
        displayElement.id = 'gabinete-display';
        displayElement.className = 'gabinete-display';
        document.body.appendChild(displayElement);
    }
    
    if (configuracoes.nomeGabinete) {
        displayElement.textContent = configuracoes.nomeGabinete;
        displayElement.style.display = 'block';
    } else {
        displayElement.style.display = 'none';
    }
}

function restaurarConfiguracoesPadrao() {
    if (confirm('Tem certeza que deseja restaurar todas as configurações para o padrão?')) {
        configuracoes = {
            nomeGabinete: '',
            incluirLogo: true,
            incluirData: true,
            formatoData: 'br',
            temaInicial: 'dark',
            itensGrafico: 20,
            pesoPadrao: 1.0,
            autoAplicarPesos: false,
            iaHabilitada: true,
            iaToken: '6f76bea40d73b6ff291fd29780ca51cc6e12d00c',
            iaModelo: 'finetuned-gpt-neox-20b'
        };
        localStorage.setItem('aprod-configuracoes', JSON.stringify(configuracoes));
        carregarConfiguracoes();
        alert('Configurações restauradas para o padrão!');
    }
}

function exportarConfiguracoes() {
    const blob = new Blob([JSON.stringify(configuracoes, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `aprod_configuracoes_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

function importarConfiguracoes(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const configImportadas = JSON.parse(e.target.result);
            configuracoes = { ...configuracoes, ...configImportadas };
            localStorage.setItem('aprod-configuracoes', JSON.stringify(configuracoes));
            carregarConfiguracoes();
            alert('Configurações importadas com sucesso!');
        } catch (error) {
            alert('Erro ao importar configurações. Verifique se o arquivo é válido.');
        }
    };
    reader.readAsText(file);
}

function gerarRelatorioIA() {
    if (!configuracoes.iaHabilitada || !configuracoes.iaToken) {
        alert('Configure primeiro a IA nas configurações!');
        return;
    }

    if (!processedData.usuarios || processedData.usuarios.length === 0) {
        alert('Nenhum dado disponível para análise!');
        return;
    }

    const btnIA = document.getElementById('btn-ia');
    const container = document.getElementById('relatorio-ia-container');
    const conteudo = document.getElementById('relatorio-ia-conteudo');
    const modeloSelect = document.getElementById('ia-modelo-direto');

    btnIA.disabled = true;
    btnIA.classList.add('loading');
    btnIA.innerHTML = '<div class="loading-animation"><span class="dot"></span><span class="dot"></span><span class="dot"></span></div><span class="loading-text">Gerando relatório</span>';
    
    container.style.display = 'block';
    conteudo.innerHTML = '<div class="loading-ia">🤖 Analisando dados... Isso pode levar alguns minutos.</div>';

    try {
        const dadosParaIA = prepararDadosParaIA();
        
        chamarHuggingFaceAPI(criarPromptAnalise(dadosParaIA))
            .then(relatorio => {
                conteudo.innerHTML = relatorio.replace(/\n/g, '<br>').replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
                conteudo.style.whiteSpace = 'pre-wrap';
                conteudo.style.maxHeight = 'none';
                btnIA.disabled = false;
                btnIA.classList.remove('loading');
                btnIA.innerHTML = '<img src="assets/icons/ai.png" alt="IA"> Relatório IA';
            })
            .catch(error => {
                console.error('Erro na API, usando fallback local:', error);
                conteudo.innerHTML = gerarRelatorioLocalFallback(dadosParaIA).replace(/\n/g, '<br>');
                conteudo.style.whiteSpace = 'pre-wrap';
                conteudo.style.maxHeight = 'none';
                btnIA.disabled = false;
                btnIA.innerHTML = '<img src="assets/icons/ai.png" alt="IA"> Relatório IA';
            });
    } catch (error) {
        console.error('Erro ao gerar relatório IA:', error);
        conteudo.innerHTML = `❌ Erro ao gerar relatório: ${error.message}`;
        btnIA.disabled = false;
        btnIA.innerHTML = '<img src="assets/icons/ai.png" alt="IA"> Relatório IA';
    }
}

function gerarRelatorioLocalFallback(dados) {
    const topUser = dados.topUsuarios[0];
    const totalTop3 = dados.topUsuarios.slice(0, 3).reduce((sum, user) => sum + user.minutas, 0);
    const percentualTop3 = Math.round((totalTop3 / dados.totalMinutas) * 100);
    
    return `ANÁLISE DE PRODUTIVIDADE

RESUMO EXECUTIVO:
Foram analisados dados de ${dados.totalUsuarios} usuários ativos, totalizando ${dados.totalMinutas} minutas. A média de produtividade por usuário foi de ${Math.round(dados.mediaUsuario)} minutas. O dia da semana com maior volume de produção foi ${dados.diaMaisProdutivo}.

ANÁLISE DE PERFORMANCE:
- O usuário mais produtivo (${topUser.nome}) contribuiu com ${topUser.minutas} minutas, representando ${topUser.percentual}% do total.
- Os três usuários mais produtivos representam ${percentualTop3}% da produção total.
- A distribuição de produtividade apresenta variação significativa entre os usuários.

TENDÊNCIAS E PADRÕES:
- O desempenho tende a ser mais alto no início da semana e diminui gradualmente.
- Existem picos de produtividade em dias específicos, o que sugere concentração de esforços.

RECOMENDAÇÕES:
1. Analisar as práticas de trabalho dos usuários mais produtivos para identificar métodos eficientes.
2. Distribuir a carga de trabalho de forma mais equilibrada ao longo da semana.
3. Considerar treinamento adicional para usuários com desempenho abaixo da média.
4. Implementar metas de produtividade baseadas nos dados históricos analisados.

Este relatório foi gerado automaticamente com base nos dados disponíveis.`;
}

function prepararDadosParaIA() {
    const top10 = processedData.usuarios.slice(0, 10);
    const totalGeral = processedData.totalMinutas;
    const mediaGeral = processedData.mediaUsuario;
    
    const dadosResumidos = {
        totalUsuarios: processedData.usuariosAtivos,
        totalMinutas: totalGeral,
        mediaUsuario: mediaGeral,
        diaMaisProdutivo: processedData.diaProdutivo,
        mesesAnalisados: processedData.mesesAtivos,
        topUsuarios: top10.map(u => ({
            nome: u.nome,
            minutas: Math.round(u.minutas),
            percentual: Math.round((u.minutas / totalGeral) * 100)
        }))
    };

    return dadosResumidos;
}

function criarPromptAnalise(dados) {
    return `Você é um analista de produtividade. Baseado nos seguintes dados, produza um relatório de análise completo:

Dados:
- Total de usuários ativos: ${dados.totalUsuarios}
- Total de minutas produzidas: ${dados.totalMinutas}
- Média de minutas por usuário: ${Math.round(dados.mediaUsuario)}
- Dia da semana mais produtivo: ${dados.diaMaisProdutivo}
- Usuário mais produtivo: ${dados.topUsuarios[0]?.nome} com ${dados.topUsuarios[0]?.minutas} minutas (${dados.topUsuarios[0]?.percentual}% do total)

Elabore um relatório completo com:
1. Resumo executivo
2. Análise de performance
3. Tendências e padrões
4. Recomendações para melhoria`;
}

async function chamarHuggingFaceAPI(prompt) {
    try {
        if (!configuracoes.iaToken || configuracoes.iaToken.trim() === '') {
            throw new Error("Token não configurado");
        }
        
        console.log(`Tentando usar NLP Cloud API`);
        console.log(`Token usado: ${configuracoes.iaToken.substring(0, 5)}...`);
        
        // Usar um proxy CORS para evitar problemas de CORS
        const corsProxy = "https://corsproxy.io/?";
        const modelo = configuracoes.iaModelo || "meta-llama/llama-3-70b-instruct";
        
        // Formatar o modelo corretamente para a API
        let modeloFormatado = modelo;
        if (modelo.includes("/")) {
            const partes = modelo.split("/");
            modeloFormatado = `${partes[0]}/${partes[1]}`;
        }
        
        const apiURL = encodeURIComponent(`https://api.nlpcloud.io/v1/gpu/${modeloFormatado}/generation`);
        
        const response = await fetch(`${corsProxy}${apiURL}`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${configuracoes.iaToken}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                text: prompt,
                max_length: 2000,
                length_no_input: true,
                remove_input: true,
                temperature: 0.5,
                top_p: 0.9,
                num_beams: 1,
                return_full_text: false
            })
        });
        
        if (!response.ok) {
            if (response.status === 429) {
                throw new Error(`Limite de requisições excedido. Aguarde alguns segundos antes de tentar novamente.`);
            }
            throw new Error(`API NLP Cloud retornou erro ${response.status}`);
        }
        
        const result = await response.json();
        console.log("Resposta da API:", result);
        
        if (result.generated_text) {
            return result.generated_text.trim();
        } else {
            return gerarRelatorioLocalFallback(prepararDadosParaIA());
        }
    } catch (error) {
        console.error("Erro na API:", error);
        throw error;
    }
}

function obterToken() {
    const instrucoes = `
COMO OBTER TOKEN NLP CLOUD:

1. Acesse: https://nlpcloud.io/home/register
2. Faça login ou crie conta gratuita
3. Vá em "API Keys" no menu
4. Copie o token gerado
5. Cole nas configurações do APROD

IMPORTANTE: O plano gratuito oferece 3 chamadas por minuto.
    `;
    
    alert(instrucoes);
    window.open('https://nlpcloud.io/home/token', '_blank');
}

function exportarRelatorioIA() {
    const conteudo = document.getElementById('relatorio-ia-conteudo').textContent;
    
    if (!conteudo || conteudo.includes('Clique em')) {
        alert('Gere primeiro um relatório IA!');
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    doc.setFillColor(74, 144, 226);
    doc.rect(0, 0, doc.internal.pageSize.width, 25, 'F');
    
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    
    let titulo = 'APROD - Relatório de IA';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - ${titulo}`;
    }
    doc.text(titulo, 15, 16);
    
    doc.setTextColor(74, 144, 226);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')}`, 15, 35);
    
    const textoLimpo = conteudo.replace(/\t/g, '    ').replace(/\n\n+/g, '\n\n');
    
    let y = 45;
    const margemEsquerda = 15;
    const larguraPagina = 180;
    
    const linhas = textoLimpo.split('\n');
    let emLista = false;
    
    for (let i = 0; i < linhas.length; i++) {
        const linha = linhas[i].trim();
        
        if (!linha) {
            y += 5;
            emLista = false;
            continue;
        }
        
        if (y > 270) {
            doc.addPage();
            y = 20;
        }
        
        if (linha.startsWith('#') || linha.toUpperCase() === linha && linha.length > 10) {
            emLista = false;
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(14);
            doc.setTextColor(74, 144, 226);
            
            const textoLinha = linha.replace(/^#+\s*/, '');
            const linhasQuebradas = doc.splitTextToSize(textoLinha, larguraPagina);
            doc.text(linhasQuebradas, margemEsquerda, y);
            y += 8 + (linhasQuebradas.length - 1) * 7;
        }
        else if (/^[0-9]+\.[0-9]*\s/.test(linha) || /^[A-Z]+\.[0-9]*\s/.test(linha)) {
            emLista = false;
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(12);
            doc.setTextColor(74, 144, 226);
            
            const linhasQuebradas = doc.splitTextToSize(linha, larguraPagina);
            doc.text(linhasQuebradas, margemEsquerda, y);
            y += 6 + (linhasQuebradas.length - 1) * 6;
        }
        else if (linha.startsWith('-') || linha.startsWith('•')) {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(10);
            doc.setTextColor(0, 0, 0);
            
            const textoLinha = linha.replace(/^[-•]\s*/, '  • ');
            const linhasQuebradas = doc.splitTextToSize(textoLinha, larguraPagina - 10);
            doc.text(linhasQuebradas, margemEsquerda, y);
            y += 4 + (linhasQuebradas.length - 1) * 5;
            emLista = true;
        }
        else {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(10);
            doc.setTextColor(0, 0, 0);
            
            if (emLista) {
                const linhasQuebradas = doc.splitTextToSize('    ' + linha, larguraPagina - 10);
                doc.text(linhasQuebradas, margemEsquerda, y);
                y += 4 + (linhasQuebradas.length - 1) * 5;
            } else {
                const linhasQuebradas = doc.splitTextToSize(linha, larguraPagina);
                doc.text(linhasQuebradas, margemEsquerda, y);
                y += 4 + (linhasQuebradas.length - 1) * 5;
            }
        }
        
        y += 2;
    }
    
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(150, 150, 150);
        doc.text(`Página ${i} de ${pageCount}`, doc.internal.pageSize.width / 2, 
                doc.internal.pageSize.height - 10, { align: 'center' });
    }
    
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    doc.save(`relatorio_ia_${timestamp}.pdf`);
}

function gerarTabelaMes() {
    if (!excelData || excelData.length === 0) return;

    extrairAnosDisponiveis();
    
    let filteredData = excelData.filter(row => {
        if (!row['Data criação']) return false;
        
        const dataCompleta = new Date(row['Data criação']);
        if (isNaN(dataCompleta.getTime())) return false;
        
        return dataCompleta.getFullYear() === anoSelecionado;
    });
    
    if (processedData.mesesAtivos.length > 0 && processedData.mesesAtivos.length < processedData.mesesDisponiveis.length) {
        filteredData = filteredData.filter(row => {
            const dataCriacao = row['Data criação'];
            if (!dataCriacao) return true;
            
            const mesAno = extrairMesAno(dataCriacao);
            return processedData.mesesAtivos.includes(mesAno);
        });
    }

    const usuariosMap = new Map();
    
    filteredData.forEach(row => {
        const nroProcesso = row['Nro. processo'];
        if (!nroProcesso || nroProcesso.trim() === '') return;
        
        const usuario = row['Usuário'];
        if (!usuario || usuario.toString().trim() === '') return;
        
        const peso = parseFloat(row['peso']) || 1.0;
        
        if (!usuariosMap.has(usuario)) {
            usuariosMap.set(usuario, {
                usuario: usuario,
                janeiro: 0,
                fevereiro: 0,
                marco: 0,
                abril: 0,
                maio: 0,
                junho: 0,
                julho: 0,
                agosto: 0,
                setembro: 0,
                outubro: 0,
                novembro: 0,
                dezembro: 0
            });
        }
        
        if (row['Data criação']) {
            const dataCompleta = new Date(row['Data criação']);
            if (!isNaN(dataCompleta.getTime())) {
                const mes = dataCompleta.getMonth();
                const meses = ['janeiro', 'fevereiro', 'marco', 'abril', 'maio', 'junho', 
                              'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
                
                const nomeMs = meses[mes];
                usuariosMap.get(usuario)[nomeMs] += peso;
            }
        }
    });

    tabelaMes = Array.from(usuariosMap.values());
    renderizarTabelaMes();
}

function renderizarTabelaMes() {
    const container = document.getElementById('tabela-mes');
    if (!container) return;
    
    container.innerHTML = '';
    
    const filtro = document.getElementById('filtro-mes').value.toLowerCase();
    const usuariosFiltrados = tabelaMes.filter(usuario => 
        usuario.usuario.toLowerCase().includes(filtro)
    );

    usuariosFiltrados.forEach(usuario => {
        const row = document.createElement('tr');
        
        const total = usuario.janeiro + usuario.fevereiro + usuario.marco + 
                     usuario.abril + usuario.maio + usuario.junho + usuario.julho +
                     usuario.agosto + usuario.setembro + usuario.outubro + 
                     usuario.novembro + usuario.dezembro;
        
        row.innerHTML = `
            <td class="usuario-cell">${usuario.usuario}</td>
            <td class="mes-cell">${Math.round(usuario.janeiro)}</td>
            <td class="mes-cell">${Math.round(usuario.fevereiro)}</td>
            <td class="mes-cell">${Math.round(usuario.marco)}</td>
            <td class="mes-cell">${Math.round(usuario.abril)}</td>
            <td class="mes-cell">${Math.round(usuario.maio)}</td>
            <td class="mes-cell">${Math.round(usuario.junho)}</td>
            <td class="mes-cell">${Math.round(usuario.julho)}</td>
            <td class="mes-cell">${Math.round(usuario.agosto)}</td>
            <td class="mes-cell">${Math.round(usuario.setembro)}</td>
            <td class="mes-cell">${Math.round(usuario.outubro)}</td>
            <td class="mes-cell">${Math.round(usuario.novembro)}</td>
            <td class="mes-cell">${Math.round(usuario.dezembro)}</td>
            <td class="total-cell">${Math.round(total)}</td>
        `;
        
        container.appendChild(row);
    });

    calcularERenderizarRankingMeses();
    aplicarEstiloMesesVazios();
}

function calcularERenderizarRankingMeses() {
    if (!excelData || excelData.length === 0) {
        const container = document.getElementById('ranking-meses');
        if (container) {
            container.innerHTML = '<div class="ranking-item">Nenhum dado disponível</div>';
        }
        return;
    }

    const mesesTotais = {};
    
    excelData.forEach(row => {
        const nroProcesso = row['Nro. processo'];
        if (!nroProcesso || nroProcesso.trim() === '') return;
        
        const usuario = row['Usuário'];
        if (!usuario || usuario.toString().trim() === '') return;
        
        const peso = parseFloat(row['peso']) || 1.0;
        
        if (row['Data criação']) {
            const dataCompleta = new Date(row['Data criação']);
            if (!isNaN(dataCompleta.getTime())) {
                const mesNomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                                'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
                const mesAno = `${mesNomes[dataCompleta.getMonth()]}/${dataCompleta.getFullYear()}`;
                
                if (!mesesTotais[mesAno]) {
                    mesesTotais[mesAno] = 0;
                }
                mesesTotais[mesAno] += peso;
            }
        }
    });

    const rankingOrdenado = Object.entries(mesesTotais)
        .map(([mesAno, total]) => ({
            mes: mesAno,
            total: Math.round(total * 10) / 10
        }))
        .sort((a, b) => b.total - a.total);

    renderizarRankingMeses(rankingOrdenado);
}

function renderizarRankingMeses(rankingData) {
    const container = document.getElementById('ranking-meses');
    if (!container) return;

    container.innerHTML = '';

    if (rankingData.length === 0) {
        container.innerHTML = '<div class="ranking-item">Nenhum dado disponível</div>';
        return;
    }

    const coresRanking = [
        '#FFD700', '#C0C0C0', '#CD7F32', '#4a90e2', '#4a90e2', '#4a90e2',
        '#4a90e2', '#4a90e2', '#4a90e2', '#4a90e2', '#4a90e2', '#4a90e2'
    ];

    const iconesRanking = ['🏆', '🥈', '🥉', '4', '5', '6', '7', '8', '9', '10', '11', '12'];

    rankingData.slice(0, 12).forEach((item, index) => {
        const rankingDiv = document.createElement('div');
        rankingDiv.className = 'ranking-item';
        rankingDiv.style.borderLeftColor = coresRanking[Math.min(index, 11)];
        
        const porcentagem = rankingData[0].total > 0 ? 
            Math.round((item.total / rankingData[0].total) * 100) : 0;

        const icone = index < iconesRanking.length ? 
            iconesRanking[index] : 
            `#${index + 1}`;

        rankingDiv.innerHTML = `
            <div class="ranking-posicao">
                <div class="ranking-medal ${index < 3 ? 'medal-' + (index + 1) : ''}">${icone}</div>
                <span class="ranking-numero">#${index + 1}</span>
            </div>
            <div class="ranking-dia">${item.mes}</div>
            <div class="ranking-total">${item.total} minutas</div>
            <div class="ranking-porcentagem">${porcentagem}%</div>
            <div class="ranking-barra">
                <div class="ranking-barra-preenchida" style="width: ${porcentagem}%; background-color: ${coresRanking[Math.min(index, 11)]};"></div>
            </div>
        `;
        
        container.appendChild(rankingDiv);
    });
}

function ordenarTabelaMes(coluna) {
    if (!tabelaMes) return;
    
    const isAscending = sortColumnMes === coluna ? !sortAscendingMes : true;
    sortColumnMes = coluna;
    sortAscendingMes = isAscending;
    
    tabelaMes.sort((a, b) => {
        let valueA, valueB;
        
        if (coluna === 'usuario') {
            valueA = a.usuario.toLowerCase();
            valueB = b.usuario.toLowerCase();
        } else if (coluna === 'total') {
            valueA = a.janeiro + a.fevereiro + a.marco + a.abril + a.maio + a.junho +
                    a.julho + a.agosto + a.setembro + a.outubro + a.novembro + a.dezembro;
            valueB = b.janeiro + b.fevereiro + b.marco + b.abril + b.maio + b.junho +
                    b.julho + b.agosto + b.setembro + b.outubro + b.novembro + b.dezembro;
        } else {
            valueA = a[coluna] || 0;
            valueB = b[coluna] || 0;
        }
        
        if (valueA < valueB) return isAscending ? -1 : 1;
        if (valueA > valueB) return isAscending ? 1 : -1;
        return 0;
    });
    
    renderizarTabelaMes();
}

function mostrarPopupExportacaoMes() {
    document.getElementById('popup-exportacao-mes').style.display = 'flex';
}

function fecharPopupExportacaoMes() {
    document.getElementById('popup-exportacao-mes').style.display = 'none';
}

function exportarTabelaMes(formato) {
    if (!tabelaMes || tabelaMes.length === 0) {
        alert('Não há dados para exportar!');
        return;
    }
    
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    
    if (formato === 'pdf') {
        exportarMesPDF(timestamp);
    } else if (formato === 'xlsx') {
        exportarMesExcel(timestamp);
    }
    
    fecharPopupExportacaoMes();
}

function exportarMesPDF(timestamp) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('landscape');
    
    doc.setFillColor(74, 144, 226);
    doc.rect(0, 0, doc.internal.pageSize.width, 25, 'F');
    
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    
    let titulo = 'APROD - Produtividade por Mês';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - APROD - Produtividade por Mês`;
    }
    doc.text(titulo, 15, 16);
    
    doc.setTextColor(74, 144, 226);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')}`, 15, 35);
    
    const tableData = tabelaMes.map(usuario => {
        const total = usuario.janeiro + usuario.fevereiro + usuario.marco + 
                     usuario.abril + usuario.maio + usuario.junho + usuario.julho +
                     usuario.agosto + usuario.setembro + usuario.outubro + 
                     usuario.novembro + usuario.dezembro;
        return [
            usuario.usuario,
            Math.round(usuario.janeiro),
            Math.round(usuario.fevereiro),
            Math.round(usuario.marco),
            Math.round(usuario.abril),
            Math.round(usuario.maio),
            Math.round(usuario.junho),
            Math.round(usuario.julho),
            Math.round(usuario.agosto),
            Math.round(usuario.setembro),
            Math.round(usuario.outubro),
            Math.round(usuario.novembro),
            Math.round(usuario.dezembro),
            Math.round(total)
        ];
    });
    
    doc.autoTable({
        head: [['Usuário', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Total']],
        body: tableData,
        startY: 45,
        styles: {
            fontSize: 7,
            cellPadding: 2,
            textColor: [35, 41, 70],
            lineColor: [74, 144, 226],
            lineWidth: 0.5
        },
        headStyles: {
            fillColor: [74, 144, 226],
            textColor: [255, 255, 255],
            fontStyle: 'bold',
            fontSize: 8
        },
        alternateRowStyles: {
            fillColor: [248, 249, 250]
        },
        columnStyles: {
            0: { cellWidth: 25, fontStyle: 'bold' },
            13: { fillColor: [227, 242, 253], fontStyle: 'bold' }
        },
        margin: { left: 15, right: 15 }
    });
    
    doc.save(`tabela_mes_${timestamp}.pdf`);
}

function exportarMesExcel(timestamp) {
    const wb = XLSX.utils.book_new();
    const ws = {};
    
    let titulo = 'APROD - Produtividade por Mês';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - APROD - Produtividade por Mês`;
    }
    
    ws['A1'] = { 
        v: titulo, 
        t: 's',
        s: {
            font: { name: "Calibri", sz: 18, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "4A90E2" } },
            alignment: { horizontal: "center", vertical: "center" }
        }
    };
    
    const headers = ['Usuário', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro', 'Total'];
    const headerCells = ['A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3', 'K3', 'L3', 'M3', 'N3'];
    
    headers.forEach((header, index) => {
        ws[headerCells[index]] = {
            v: header,
            t: 's',
            s: {
                font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "4A90E2" } },
                alignment: { horizontal: "center", vertical: "center" }
            }
        };
    });
    
    let row = 4;
    tabelaMes.forEach(usuario => {
        const total = usuario.janeiro + usuario.fevereiro + usuario.marco + 
                     usuario.abril + usuario.maio + usuario.junho + usuario.julho +
                     usuario.agosto + usuario.setembro + usuario.outubro + 
                     usuario.novembro + usuario.dezembro;
        
        const rowData = [
            usuario.usuario,
            Math.round(usuario.janeiro),
            Math.round(usuario.fevereiro),
            Math.round(usuario.marco),
            Math.round(usuario.abril),
            Math.round(usuario.maio),
            Math.round(usuario.junho),
            Math.round(usuario.julho),
            Math.round(usuario.agosto),
            Math.round(usuario.setembro),
            Math.round(usuario.outubro),
            Math.round(usuario.novembro),
            Math.round(usuario.dezembro),
            Math.round(total)
        ];
        
        const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
        
        rowData.forEach((value, colIndex) => {
            const cellAddress = `${columns[colIndex]}${row}`;
            ws[cellAddress] = {
                v: value,
                t: colIndex === 0 ? 's' : 'n',
                s: {
                    font: { name: "Calibri", sz: 11 },
                    alignment: { horizontal: colIndex === 0 ? "left" : "center", vertical: "center" }
                }
            };
        });
        
        row++;
    });
    
    ws['!ref'] = `A1:N${row - 1}`;
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 13 } }];
    
    XLSX.utils.book_append_sheet(wb, ws, "Produtividade Mensal");
    XLSX.writeFile(wb, `tabela_mes_${timestamp}.xlsx`);
}

function alterarModeloIA(valor) {
    configuracoes.iaModelo = valor;
    
    const modeloAtual = document.getElementById('modelo-atual');
    const modeloDireto = document.getElementById('ia-modelo-direto');
    const modeloConfig = document.getElementById('ia-modelo');
    
    if (modeloAtual) {
        modeloAtual.textContent = modeloDireto.options[modeloDireto.selectedIndex].text;
    }
    
    if (modeloConfig && modeloConfig.value !== valor) {
        modeloConfig.value = valor;
    }
    
    localStorage.setItem('aprod-configuracoes', JSON.stringify(configuracoes));
}

function verificarMesesComDados() {
    if (!excelData || excelData.length === 0) return [];
    
    const mesesComDados = new Set();
    
    excelData.forEach(row => {
        if (row['Data criação']) {
            const dataCompleta = new Date(row['Data criação']);
            if (!isNaN(dataCompleta.getTime())) {
                mesesComDados.add(dataCompleta.getMonth());
            }
        }
    });
    
    return Array.from(mesesComDados);
}

function verificarMesesComDadosAno() {
    if (!excelData || excelData.length === 0) return [];
    
    const mesesComDados = new Set();
    
    excelData.forEach(row => {
        if (row['Data criação']) {
            const dataCompleta = new Date(row['Data criação']);
            if (!isNaN(dataCompleta.getTime()) && dataCompleta.getFullYear() === anoSelecionado) {
                mesesComDados.add(dataCompleta.getMonth());
            }
        }
    });
    
    return Array.from(mesesComDados);
}

function aplicarEstiloMesesVazios() {
    const mesesComDados = verificarMesesComDadosAno();
    const meses = ['janeiro', 'fevereiro', 'marco', 'abril', 'maio', 'junho', 
                  'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
    
    meses.forEach((mes, index) => {
        const th = document.querySelector(`th[onclick*="${mes}"]`);
        if (th) {
            if (mesesComDados.includes(index)) {
                th.classList.remove('mes-vazio');
            } else {
                th.classList.add('mes-vazio');
            }
        }
    });
    
    const tbody = document.getElementById('tabela-mes');
    if (tbody) {
        tbody.querySelectorAll('tr').forEach(row => {
            const cells = row.querySelectorAll('td');
            cells.forEach((cell, cellIndex) => {
                if (cellIndex > 0 && cellIndex < 13) {
                    const mesIndex = cellIndex - 1;
                    if (mesesComDados.includes(mesIndex)) {
                        cell.classList.remove('mes-vazio');
                    } else {
                        cell.classList.add('mes-vazio');
                        cell.textContent = '';
                    }
                }
            });
        });
    }
}

function extrairAnosDisponiveis() {
    if (!excelData || excelData.length === 0) {
        anosDisponiveis = [];
        return;
    }
    
    const anos = new Set();
    
    excelData.forEach(row => {
        if (row['Data criação']) {
            const dataCompleta = new Date(row['Data criação']);
            if (!isNaN(dataCompleta.getTime())) {
                anos.add(dataCompleta.getFullYear());
            }
        }
    });
    
    anosDisponiveis = Array.from(anos).sort((a, b) => b - a);
    
    if (anosDisponiveis.length > 0 && !anosDisponiveis.includes(anoSelecionado)) {
        anoSelecionado = anosDisponiveis[0];
    }
    
    renderizarSeletorAno();
}

function renderizarSeletorAno() {
    const container = document.getElementById('seletor-ano');
    if (!container) return;
    
    container.innerHTML = '';
    
    if (anosDisponiveis.length <= 1) {
        container.style.display = 'none';
        return;
    }
    
    container.style.display = 'flex';
    
    anosDisponiveis.forEach(ano => {
        const anoButton = document.createElement('button');
        anoButton.className = `ano-chip ${ano === anoSelecionado ? 'active' : ''}`;
        anoButton.textContent = ano;
        anoButton.onclick = () => selecionarAno(ano);
        container.appendChild(anoButton);
    });
}

function selecionarAno(novoAno) {
    if (anoSelecionado === novoAno) return;
    
    anoSelecionado = novoAno;
    
    const tabela = document.querySelector('.mes-table tbody');
    if (tabela) {
        tabela.style.opacity = '0.3';
        tabela.style.transform = 'scale(0.95)';
        
        setTimeout(() => {
            gerarTabelaMes();
            renderizarSeletorAno();
            
            tabela.style.opacity = '1';
            tabela.style.transform = 'scale(1)';
        }, 300);
    } else {
        gerarTabelaMes();
        renderizarSeletorAno();
    }
}

async function gerarDadosAssuntos() {
    if (!excelData || excelData.length === 0) {
        alert('Nenhum dado foi importado ainda. Por favor, importe um arquivo Excel primeiro.');
        return;
    }

    console.log('Iniciando processamento de assuntos...');
    
    assuntosProcessados = {
        minutasPorAssunto: new Map(),
        processosPorAssunto: new Map(),
        codigosAssuntos: new Map(),
        totalMinutasAssuntos: 0,
        totalProcessosAssuntos: 0,
        assuntosPorMes: new Map()
    };

    let filteredData = excelData;
    
    if (processedData.mesesAtivos.length > 0 && processedData.mesesAtivos.length < processedData.mesesDisponiveis.length) {
        const mesesAtivosSet = new Set(processedData.mesesAtivos);
        filteredData = excelData.filter(row => {
            const data = row['Data criação'];
            if (!data) return false;
            const mesAno = extrairMesAno(data);
            return mesAno && mesesAtivosSet.has(mesAno);
        });
    }

    filteredData.forEach(row => {
        const codigoAssunto = row['Cod. assunto'];
        const usuario = row['Usuário'];
        const data = row['Data criação'];
        
        if (!codigoAssunto || !usuario) return;

        const peso = pesosAtivos && pesosAtuais[row['Tipo']] ? pesosAtuais[row['Tipo']] : 1;
        const minutas = peso;

        if (!assuntosProcessados.minutasPorAssunto.has(codigoAssunto)) {
            assuntosProcessados.minutasPorAssunto.set(codigoAssunto, 0);
            assuntosProcessados.processosPorAssunto.set(codigoAssunto, new Set());
            assuntosProcessados.codigosAssuntos.set(codigoAssunto, {
                codigo: codigoAssunto,
                descricao: obterDescricaoAssunto(codigoAssunto),
                usuarios: new Map(),
                porMes: new Map()
            });
        }

        assuntosProcessados.minutasPorAssunto.set(
            codigoAssunto,
            assuntosProcessados.minutasPorAssunto.get(codigoAssunto) + minutas
        );

        assuntosProcessados.processosPorAssunto.get(codigoAssunto).add(row['Nro. processo']);

        const dadosAssunto = assuntosProcessados.codigosAssuntos.get(codigoAssunto);
        if (!dadosAssunto.usuarios.has(usuario)) {
            dadosAssunto.usuarios.set(usuario, 0);
        }
        dadosAssunto.usuarios.set(usuario, dadosAssunto.usuarios.get(usuario) + minutas);

        const mesAno = extrairMesAno(data);
        if (mesAno) {
            if (!dadosAssunto.porMes.has(mesAno)) {
                dadosAssunto.porMes.set(mesAno, 0);
            }
            dadosAssunto.porMes.set(mesAno, dadosAssunto.porMes.get(mesAno) + minutas);
        }

        assuntosProcessados.totalMinutasAssuntos += minutas;
    });

    assuntosProcessados.totalProcessosAssuntos = new Set(filteredData.map(row => row['Nro. processo'])).size;

    assuntosData = Array.from(assuntosProcessados.minutasPorAssunto.entries())
        .map(([codigo, minutas]) => {
            const dadosAssunto = assuntosProcessados.codigosAssuntos.get(codigo);
            return {
                codigo,
                descricao: dadosAssunto.descricao,
                minutas,
                processos: assuntosProcessados.processosPorAssunto.get(codigo).size,
                percentualMinutas: (minutas / assuntosProcessados.totalMinutasAssuntos) * 100,
                percentualProcessos: (assuntosProcessados.processosPorAssunto.get(codigo).size / assuntosProcessados.totalProcessosAssuntos) * 100,
                topUsuarios: Array.from(dadosAssunto.usuarios.entries())
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 3),
                porMes: dadosAssunto.porMes
            };
        })
        .sort((a, b) => b.minutas - a.minutas);

    console.log('Dados de assuntos processados:', assuntosData);
}

async function carregarAssuntosJSON() {
    try {
        const response = await fetch('assets/data/assuntos.json');
        if (!response.ok) {
            console.warn('Arquivo assuntos.json não encontrado, usando modo básico');
            assuntosJSON = null;
            return;
        }
        assuntosJSON = await response.json();
        console.log('Arquivo assuntos.json carregado com sucesso');
    } catch (error) {
        console.warn('Não foi possível carregar assuntos.json, usando modo básico:', error.message);
        assuntosJSON = null;
    }
}

function obterDescricaoAssunto(codigo) {
    if (!codigo) {
        return 'Assunto não informado';
    }

    codigo = codigo.toString().trim();

    if (codigo.includes(',')) {
        const codigos = codigo.split(',').map(c => c.trim().replace(/^["']|["']$/g, ''));
        const descricoes = codigos.map(c => {
            if (assuntosJSON && assuntosJSON["Assunto Judicial"]) {
                const assunto = assuntosJSON["Assunto Judicial"].find(item => 
                    item["Cod. Assunto"] && item["Cod. Assunto"].toString().trim() === c
                );
                return assunto ? assunto.Assunto : `Assunto ${c}`;
            }
            return `Assunto ${c}`;
        });
        return descricoes.join(' + ');
    }

    codigo = codigo.replace(/^["']|["']$/g, '');
    
    if (assuntosJSON && assuntosJSON["Assunto Judicial"]) {
        const assunto = assuntosJSON["Assunto Judicial"].find(item => 
            item["Cod. Assunto"] && item["Cod. Assunto"].toString().trim() === codigo
        );
        return assunto ? assunto.Assunto : `Assunto ${codigo}`;
    }

    return `Assunto ${codigo}`;
}

function initializeApp() {
    carregarConfiguracoes();
    updateTheme();
    updatePesosButton();
    createChart();
    renderMonths();
    updateKPIs();
    carregarAssuntosJSON();
}

function gerarTabelaAssuntos() {
    if (!assuntosData || assuntosData.length === 0) {
        document.getElementById('tabela-assuntos').innerHTML = 
            '<tr><td colspan="5" style="text-align: center; padding: 40px; color: var(--text-secondary);">Nenhum dado de assuntos encontrado.</td></tr>';
        return;
    }

    renderizarTabelaAssuntos();
}

function renderizarTabelaAssuntos() {
    const container = document.getElementById('tabela-assuntos');
    if (!container || !assuntosData) return;
    
    container.innerHTML = '';
    
    const filtro = document.getElementById('filtro-assuntos')?.value.toLowerCase() || '';
    let assuntosFiltrados = assuntosData;
    
    if (filtro) {
        assuntosFiltrados = assuntosData.filter(assunto => 
            assunto.descricao.toLowerCase().includes(filtro) ||
            assunto.codigo.toString().toLowerCase().includes(filtro)
        );
    }

    const limitePadrao = 20;
    const mostrarTodos = container.getAttribute('data-expandido') === 'true';
    const assuntosParaMostrar = mostrarTodos ? assuntosFiltrados : assuntosFiltrados.slice(0, limitePadrao);

    assuntosParaMostrar.forEach((assunto, index) => {
        const posicaoReal = assuntosFiltrados.findIndex(a => a === assunto) + 1;
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="posicao-cell">${posicaoReal}</td>
            <td class="descricao-cell" title="${assunto.descricao}">${assunto.descricao}</td>
            <td class="minutas-cell">${Math.round(assunto.minutas)}</td>
            <td class="processos-cell">${assunto.processos}</td>
            <td class="percentuais-cell">
                <div class="percentual-minutas">${assunto.percentualMinutas.toFixed(1)}%</div>
                <div class="percentual-processos">${assunto.percentualProcessos.toFixed(1)}%</div>
            </td>
        `;
        container.appendChild(row);
    });

    if (assuntosFiltrados.length > limitePadrao && !filtro) {
        const expandirRow = document.createElement('tr');
        expandirRow.className = 'expandir-row';
        expandirRow.innerHTML = `
            <td colspan="5" class="expandir-cell">
                <button class="btn-expandir" onclick="toggleExpandirTabela()">
                    ${mostrarTodos ? 
                        `🔼 Mostrar apenas top ${limitePadrao}` : 
                        `🔽 Mostrar todos os ${assuntosFiltrados.length} assuntos`
                    }
                </button>
            </td>
        `;
        container.appendChild(expandirRow);
    }

    atualizarIconesOrdenacao(document.querySelector('.assuntos-table'), sortColumnAssuntos, sortAscendingAssuntos);
}

function toggleExpandirTabela() {
    const container = document.getElementById('tabela-assuntos');
    const expandido = container.getAttribute('data-expandido') === 'true';
    container.setAttribute('data-expandido', !expandido);
    renderizarTabelaAssuntos();
}

function renderizarResumoAssuntos() {
    const container = document.getElementById('resumo-assuntos');
    if (!container || !assuntosData || assuntosData.length === 0) return;

    const top5Assuntos = assuntosData.slice(0, 5);
    const topUsuarios = gerarTopUsuariosPorAssunto();

    container.innerHTML = `
        <div class="resumo-item top-assuntos">
            <div class="resumo-label">Top 5 Assuntos</div>
            <div class="top-list">
                ${top5Assuntos.map((assunto, index) => `
                    <div class="top-item">
                        <span class="top-posicao">${index + 1}º</span>
                        <span class="top-nome" title="${assunto.descricao}">${assunto.descricao}</span>
                        <span class="top-minutas">${Math.round(assunto.minutas)}</span>
                    </div>
                `).join('')}
            </div>
        </div>

        <div class="resumo-item top-usuarios">
            <div class="resumo-label">Top 3 Usuários por Assunto</div>
            <div class="usuarios-list">
                ${topUsuarios.slice(0, 3).map(usuario => `
                    <div class="usuario-assunto-item">
                        <div class="usuario-nome-assunto">${usuario.nome}</div>
                        <div class="usuario-top-assuntos">
                            ${usuario.topAssuntos.slice(0, 3).map((assunto, index) => `
                                <div class="assunto-item-linha">
                                    <span class="assunto-posicao-badge">${index + 1}</span>
                                    <span class="assunto-nome-mini" title="${assunto.descricao}">${assunto.descricao}</span>
                                    <span class="assunto-minutas-mini">${Math.round(assunto.minutas)}</span>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                `).join('')}
            </div>
        </div>

        <div class="grafico-assuntos-container">
            <div class="grafico-header">
                <h3>Variação de Minutas por Assunto ao Longo dos Meses</h3>
            </div>
            <div class="grafico-chart-container">
                <canvas id="chart-variacao-assuntos"></canvas>
            </div>
        </div>
    `;
}

function gerarTopUsuariosPorAssunto() {
    if (!excelData || excelData.length === 0) return [];

    const usuariosAssuntos = new Map();

    let filteredData = excelData;
    if (processedData.mesesAtivos.length > 0 && processedData.mesesAtivos.length < processedData.mesesDisponiveis.length) {
        const mesesAtivosSet = new Set(processedData.mesesAtivos);
        filteredData = excelData.filter(row => {
            const data = row['Data criação'];
            if (!data) return false;
            const mesAno = extrairMesAno(data);
            return mesAno && mesesAtivosSet.has(mesAno);
        });
    }

    filteredData.forEach(row => {
        const usuario = row['Usuário'];
        const codigoAssunto = row['Cod. assunto'];
        
        if (!usuario || !codigoAssunto) return;

        const peso = pesosAtivos && pesosAtuais[row['Tipo']] ? pesosAtuais[row['Tipo']] : 1;

        if (!usuariosAssuntos.has(usuario)) {
            usuariosAssuntos.set(usuario, new Map());
        }

        const assuntosUsuario = usuariosAssuntos.get(usuario);
        if (!assuntosUsuario.has(codigoAssunto)) {
            assuntosUsuario.set(codigoAssunto, {
                codigo: codigoAssunto,
                descricao: obterDescricaoAssunto(codigoAssunto),
                minutas: 0
            });
        }

        const assunto = assuntosUsuario.get(codigoAssunto);
        assunto.minutas += peso;
    });

    return Array.from(usuariosAssuntos.entries())
        .map(([nome, assuntos]) => ({
            nome,
            totalMinutas: Array.from(assuntos.values()).reduce((sum, a) => sum + a.minutas, 0),
            topAssuntos: Array.from(assuntos.values()).sort((a, b) => b.minutas - a.minutas)
        }))
        .sort((a, b) => b.totalMinutas - a.totalMinutas);
}

function toggleTodosAssuntos(checked) {
    const checkboxes = document.querySelectorAll('#assuntos-checkboxes input[type="checkbox"]:not(#todos-assuntos)');
    checkboxes.forEach(checkbox => {
        checkbox.checked = checked;
    });
    atualizarGraficoAssuntos();
}

function atualizarGraficoAssuntos() {
    const checkboxesTodos = document.getElementById('todos-assuntos');
    const checkboxes = document.querySelectorAll('#assuntos-checkboxes input[type="checkbox"]:not(#todos-assuntos)');
    const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
    
    if (checkboxesTodos) {
        checkboxesTodos.checked = checkedCount === checkboxes.length;
        checkboxesTodos.indeterminate = checkedCount > 0 && checkedCount < checkboxes.length;
    }
    
    criarGraficoVariacaoAssuntos();
}

function criarGraficoVariacaoAssuntos() {
    const container = document.querySelector('.grafico-assuntos-container');
    if (!container || !assuntosData || assuntosData.length === 0) return;

    if (chartAssuntos) {
        chartAssuntos.destroy();
    }

    const canvasId = 'chart-variacao-assuntos';
    container.innerHTML = `
        <div class="grafico-header">
            <h3>Variação de Minutas por Assunto ao Longo dos Meses</h3>
        </div>
        <div class="grafico-chart-container">
            <canvas id="${canvasId}"></canvas>
        </div>
    `;

    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    const top6Assuntos = assuntosData.slice(0, 6);
    const mesesOrdenados = processedData.mesesDisponiveis.sort((a, b) => {
        const [mesA, anoA] = a.split('/');
        const [mesB, anoB] = b.split('/');
        const dataA = new Date(parseInt(anoA), obterNumeroMes(mesA), 1);
        const dataB = new Date(parseInt(anoB), obterNumeroMes(mesB), 1);
        return dataA - dataB;
    });

    const cores = [
        '#1565c0', '#43a047', '#e53935', '#8e24aa', '#fb8c00', '#00acc1'
    ];

    const datasets = top6Assuntos.map((assunto, index) => {
        const dados = mesesOrdenados.map(mes => {
            const minutas = assunto.porMes.get(mes) || 0;
            return Math.round(minutas);
        });

        return {
            label: assunto.descricao.length > 30 ? 
                   assunto.descricao.substring(0, 30) + '...' : 
                   assunto.descricao,
            data: dados,
            borderColor: cores[index],
            backgroundColor: cores[index] + '20',
            borderWidth: 3,
            fill: false,
            tension: 0.4,
            pointRadius: 6,
            pointHoverRadius: 8,
            pointBackgroundColor: cores[index],
            pointBorderColor: '#ffffff',
            pointBorderWidth: 2
        };
    });

    const isDark = document.documentElement.getAttribute('data-theme') !== 'light';

    chartAssuntos = new Chart(ctx, {
        type: 'line',
        data: {
            labels: mesesOrdenados,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                title: {
                    display: false
                },
                legend: {
                    display: true,
                    position: 'bottom',
                    labels: {
                        color: isDark ? '#ffffff' : '#232946',
                        font: {
                            size: 12
                        },
                        padding: 20,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                tooltip: {
                    backgroundColor: isDark ? '#2d2d2d' : '#ffffff',
                    titleColor: isDark ? '#ffffff' : '#232946',
                    bodyColor: isDark ? '#ffffff' : '#232946',
                    borderColor: isDark ? '#4a90e2' : '#3cb3e6',
                    borderWidth: 1,
                    cornerRadius: 8,
                    displayColors: true,
                    callbacks: {
                        title: function(context) {
                            return `Mês: ${context[0].label}`;
                        },
                        label: function(context) {
                            return `${context.dataset.label}: ${context.parsed.y} minutas`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        color: isDark ? '#404040' : '#e0e0e0',
                        drawBorder: false
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946',
                        font: {
                            size: 11
                        },
                        maxRotation: 45
                    }
                },
                y: {
                    beginAtZero: true,
                    grid: {
                        color: isDark ? '#404040' : '#e0e0e0',
                        drawBorder: false
                    },
                    ticks: {
                        color: isDark ? '#ffffff' : '#232946',
                        font: {
                            size: 11
                        },
                        callback: function(value) {
                            return Math.round(value);
                        }
                    }
                }
            },
            animation: {
                duration: 1500,
                easing: 'easeInOutQuart'
            }
        }
    });
}

function obterNumeroMes(nomeAbreviado) {
    const meses = {
        'jan': 0, 'fev': 1, 'mar': 2, 'abr': 3, 'mai': 4, 'jun': 5,
        'jul': 6, 'ago': 7, 'set': 8, 'out': 9, 'nov': 10, 'dez': 11
    };
    return meses[nomeAbreviado] || 0;
}

function gerarDadosVariacaoMensal(assuntosSelecionados) {
    const mesesDisponiveis = processedData.mesesDisponiveis || [];
    const dadosPorAssunto = {};
    
    assuntosSelecionados.forEach(assunto => {
        dadosPorAssunto[assunto] = {};
        mesesDisponiveis.forEach(mes => {
            dadosPorAssunto[assunto][mes] = 0;
        });
    });
    
    if (!excelData || excelData.length === 0) {
        return { labels: mesesDisponiveis, dados: dadosPorAssunto };
    }

    excelData.forEach(row => {
        const codigoAssunto = row['Cod. assunto'];
        const tipo = row['Tipo'];
        const dataRow = row['Data criação'];
        
        if (!codigoAssunto || !tipo || !dataRow) return;
        
        const descricaoAssunto = obterDescricaoAssunto(codigoAssunto);
        const chaveAssunto = `${codigoAssunto} - ${descricaoAssunto}`;
        
        if (!assuntosSelecionados.includes(chaveAssunto)) return;
        
        const data = new Date(dataRow);
        if (isNaN(data.getTime())) return;
        
        const mesAno = `${String(data.getMonth() + 1).padStart(2, '0')}/${data.getFullYear()}`;
        
        if (mesesDisponiveis.includes(mesAno)) {
            const peso = pesosAtivos && pesosAtuais[tipo] ? pesosAtuais[tipo] : 1;
            dadosPorAssunto[chaveAssunto][mesAno] += peso;
        }
    });
    
    return { labels: mesesDisponiveis, dados: dadosPorAssunto };
}

function ordenarTabelaAssuntos(coluna) {
    if (!assuntosData) return;
    
    const isAscending = sortColumnAssuntos === coluna ? !sortAscendingAssuntos : false;
    sortColumnAssuntos = coluna;
    sortAscendingAssuntos = isAscending;
    
    const filtro = document.getElementById('filtro-assuntos')?.value.toLowerCase() || '';
    let dadosParaOrdenar = assuntosData;
    
    if (filtro) {
        dadosParaOrdenar = assuntosData.filter(assunto => 
            assunto.descricao.toLowerCase().includes(filtro) ||
            assunto.codigo.toString().toLowerCase().includes(filtro)
        );
    }
    
    dadosParaOrdenar.sort((a, b) => {
        let valorA, valorB;
        
        switch(coluna) {
            case 'posicao':
                return isAscending ? 1 : -1;
            case 'descricao':
                valorA = a.descricao.toLowerCase();
                valorB = b.descricao.toLowerCase();
                break;
            case 'minutas':
                valorA = a.minutas;
                valorB = b.minutas;
                break;
            case 'processos':
                valorA = a.processos;
                valorB = b.processos;
                break;
            default:
                return 0;
        }
        
        if (typeof valorA === 'string') {
            return isAscending ? valorA.localeCompare(valorB) : valorB.localeCompare(valorA);
        } else {
            return isAscending ? valorA - valorB : valorB - valorA;
        }
    });
    
    if (filtro) {
        assuntosData = dadosParaOrdenar.concat(
            assuntosData.filter(assunto => 
                !assunto.descricao.toLowerCase().includes(filtro) &&
                !assunto.codigo.toString().toLowerCase().includes(filtro)
            )
        );
    } else {
        assuntosData = dadosParaOrdenar;
    }
    
    renderizarTabelaAssuntos();
}

function atualizarIconesOrdenacao(tabela, colunaAtiva, ascending) {
    if (!tabela) return;
    
    const headers = tabela.querySelectorAll('th.sortable');
    headers.forEach(header => {
        const icon = header.querySelector('.sort-icon');
        if (icon) {
            icon.textContent = '↕';
            icon.style.opacity = '0.5';
        }
    });
    
    const activeHeader = tabela.querySelector(`th[onclick*="${colunaAtiva}"]`);
    if (activeHeader) {
        const icon = activeHeader.querySelector('.sort-icon');
        if (icon) {
            icon.textContent = ascending ? '↑' : '↓';
            icon.style.opacity = '1';
        }
    }
}

function mostrarPopupExportacaoAssuntos() {
    document.getElementById('popup-exportacao-assuntos').style.display = 'flex';
}

function fecharPopupExportacaoAssuntos() {
    document.getElementById('popup-exportacao-assuntos').style.display = 'none';
}

function exportarTabelaAssuntos(formato) {
    if (!assuntosData || assuntosData.length === 0) {
        alert('Não há dados para exportar!');
        return;
    }
    
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    
    if (formato === 'pdf') {
        exportarAssuntosPDF(timestamp);
    } else if (formato === 'xlsx') {
        exportarAssuntosExcel(timestamp);
    }
    
    fecharPopupExportacaoAssuntos();
}

function exportarAssuntosPDF(timestamp) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('landscape');
    
    doc.setFillColor(74, 144, 226);
    doc.rect(0, 0, doc.internal.pageSize.width, 25, 'F');
    
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    
    let titulo = 'APROD - Produtividade por Assunto';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - APROD - Produtividade por Assunto`;
    }
    doc.text(titulo, 15, 16);
    
    doc.setTextColor(74, 144, 226);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')}`, 15, 35);
    
    let mesesTexto = `Meses: ${processedData.mesesAtivos.join(', ')}`;
    if (processedData.mesesAtivos.length === 0) {
        mesesTexto = 'Meses: Todos os meses disponíveis';
    }
    doc.text(mesesTexto, 15, 42);
    
    const tableData = assuntosData.map((assunto, index) => [
        index + 1,
        assunto.descricao.length > 70 ? assunto.descricao.substring(0, 67) + '...' : assunto.descricao,
        Math.round(assunto.minutas),
        assunto.processos,
        `${assunto.percentualMinutas.toFixed(1)}% / ${assunto.percentualProcessos.toFixed(1)}%`
    ]);
    
    doc.autoTable({
        head: [['#', 'Descrição do Assunto', 'Minutas', 'Processos', '% Min / % Proc']],
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
            0: { cellWidth: 15, halign: 'center' },
            1: { cellWidth: 110 },
            2: { cellWidth: 25, halign: 'center', fontStyle: 'bold' },
            3: { cellWidth: 25, halign: 'center' },
            4: { cellWidth: 35, halign: 'center', fontSize: 7 }
        },
        margin: { left: 15, right: 15 }
    });
    
    const totalGeral = Math.round(assuntosProcessados.totalMinutasAssuntos);
    const totalProcessos = assuntosProcessados.totalProcessosAssuntos;
    
    const finalY = doc.previousAutoTable.finalY + 10;
    doc.setFillColor(227, 242, 253);
    doc.rect(15, finalY, doc.internal.pageSize.width - 30, 15, 'F');
    
    doc.setTextColor(74, 144, 226);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(`Totais: ${totalGeral} minutas • ${totalProcessos} processos • ${assuntosData.length} assuntos`, 20, finalY + 10);
    
    doc.save(`assuntos_${timestamp}.pdf`);
}

function exportarAssuntosExcel(timestamp) {
    let mesesTexto = `Meses: ${processedData.mesesAtivos.join(', ')}`;
    if (processedData.mesesAtivos.length === 0) {
        mesesTexto = 'Período: Todos os meses disponíveis';
    }
    
    const wb = XLSX.utils.book_new();
    const ws = {};
    
    let titulo = 'APROD - Produtividade por Assunto';
    if (configuracoes.nomeGabinete) {
        titulo = `${configuracoes.nomeGabinete} - ${titulo}`;
    }
    
    ws['A1'] = { 
        v: titulo, 
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
    
    const headers = ['Posição', 'Descrição do Assunto', 'Minutas', 'Processos', '% Minutas', '% Processos'];
    const headerCells = ['A6', 'B6', 'C6', 'D6', 'E6', 'F6'];
    
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
    assuntosData.forEach((assunto, index) => {
        const rowData = [
            index + 1,
            assunto.descricao,
            Math.round(assunto.minutas),
            assunto.processos,
            `${assunto.percentualMinutas.toFixed(1)}%`,
            `${assunto.percentualProcessos.toFixed(1)}%`
        ];
        
        const columns = ['A', 'B', 'C', 'D', 'E', 'F'];
        
        rowData.forEach((value, colIndex) => {
            const cellAddress = `${columns[colIndex]}${row}`;
            ws[cellAddress] = {
                v: value,
                t: colIndex === 1 ? 's' : (colIndex === 4 || colIndex === 5 ? 's' : 'n'),
                s: {
                    font: { name: "Calibri", sz: 11 },
                    alignment: { horizontal: colIndex === 1 ? "left" : "center", vertical: "center" },
                    border: {
                        top: { style: "thin", color: { rgb: "E0E0E0" } },
                        bottom: { style: "thin", color: { rgb: "E0E0E0" } },
                        left: { style: "thin", color: { rgb: "E0E0E0" } },
                        right: { style: "thin", color: { rgb: "E0E0E0" } }
                    }
                }
            };
        });
        
        row++;
    });
    
    const totalGeral = Math.round(assuntosProcessados.totalMinutasAssuntos);
    const totalProcessos = assuntosProcessados.totalProcessosAssuntos;
    
    row += 1;
    const totalColumns = ['A', 'B', 'C', 'D', 'E', 'F'];
    const totalValues = ['TOTAIS:', `${assuntosData.length} assuntos`, totalGeral, totalProcessos, '100%', '100%'];
    
    totalValues.forEach((value, colIndex) => {
        const cellAddress = `${totalColumns[colIndex]}${row}`;
        ws[cellAddress] = {
            v: value,
            t: colIndex === 1 ? 's' : (colIndex === 4 || colIndex === 5 ? 's' : 'n'),
            s: {
                font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "4A90E2" } },
                fill: { fgColor: { rgb: "E3F2FD" } },
                alignment: { horizontal: colIndex === 1 ? "left" : "center", vertical: "center" },
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
    ws['!ref'] = `A1:F${lastRow}`;
    
    ws['!cols'] = [
        { wch: 10 }, { wch: 60 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }
    ];
    
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
    
    ws['!rows'] = [
        { hpt: 30 },
        { hpt: 15 },
        { hpt: 18 },
        { hpt: 18 },
        { hpt: 15 },
        { hpt: 25 }
    ];
    
    for (let i = 6; i < lastRow; i++) {
        ws[`!rows`][i] = { hpt: 20 };
    }
    
    XLSX.utils.book_append_sheet(wb, ws, "Produtividade por Assunto");
    XLSX.writeFile(wb, `assuntos_${timestamp}.xlsx`);
}

window.onclick = function(event) {
    const popup = document.getElementById('popup-exportacao');
    const popupDashboard = document.getElementById('popup-exportacao-dashboard');
    const popupMes = document.getElementById('popup-exportacao-mes');
    const popupAssuntos = document.getElementById('popup-exportacao-assuntos');
    if (event.target === popup) {
        fecharPopupExportacao();
    }
    if (event.target === popupDashboard) {
        fecharPopupExportacaoDashboard();
    }
    if (event.target === popupMes) {
        fecharPopupExportacaoMes();
    }
    if (event.target === popupAssuntos) {
        fecharPopupExportacaoAssuntos();
    }
}

window.addEventListener('resize', function() {
    if (chartInstance) {
        chartInstance.resize();
    }
});
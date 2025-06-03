let chartInstance = null;
let isDarkTheme = true;
let appData = {
    totalMinutas: 0,
    diaProdutivo: '--',
    topUsers: ['Carregando...'],
    mesesAtivos: ['Jan 2025', 'Fev 2025', 'Mar 2025', 'Abr 2025', 'Mai 2025'],
    chartData: {
        labels: ['Ana Silva', 'Carlos Santos', 'Maria Oliveira', 'João Pedro', 'Lucia Costa'],
        values: [145, 132, 128, 115, 98]
    }
};

function initializeApp() {
    updateTheme();
    loadMockData();
    createChart();
    renderMonths();
    
    setTimeout(() => {
        updateKPIs();
    }, 1000);
}

function loadMockData() {
    appData = {
        totalMinutas: 2547,
        diaProdutivo: 'Terça-feira\n(432 min)',
        topUsers: [
            '1º Ana Silva (145 min)',
            '2º Carlos Santos (132 min)', 
            '3º Maria Oliveira (128 min)'
        ],
        mesesAtivos: ['Jan 2025', 'Fev 2025', 'Mar 2025', 'Abr 2025', 'Mai 2025'],
        mesesDisponiveis: ['Dez 2024', 'Jan 2025', 'Fev 2025', 'Mar 2025', 'Abr 2025', 'Mai 2025', 'Jun 2025'],
        chartData: {
            labels: ['Ana Silva', 'Carlos Santos', 'Maria Oliveira', 'João Pedro', 'Lucia Costa', 'Pedro Alves', 'Carla Lima'],
            values: [145, 132, 128, 115, 98, 87, 76]
        }
    };
}

function updateKPIs() {
    document.getElementById('total-minutas').textContent = appData.totalMinutas.toLocaleString();
    document.getElementById('dia-produtivo').textContent = appData.diaProdutivo;
    
    const topUsersContainer = document.getElementById('top-users');
    topUsersContainer.innerHTML = '';
    
    appData.topUsers.forEach(user => {
        const userDiv = document.createElement('div');
        userDiv.className = 'user-item';
        userDiv.textContent = user;
        topUsersContainer.appendChild(userDiv);
    });
}

function renderMonths() {
    const monthsGrid = document.getElementById('months-grid');
    monthsGrid.innerHTML = '';
    
    appData.mesesDisponiveis.forEach(month => {
        const isActive = appData.mesesAtivos.includes(month);
        
        const monthChip = document.createElement('div');
        monthChip.className = `month-chip ${isActive ? 'active' : ''}`;
        monthChip.textContent = month;
        monthChip.onclick = () => toggleMonth(month);
        
        monthsGrid.appendChild(monthChip);
    });
}

function toggleMonth(month) {
    const index = appData.mesesAtivos.indexOf(month);
    
    if (index > -1) {
        appData.mesesAtivos.splice(index, 1);
    } else {
        appData.mesesAtivos.push(month);
    }
    
    renderMonths();
    updateChart();
    updateKPIs();
}

function createChart() {
    const ctx = document.getElementById('produtividade-chart').getContext('2d');
    
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
    
    chartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: appData.chartData.labels,
            datasets: [{
                label: 'Minutas',
                data: appData.chartData.values,
                backgroundColor: isDark ? '#4a90e2' : '#3cb3e6',
                borderColor: isDark ? '#4a90e2' : '#3cb3e6',
                borderWidth: 0,
                borderRadius: 8
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
                            size: 12
                        }
                    }
                }
            },
            elements: {
                bar: {
                    borderRadius: 6
                }
            }
        }
    });
}

function updateChart() {
    if (chartInstance) {
        const activeMonthsCount = appData.mesesAtivos.length;
        const multiplier = activeMonthsCount > 0 ? activeMonthsCount * 0.8 : 0.1;
        
        const newValues = appData.chartData.values.map(value => 
            Math.round(value * multiplier)
        );
        
        chartInstance.data.datasets[0].data = newValues;
        chartInstance.update();
        
        const newTotal = newValues.reduce((sum, val) => sum + val, 0);
        appData.totalMinutas = newTotal;
        updateKPIs();
    }
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

function importData() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls,.csv';
    
    input.onchange = function(event) {
        const file = event.target.files[0];
        if (file) {
            processFile(file);
        }
    };
    
    input.click();
}

function processFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            alert(`Arquivo "${file.name}" carregado com sucesso!\n\nEm uma implementação completa, os dados seriam processados aqui.`);
            
            simulateDataUpdate();
            
        } catch (error) {
            alert('Erro ao processar arquivo: ' + error.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function simulateDataUpdate() {
    const newData = {
        totalMinutas: Math.floor(Math.random() * 3000) + 2000,
        diaProdutivo: ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira'][Math.floor(Math.random() * 5)] + `\n(${Math.floor(Math.random() * 200) + 300} min)`,
        topUsers: [
            `1º Usuário A (${Math.floor(Math.random() * 50) + 100} min)`,
            `2º Usuário B (${Math.floor(Math.random() * 40) + 90} min)`,
            `3º Usuário C (${Math.floor(Math.random() * 30) + 80} min)`
        ],
        chartData: {
            labels: ['Usuário A', 'Usuário B', 'Usuário C', 'Usuário D', 'Usuário E', 'Usuário F'],
            values: Array.from({length: 6}, () => Math.floor(Math.random() * 100) + 50)
        }
    };
    
    Object.assign(appData, newData);
    updateKPIs();
    createChart();
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
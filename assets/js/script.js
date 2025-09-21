let dadosAgiotagem = {};
let mesAtual = '';
const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

function getMesNome(data) {
    return `${meses[data.getMonth()]} ${data.getFullYear()}`;
}

function carregarOpcoesMeses() {
    const mesInicioSelect = document.getElementById('mesInicio');
    mesInicioSelect.innerHTML = '';
    
    const hoje = new Date();
    const anoAtual = hoje.getFullYear();
    
    const startYear = anoAtual - 2; 
    const endYear = anoAtual + 2;
    
    for (let ano = startYear; ano <= endYear; ano++) {
        for (let mes = 0; mes < 12; mes++) {
            const data = new Date(ano, mes, 1);
            const nomeMes = getMesNome(data);
            const option = document.createElement('option');
            option.value = nomeMes;
            option.textContent = nomeMes;
            mesInicioSelect.appendChild(option);
        }
    }
}

window.onload = function() {
    carregarOpcoesMeses();
    const hoje = new Date();
    const mesAtualNome = getMesNome(hoje);
    mesAtual = mesAtualNome;
    gerarBotoesMeses();
    carregarParcelasMes(mesAtual);
};

// Funções de Modal
function openModal(modalId) {
    document.getElementById(modalId).style.display = 'block';
}

function closeModal(modalId) {
    document.getElementById(modalId).style.display = 'none';
}

// Funções de Alerta
function mostrarAlerta(mensagem, tipo = 'success') {
    const alertContainer = document.getElementById('alertContainer');
    const alertClass = tipo === 'success' ? 'alert-success' : 'alert-error';
    alertContainer.innerHTML = `<div class="alert ${alertClass}">${mensagem}</div>`;
    setTimeout(() => {
        alertContainer.innerHTML = '';
    }, 5000);
}

// Carregar arquivo Excel
document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheetName = 'AGIOTAGEM PARA O JEAN';
            
            if (!workbook.Sheets[sheetName]) {
                mostrarAlerta('Planilha "AGIOTAGEM PARA O JEAN" não encontrada!', 'error');
                return;
            }

            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            
            processarDadosExcel(jsonData);
            mostrarAlerta('Arquivo carregado com sucesso!');
        } catch (error) {
            mostrarAlerta('Erro ao ler o arquivo: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
});

// Processar dados do Excel - Lógica Refatorada
function processarDadosExcel(data) {
    dadosAgiotagem = {};
    const itemsUnicos = {};
    const allParcels = [];

    let mesAtualLido = null;
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        
        if (row[0] && !row[1] && !row[2] && !row[3] && !row[4]) {
            mesAtualLido = row[0];
            continue;
        }
        
        if (row[0] === 'Item' || row[0] === 'Total' || !row[0]) {
            continue;
        }
        
        if (mesAtualLido && row[0] && row[1]) {
            const [parcelaAtual, totalParcelas] = row[1].split('/').map(Number);
            const [mesString, anoString] = mesAtualLido.split(' ');
            
            allParcels.push({
                item: row[0],
                parcelaAtual: parcelaAtual,
                totalParcelas: totalParcelas,
                valorTotal: parseFloat(row[2]) || 0,
                valorParcela: parseFloat(row[3]) || 0,
                pago: row[4] === true || row[4] === 'True' || row[4] === 'TRUE',
                data: new Date(parseInt(anoString), meses.indexOf(mesString), 1)
            });
        }
    }
    
    // Agrupar e processar os itens
    allParcels.forEach(parcela => {
        const itemKey = `${parcela.item}-${parcela.totalParcelas}`;
        
        if (!itemsUnicos[itemKey]) {
            itemsUnicos[itemKey] = {
                item: parcela.item,
                valorTotal: parcela.valorTotal,
                valorParcela: parcela.valorParcela,
                totalParcelas: parcela.totalParcelas,
                dataInicio: parcela.data,
                pagas: {}
            };
        }
        
        // Encontrar o mês de início real
        if (parcela.parcelaAtual === 1) {
            itemsUnicos[itemKey].dataInicio = parcela.data;
        } else if (parcela.data < itemsUnicos[itemKey].dataInicio) {
            const dataRealInicio = new Date(parcela.data.getFullYear(), parcela.data.getMonth() - (parcela.parcelaAtual - 1), 1);
            itemsUnicos[itemKey].dataInicio = dataRealInicio;
        }
        
        itemsUnicos[itemKey].pagas[parcela.parcelaAtual] = parcela.pago;
    });
    
    // Gerar todas as parcelas dinamicamente a partir dos itens únicos
    Object.values(itemsUnicos).forEach(item => {
        for (let i = 1; i <= item.totalParcelas; i++) {
            const dataParcela = new Date(item.dataInicio.getFullYear(), item.dataInicio.getMonth() + (i - 1), 1);
            const mesChave = getMesNome(dataParcela);
            
            if (!dadosAgiotagem[mesChave]) {
                dadosAgiotagem[mesChave] = [];
            }

            dadosAgiotagem[mesChave].push({
                item: item.item,
                parcelamento: `${i}/${item.totalParcelas}`,
                valorTotal: item.valorTotal,
                valorParcela: item.valorParcela,
                pago: item.pagas[i] || false
            });
        }
    });

    // Ordenar as parcelas de cada mês pelo nome do item
    Object.values(dadosAgiotagem).forEach(parcelas => {
        parcelas.sort((a, b) => a.item.localeCompare(b.item));
    });

    const hoje = new Date();
    const mesAtualNome = getMesNome(hoje);
    mesAtual = Object.keys(dadosAgiotagem).find(mes => mes === mesAtualNome) || Object.keys(dadosAgiotagem)[0] || mesAtualNome;
    
    atualizarResumo();
    gerarBotoesMeses();
    carregarParcelasMes(mesAtual);
}

// Geração dinâmica de botões de mês
function gerarBotoesMeses() {
    const mesSelector = document.querySelector('.mes-selector');
    mesSelector.innerHTML = '';
    
    const mesesDisponiveis = Object.keys(dadosAgiotagem).sort((a, b) => {
        const [mesA, anoA] = a.split(' ');
        const [mesB, anoB] = b.split(' ');
        const dataA = new Date(anoA, meses.indexOf(mesA), 1);
        const dataB = new Date(anoB, meses.indexOf(mesB), 1);
        return dataA - dataB;
    });
    
    mesesDisponiveis.forEach(mes => {
        const btn = document.createElement('div');
        btn.className = 'mes-btn';
        if (mes === mesAtual) {
            btn.classList.add('active');
        }
        btn.setAttribute('data-mes', mes);
        btn.setAttribute('onclick', `selecionarMes(this, '${mes}')`);
        btn.textContent = mes;
        mesSelector.appendChild(btn);
    });
}

// Seleção de Mês
function selecionarMes(elemento, mes) {
    document.querySelectorAll('.mes-btn').forEach(btn => btn.classList.remove('active'));
    elemento.classList.add('active');
    mesAtual = mes;
    carregarParcelasMes(mes);
}

// Carregar parcelas do mês
function carregarParcelasMes(mes) {
    const tbody = document.getElementById('tabelaDados');
    
    if (!dadosAgiotagem[mes] || dadosAgiotagem[mes].length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center;">Nenhuma parcela encontrada para este mês.</td></tr>';
        return;
    }
    
    let totalMensal = 0;

    const linhasTabela = dadosAgiotagem[mes].map((parcela, index) => {
        totalMensal += parcela.valorParcela;
        return `
            <tr>
                <td>${parcela.item}</td>
                <td>${parcela.parcelamento}</td>
                <td>R$ ${parcela.valorTotal.toFixed(2).replace('.', ',')}</td>
                <td>R$ ${parcela.valorParcela.toFixed(2).replace('.', ',')}</td>
                <td>
                    <span class="${parcela.pago ? 'status-pago' : 'status-pendente'}">
                        ${parcela.pago ? 'Pago' : 'Pendente'}
                    </span>
                </td>
                <td>
                    <button class="btn ${parcela.pago ? 'btn-danger' : 'btn-success'}" 
                            onclick="alterarStatusPagamento('${mes}', ${index})">
                        ${parcela.pago ? 'Marcar Pendente' : 'Marcar Pago'}
                    </button>
                </td>
                <td>
                    <button class="btn btn-danger" onclick="excluirItem('${mes}', ${index})">
                        Excluir
                    </button>
                </td>
            </tr>
        `;
    }).join('');
    
    const linhaTotal = `
        <tr class="total-row">
            <td colspan="3" style="text-align: right; font-weight: bold;">Total Acumulado:</td>
            <td colspan="4" style="font-weight: bold;">R$ ${totalMensal.toFixed(2).replace('.', ',')}</td>
        </tr>
    `;

    tbody.innerHTML = linhasTabela + linhaTotal;
}

// Alterar status de pagamento
function alterarStatusPagamento(mes, index) {
    if (dadosAgiotagem[mes] && dadosAgiotagem[mes].length > index) {
        dadosAgiotagem[mes][index].pago = !dadosAgiotagem[mes][index].pago;
        carregarParcelasMes(mesAtual);
        atualizarResumo();
        mostrarAlerta('Status de pagamento alterado!');
    }
}

// Excluir item
function excluirItem(mes, index) {
    if (confirm('Tem certeza que deseja excluir este item?')) {
        if (dadosAgiotagem[mes] && dadosAgiotagem[mes].length > index) {
            dadosAgiotagem[mes].splice(index, 1);
            if (dadosAgiotagem[mes].length === 0) {
                delete dadosAgiotagem[mes];
                const mesesDisponiveis = Object.keys(dadosAgiotagem);
                if (mesesDisponiveis.length > 0) {
                    mesAtual = mesesDisponiveis[0];
                } else {
                    mesAtual = getMesNome(new Date());
                }
                gerarBotoesMeses();
            }
            carregarParcelasMes(mesAtual);
            atualizarResumo();
            mostrarAlerta('Item excluído com sucesso!', 'success');
        }
    }
}

// Atualizar resumo
function atualizarResumo() {
    let totalItens = 0;
    let totalPendente = 0;
    let totalPago = 0;
    
    Object.values(dadosAgiotagem).forEach(parcelas => {
        parcelas.forEach(parcela => {
            totalItens++;
            if (parcela.pago) {
                totalPago += parcela.valorParcela;
            } else {
                totalPendente += parcela.valorParcela;
            }
        });
    });
    
    document.getElementById('totalItens').textContent = totalItens;
    document.getElementById('totalPendente').textContent = 'R$ ' + totalPendente.toFixed(2).replace('.', ',');
    document.getElementById('totalPago').textContent = 'R$ ' + totalPago.toFixed(2).replace('.', ',');
}

// Adicionar novo item
document.getElementById('addItemForm').addEventListener('submit', function(e) {
    e.preventDefault();
    
    const nome = document.getElementById('nomeItem').value;
    const valorTotal = parseFloat(document.getElementById('valorTotal').value);
    const parcelasTotal = parseInt(document.getElementById('parcelasTotal').value);
    const mesInicio = document.getElementById('mesInicio').value;
    
    const valorParcela = valorTotal / parcelasTotal;

    const [mesString, anoString] = mesInicio.split(' ');
    const mesIndex = meses.indexOf(mesString);
    const ano = parseInt(anoString);
    
    const dataInicio = new Date(ano, mesIndex, 1);
    
    for (let i = 0; i < parcelasTotal; i++) {
        const dataParcela = new Date(dataInicio.getFullYear(), dataInicio.getMonth() + i, 1);
        const mesChave = getMesNome(dataParcela);
        
        if (!dadosAgiotagem[mesChave]) {
            dadosAgiotagem[mesChave] = [];
        }
        
        dadosAgiotagem[mesChave].push({
            item: nome,
            parcelamento: `${i + 1}/${parcelasTotal}`,
            valorTotal: valorTotal,
            valorParcela: valorParcela,
            pago: false
        });
    }
    
    closeModal('addItemModal');
    document.getElementById('addItemForm').reset();
    atualizarResumo();
    gerarBotoesMeses();
    carregarParcelasMes(mesAtual);
    mostrarAlerta('Item adicionado com sucesso!');
});

// Exportar dados para Excel
function exportarDados() {
    const wb = XLSX.utils.book_new();
    const wsData = [];
    
    const mesesParaExportar = Object.keys(dadosAgiotagem).sort((a, b) => {
        const [mesA, anoA] = a.split(' ');
        const [mesB, anoB] = b.split(' ');
        const dataA = new Date(anoA, meses.indexOf(mesA), 1);
        const dataB = new Date(anoB, meses.indexOf(mesB), 1);
        return dataA - dataB;
    });
    
    mesesParaExportar.forEach(mes => {
        if (dadosAgiotagem[mes] && dadosAgiotagem[mes].length > 0) {
            wsData.push([mes, '', '', '', '']);
            wsData.push(['Item', 'Parcelamento', 'Valor Total (R$)', 'Valor da Parcela (R$)', 'Pago']);
            
            dadosAgiotagem[mes].forEach(parcela => {
                wsData.push([
                    parcela.item,
                    parcela.parcelamento,
                    parcela.valorTotal,
                    parcela.valorParcela,
                    parcela.pago ? 'TRUE' : 'FALSE'
                ]);
            });

            let totalMes = dadosAgiotagem[mes].reduce((acc, curr) => acc + curr.valorParcela, 0);
            wsData.push(['Total', '', '', totalMes, '']);
            wsData.push(['', '', '', '', '']);
        }
    });
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'AGIOTAGEM PARA O JEAN');
    
    XLSX.writeFile(wb, 'agiotagem_atualizada.xlsx');
    mostrarAlerta('Arquivo Excel exportado com sucesso!');
}

// Limpar dados
function limparDados() {
    if (confirm('Tem certeza que deseja limpar todos os dados?')) {
        dadosAgiotagem = {};
        document.getElementById('excelFile').value = '';
        mostrarAlerta('Dados limpos com sucesso!');
        gerarBotoesMeses();
        carregarParcelasMes(getMesNome(new Date()));
        atualizarResumo();
    }
}
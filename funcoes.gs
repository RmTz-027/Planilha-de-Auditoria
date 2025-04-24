function doGet(e) {
  const htmlTemplate = HtmlService.createTemplateFromFile('interface');
  SpreadsheetApp.flush();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BANCO DE DADOS');
  const data = sheet.getDataRange().getValues();
  htmlTemplate.dadosBanco = JSON.stringify(data);
  
  return htmlTemplate.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // permite embutir em iframes se quiser
    .setTitle("Auditoria de Estoque")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME); // ou .NATIVE dependendo das necessidades
}


function getEnderecos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BANCO DE DADOS');
  const range = sheet.getRange("E2:E" + sheet.getLastRow());
  const enderecos = range.getValues().flat().filter(String);
  const enderecosUnicos = [...new Set(enderecos)];
  return enderecosUnicos;
}

function getCFsByEndereco(enderecoSelecionado) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BANCO DE DADOS');
  const data = sheet.getRange("A2:E" + sheet.getLastRow()).getValues();

  const cfs = {};

  data.forEach(row => {
    const cf = row[1]; 
    const endereco = row[4]; 
    const nome = row[3]; 

    if (endereco === enderecoSelecionado && cf && nome) {
      if (!cfs[cf]) {
        cfs[cf] = { cf: cf, nome: nome, qtd: 0 };
      }
      cfs[cf].qtd++;
    }
  });

  return Object.values(cfs);
}

function buscarMercadoriaPorCodigo(codigo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BANCO DE DADOS');
  const data = sheet.getRange("A2:E" + sheet.getLastRow()).getValues();

  let mercadoria = data.find(row => row[0] === codigo);

  if (mercadoria) {
    return [{
      codigo: mercadoria[0],
      cf: mercadoria[1],
      final: mercadoria[2],
      descricao: mercadoria[3],
      endereco: mercadoria[4]
    }];
  } else {
    return [];
  }
}

function registrarCodigoNaBase(codigo, endereco) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASE DO SCRIPT');
  const lastRow = sheet.getLastRow();
  
  // Verifica se a aba tem um cabeçalho para o endereço
  if (sheet.getLastColumn() < 2 && lastRow > 0) {
    // Adiciona o cabeçalho para a coluna de endereço se não existir
    sheet.getRange(1, 2).setValue("ENDEREÇO");
  }

  // Verificar se o código já existe na planilha para evitar duplicações
  if (lastRow > 1) {
    const dadosExistentes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const codigoJaExiste = dadosExistentes.some(row => String(row[0]) === String(codigo));
    
    if (codigoJaExiste) {
      Logger.log("Código já existe na BASE DO SCRIPT, ignorando duplicação: " + codigo);
      return false; // Informa que o código já existe e não foi adicionado novamente
    }
  }

  // Salva o código e o endereço apenas se não existir
  const newRow = lastRow + 1;
  sheet.getRange(newRow, 1, 1, 2).setValues([[codigo, endereco || ""]]);
  
  Logger.log("Registrado código na base: " + codigo + " - Endereço: " + (endereco || "não informado"));
  return true;
}

function contarCFsNaBase(enderecoSelecionado) {
  const baseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASE DO SCRIPT');
  const bancoDadosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BANCO DE DADOS');

  // Verifica se a planilha BASE DO SCRIPT tem uma coluna de endereço
  const temColunaEndereco = baseSheet.getLastColumn() >= 2;
  
  // Obtém os dados de BASE DO SCRIPT
  let dadosBase;
  if (temColunaEndereco) {
    // Se tiver coluna de endereço, pega código e endereço
    dadosBase = baseSheet.getRange(2, 1, baseSheet.getLastRow() - 1, 2).getValues();
  } else {
    // Se não tiver, pega só o código
    dadosBase = baseSheet.getRange(2, 1, baseSheet.getLastRow() - 1, 1).getValues();
  }

  // Obtém os dados do banco
  const dadosBanco = bancoDadosSheet.getRange("A2:E" + bancoDadosSheet.getLastRow()).getValues();

  let contagem = {};

  // Processa cada código na BASE DO SCRIPT
  dadosBase.forEach(row => {
    const codigo = row[0];
    const enderecoItem = temColunaEndereco ? row[1] : null;
    
    // Se temos endereço na BASE DO SCRIPT e um endereço selecionado foi informado,
    // verificamos se o endereço do item corresponde ao endereço selecionado
    if (temColunaEndereco && enderecoSelecionado && enderecoItem !== enderecoSelecionado) {
      return; // Pula este item se o endereço não corresponder
    }

    // Busca o item no banco de dados
    let linhaBanco = dadosBanco.find(row => row[0] === codigo);
    if (linhaBanco) {
      let cf = linhaBanco[1];
      
      // Se temos endereço no banco, verificamos também
      if (enderecoSelecionado && linhaBanco[4] !== enderecoSelecionado) {
        return; // Pula este item se o endereço no banco não corresponder
      }
      
      contagem[cf] = (contagem[cf] || 0) + 1;
    }
  });

  return contagem;
}

function salvarCodigosInseridos(dados) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASE DO SCRIPT');
  const lastRow = sheet.getLastRow();
  
  // Verifica se a aba tem cabeçalhos
  if (lastRow === 0) {
    // Adiciona cabeçalhos se a planilha estiver vazia
    sheet.getRange(1, 1, 1, 2).setValues([["CÓDIGO", "ENDEREÇO"]]);
  }
  
  // Obter os códigos existentes para verificar duplicações
  const codigosExistentes = lastRow > 1 ? 
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => String(row[0]).trim()) : [];
  
  // Normaliza os códigos novos e filtra duplicados
  const dadosNovos = dados.filter(item => {
    const codigoNormalizado = String(item.codigo).trim();
    return !codigosExistentes.includes(codigoNormalizado);
  });
  
  if (dadosNovos.length === 0) {
    return "Todos os códigos já existem na BASE DO SCRIPT. Nenhum novo código foi adicionado.";
  }
  
  // Preparar os dados para inserção (usando códigos normalizados)
  const valoresParaSalvar = dadosNovos.map(item => [
    String(item.codigo).trim(),
    item.endereco || ""
  ]);
  
  // Salvar os dados filtrados
  sheet.getRange(lastRow + 1, 1, valoresParaSalvar.length, 2).setValues(valoresParaSalvar);
  
  return `Dados salvos com sucesso! ${dadosNovos.length} novos códigos adicionados. ${dados.length - dadosNovos.length} códigos eram duplicados e foram ignorados.`;
}

function adicionarNovoRegistro(novoRegistro) {
  try {
    Logger.log("Início da função adicionarNovoRegistro - Dados recebidos: " + JSON.stringify(novoRegistro));
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BANCO DE DADOS');
    
    if (!sheet) {
      Logger.log("Erro: Aba 'BANCO DE DADOS' não encontrada!");
      throw new Error("Aba 'BANCO DE DADOS' não encontrada!");
    }
    
    // Adiciona a nova linha
    sheet.appendRow([
      novoRegistro.codigo,
      novoRegistro.cf,
      novoRegistro.final,
      novoRegistro.descricao,
      novoRegistro.endereco
    ]);
    
    // Força a atualização do cache antes de ler os dados
    SpreadsheetApp.flush();
    
    // Obtém todos os dados atualizados
    const novosDados = sheet.getDataRange().getValues();
    
    Logger.log("Dados atualizados obtidos, retornando " + novosDados.length + " linhas");
    
    // Verifica se o item foi adicionado corretamente
    const itemAdicionado = novosDados.some(row => 
      row[0] === novoRegistro.codigo && 
      row[1] === novoRegistro.cf && 
      row[4] === novoRegistro.endereco
    );
    
    if (!itemAdicionado) {
      Logger.log("AVISO: O item adicionado não foi encontrado nos dados retornados!");
    } else {
      Logger.log("Item adicionado com sucesso e encontrado nos dados retornados.");
    }
    
    return JSON.stringify(novosDados);
    
  } catch (erro) {
    Logger.log("ERRO em adicionarNovoRegistro: " + erro.message);
    throw erro;
  }
}

function verificarAuditoriaPorEndereco(enderecoSelecionado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bancoDados = ss.getSheetByName('BANCO DE DADOS');
  const dadosBanco = bancoDados.getRange(2, 1, bancoDados.getLastRow() - 1, 5).getValues(); 
  // Colunas: 0-Código, 1-CF, 2-Final, 3-Descrição, 4-Endereço

  const auditoria = ss.getSheetByName('AUDITORIA');
  const dadosAuditoria = auditoria.getRange(2, 1, auditoria.getLastRow() - 1, 1).getValues().flat(); 
  // Coluna A: Código Produto auditado

  const resultado = dadosBanco
    .filter(linha => linha[4] === enderecoSelecionado) // Filtra pelo endereço
    .map(linha => ({
      codigo: linha[0],
      cf: linha[1],
      final: linha[2],
      descricao: linha[3],
      status: dadosAuditoria.includes(linha[0]) ? 'Auditado' : 'Não Auditado'
    }));

  return resultado;
}

function getListas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const divergenciaSheet = ss.getSheetByName('DIVERGENCIA');
  
  if (!divergenciaSheet) {
    return ["Aba Divergencia não encontrada"];
  }
  
  // Supondo que as listas estejam na coluna A
  const lastRow = divergenciaSheet.getLastRow();
  
  if (lastRow < 2) {
    return ["Sem listas disponíveis"];
  }
  
  const listas = divergenciaSheet.getRange(2, 1, lastRow - 1).getValues();
  
  // Filtra valores vazios e remove duplicatas
  const listasUnicas = [...new Set(listas.flat().filter(lista => lista))];
  
  return listasUnicas;
}

function SalvarDivergencia(dados) {
  try {
    // 1. Acesse a outra planilha pelo ID dela
    const outraPlanilhaID = '1WGm_PdDJUb9Jf7YABuotMdU7XWv9ckyH4BgWzoxi2ws'; // <- Pegue no link da planilha
    const outraPlanilha = SpreadsheetApp.openById(outraPlanilhaID);

    // 2. Selecione a aba onde vai gravar
    const abaDestino = outraPlanilha.getSheetByName('2025'); // ex: 'Divergencias Recebidas'

    if (!abaDestino) {
      throw new Error('Aba de destino não encontrada!');
    }

    // 3. Prepare a linha que vai adicionar
    const novaLinha = [
      
      dados.Data,
      dados.CF,
      dados.Final,     
      dados.CdMercadoria,
      dados.Descricao,
      dados.Qtd,      
      dados.EndeAud,
       dados.Lista
    ];

    // 4. Adiciona a linha na próxima linha vazia
    abaDestino.appendRow(novaLinha);

    return 'Divergência registrada com sucesso na outra planilha!';
  } catch (erro) {
    Logger.log('Erro ao salvar divergência: ' + erro);
    throw erro; // devolve o erro para o front
  }
}

function carregarEnderecos() {
  google.script.run
    .withSuccessHandler(function(enderecos) {
      const dropdown = document.getElementById("enderecoDropdown");
      dropdown.innerHTML = '<option value="">Selecione um endereço</option>';
      
      if (Array.isArray(enderecos) && enderecos.length > 0) {
        enderecos.forEach(endereco => {
          const option = document.createElement("option");
          option.value = endereco;
          option.textContent = endereco;
          dropdown.appendChild(option);
        });
      } else {
        const option = document.createElement("option");
        option.disabled = true;
        option.textContent = "Nenhum endereço disponível";
        dropdown.appendChild(option);
      }
    })
    .withFailureHandler(function(erro) {
      console.error("Erro ao carregar endereços:", erro);
      const dropdown = document.getElementById("enderecoDropdown");
      dropdown.innerHTML = '<option value="">Erro ao carregar endereços</option>';
    })
    .getEnderecos();
}

function carregarEnderecosVerificacao() {
  google.script.run
    .withSuccessHandler(function(enderecos) {
      const dropdown = document.getElementById("enderecoVerificacaoDropdown");
      dropdown.innerHTML = '<option value="">Selecione um endereço</option>';
      
      enderecos.forEach(endereco => {
        const option = document.createElement("option");
        option.value = endereco;
        option.textContent = endereco;
        dropdown.appendChild(option);
      });
    })
    .getEnderecos();
}

function getCodigosAuditados() {
  try {
    Logger.log("Obtendo códigos auditados da aba BASE DO SCRIPT");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASE DO SCRIPT');
    
    if (!sheet) {
      Logger.log("Aba BASE DO SCRIPT não encontrada");
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("Nenhum código auditado encontrado (sheet vazia)");
      return [];
    }
    
    // Obtém todos os códigos da coluna A, começando da linha 2
    const codigosRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const codigos = codigosRange.getValues().flat().filter(codigo => codigo); // Remove vazios
    
    Logger.log("Recuperados " + codigos.length + " códigos auditados");
    return codigos;
  } catch (erro) {
    Logger.log("Erro ao obter códigos auditados: " + erro.message);
    return [];
  }
}

function salvarNovoRegistro() {
  try {
    const codigo = document.getElementById("novoCodigo").value.trim();
    const cf = document.getElementById("novoCF").value.trim();
    const final = document.getElementById("novoFinal").value.trim();
    const descricao = document.getElementById("novoDescricao").value.trim();
    const endereco = document.getElementById("novoEndereco").value.trim();

    if (!codigo || !cf || !final || !descricao || !endereco) {
      alert("Preencha todos os campos.");
      return;
    }

    const novoRegistro = { codigo, cf, final, descricao, endereco };

    const btnSalvar = document.getElementById("btnSalvarNovo");
    btnSalvar.disabled = true;
    btnSalvar.textContent = "Salvando...";

    google.script.run
      .withSuccessHandler((novosDados) => {
        try {
          console.log("Dados recebidos do servidor:", novosDados);
          
          // Atualiza o banco de dados local
          window.dadosBanco = JSON.parse(novosDados);
          console.log("Banco de dados atualizado:", window.dadosBanco);
          
          // Limpa a lista de códigos inseridos
          codigosInseridos = [];
          document.getElementById("mercadoriasLista").innerHTML = "";
          
          // Fecha o modal
          fecharModal();
          
          // Limpa o campo de código
          document.getElementById("codigo").value = "";
          
          // Recarrega as interfaces
          carregarEnderecos();
          carregarCFs();
          
          // Força uma nova busca do código recém-adicionado
          const codigoInput = document.getElementById("codigo");
          codigoInput.value = codigo;
          buscarCodigo({ key: "Enter" });
          
          alert("Registro adicionado com sucesso!");
          
        } catch (erro) {
          console.error("Erro ao processar resposta:", erro);
          alert("Erro ao processar resposta: " + erro.message);
        } finally {
          btnSalvar.disabled = false;
          btnSalvar.textContent = "Salvar";
        }
      })
      .withFailureHandler((erro) => {
        alert("Erro ao adicionar registro: " + erro.message);
        btnSalvar.disabled = false;
        btnSalvar.textContent = "Salvar";
      })
      .adicionarNovoRegistro(novoRegistro);
  } catch (erro) {
    alert("Erro inesperado: " + erro.message);
  }
}

function buscarCodigo(event) {
  if (event.key !== "Enter") return;

  const codigoInput = document.getElementById("codigo");
  const codigo = codigoInput.value.trim();

  if (!codigo || !enderecoAtual) {
    alert("Digite o código e selecione um endereço.");
    return;
  }

  if (codigosInseridos.some(item => item.codigo === codigo)) {
    alert("Código já inserido.");
    return;
  }

  // Certifique-se de que dadosBanco está definido
  if (!window.dadosBanco) {
    alert("Erro: Banco de dados não carregado corretamente.");
    return;
  }

  console.log("Procurando código:", codigo);
  console.log("Banco de dados atual:", window.dadosBanco);

  // Procura a mercadoria no banco de dados atualizado
  const mercadoria = window.dadosBanco.find(row => String(row[0]) === String(codigo));
  
  if (mercadoria) {
    console.log("Mercadoria encontrada:", mercadoria);
  } else {
    console.log("Mercadoria não encontrada para o código:", codigo);
  }

  const tabela = document.getElementById("mercadoriasLista");
  const row = document.createElement("tr");

  if (mercadoria) {
    const [cod, cf, final, descricao, endereco] = mercadoria;

    row.innerHTML = `
      <td>${cod}</td>
      <td>${cf}</td>
      <td>${final}</td>
      <td>${endereco}</td>
      <td><button class="remove-btn" onclick="removerItem('${cod}', this)">Remover</button></td>
    `;

    // Verifica se o endereço da mercadoria é diferente do endereço selecionado
    const enderecoErrado = (endereco !== enderecoAtual);
    
    if (enderecoErrado) {
      row.classList.add("linha-vermelha");
      row.title = "Este item está em um endereço diferente do selecionado";
    }

    // Adiciona à lista de códigos inseridos
    codigosInseridos.unshift({ 
      codigo: cod, 
      cf: cf, 
      enderecoErrado: enderecoErrado 
    });
  } else {
    row.innerHTML = `
      <td>${codigo}</td>
      <td>Fora do Banco de Dados</td>
      <td>-</td>
      <td>-</td>
      <td>
        <button class="remove-btn" onclick="removerItem('${codigo}', this)">Remover</button>
        <button onclick="abrirModalAdicionar('${codigo}')">Adicionar</button>
      </td>
    `;
    row.classList.add("linha-vermelha");
    row.classList.add("fora-banco");
    row.title = "Este item não foi encontrado no banco de dados";

    codigosInseridos.unshift({ 
      codigo: codigo, 
      cf: "Fora do Banco de Dados", 
      enderecoErrado: true 
    });
  }

  tabela.prepend(row);
  carregarCFs();

  codigoInput.value = "";
  codigoInput.focus();
}


function calcularRendimento() {
    var aba = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1oA7-jZs7KUPDMixt9FLECo-rBQoJ1DaT0Kzqj5B7sQA/edit#gid=0').getSheetByName('Lista_Ativos');
    var numeroAtivos = aba.getRange(2, 13).getValue();
    var array_btc_possuido = [];
    var array_btc_custo = [];
    for (let i = 2; i <= numeroAtivos + 1; i++) {
      Logger.log(i);
      var valor_aplicado = aba.getRange(i,10).getValue();
      if (aba.getRange(i, 5).getValue() == 'Pré') {
      var intervalo_pre_rentabilidade_aa = aba.getRange(i,6);
      var pre_rentabilidade_aa = intervalo_pre_rentabilidade_aa.getValue();
      var pre_rentabilidade_ad = (pre_rentabilidade_aa + 1) ** (1/365) - 1 // converte a rentabilidade anual em diária considerando 365 dias corridos no ano
      var dias_corridos = aba.getRange(i,9).getValue();
      var pre_rentabilidade_acumulada = (1 + pre_rentabilidade_ad) ** (dias_corridos);
      var valor_atual = pre_rentabilidade_acumulada * valor_aplicado;
      aba.getRange(i,11).setValue(valor_atual);
      }
      else if (aba.getRange(i, 5).getValue() == 'CDI%') {
        const codigo_serie = 11 // https://dadosabertos.bcb.gov.br/dataset/11-taxa-de-juros---selic/resource/b73edc07-bbac-430c-a2cb-b1639e605fa8
        dataInicial = aba.getRange(i,7).getValue();
        dataFinal = aba.getRange(i,8).getValue();
        dataInicial = Moment.moment(dataInicial).format('DD-MM-YYYY'); // https://stackoverflow.com/questions/22410210/how-do-i-use-momentsjs-in-google-apps-script
        dataFinal = Moment.moment(dataFinal).format('DD-MM-YYYY');  // Podem ser encontrados guias de como usar biblioteca moment em https://momentjs.com/
        endereco = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.' + codigo_serie + '/dados?formato=json&dataInicial=' + dataInicial + '&dataFinal=' + dataFinal;
        var resposta = UrlFetchApp.fetch(endereco);
        var respostaJson = resposta.getContentText();
        var dadosJson = JSON.parse(respostaJson); // converte o objeto JSON em objeto JS
        var pos_rentabilidade_aa = aba.getRange(i, 6).getValue();
        var lista_rentabilidade_ad = [];
        for (let k = 0; k <= dadosJson.length - 1; k++) {
          lista_rentabilidade_ad.push((dadosJson[k].valor * 0.01) * pos_rentabilidade_aa);  // acrescenta à lista o valor da rentabilidade da selic diária em decimal multiplicada pela rentabilidade
          // (em termos de %CDI) do CDB
        }
        var dias_uteis = dadosJson.length
        var pos_rentabilidade_acumulada = 1
        for (let l = 0; l <= dias_uteis - 1; l++) {
          pos_rentabilidade_acumulada *= 1 + lista_rentabilidade_ad[l];  // calcula a rentabilidade acumulada realizando o produtório das rentabilidades diárias
        }
        var valor_atual = valor_aplicado * pos_rentabilidade_acumulada;
        aba.getRange(i, 11).setValue(valor_atual);
      
      }
      else if (aba.getRange(i,5).getValue() == 'CDI+') {
        const codigo_serie = 11 // https://dadosabertos.bcb.gov.br/dataset/11-taxa-de-juros---selic/resource/b73edc07-bbac-430c-a2cb-b1639e605fa8
          dataInicial = aba.getRange(i,7).getValue();
          dataFinal = aba.getRange(i,8).getValue();
          dataInicial = Moment.moment(dataInicial).format('DD-MM-YYYY'); // https://stackoverflow.com/questions/22410210/how-do-i-use-momentsjs-in-google-apps-script
          dataFinal = Moment.moment(dataFinal).format('DD-MM-YYYY');  // Podem ser encontrados guias de como usar biblioteca moment em https://momentjs.com/
          endereco = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.' + codigo_serie + '/dados?formato=json&dataInicial=' + dataInicial + '&dataFinal=' + dataFinal;
          var resposta = UrlFetchApp.fetch(endereco);
          var respostaJson = resposta.getContentText();
          var dadosJson = JSON.parse(respostaJson); // converte o objeto JSON em objeto JS
          var pos_rentabilidade_aa = aba.getRange(i, 6).getValue();  // Valor de taxa prefixada anual somada ao CDI
          var pos_rentabilidade_ad = (pos_rentabilidade_aa + 1) ** (1/252) - 1  // Conversão de taxa anual em diária considerando-se 252 dias úteis no ano
          var lista_rentabilidade_ad = [];
          for (let k = 0; k <= dadosJson.length - 1; k++) {
            lista_rentabilidade_ad.push((dadosJson[k].valor * 0.01) + pos_rentabilidade_ad);  // acrescenta à lista o valor da rentabilidade da selic diária em decimal acrescida da rentabilidade
            // (em termos de CDI + taxa prefixada) do CDB
          }
          var dias_uteis = dadosJson.length
          var pos_rentabilidade_acumulada = 1
          for (let l = 0; l <= dias_uteis - 1; l++) {
            pos_rentabilidade_acumulada *= 1 + lista_rentabilidade_ad[l];  // calcula a rentabilidade acumulada realizando o produtório das rentabilidades diárias
          }
          var valor_atual = valor_aplicado * pos_rentabilidade_acumulada;
          aba.getRange(i,11).setValue(valor_atual);  
      }
      else if (aba.getRange(i, 2).getValue() == 'BTC') {
        var btc_possuido = aba.getRange(i, 10).getValue();
        array_btc_possuido.push(btc_possuido);
        var btc_cotacao = aba.getRange(10, 12).getValue();
        var valor_atual = btc_possuido * btc_cotacao;
        aba.getRange(i, 11).setValue(valor_atual);
        var btc_custo = aba.getRange(i, 6).getValue();
        array_btc_custo.push(btc_custo);
      }
    }
     var array_weighted_btc_custo = [];
     for (let x = 0; x <= array_btc_custo.length - 1; x++) {
      var weighted_btc_custo = array_btc_custo[x] * array_btc_possuido[x];
      array_weighted_btc_custo.push(weighted_btc_custo);
    }
    var sum_weighted_btc_custo = 0;
      for (let x = 0; x <= array_weighted_btc_custo.length - 1; x++) {
        sum_weighted_btc_custo += array_weighted_btc_custo[x];
      }
    var sum_btc_possuido = 0;
      for (let x = 0; x <= array_btc_possuido.length - 1; x++) {
        sum_btc_possuido += array_btc_possuido[x];
      }
    var pmc = sum_weighted_btc_custo / sum_btc_possuido;
    aba.getRange(12, 12).setValue(pmc);
  }
  
  /*
  function calcularRentabilidade() {
  var array_valores_CDI_percentual = [];  // Array que armazena os valores investidos em CDB pós-fixados remunerados por CDI%
   var aba = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1oA7-jZs7KUPDMixt9FLECo-rBQoJ1DaT0Kzqj5B7sQA/edit#gid=0').getSheetByName('Lista_Ativos');
    
      var diaAtual = Moment.moment().date();  // dia do mês de hoje
      var dataAtual = Moment.moment().add(-diaAtual + 1, 'day').format('YYYY-MM-DD')  // data no formato 'YYYY-MM-01' para que a comparação possa ser feita
      var dataInicial = '10-01-2023';  // data do início da simulação. Formato de data 'DD-MM-YYYY'
      var dataInicial = Moment.moment(dataInicial).format('YYYY-MM-DD'); // Notação em 'YYYY-MM-DD' para que a comparação possa ser feita
    
      const arrayDatas = [];
      while (dataInicial <= dataFinal || (Moment.moment(dataAtual).year() == Moment.moment(dataFinal).year() && Moment.moment(dataAtual).month() == Moment.moment(dataFinal).month())) {
        arrayDatas.push(Moment.moment(dataAtual).format('YYYY-MM-DD'));
        
        // Incrementa a data para o próximo mês
        dataAtual = Moment.moment(dataAtual).add(1, 'month').format('YYYY-MM-DD');
      }
    
      const arrayValores = [];
    
      for (let z = 0; z < arrayDatas.length; z++) { // loop no número de meses z
        Logger.log(z);
        var variacaoTotalInvestido = 0;
        var numeroAtivos = aba.getRange(2, 13).getValue();
        for (let i = 2; i <= numeroAtivos + 1; i++) {  // loop no número de ativos
        if (aba.getRange(i, 5).getValue() == 'Pré') {  // Títulos pré-fixados
  
        var intervalo_pre_rentabilidade_aa = aba.getRange(i,6);
        var pre_rentabilidade_aa = intervalo_pre_rentabilidade_aa.getValue();
        var pre_rentabilidade_ad = (pre_rentabilidade_aa + 1) ** (1/365) - 1 // converte a rentabilidade anual em diária, considerando dias corridos
        
        var data_aplicacao = aba.getRange(i,7).getValue();  // Data em que o valor foi aplicado
        var data_aplicacao = Moment.moment(data_aplicacao).format('YYYY-MM-DD');
        var dia_final_mes = Moment.moment(arrayDatas[z]).endOf('month').format('YYYY-MM-DD');
        var dias_dif = dia_final_mes.diff(data_aplicacao, 'days');  // Retorna o número de dias corridos entre o último dia do z-ésimo mês (mês da iteração atual) e o dia da aplicação
  
        if (data_aplicacao >= dataInicial) {
          if (Moment.moment(data_aplicacao).month() == Moment.moment(arrayDatas[z]).month() && Moment.moment(data_aplicacao).year() == Moment.moment(arrayDatas[z]).year()) {  // verifica se a data de aplicação do CDB é no mesmo mês e ano que a data z-ésimo mês (mês da iteração atual)
            var valor_investido = aba.getRange(i, 10).getValue();
            var valor_atual = valor_investido * (1 + pre_rentabilidade_ad) ** (dias_dif)
  
          }
        }
        var valor_investido = aba.getRange(i,10).getValue();
        var valor_atual = valor_investido * (1 + pre_rentabilidade_am) ** (z + 1);  // valor atual do título no dia de hoje atualizado pelos juros do loop em z
  
        var variacaoTotalInvestidoPre = valor_atual - valor_atual / (1 + pre_rentabilidade_am);
        variacaoTotalInvestido += variacaoTotalInvestidoPre; // ganho de juros ao longo do tempo
        }
        
  
        else if (aba.getRange(i, 5).getValue() == 'CDI%') {  // Títulos Pós-fixados que pagam % do CDI
          
          var dataSelicMesAtual = arrayDatas[z];
          const codigo_serie = 11 // https://dadosabertos.bcb.gov.br/dataset/11-taxa-de-juros---selic/resource/b73edc07-bbac-430c-a2cb-b1639e605fa8
          endereco = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.' + codigo_serie + '/dados?formato=json&dataInicial=' + dataSelicMesAtual + '&dataFinal=' + dataSelicMesAtual;
          var resposta = UrlFetchApp.fetch(endereco);
          var respostaJson = resposta.getContentText();
          var dadosJson = JSON.parse(respostaJson); // converte o objeto JSON em objeto JS
  
          var pos_rentabilidade_aa = aba.getRange(i, 6).getValue(); // % CDI do CDB
          var pos_rentabilidade_am = (dadosJson * 0.01 * pos_rentabilidade_aa + 1) ** 30 - 1; // converte a rentabilidade anual em mensal
  
          if (z == 0) {
            var valor_atual = aba.getRange(i,11).getValue() * (1 + pos_rentabilidade_am);  // valor atual do título no dia de hoje atualizado pelos juros do loop em z
            array_valores_CDI_percentual.push(valor_atual);
            var variacaoTotalInvestidoPos = valor_atual - valor_atual / (1 + pre_rentabilidade_am);
          }
          
          else if (z > 0) {
            var valor_atual = array_valores_CDI_percentual[z - 1] * (1 + pos_rentabilidade_am);
            array_valores_CDI_percentual.push(valor_atual);
            var variacaoTotalInvestidoPos = array_valores_CDI_percentual[z] - array_valores_CDI_percentual[z - 1];
          }
  
          variacaoTotalInvestido += variacaoTotalInvestidoPos; // ganho de juros ao longo do tempo
        }
  
        else if (aba.getERange(i, 5).getValue() == 'CDI+') {  // Títulos Pós-fixados que pagam CDI + taxa fixa
          var dataSelicMesAtual = arrayDatas[z];
          const codigo_serie = 11 // https://dadosabertos.bcb.gov.br/dataset/11-taxa-de-juros---selic/resource/b73edc07-bbac-430c-a2cb-b1639e605fa8
          endereco = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.' + codigo_serie + '/dados?formato=json&dataInicial=' + dataSelicMesAtual + '&dataFinal=' + dataSelicMesAtual;
          var resposta = UrlFetchApp.fetch(endereco);
          var respostaJson = resposta.getContentText();
          var dadosJson = JSON.parse(respostaJson); // converte o objeto JSON em objeto JS
  
  
          var pos_rentabilidade_aa = aba.getRange(i, 6).getValue(); // taxa fixa somada ao CDI do CDB
          var pos_rentabilidade_am = (dadosJson * 0.01 * pos_rentabilidade_aa + 1) ** 30 - 1 + ((pos_rentabilidade_aa + 1) ** (1 / 12) - 1); // converte as rentabilidades anual e diária em mensal
  
          if (z == 0) {
            var valor_atual = aba.getRange(i,11).getValue() * (1 + pos_rentabilidade_am);  // valor atual do título no dia de hoje atualizado pelos juros do loop em z
            array_valores_CDI_percentual.push(valor_atual);
            var variacaoTotalInvestidoPos = valor_atual - valor_atual / (1 + pre_rentabilidade_am);
          }
          
          else if (z > 0) {
            var valor_atual = array_valores_CDI_percentual[z - 1] * (1 + pos_rentabilidade_am);
            array_valores_CDI_percentual.push(valor_atual);
            var variacaoTotalInvestidoPos = array_valores_CDI_percentual[z] - array_valores_CDI_percentual[z - 1];
          }
  
          variacaoTotalInvestido += variacaoTotalInvestidoPos; // ganho de juros ao longo do tempo
  
        }
  
        else if (aba.getRange(i, 2).getValue() == 'BTC') { // Compras de Bitcoin
          
  
        }
      }
      if (z == 0) {
        arrayValores.push(parseFloat(totalInvestido).toFixed(2));
      } else if (z > 0) {
        arrayValores.push((parseFloat(arrayValores[z - 1]) + parseFloat(variacaoTotalInvestido)).toFixed(2));
      }
      
      Logger.log(arrayDatas[z]);
      Logger.log(arrayValores[z]);
    }
      // Verifica se os arrays têm o mesmo comprimento
      Logger.log(arrayDatas.length);
      Logger.log(arrayDatas);
      Logger.log(arrayValores.length);
      Logger.log(arrayValores);
  
      var abaFGC = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1oA7-jZs7KUPDMixt9FLECo-rBQoJ1DaT0Kzqj5B7sQA/edit#gid=0').getSheetByName('FGC');
  
      for (let m = 0; m < arrayDatas.length; m++) {
        abaFGC.getRange(m+2,1).setValue(arrayDatas[m]);
        abaFGC.getRange(m+2,2).setValue(parseFloat(arrayValores[m]));
      }
  }
  */
  
  function alterarGrafico() {
    var aba = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1oA7-jZs7KUPDMixt9FLECo-rBQoJ1DaT0Kzqj5B7sQA/edit#gid=0').getSheetByName('Lista_Ativos');
  
    // Capturar gráfico em tela na planilha
    var graficos = aba.getCharts()[0];  // seleciona o gráfico em tela
  
    // Define o título do gráfico
    var totalInvestido = aba.getRange(2,12).getValue().toFixed(2);
    var totalInvestidoFMT = Intl.NumberFormat('pt-BR').format(totalInvestido); // O método Intl.NumberFormat formata o valor de moeda para o local desejado.
    var novoTitulo = 'Total investido: R$' + totalInvestidoFMT;
  
    var grafico = graficos.modify().setOption('title', novoTitulo).build();
    aba.updateChart(grafico);
  }
  
  function coberturaFGC() {
      var aba = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1oA7-jZs7KUPDMixt9FLECo-rBQoJ1DaT0Kzqj5B7sQA/edit#gid=0').getSheetByName('Lista_Ativos');
      var totalInvestido = aba.getRange(8,12).getValue().toFixed(2);  // Total Investido no Banco Master
      const juroFuturo = 0.1040  // estimativa para jan-2032 baseada no contrato futuro DI1F32, disponível na página https://www.infomoney.com.br/ferramentas/juros-futuros-di/, para transação efetuada em 12/01/2024 // 05:52
    
      var diaAtual = Moment.moment().date();  // dia do mês de hoje
      var dataAtual = Moment.moment().add(-diaAtual + 1, 'day').format('YYYY-MM-DD')  // data no formato 'YYYY-MM-01' para que a comparação possa ser feita
      var dataFinal = '08-06-2030'  // data do final da simulação (data do título de vencimento mais longo do Banco Master).
      var dataFinal = Moment.moment(dataFinal).format('YYYY-MM-DD');  // Notação em 'YYYY-MM-DD' para que a comparação possa ser feita
    
      const arrayDatas = [];
      while (dataAtual <= dataFinal || (Moment.moment(dataAtual).year() == Moment.moment(dataFinal).year() && Moment.moment(dataAtual).month() == Moment.moment(dataFinal).month())) {
        arrayDatas.push(Moment.moment(dataAtual).format('YYYY-MM-DD'));
        
        // Incrementa a data para o próximo mês
        dataAtual = Moment.moment(dataAtual).add(1, 'month').format('YYYY-MM-DD');
      }
    
      const arrayValores = [];
    
      for (let z = 0; z < arrayDatas.length; z++) { // loop no número de meses z
        var variacaoTotalInvestido = 0;
        var numeroAtivos = aba.getRange(2, 13).getValue();
        for (let i = 2; i <= numeroAtivos + 1; i++) {
        if (aba.getRange(i, 5).getValue() == 'Pré' && aba.getRange(i,4).getValue() == 'Banco Master') { // Título Prefixado do Banco Master
        
        var intervalo_pre_rentabilidade_aa = aba.getRange(i,6);
        var pre_rentabilidade_aa = intervalo_pre_rentabilidade_aa.getValue();
        var pre_rentabilidade_am = (pre_rentabilidade_aa + 1) ** (1/12) - 1 // converte a rentabilidade anual em mensal
        
        var valor_atual = aba.getRange(i,11).getValue() * (1 + pre_rentabilidade_am) ** (z + 1) ;  // valor atual do título no dia de hoje atualizado pelos juros do loop em z
  
        var dataVencimento = aba.getRange(i,8).getValue();
        dataVencimento = Moment.moment(dataVencimento).format('YYYY-MM-DD'); // https://stackoverflow.com/questions/22410210/how-do-i-use-momentsjs-in-google-apps-script
    
        if (Moment.moment(arrayDatas[z]) < Moment.moment(dataVencimento).add(-1, 'month')) {
          var variacaoTotalInvestidoPre = valor_atual - valor_atual / (1 + pre_rentabilidade_am);
          variacaoTotalInvestido += variacaoTotalInvestidoPre; // ganho de juros ao longo do tempo
        } else if (Moment.moment(arrayDatas[z]).month() == Moment.moment(dataVencimento).month() && Moment.moment(arrayDatas[z]).year() == Moment.moment(dataVencimento).year()) {
          var variacaoTotalInvestidoPre = -valor_atual * (1 + pre_rentabilidade_am);  // saque do ativo no mês do vencimento
          variacaoTotalInvestido += variacaoTotalInvestidoPre;
        }
        }
        else if (aba.getRange(i, 5).getValue() == 'CDI%' && aba.getRange(i,4).getValue() == 'Banco Master') {  // Título Posfixado do Banco Master
          
          var pre_rentabilidade_aa = juroFuturo;
          var pos_rentabilidade_aa = aba.getRange(i, 6).getValue(); // % CDI do CDB
          var pre_rentabilidade_am = (pre_rentabilidade_aa * pos_rentabilidade_aa + 1) ** (1/12) - 1 // converte a rentabilidade anual em mensal
  
          var valor_atual = aba.getRange(i,11).getValue() * (1 + pre_rentabilidade_am) ** z ;  // valor atual do título no dia de hoje atualizado pelos juros do loop em z
  
          var dataVencimento = aba.getRange(i,8).getValue();
          dataVencimento = Moment.moment(dataVencimento).format('YYYY-MM-DD'); // https://stackoverflow.com/questions/22410210/how-do-i-use-momentsjs-in-google-apps-script
          
        if (Moment.moment(arrayDatas[z]) < Moment.moment(dataVencimento).add(-1, 'month')) {
          var variacaoTotalInvestidoPos = valor_atual - valor_atual / (1 + pre_rentabilidade_am);
          variacaoTotalInvestido += variacaoTotalInvestidoPos; // ganho de juros ao longo do tempo
        } else if (Moment.moment(arrayDatas[z]).month() == Moment.moment(dataVencimento).month() && Moment.moment(arrayDatas[z]).year() == Moment.moment(dataVencimento).year()) {
          var variacaoTotalInvestido = -valor_atual * (1 + pre_rentabilidade_am);  // saque do ativo no mês do vencimento
          variacaoTotalInvestido += variacaoTotalInvestidoPos;
        }
        }
      }
      if (z == 0) {
        arrayValores.push(parseFloat(totalInvestido).toFixed(2));
      } else if (z > 0) {
        arrayValores.push((parseFloat(arrayValores[z - 1]) + parseFloat(variacaoTotalInvestido)).toFixed(2));
      }
      
      Logger.log(arrayDatas[z]);
      Logger.log(arrayValores[z]);
    }
      // Verifica se os arrays têm o mesmo comprimento
      Logger.log(arrayDatas.length);
      Logger.log(arrayDatas);
      Logger.log(arrayValores.length);
      Logger.log(arrayValores);
  
      var abaFGC = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1oA7-jZs7KUPDMixt9FLECo-rBQoJ1DaT0Kzqj5B7sQA/edit#gid=0').getSheetByName('FGC');
  
      for (let m = 0; m < arrayDatas.length; m++) {
        abaFGC.getRange(m+2,1).setValue(arrayDatas[m]);
        abaFGC.getRange(m+2,2).setValue(parseFloat(arrayValores[m]));
      }
  }
  
  function pintarRentabilidade() {
    aba = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1oA7-jZs7KUPDMixt9FLECo-rBQoJ1DaT0Kzqj5B7sQA/edit#gid=0').getSheetByName('Rentabilidade')
    // intervalo_total_investido = aba.getRange(7,7,1000);
    //numero_linhas = intervalo_total_investido.getValues().length;
    // conte o número de linhas preenchidas
    counter = 0  // contador de número de linhas preenchidas na tabela 'Real'
    for (let i = 3; i <= 602; i++) {
      if (aba.getRange(i,7).getValue() != '') {
        counter += 1;
      }
    }
    for (let i = 7; i <= counter + 6; i++) { // loop pelo número de linhas preenchidas
      if (aba.getRange(i,7).getValue() < aba.getRange(i,4).getValue()) {
        aba.getRange(i,7).setBackground('red')
      }  else {
        aba.getRange(i,7).setBackground('green');
      }
      if (aba.getRange(i,8).getValue() < aba.getRange(i,3).getValue()) {
        aba.getRange(i,8).setBackground('red');
      }  else {
        aba.getRange(i,8).setBackground('green');
      }
      if (aba.getRange(i,9).getValue() < aba.getRange(i,2).getValue()) {
        aba.getRange(i,9).setBackground('red');
      }  else {
        aba.getRange(i,9).setBackground('green');
      }
      if (aba.getRange(i,10).getValue() < aba.getRange(1,5).getValue()) {
        aba.getRange(i,10).setBackground('red')
      }  else {
        aba.getRange(i,10).setBackground('green')
      }
    }
  }
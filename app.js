const xlsx = require("xlsx")
const readline = require('readline')
var fs = require('fs');

var primeiraLinha = 'Processar,Usuario,Nome,Sobrenome,Iniciais,Descricao,Escritorio,Telefone,Email,Endereco,Caixa Postal,Cidade,Estado,CEP,Pais,Caminho do perfil,Script de logon,Caminho local,Conectar,A,CPF,Empresa,Departamento,Cargo,Senha nunca expira,Conta habilitada,OU destino,ProxyAddresses'

const data = new Date()
const nome_arquivo = String(data.getDate() +''+ data.getMonth() +'' +data.getFullYear())

const arquivo = xlsx.readFile('./excel/alunos.xlsx',{cellDates: true})
const dados = arquivo.Sheets["Alunos_AD_CPI"]

const informacoes = xlsx.utils.sheet_to_csv(dados)

fs.appendFile(`./cache/${nome_arquivo}.csv`, String(informacoes), function (err) {
    if (err) throw err;
    console.log('Criado com sucesso!');
});

fs.appendFile(`./gerado/script.csv`, primeiraLinha, function (err) {
    if (err) throw err;
    console.log('Criado com sucesso!');
});

const readable = fs.createReadStream(`./cache/${nome_arquivo}.csv`)
const rl = readline.createInterface({
    input: readable
    // output: process.stdout
})

rl.on('line',(line)=>{
    var dados = line.split(',',3);
    var nome = dados[0].split(' ',1)
    var sobrenome = String(dados[0]).replace(nome,"")

    // tem que muda essa linha
    if(dados[0] === 'NOME' ){
        console.log('Linha removida!')
    }else{
        var linhaCSV = `,Sim,${dados[2]}, ${nome},${sobrenome},,,,,,,,,,,,,,,,,${dados[1]},,,,sim,sim,,,,,,,OU=Alunos,,,,,,,,`
        fs.appendFile("./gerado/script.csv", "\n"+linhaCSV, function(err){
            if(err) console.log('Erro ao adicionar usuário ao script ' + err);
            else console.log(nome + ' Adicionado no script')
        });
    }
    
})

// 1. Loop "while": Solicitar número até que seja digitado o número 3.
function loopWhile() {
    let numero = parseInt(prompt("Digite um número:"));
    while (numero !== 3) {
        console.log("Você digitou:", numero);
        numero = parseInt(prompt("Digite outro número:"));
    }
    console.log("Número 3 digitado. Encerrando o loop.");
}

// 2. Loop "do...while": Solicitar senha com no máximo 3 tentativas.
function validaSenha() {
    const senhaCorreta = "123456";
    let tentativas = 0;
    let acesso = false;
    let tentativa;
    do {
        tentativa = prompt("Digite a senha de acesso:");
        tentativas++;
        if (tentativa === senhaCorreta) {
            acesso = true;
            break;
        }
    } while (tentativas < 3);

    if (acesso) {
        console.log("Acesso concedido!");
    } else {
        console.log("Senha bloqueada!");
    }
}

// 3. Exibir uma lista com 4 números.
function exibirListaFixa() {
    const lista = [10, 20, 30, 40];
    console.log("Lista fixa de 4 números:");
    lista.forEach((numero) => console.log(numero));
}

// 4. Solicitar 5 números do usuário e exibi-los.
function solicitarCincoNumeros() {
    const numeros = [];
    for (let i = 0; i < 5; i++) {
        let num = parseFloat(prompt(`Digite o ${i + 1}º número:`));
        numeros.push(num);
    }
    console.log("Números digitados:");
    numeros.forEach((numero) => console.log(numero));
}

// 5. Função sem parâmetros que retorna uma mensagem personalizada.
function mensagemPersonalizada() {
    return "Bem-vindo ao programa de Ciência de Dados!";
}

// 6. Função que recebe um nome como parâmetro e retorna uma saudação.
function saudacao(nome) {
    return `Olá, ${nome}, que bom ter você no programa Trilhas.`;
}

// 7. Função calcularQuadrado: retorna o quadrado de um número.
function calcularQuadrado(numero) {
    return numero ** 2;
}

// 8. Função Subtracao: retorna a subtração entre dois números.
function Subtracao(a, b) {
    return a - b;
}

// Execução das funções para demonstrar os resultados:
console.log("=== LOOP WHILE ===");
loopWhile();

console.log("\n=== VALIDAÇÃO DE SENHA (do...while) ===");
validaSenha();

console.log("\n=== LISTA FIXA ===");
exibirListaFixa();

console.log("\n=== SOLICITAÇÃO DE 5 NÚMEROS ===");
solicitarCincoNumeros();

console.log("\n=== FUNÇÕES ===");
console.log(mensagemPersonalizada());
console.log(saudacao("Ana"));
console.log("Quadrado de 5:", calcularQuadrado(5));
console.log("Subtração de 10 e 3:", Subtracao(10, 3));

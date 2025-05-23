# Query2Xlsx 📊

![Node.js](https://img.shields.io/badge/Node.js-v16+-green) ![Excel4Node](https://img.shields.io/badge/Excel4Node-v1.8-blue) ![License](https://img.shields.io/badge/license-MIT-brightgreen)

**Query2Xlsx** é um microsserviço desenvolvido para facilitar a geração de planilhas Excel a partir de queries no ambiente Protheus (AdvPL). Este projeto, inspirado no [zQry2Excel](https://github.com/dan-atilio/AdvPL/blob/master/Fontes/zQry2Excel.prw) do usuário [dan-atilio](https://github.com/dan-atilio), foi criado por **Denis Varella** em **Maio de 2025** utilizando **Node.js** e a biblioteca [excel4node](https://www.npmjs.com/package/excel4node) para geração de planilhas.

🚀 O objetivo é oferecer uma solução leve e eficiente para exportar dados do Protheus para o formato Excel, integrando-se facilmente ao ambiente Protheus_Data.

---

## 🛠️ Tecnologias Utilizadas

-   **Node.js**: Backend para o microsserviço.
-   **excel4node**: Biblioteca para criação de arquivos Excel.
-   **Protheus (AdvPL)**: Integração com o ERP Protheus.

---

## 📋 Pré-requisitos

Para utilizar o Query2Xlsx, você precisará ter instalado:

-   [Node.js](https://nodejs.org/) (versão 16 ou superior).
-   [Protheus](https://www.totvs.com/protheus/) configurado no ambiente.
-   [npm](https://www.npmjs.com/) para gerenciar dependências.

---

## ⚙️ Instalação

Siga os passos abaixo para configurar o projeto:

1. **Clone o repositório**:

    ```bash
    git clone https://github.com/denisvarella/query2xlsx.git
    ```

2. **Instale as dependências**:
   Navegue até a pasta do projeto e execute:

    ```bash
    npm install
    ```

3. **Configure as variáveis de ambiente**:

    - Na raiz do projeto, há um arquivo `.env-model`. Copie-o e renomeie para `.env`.
    - Preencha as variáveis de ambiente conforme o modelo (ex.: conexão com o Protheus, portas, etc.).

4. **Instale na pasta Protheus_Data**:

    - Copie a pasta do projeto para o diretório `Protheus_Data`.
    - Acesse a pasta via terminal e inicie o serviço:
        ```bash
        npm run dev
        ```

5. **Configuração do serviço no Windows** (opcional):
    - Utilize o arquivo `Instalar_Serviço.bat` para criar um serviço no Windows que inicializa automaticamente o Query2Xlsx.
    - Execute o arquivo como administrador:
        ```bash
        Instalar_Serviço.bat
        ```

---

## 🚀 Como Usar no Protheus

Para integrar o Query2Xlsx ao Protheus, você pode utilizar uma função específica no AdvPL. Abaixo está um modelo que você pode preencher com os detalhes da sua implementação:

```advpl
// Função para integração com o Query2Xlsx
User Function Query2Xlsx(cQuery)
    // Adicione aqui o código para chamar o microsserviço
    // Exemplo: enviar a query para o endpoint do Node.js
    // e processar o retorno para gerar a planilha
Return
```

📝 **Nota**: Preencha a função acima com os detalhes específicos da sua integração, como o endpoint do microsserviço e os parâmetros necessários.

---

## 🌟 Inspiração

O Query2Xlsx foi inspirado no projeto [zQry2Excel](https://github.com/dan-atilio/AdvPL/blob/master/Fontes/zQry2Excel.prw), desenvolvido pelo usuário [dan-atilio](https://github.com/dan-atilio). Agradecemos pela base fornecida, que serviu como ponto de partida para esta solução.

---

## 📄 Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

---

## 🤝 Contribuições

Contribuições são bem-vindas! 😊 Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias, correções ou novas funcionalidades.

---

## 📧 Contato

Desenvolvido por **Denis Varella**  
📅 **Maio de 2025**  
📫 Entre em contato: [seu-email@example.com](mailto:seu-email@example.com)  
🔗 [GitHub](https://github.com/denisvarella)

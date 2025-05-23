# Query2Xlsx 📊

![Node.js](https://img.shields.io/badge/Node.js-v16+-green) ![Excel4Node](https://img.shields.io/badge/Excel4Node-v1.8-blue) ![License](https://img.shields.io/badge/license-MIT-brightgreen)

**Query2Xlsx** é um microsserviço desenvolvido para facilitar a geração de planilhas em XLSX a partir de queries no ambiente Protheus (AdvPL). Este projeto, inspirado no [zQry2Excel](https://github.com/dan-atilio/AdvPL/blob/master/Fontes/zQry2Excel.prw) do usuário [dan-atilio](https://github.com/dan-atilio), foi criado por **Denis Varella** em **Maio de 2025** utilizando **Node.js** e a biblioteca [excel4node](https://www.npmjs.com/package/excel4node) para geração de planilhas.

🚀 O objetivo é oferecer uma solução leve e eficiente para exportar dados do Protheus para o formato XLSX, integrando-se facilmente ao Protheus.

---

## 🌟 Inspiração

O Query2Xlsx foi inspirado no projeto [zQry2Excel](https://github.com/dan-atilio/AdvPL/blob/master/Fontes/zQry2Excel.prw), desenvolvido pelo usuário [dan-atilio](https://github.com/dan-atilio). Agradecemos pela base fornecida, que serviu como ponto de partida para esta solução.

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

-   Acesse o diretório da protheus_data via terminal e execute:
    ```bash
    git clone https://github.com/denisvarella/query2xlsx.git
    ```

2. **Instale as dependências**:
   Navegue até a pasta do projeto e execute:

    - Ainda no terminal, digite "cd queryxlsx" e execute:

    ```bash
    npm install
    ```

3. **Configure as variáveis de ambiente**:

    - Acesse o projeto com o comando "code ." ou abra diretamente no VSCode.
    - Na raiz do projeto, há um arquivo `.env-model`. Copie-o e renomeie para `.env`.
    - Preencha as variáveis de ambiente conforme o modelo (ex.: conexão com o Protheus, portas, etc.).

4. **Configuração do serviço no Windows** (opcional):
    - Utilize o arquivo `Instalar_Serviço.bat` para criar um serviço no Windows que inicializa automaticamente o Query2Xlsx.
    - Execute o arquivo como administrador:
        ```bash
        Instalar_Serviço.bat
        ```

---

## 🚀 Como Usar no Protheus

O **Query2Xlsx** funciona como um microsserviço que se integra ao Protheus para gerar planilhas Excel a partir de queries SQL. A integração é feita por meio da função de usuário `U_Qry2Xlsx`, que envia os dados necessários para o serviço Node.js e processa o arquivo gerado.

### Fluxo de Funcionamento

1. **No Protheus**:

    - A função `U_Qry2Xlsx` é chamada com os parâmetros necessários para definir a query e a formatação da planilha.
    - A função envia uma requisição **POST** para o endpoint do microsserviço (`http://localhost:3333/generate/excel`), incluindo:
        - A query SQL (`cQry`).
        - O nome do arquivo Excel (`cName`).
        - Títulos das colunas (`aTitulos`).
        - Formatações das colunas (`aFormat`).
        - Tamanhos das colunas (`aSizes`).
        - Um flag para abrir o arquivo automaticamente (`lOpen`).

2. **No microsserviço**:

    - O serviço, conectado ao banco de dados do Protheus, executa a query recebida.
    - Utilizando a biblioteca **excel4node**, o serviço gera o arquivo `.xlsx` com base nos parâmetros fornecidos.
    - O arquivo é salvo na pasta `public/uploads` (ex.: `\B2B\APPNode\public\uploads\`).
    - O serviço retorna um objeto JSON contendo o caminho (`path`) e o nome (`name`) do arquivo gerado.

3. **No Protheus (pós-processamento)**:
    - A função `U_Qry2Xlsx` verifica o status da requisição (200 para sucesso).
    - O arquivo é copiado para a pasta temporária do sistema.
    - Se o parâmetro `lOpen` for `.T.`, o arquivo é aberto automaticamente:
        - Se o Microsoft Excel estiver instalado, ele é usado via automação OLE.
        - Caso contrário, tenta abrir com o LibreOffice (se instalado).
        - Como última opção, o arquivo é aberto pelo programa padrão do sistema.
    - A função retorna um array com o caminho e o nome do arquivo.

### Modelo da Função `U_Qry2Xlsx`

O código fonte da função `U_Qry2Xlsx` está disponível no repositório [inserir link para o projeto no GitHub]. Ele contém a implementação completa para integração com o microsserviço, incluindo a construção da requisição REST e o tratamento do arquivo gerado.

### Parâmetros da Função

-   **cQry** (string): Query SQL a ser executada pelo microsserviço.
-   **cName** (string): Nome do arquivo Excel a ser gerado (sem extensão).
-   **aTitulos** (array): Array com os títulos das colunas da planilha.
-   **aFormat** (array): Array com as formatações das colunas (ex.: tipo de dados, estilos).
-   **aSizes** (array): Array com os tamanhos das colunas.
-   **lOpen** (lógico): Define se o arquivo Excel deve ser aberto automaticamente (`.T.` para abrir, `.F.` para não abrir).

### Exemplo de Uso

```advpl
Local cQuery := "SELECT COD, NOME FROM SX5010 WHERE D_E_L_E_T_ = ''"
Local cFileName := "Clientes_" + DtoS(Date())
Local aTitulos := {"Código", "Nome"}
Local aFormat := {"string", "string"}
Local aSizes := {10, 30}
Local lOpen := .T.
Local aResult := U_Qry2Xlsx(cQuery, cFileName, aTitulos, aFormat, aSizes, lOpen)
If !Empty(aResult[1])
    ConOut("Arquivo gerado: " + aResult[1])
Else
    ConOut("Erro ao gerar o arquivo")
EndIf
```

📝 **Nota**: Certifique-se de que o microsserviço está em execução (`npm run dev`) e que o endpoint (`http://localhost:3333/generate/excel`) está acessível. Consulte o repositório [inserir link para o projeto no GitHub] para obter o código fonte completo da função `U_Qry2Xlsx`.

---

## 🔒 Segurança

Para garantir a segurança da aplicação, siga estas recomendações:

-   **Usuário do banco de dados**:

    -   Crie um usuário no SQL com **permissões apenas de leitura** para a conexão do microsserviço com o banco de dados do Protheus. Isso minimiza riscos de alterações indevidas.
    -   Exemplo de configuração no SQL Server:
        ```sql
        CREATE USER query2xlsx_user WITH PASSWORD = 'sua_senha_segura';
        GRANT SELECT ON DATABASE::nome_do_banco TO query2xlsx_user;
        ```

-   **Arquivo .env**:
    -   O arquivo `.env` contém informações sensíveis, como credenciais de banco de dados e configurações de portas.
    -   **Não compartilhe** o arquivo `.env` nem o inclua em repositórios públicos no GitHub.
    -   Adicione `.env` ao arquivo `.gitignore` para evitar commits acidentais.

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
📫 Entre em contato: [denisvarella@gmail.com](mailto:denisvarella@gmail.com)

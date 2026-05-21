# ⚡ Gestão Estratégica de Obras no Setor Elétrico

## 🚀 Visão Geral do Projeto

O **Operação DCMD** é uma solução robusta desenvolvida em **VBA (Visual Basic for Applications)** para Microsoft Excel, projetada para otimizar a gestão operacional de projetos de construção e manutenção de redes elétricas. Este sistema centraliza o planejamento, execução e acompanhamento de obras, oferecendo uma visão 360 graus das operações, proporcionando eficiência, controle financeiro e visibilidade em tempo real.

## ✨ Funcionalidades Principais

*   **Dashboard Gerencial:** Painel intuitivo com indicadores financeiros, metas de desempenho e eficiência operacional por equipe.
*   **Programação Detalhada:** Registro e acompanhamento de todas as atividades programadas, com histórico e projeções futuras.
*   **Carteira de Obras:** Gestão completa do portfólio de obras, incluindo rastreamento financeiro e status de execução.
*   **Calendário Interativo:** Grade mensal para alocação de equipes e visualização do status das atividades.
*   **Formulários Dinâmicos:** Geração de formulários de Ponto de Serviço (PS) diretamente via VBA.
*   **Gestão de Composição e Medição:** Ferramentas para controle de composição de equipes e medição de serviços.

## 🛠️ Tecnologias e Ferramentas

Este projeto aproveita o poder do VBA para criar uma aplicação rica em funcionalidades dentro do ambiente familiar do Excel, complementado por integrações estratégicas:

*   **Microsoft Excel:** Plataforma principal para a interface do usuário e armazenamento de dados.
*   **VBA (Visual Basic for Applications):** Linguagem de programação para automação, lógica de negócios e interação com APIs.
*   **Microsoft Graph API:** Utilizada para interagir com serviços da Microsoft 365, como o SharePoint, para leitura e escrita de dados.
*   **Microsoft SharePoint:** Serve como um robusto backend para armazenamento de dados e controle de versão das informações do projeto.
*   **OAuth 2.0 Device Code Flow:** Implementado para um processo de autenticação seguro e eficiente com a Microsoft Graph API, permitindo que o aplicativo obtenha tokens de acesso sem expor credenciais.
*   **MSXML2.XMLHTTP:** Objeto utilizado para realizar requisições HTTP (GET, POST) para as APIs RESTful.
*   **JsonConverter (VBA-JSON):** Biblioteca VBA para serialização e desserialização de objetos JSON, essencial para o tratamento de respostas da API Graph. Biblioteca JSON Converter for VBA, desenvolvido por Tim Hall
*   **Registro do Windows:** Utilizado para persistir de forma segura os tokens de acesso e refresh evitando a necessidade de reautenticação constante.

## 🏗️ Arquitetura e Integrações

Projetado com uma arquitetura que maximiza a funcionalidade do Excel enquanto se integra a serviços corporativos essenciais:

*   **Frontend (Excel/VBA):** A interface do usuário é totalmente construída no Excel, utilizando formulários (UserForms) e controles ActiveX para uma experiência interativa.
*   **Backend (SharePoint):** Listas e bibliotecas do SharePoint são utilizadas para armazenar os dados operacionais, financeiros e de programação, garantindo escalabilidade e acesso colaborativo.
*   **Camada de Comunicação (Microsoft Graph API):** Todas as interações com o SharePoint são realizadas através da Microsoft Graph API, proporcionando uma comunicação padronizada e segura.

### Fluxo de Autenticação (OAuth 2.0 Device Code Flow)

Para garantir a segurança e a conformidade, o projeto implementa o fluxo de autenticação OAuth 2.0 Device Code. Este método é ideal para aplicações que não possuem uma interface de navegador completa, como é o caso de aplicações VBA. O fluxo ocorre da seguinte forma:

1.  O aplicativo solicita um código de dispositivo e um URI de verificação à Microsoft Identity Platform.
2.  O usuário é instruído a acessar o URI em um navegador e inserir o código fornecido.
3.  Após a autenticação bem-sucedida no navegador, o aplicativo recebe um token de acesso e um token de refresh.
4.  Os tokens são armazenados no Registro do Windows para uso futuro, minimizando interrupções.
5.  O token de refresh é utilizado para obter novos tokens de acesso quando o token atual expira, garantindo acesso contínuo aos recursos do SharePoint.

## 📂 Estrutura do Projeto

O repositório contém os seguintes módulos VBA e arquivos de formulário:

*   `AtualizarConexão.bas`: Lógica para atualização de conexões de dados.
*   `Calendario.bas`: Funções relacionadas ao calendário.
*   `GetToken.bas`: Implementação do fluxo de autenticação OAuth 2.0 Device Code e gerenciamento de tokens.
*   `JsonConverter.bas`: Módulo para manipulação de JSON.
*   `frmMain.frm`, `frmMain.frx`: Formulário principal da aplicação, responsável pela interface do usuário e orquestração das funcionalidades.
*   `frmDeviceCode.frm`, `frmDeviceCode.frx`: Formulário para exibir o código de dispositivo e o URI de verificação durante a autenticação.
*   `frmAddCompo.frm`, `frmAddCompo.frx`: Formulário para adicionar composições de equipe.
*   `frmAddItem.frm`, `frmAddItem.frx`: Formulário para adicionar itens.
*   `frmAddItemCart.frm`, `frmAddItemCart.frx`: Formulário para adicionar itens na carteira de obras.
*   `frmAddOse.frm`, `frmAddOse.frx`: Formulário para adicionar códigos de mão de obra.

## 📄 Licença

Desenvolvido por: Pedro Ricardo Moreno
https://www.linkedin.com/in/pedroricardomoreno/

JSON Converter for VBA
VBA-JSON v2.3.1
(c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
@class JsonConverter
@author tim.hall.engr@gmail.com
@license MIT (http://www.opensource.org/licenses/mit-license.php)

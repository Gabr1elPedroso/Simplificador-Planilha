Gerador de Planilhas RH - Premium

Um sistema automatizado e robusto construído em Python e Streamlit para consolidar, cruzar e padronizar bases de dados complexas de Recursos Humanos.

A aplicação lê planilhas brutas contendo múltiplas abas (Ativos, Cessão, Demitidos, FGTS, Pensão, Adipar, etc.), cruza as informações de forma inteligente usando a Matrícula do funcionário como chave, e gera um layout final padronizado com 62 colunas, pronto para importação ou análise.

Principais Funcionalidades

Acesso Restrito: Sistema de login integrado para garantir que apenas pessoas autorizadas processem os dados.

Interface Premium: Visual moderno em Dark Mode com detalhes em dourado, proporcionando uma excelente experiência de usuário (UX).

Leitura Inteligente: O algoritmo localiza automaticamente a linha de cabeçalho correta nas abas brutas, evitando erros se o Excel vier com linhas em branco no topo.

Cruzamento Automático: Mescla dados financeiros e cadastrais (eConsignado, Adipar, Experiência, FGTS) de diferentes abas em uma única linha por funcionário.

Filtro de Matrículas: Permite colar uma lista de matrículas específicas para gerar um relatório segmentado.

Exportação Rápida: Gera um arquivo .xlsx limpo e estruturado em segundos.

Como Utilizar a Aplicação (Usuário Final)

Acesse o link da aplicação hospedada (ex: via Streamlit Community Cloud).

Faça o Login utilizando seu Usuário e Senha fornecidos pelo administrador.

Faça o Upload do arquivo Excel bruto (.xlsx) clicando na área indicada ou arrastando o arquivo para a tela.

Filtre (Opcional): Se desejar o relatório apenas de algumas pessoas, cole a lista de matrículas na caixa "Filtro Específico" (uma por linha ou separadas por vírgula). Se quiser processar toda a empresa, deixe em branco.

Clique no botão  PROCESSAR DADOS E GERAR LAYOUT.

Aguarde a barra de progresso e clique em 📥 BAIXAR PLANILHA CONSOLIDADA.

🛠️ Como Rodar o Projeto Localmente (Desenvolvedores)

Se você deseja baixar o código para fazer modificações na sua máquina, siga os passos abaixo:

1. Pré-requisitos

Certifique-se de ter o Python instalado (versão 3.8 ou superior).

2. Clonar o Repositório

git clone https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git
cd SEU-REPOSITORIO


3. Instalar as Dependências

Instale as bibliotecas necessárias usando o pip:

pip install -r requirements.txt


4. Configurar as Senhas Locais (Secrets)

Para que a tela de login funcione na sua máquina, você precisa criar o arquivo de segredos.

Crie uma pasta oculta chamada .streamlit na raiz do projeto.

Dentro dela, crie um arquivo chamado secrets.toml.

Adicione seus usuários e senhas assim:

[passwords]
admin = "senha123"
rh = "premium2026"


(Atenção: O arquivo secrets.toml nunca deve ser enviado para o GitHub. Ele já está protegido caso você tenha configurado o .gitignore).

5. Executar a Aplicação

Rode o comando abaixo no seu terminal:

streamlit run app.py


O navegador abrirá automaticamente no endereço http://localhost:8501.

☁️ Deploy (Hospedagem)

Este projeto está configurado para ser hospedado facilmente no Streamlit Community Cloud.
Ao fazer o deploy, lembre-se de acessar as configurações do aplicativo no painel do Streamlit (Settings > Secrets) e colar o conteúdo do seu secrets.toml lá dentro para que o login funcione na nuvem.

Estrutura de Arquivos

/
├── .streamlit/
│   └── config.toml        # Configurações de tema (Dark Mode/Dourado)
├── app.py                 # Código principal da aplicação
├── requirements.txt       # Bibliotecas (streamlit, pandas, openpyxl)
├── .gitignore             # Arquivos ignorados pelo Git (ex: secrets.toml)
└── README.md              # Documentação do projeto


Desenvolvido para automação e excelência em processos de RH.

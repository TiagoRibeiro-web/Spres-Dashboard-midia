📄 GESTÃO MÍDIA PRO SPRES
Documentação Técnica - Versão 2.0
Agosto 2026
1. VISÃO GERAL
O Gestão Mídia Pro é um dashboard desenvolvido para a Sucos Spres que permite o monitoramento em tempo real do cronograma de mídia, integrado ao SharePoint Online.

2. TECNOLOGIAS UTILIZADAS
Tecnologia	Finalidade
Streamlit	Framework para criação do dashboard
Microsoft Graph API	Conexão com SharePoint Online
Pandas	Processamento e manipulação de dados
Plotly	Visualizações interativas e gráficos
Python-dotenv	Gerenciamento de variáveis de ambiente
Requests	Requisições HTTP para API

3. ESTRUTURA DO PROJETO
text
📁 Gestao_Nova_Spres/
├── 📄 app.py                    # Dashboard principal (2300+ linhas)
├── 📄 graph_api.py              # Conexão Microsoft Graph API
├── 📄 excel_integration.py      # Integração com Excel/SharePoint
├── 📄 .env                      # Variáveis de ambiente
├── 📄 logo.png                  # Logo da Spres
└── 📁 .venv/                    # Ambiente virtual Python

Arquivo	Função

app.py	Dashboard principal e interface do usuário
graph_api.py	Conexão com Microsoft Graph API (autenticação, leitura)
excel_integration.py	Integração e processamento de dados do Excel
.env	Configurações sensíveis (credenciais Azure)
logo.png	Logo da Sucos Spres


4. FLUXO DE FUNCIONAMENTO
text
1. Usuário acessa o dashboard
   ↓
2. Tela de login (admin/spres2026)
   ↓
3. Login rápido (< 2 segundos)
   ↓
4. Carregamento dos dados do SharePoint
   ↓
5. Exibição do dashboard com dados em tempo real
   ↓
6. Usuário pode interagir com os gráficos e tabelas
   ↓
7. Botão "Recarregar Dados" para atualizar

5. FUNCIONALIDADES DETALHADAS

5.1. Tela de Login
Fundo animado com gradiente

Partículas flutuantes

Card com efeito glassmorphism

Borda com animação colorida

Feedback visual para erros

5.2. Dashboard
Big Numbers: Total Geral, Controle Aprovado, Saldo

Distribuição por Veículo: Treemap interativo

Investimento Mensal: Gráfico de barras

Resumo por Veículo: Tabela com totais

5.3. Cronograma Completo
Tabela detalhada com todos os itens

Exportação para CSV

Linhas de total com formatação

5.4. Análises
Evolução do Investimento Mensal (área)

Maiores Investimentos por Veículo (barras)

Aprovado vs Utilizado por Mês

5.5. Sidebar
Status da integração com SharePoint

Botão "Abrir Planilha no SharePoint"

Botão "Recarregar Dados"

Informações da sessão

Botão "Sair"

6. CORES E IDENTIDADE VISUAL SPRES
Elemento	Cor	Código Hex
Header	Azul	#00325F → #004B8D → #2E7DD1
Big Numbers	Azul	#004B8D
Destaques	Amarelo	#FFD600
Alertas	Laranja	#FF8A1E
Fundo	Branco/Gradiente	#F7FAFC → #EDF4FB
Sidebar	Azul Escuro	#0A1F35 → #00325F
Texto Principal	Azul Escuro	#0A1F35
Texto Secundário	Cinza	#5B6E80

7. CREDENCIAIS DE ACESSO
Usuário	Senha
admin	spres2026
gestao	midia2026

8. VARIÁVEIS DE AMBIENTE (.env)
env
# Azure AD Configuration
AZURE_TENANT_ID=46d481f9-...
AZURE_CLIENT_ID=5852b17d-...
AZURE_CLIENT_SECRET=...

# SharePoint Configuration
SHAREPOINT_SITE_URL=https://agenciaideatore-my.sharepoint.com/personal/eduardo_zaupa_ideatoreamericas_com

# Excel Configuration
EXCEL_FILENAME=base_spres_projeto_refatorada.xlsx

# URL direta do arquivo (para o botão "Abrir Planilha")
SHAREPOINT_FILE_URL=https://agenciaideatore-my.sharepoint.com/:x:/g/personal/eduardo_zaupa_ideatoreamericas_com/IQD0Y7VvdE_eR6cD-OgH85NyAVm6vrlVGBoKUC63GKwYKnI?e=PTgTb5

# Cache
CACHE_MINUTES=1
9. INSTALAÇÃO E EXECUÇÃO
9.1. Clonar o repositório
bash
git clone [URL_DO_REPOSITORIO]
cd Gestao_Nova_Spres
9.2. Criar ambiente virtual
bash
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
.venv\Scripts\activate     # Windows
9.3. Instalar dependências
bash
pip install -r requirements.txt
9.4. Configurar .env
Criar arquivo .env na raiz do projeto com as credenciais fornecidas.

9.5. Executar
bash
streamlit run app.py
9.6. Acessar
Abra o navegador em: http://localhost:8501

10. MANUTENÇÃO
Atividade	Frequência	Responsável
Atualizar credenciais Azure	Conforme expiração	Administrador
Validar dados	Diário	Usuário
Backup .env	Mensal	Administrador
Verificar logs	Semanal	Administrador
Atualizar dependências	Trimestral	Administrador

11. CORREÇÕES REALIZADAS
#	Problema	Solução
1	Login lento (40s)	Lazy loading no graph_api.py
2	Tela de login duplicada	Limpeza de CSS no main()
3	Dados com erro de conversão	Correção da função _converter_valor_monetario
4	Valores multiplicados por 100	Correção na conversão de vírgula/ponto
5	Linhas de total duplicadas	Verificação no criar_tabela_controle
6	"undefined" nos gráficos	Filtro de valores nulos
7	Métricas principais erradas	Removido cards confusos
8	Sidebar com cores Spres	CSS personalizado
9	Botão Upload invisível	Removido (desnecessário)
10	Mensagem de sucesso fixa	Auto-remoção após 2s
11	Botão "Abrir Planilha" invisível	CSS com contraste melhorado

12. DADOS ATUAIS
Métrica	Valor
Total de linhas	49
Total de colunas	15
Meses	fev/26 a jan/27
Total Geral	R$ 419.064,96
Controle Aprovado	R$ 419.835,46
Saldo	-R$ 770,50

13. DEPENDÊNCIAS
txt
streamlit>=1.28.0
pandas>=2.0.0
plotly>=5.14.0
requests>=2.31.0
python-dotenv>=1.0.0
openpyxl>=3.1.0
numpy>=1.24.0

14. CONTATO E SUPORTE
Desenvolvido para: Sucos Spres
Data: Agosto 2026
Versão: 2.0
Ambiente: Produção

📋 Checklist Final
Item	Status
Login funcional	✅
Conexão SharePoint	✅
Dados corretos	✅
Gráficos funcionais	✅
Sidebar estilizado	✅
Botões visíveis	✅
Código organizado	✅
Documentação pronta	✅
# 📊 CNPJ Enterprise Analytics - Automator

Uma aplicação web local robusta desenvolvida em **Python** e **Streamlit** para automação de enriquecimento cadastral em massa. A ferramenta consome a API pública `cnpj.ws` e processa lotes de CNPJs a partir de planilhas Excel, lidando ativamente com limites de rede, tratamento de dados complexos e resiliência de estado.

## 🚀 Principais Funcionalidades

- **Tratamento de Rate Limit (429):** Leitura dinâmica do `Retry-After` header para pausas inteligentes e prevenção de bloqueio de IP.
- **Resiliência Anti-Queda:** Salvamento em lotes e cache de estado. Se o sistema for interrompido, ele retoma exatamente do ponto em que parou.
- **Integração Nativa OS:** Utiliza `tkinter` para invocar o Windows Explorer nativo dentro da interface web, facilitando a escolha do diretório de destino.
- **Mapeamento de Múltiplas IEs:** Identifica e mapeia dinamicamente Inscrições Estaduais (IE) separando-as ativamente por Estado (UF) em novas colunas.
- **Telemetria em Tempo Real:** Dashboard Streamlit integrado exibindo velocidade (Req/Min), ETA matemático real e logs detalhados de execução.

## 🛠️ Tecnologias Utilizadas

- [Python 3.9+](https://www.python.org/)
- [Streamlit](https://streamlit.io/) - Interface Web e UI reativa
- [Pandas](https://pandas.pydata.org/) - Manipulação de DataFrames e Excel
- [Requests](https://requests.readthedocs.io/) - Integração com API REST e gerenciamento de Sessão HTTP
- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/) - Engine de leitura/escrita nativa de arquivos `.xlsx`

## ⚙️ Como Instalar e Rodar

1. **Clone este repositório:**
   ```bash
   git clone [https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git](https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git)
   cd SEU_REPOSITORIO

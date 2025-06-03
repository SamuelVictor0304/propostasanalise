# Sistema de Análise de Propostas

Este projeto é uma aplicação web desenvolvida em [Streamlit](https://streamlit.io/) para análise de propostas comerciais a partir de uma planilha Excel. O sistema permite filtrar, visualizar e exportar relatórios detalhados sobre o desempenho dos negociadores e o status das propostas.

## Funcionalidades

- **Upload automático** da planilha de propostas (`PROPOSTAS 2025..xlsx`)
- **Filtros interativos** por negociador e mês
- **Visualização de métricas** principais (total de propostas, aprovadas, taxa de aprovação)
- **Relatório detalhado** por negociador e status
- **Exportação** do relatório filtrado para Excel

## Como usar

1. Certifique-se de que o arquivo `PROPOSTAS 2025..xlsx` está na mesma pasta do aplicativo.
2. Instale as dependências:
   ```sh
   pip install -r requirements.txt
   ```
3. Execute o aplicativo:
   ```sh
   streamlit run app_streamlit.py
   ```
4. Acesse o endereço exibido no terminal para interagir com a aplicação.

## Estrutura da Planilha

A planilha deve conter, pelo menos, as seguintes colunas (os nomes podem ser ajustados no código):

- **NEGOCIADOR**: Nome do negociador responsável pela proposta
- **STATUS**: Situação da proposta (ex: Aprovada, Recusada)
- **DATA DO ENVIO DA PROPOSTA**: Data de envio da proposta

(Por questões de segurança a planilha nao pode ser compartilhada por conter dados sigilosos)

## Dependências

- streamlit==1.32.0
- pandas==2.2.0
- openpyxl==3.1.2
- numpy

Veja o arquivo [requirements.txt](requirements.txt) para mais detalhes.

## Observações

- O relatório pode ser baixado em formato Excel diretamente pela interface.
- Caso a planilha não esteja no formato esperado, mensagens de erro serão exibidas na tela.

---

Desenvolvido para facilitar a análise e acompanhamento de propostas comerciais.

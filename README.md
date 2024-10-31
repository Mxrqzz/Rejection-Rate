# Rejection Rate

Este repositório contém scripts em Python que automatizam a análise de dados de CNPJ a partir de planilhas do Excel.

## Funcionalidades

1. **Análise de CNPJ**:
   - Mapeia CNPJs e conta ocorrências totais.
   - Identifica e contabiliza ocorrências "Indeferida".
   - Gera um relatório em uma nova planilha com as seguintes colunas:
     - CNPJ
     - Indefinidas
     - Total Analisadas
     - Índice de Indeferimento (%)

2. **Filtragem de Dados**:
   - Remove linhas onde o índice de indeferimento é inferior a 50%.
   - Salva uma nova planilha com os dados filtrados.

## Pré-requisitos

- Python 3.x
- openpyxl

```bash
pip install openpyxl

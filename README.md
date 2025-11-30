# Medição Dezembro - CAERD

Sistema para processar planilhas de medição e gerar página HTML com itens agrupados para cotação.

## Funcionalidades

- ✅ Processa planilhas Excel (.xlsx)
- ✅ Agrupa itens repetidos
- ✅ Gera página HTML com tabela de itens
- ✅ Suporta imagens dos produtos (tooltip ao passar o mouse)
- ✅ Exporta dados para CSV

## Como usar

1. Coloque a planilha Excel (`Dartagnan.xlsx`) na pasta do projeto
2. Execute o script:
   ```bash
   python agrupar_itens_cotacao.py
   ```
3. Abra o arquivo `itens_cotacao_dartagnan.html` no navegador

## Adicionar fotos dos produtos

1. Coloque as fotos na pasta `imagens/`
2. Nomeie os arquivos com números de 1 a 39:
   - `1.jpg`, `2.jpg`, `3.jpg`... até `39.jpg`
   - Ou `1.png`, `2.png`, etc.
3. As imagens aparecerão automaticamente ao passar o mouse sobre os itens com ícone 📷

## Estrutura do projeto

```
.
├── agrupar_itens_cotacao.py    # Script principal
├── itens_cotacao_dartagnan.html # Página HTML gerada
├── imagens/                      # Pasta para fotos dos produtos
└── README.md                     # Este arquivo
```

## Requisitos

- Python 3.x
- pandas
- openpyxl

Instalar dependências:
```bash
pip install pandas openpyxl
```


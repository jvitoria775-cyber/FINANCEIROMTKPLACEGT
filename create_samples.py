"""
Gera planilhas de exemplo no formato do UpSeller / Marketplaces para testes rápidos.
"""

import os
import pandas as pd

os.makedirs("exemplos", exist_ok=True)

# 1. Planilha de Exemplo de Anúncios
anuncios_data = [
    {
        "ID da Variante": "VAR_SHP_101",
        "ID do Anúncio": "ITEM_SHP_901",
        "SKU": "POLO-AZUL-M",
        "Nome do Produto": "Camisa Polo Algodão Pima Confort",
        "Variação": "Azul Marinho / M",
        "Preço Original": "R$ 139,90",
        "Preço com Desconto": "R$ 109,90",
    },
    {
        "ID da Variante": "VAR_SHP_102",
        "ID do Anúncio": "ITEM_SHP_901",
        "SKU": "POLO-PRETA-G",
        "Nome do Produto": "Camisa Polo Algodão Pima Confort",
        "Variação": "Preta / G",
        "Preço Original": "R$ 139,90",
        "Preço com Desconto": "R$ 109,90",
    },
    {
        "ID da Variante": "VAR_TT_201",
        "ID do Anúncio": "ITEM_TT_801",
        "SKU": "CALCA-JEANS-40",
        "Nome do Produto": "Calça Jeans Slim Masculina",
        "Variação": "Jeans Escuro / 40",
        "Preço Original": "R$ 180,00",
        "Preço com Desconto": "R$ 149,90",
    },
    {
        "ID da Variante": "VAR_SHN_301",
        "ID do Anúncio": "ITEM_SHN_701",
        "SKU": "VESTIDO-VERAO-P",
        "Nome do Produto": "Vestido Floral Alça Fina Verão",
        "Variação": "Floral Estampa / P",
        "Preço Original": "R$ 89,90",
        "Preço com Desconto": "R$ 69,90",
    },
]

df_anuncios = pd.DataFrame(anuncios_data)
df_anuncios.to_excel("exemplos/exemplo_anuncios_upseller.xlsx", index=False)
print("Planilha exemplos/exemplo_anuncios_upseller.xlsx gerada com sucesso!")

# 2. Planilha de Exemplo de Vendas (Export_Order...)
vendas_data = [
    {
        "Número do Pedido": "241006A1B2C3",
        "ID da Variante": "VAR_SHP_101",
        "SKU": "POLO-AZUL-M",
        "Nome do Produto": "Camisa Polo Algodão Pima Confort",
        "Quantidade": 2,
        "Preço de Produto": "R$ 85,00",  # Cliente usou cupom Shopee de R$ 24,90 por peça!
    },
    {
        "Número do Pedido": "241006D4E5F6",
        "ID da Variante": "VAR_SHP_102",
        "SKU": "POLO-PRETA-G",
        "Nome do Produto": "Camisa Polo Algodão Pima Confort",
        "Quantidade": 1,
        "Preço de Produto": "R$ 89,90",  # Cliente pagou R$ 89,90 (anunciado R$ 109,90)
    },
    {
        "Número do Pedido": "241006G7H8I9",
        "ID da Variante": "VAR_TT_201",
        "SKU": "CALCA-JEANS-40",
        "Nome do Produto": "Calça Jeans Slim Masculina",
        "Quantidade": 1,
        "Preço de Produto": "R$ 119,90",  # Cupom TikTok Shop de R$ 30,00!
    },
    {
        "Número da Venda": "241006J1K2L3",
        "ID da Variante": "VAR_SHN_301",
        "SKU": "VESTIDO-VERAO-P",
        "Nome do Produto": "Vestido Floral Alça Fina Verão",
        "Quantidade": 3,
        "Preço de Produto": "R$ 55,00",  # Cupom Shein!
    },
    {
        "Número da Venda": "241006MULTI_99",
        "ID da Variante": "VAR_SHP_101",
        "SKU": "POLO-AZUL-M",
        "Nome do Produto": "Camisa Polo Algodão Pima Confort",
        "Quantidade": 1,
        "Preço de Produto": "R$ 85,00",
    },
    {
        "Número da Venda": "241006MULTI_99",  # Segunda linha da mesma venda
        "ID da Variante": "VAR_SHP_102",
        "SKU": "POLO-PRETA-G",
        "Nome do Produto": "Camisa Polo Algodão Pima Confort",
        "Quantidade": 1,
        "Preço de Produto": "R$ 89,90",
    },
]

df_vendas = pd.DataFrame(vendas_data)
df_vendas.to_excel("exemplos/exemplo_vendas_upseller.xlsx", index=False)
print("Planilha exemplos/exemplo_vendas_upseller.xlsx gerada com sucesso!")

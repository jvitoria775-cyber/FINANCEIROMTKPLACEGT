import os
import sys
import sqlite3
import pandas as pd

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
from app import DatabaseManager, DataProcessor, ExcelExporter, DEFAULT_CHANNEL_RATES


def run_tests():
    print("=" * 60)
    print("INICIANDO TESTES DO MOTOR FINANCEIRO E BANCO DE DADOS")
    print("=" * 60)

    # 1. Teste de Inicialização do Banco
    DatabaseManager.init_db()
    assert os.path.exists("ecommerce_financeiro.db"), "Banco SQLite não foi criado!"
    print("✅ 1. Banco SQLite inicializado com sucesso.")

    # 2. Teste de Taxas Padrão
    DatabaseManager.update_channel_config("Shopee", 20.0, 5.0, 8.0)
    configs = DatabaseManager.get_channel_configs()
    assert "Shopee" in configs, "Shopee não encontrada em configuracoes_canais"
    assert configs["Shopee"]["taxa_comissao_pct"] == 20.0
    assert configs["Shopee"]["taxa_fixa_unidade"] == 5.0
    assert configs["Shopee"]["aliquota_imposto_pct"] == 8.0

    assert "TikTok Shop" in configs
    assert configs["TikTok Shop"]["taxa_comissao_pct"] == 22.0
    assert configs["TikTok Shop"]["taxa_fixa_unidade"] == 5.0

    assert "Shein" in configs
    assert configs["Shein"]["taxa_comissao_pct"] == 20.0
    assert configs["Shein"]["taxa_fixa_unidade"] == 4.0
    print("✅ 2. Taxas padrão verificadas com sucesso.")

    # 3. Teste de UPSERT de Produtos e Preservação de CMV
    sample_products = [
        {
            "id_variante": "VAR_SHP_001",
            "canal": "Shopee",
            "id_anuncio": "ITEM_SHP_001",
            "sku": "CAMISA-PRETA-G",
            "nome_anuncio": "Camisa Polo Algodão Premium",
            "variacao": "Preta / G",
            "preco_cheio": 120.00,
            "preco_desconto": 99.90,
            "preco_base": 99.90,
            "custo_cmv": 30.00,
        },
        {
            "id_variante": "VAR_TT_002",
            "canal": "TikTok Shop",
            "id_anuncio": "ITEM_TT_002",
            "sku": "VESTIDO-FLORAL-M",
            "nome_anuncio": "Vestido Midi Floral",
            "variacao": "Floral / M",
            "preco_cheio": 150.00,
            "preco_desconto": 0.0,
            "preco_base": 150.00,
            "custo_cmv": 45.00,
        },
    ]

    novos, atus = DatabaseManager.upsert_products(sample_products)
    assert novos + atus == 2, f"Esperado 2 produtos processados, obtido novos={novos}, atus={atus}"
    print("✅ 3. Inserção de produtos realizada com sucesso.")

    # Teste de reimportação com preço atualizado mas sem CMV (deve manter o CMV de 30.00)
    updated_product = [
        {
            "id_variante": "VAR_SHP_001",
            "canal": "Shopee",
            "id_anuncio": "ITEM_SHP_001",
            "sku": "CAMISA-PRETA-G",
            "nome_anuncio": "Camisa Polo Algodão Premium 2026",
            "variacao": "Preta / G",
            "preco_cheio": 130.00,
            "preco_desconto": 105.00,
            "preco_base": 105.00,
            "custo_cmv": 0.0,  # Importação de anúncio não costuma ter CMV
        }
    ]
    novos_reimp, atus_reimp = DatabaseManager.upsert_products(updated_product)
    assert atus_reimp == 1, "Deveria ter atualizado 1 produto"
    
    prods_df = DatabaseManager.get_products_dataframe()
    p_shp = prods_df[prods_df["id_variante"] == "VAR_SHP_001"].iloc[0]
    assert p_shp["preco_base"] == 105.00, "Preço base não foi atualizado"
    assert p_shp["custo_cmv"] == 30.00, f"CMV foi sobrescrito indevidamente! Valor atual: {p_shp['custo_cmv']}"
    print("✅ 4. Idempotência e preservação de CMV no UPSERT verificada.")

    # 4. Teste de Conciliação de Vendas (com cupom do cliente)
    # Suponha que o UpSeller registrou que o cliente pagou R$ 79,90 (com cupom de R$ 25,10 da plataforma)
    # Mas o produto anunciado é R$ 105,00
    sample_sales_df = pd.DataFrame([
        {
            "Número do Pedido": "PED_2026_001",
            "ID da Variante": "VAR_SHP_001",
            "SKU": "CAMISA-PRETA-G",
            "Nome do Produto": "Camisa Polo",
            "Quantidade": 2,
            "Preço de Produto": 79.90,  # Valor distorcido pelo cupom da plataforma
        },
        {
            "Número do Pedido": "PED_2026_002",
            "ID da Variante": "",  # Sem ID de variante no relatório -> testando fallback de SKU
            "SKU": "VESTIDO-FLORAL-M",
            "Nome do Produto": "Vestido Midi Floral",
            "Quantidade": 1,
            "Preço de Produto": 120.00,  # Cliente pagou 120 com cupom, anunciado é 150
        },
    ])

    processed_sales, stats = DataProcessor.reconcile_sales_batch(
        df_sales=sample_sales_df,
        canal="Shopee",
        channel_configs=configs,
        products_df=prods_df,
    )

    assert stats["itens_conciliados"] == 2, "Deveria ter conciliado ambos os itens (1 por variante, 1 por fallback SKU)"
    print("✅ 5. Conciliação com matching por ID de Variante e fallback por SKU verificada.")

    # Validar Matemática do Pedido 1 (Shopee: 20% comissão, R$ 5 fixa, 8% imposto)
    # Qtd: 2, Preço Anunciado: 105.00
    # Fat Bruto: 105.00 * 2 = 210.00
    # Comissão (20%): 210 * 0.20 = 42.00
    # Fixa (R$ 5 * 2): 10.00
    # Total Taxas: 52.00
    # Imposto (8%): 210 * 0.08 = 16.80
    # Saque Líquido: 210.00 - 52.00 - 16.80 = 141.20
    # CMV Total (30.00 * 2): 60.00
    # Lucro Real: 141.20 - 60.00 = 81.20
    p1 = processed_sales[0]
    assert p1["preco_base_anuncio"] == 105.00, f"Preço base incorreto: {p1['preco_base_anuncio']}"
    assert p1["faturamento_bruto"] == 210.00, f"Fat bruto incorreto: {p1['faturamento_bruto']}"
    assert p1["total_taxas_canal"] == 52.00, f"Total taxas incorreto: {p1['total_taxas_canal']}"
    assert p1["valor_imposto"] == 16.80, f"Imposto incorreto: {p1['valor_imposto']}"
    assert p1["saque_liquido_previsto"] == 141.20, f"Saque líquido incorreto: {p1['saque_liquido_previsto']}"
    assert p1["custo_cmv_total"] == 60.00, f"CMV total incorreto: {p1['custo_cmv_total']}"
    assert p1["lucro_liquido_real"] == 81.20, f"Lucro líquido real incorreto: {p1['lucro_liquido_real']}"
    print("✅ 6. Fórmulas de Faturamento, Comissão, Imposto, Saque e Lucro validadas com 100% de exatidão.")

    # 5. Teste de Gravação de Vendas
    novas_v, atus_v = DatabaseManager.upsert_sales(processed_sales)
    assert novas_v + atus_v == 2, "Deveria ter processado 2 vendas"
    
    # Gravar novamente para testar idempotência
    novas_v2, atus_v2 = DatabaseManager.upsert_sales(processed_sales)
    assert novas_v2 == 0 and atus_v2 == 2, "Deveria ter atualizado 2 vendas sem duplicar"
    print("✅ 7. Gravação idempotente de vendas no SQLite validada.")

    # 6. Teste de Recálculo Retroativo
    # Alterar taxa de comissão da Shopee para 18.0%
    DatabaseManager.update_channel_config("Shopee", 18.0, 5.0, 8.0)
    recalc_count = DatabaseManager.recalculate_all_sales()
    assert recalc_count >= 2, f"Esperado ao menos 2 vendas recalculadas, obtido {recalc_count}"

    sales_df_updated = DatabaseManager.get_sales_dataframe("Shopee")
    p1_updated = sales_df_updated[sales_df_updated["pedido_id"] == "PED_2026_001"].iloc[0]
    # Nova Comissão (18% de 210): 37.80 + 10.00 fixa = 47.80
    # Saque Líquido: 210 - 47.80 - 16.80 = 145.40
    # Lucro Real: 145.40 - 60.00 = 85.40
    assert p1_updated["total_taxas_canal"] == 47.80, f"Recálculo de taxas incorreto: {p1_updated['total_taxas_canal']}"
    assert p1_updated["saque_liquido_previsto"] == 145.40, f"Recálculo de saque incorreto: {p1_updated['saque_liquido_previsto']}"
    print("✅ 8. Recálculo retroativo após alteração de taxas validado.")

    # 7. Teste de Exportação Excel (.xlsx)
    excel_bytes = ExcelExporter.export_sales_to_excel(sales_df_updated)
    assert len(excel_bytes) > 1000, "Arquivo Excel gerado está vazio ou truncado"
    print("✅ 9. Geração de arquivo Excel formatado via OpenPyXL validada.")

    # 8. Teste de Coluna de Data e Filtro por Data
    assert "data_pedido" in p1, "Campo data_pedido não foi gerado no processamento da venda"
    assert "data_pedido" in sales_df_updated.columns, "Coluna data_pedido não existe na tabela de vendas"
    dt_val = DataProcessor.parse_date_str("06/10/2026 15:30:00")
    assert "2026-10-06" in dt_val, f"Falha no parsing de data BR: {dt_val}"
    print("✅ 10. Parsing de datas e persistência de data_pedido validados.")

    # 9. Teste de Taggear Número da Venda em Pedidos Multilinhas (e detecção de "Número da Venda")
    multi_sales_df = pd.DataFrame([
        {
            "Número da Venda": "VENDA_MULTILINHA_888",
            "ID da Variante": "VAR_SHP_001",
            "SKU": "CAMISA-PRETA-G",
            "Nome do Produto": "Camisa Polo Algodão",
            "Quantidade": 1,
            "Preço de Produto": 80.00,
        },
        {
            "Número da Venda": "VENDA_MULTILINHA_888",  # Segunda linha da mesma venda
            "ID da Variante": "VAR_TT_002",
            "SKU": "VESTIDO-FLORAL-M",
            "Nome do Produto": "Vestido Midi Floral",
            "Quantidade": 1,
            "Preço de Produto": 110.00,
        },
        {
            "Número da Venda": "VENDA_UNICA_999",  # Venda de item único
            "ID da Variante": "VAR_SHP_001",
            "SKU": "CAMISA-PRETA-G",
            "Nome do Produto": "Camisa Polo Algodão",
            "Quantidade": 1,
            "Preço de Produto": 80.00,
        },
    ])

    processed_multisales, stats_multi = DataProcessor.reconcile_sales_batch(
        df_sales=multi_sales_df,
        canal="Shopee",
        channel_configs=configs,
        products_df=prods_df,
    )

    assert len(processed_multisales) == 3, "Deveriam ter 3 linhas processadas"
    
    # Valida linha 1 da venda multilinhas
    l1 = processed_multisales[0]
    assert l1["numero_venda"] == "VENDA_MULTILINHA_888", f"Número da venda incorreto: {l1['numero_venda']}"
    assert l1["tag_item"] == "Item 1 de 2", f"Tag do item incorreta: {l1['tag_item']}"
    assert l1["pedido_id"] == "VENDA_MULTILINHA_888-01", f"ID taggeado incorreto: {l1['pedido_id']}"

    # Valida linha 2 da venda multilinhas
    l2 = processed_multisales[1]
    assert l2["numero_venda"] == "VENDA_MULTILINHA_888", f"Número da venda incorreto: {l2['numero_venda']}"
    assert l2["tag_item"] == "Item 2 de 2", f"Tag do item incorreta: {l2['tag_item']}"
    assert l2["pedido_id"] == "VENDA_MULTILINHA_888-02", f"ID taggeado incorreto: {l2['pedido_id']}"

    # Valida venda de item único
    l3 = processed_multisales[2]
    assert l3["numero_venda"] == "VENDA_UNICA_999", f"Número da venda incorreto: {l3['numero_venda']}"
    assert l3["tag_item"] == "Item Único", f"Tag do item incorreta: {l3['tag_item']}"
    assert l3["pedido_id"] == "VENDA_UNICA_999", f"ID de item único incorreto: {l3['pedido_id']}"

    # Gravar no banco e verificar persistência
    DatabaseManager.upsert_sales(processed_multisales)
    df_db_check = DatabaseManager.get_sales_dataframe()
    row_l1 = df_db_check[df_db_check["pedido_id"] == "VENDA_MULTILINHA_888-01"].iloc[0]
    assert row_l1["numero_venda"] == "VENDA_MULTILINHA_888"
    assert row_l1["tag_item"] == "Item 1 de 2"

    # Reimportar exatamente o mesmo lote para testar prevenção de duplicação
    novas_reimp, atus_reimp = DatabaseManager.upsert_sales(processed_multisales)
    assert novas_reimp == 0, "Reimportação não deveria criar novas vendas"
    assert atus_reimp == 3, "Deveria ter atualizado as 3 vendas existentes sem duplicar"
    print("✅ 11. Taggeamento de pedidos multilinhas e prevenção de duplicações validados com sucesso.")

    # 12. Teste de Separação Automática de Canais em Planilha Unificada (UpSeller)
    uni_sales_df = pd.DataFrame([
        {
            "Plataforma": "Shopee Brasil",
            "Número da Venda": "PED_UNI_SHP_101",
            "ID da Variante": "VAR_SHP_001",
            "SKU": "CAMISA-PRETA-G",
            "Nome do Produto": "Camisa Polo Algodão",
            "Quantidade": 1,
            "Preço de Produto": 85.00,
        },
        {
            "Plataforma": "TikTok Shop Oficial",
            "Número da Venda": "PED_UNI_TTS_102",
            "ID da Variante": "VAR_TT_002",
            "SKU": "VESTIDO-FLORAL-M",
            "Nome do Produto": "Vestido Midi Floral",
            "Quantidade": 1,
            "Preço de Produto": 110.00,
        },
        {
            "Plataforma": "SHEIN",
            "Número da Venda": "PED_UNI_SHN_103",
            "ID da Variante": "VAR_SHP_001",
            "SKU": "CAMISA-PRETA-G",
            "Nome do Produto": "Camisa Polo Algodão",
            "Quantidade": 1,
            "Preço de Produto": 90.00,
        },
        {
            "Plataforma": "Mercado Livre",
            "Número da Venda": "PED_UNI_MLB_104",
            "ID da Variante": "VAR_TT_002",
            "SKU": "VESTIDO-FLORAL-M",
            "Nome do Produto": "Vestido Midi Floral",
            "Quantidade": 1,
            "Preço de Produto": 130.00,
        },
    ])

    current_configs = DatabaseManager.get_channel_configs()
    processed_uni, stats_uni = DataProcessor.reconcile_sales_batch(
        df_sales=uni_sales_df,
        canal="Automático (Detectar pela Planilha Unificada)",
        channel_configs=current_configs,
        products_df=prods_df,
    )

    assert len(processed_uni) == 4, f"Deveriam ter sido processadas 4 vendas unificadas, obtido {len(processed_uni)}"
    assert stats_uni["canais_detectados"]["Shopee"] == 1
    assert stats_uni["canais_detectados"]["TikTok Shop"] == 1
    assert stats_uni["canais_detectados"]["Shein"] == 1
    assert stats_uni["canais_detectados"]["Mercado Livre"] == 1

    # Valida taxas individuais de cada canal
    u_shp = processed_uni[0]
    assert u_shp["canal"] == "Shopee"
    # Preço base: 105.00, Com: 18% (18.90), Fixa: 5.00 -> total taxas: 23.90
    assert u_shp["total_taxas_canal"] == 23.90, f"Taxas Shopee incorretas: {u_shp['total_taxas_canal']}"

    u_tts = processed_uni[1]
    assert u_tts["canal"] == "TikTok Shop"
    # Preço base: 150.00, Com: 22% (33.00), Fixa: 5.00 -> total taxas: 38.00
    assert u_tts["total_taxas_canal"] == 38.00, f"Taxas TikTok Shop incorretas: {u_tts['total_taxas_canal']}"

    u_shn = processed_uni[2]
    assert u_shn["canal"] == "Shein"
    # Preço base: 105.00, Com: 20% (21.00), Fixa: 4.00 -> total taxas: 25.00
    assert u_shn["total_taxas_canal"] == 25.00, f"Taxas Shein incorretas: {u_shn['total_taxas_canal']}"

    u_mlb = processed_uni[3]
    assert u_mlb["canal"] == "Mercado Livre"
    # Preço base: 150.00, Com: 19% (28.50), Fixa: 6.00 -> total taxas: 34.50
    assert u_mlb["total_taxas_canal"] == 34.50, f"Taxas Mercado Livre incorretas: {u_mlb['total_taxas_canal']}"

    # Gravar no banco e testar persistência e auto-cadastro do canal Mercado Livre
    novas_u, rep_u = DatabaseManager.upsert_sales(processed_uni, ignorar_existentes=True)
    assert novas_u == 4 and rep_u == 0, f"Esperado 4 novas vendas unificadas inseridas, obtido novas={novas_u}, rep={rep_u}"

    all_channels_db = DatabaseManager.get_all_channels()
    assert "Mercado Livre" in all_channels_db, "Mercado Livre deveria estar cadastrado nos canais do banco"

    # Reimportar exatamente o mesmo lote com proteção antiduplicação ativa
    novas_u2, rep_u2 = DatabaseManager.upsert_sales(processed_uni, ignorar_existentes=True)
    assert novas_u2 == 0 and rep_u2 == 4, f"Esperado 0 novas e 4 repetidas ignoradas, obtido novas={novas_u2}, rep={rep_u2}"
    print("✅ 12. Separação multicanal automática em planilha unificada e proteção antiduplicação validadas com 100% de sucesso.")

    print("=" * 60)
    print("TODOS OS TESTES PASSARAM COM SUCESSO! 🎉")
    print("=" * 60)


if __name__ == "__main__":
    run_tests()

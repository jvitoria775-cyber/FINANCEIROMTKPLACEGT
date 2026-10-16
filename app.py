"""
=============================================================================
SISTEMA DE CONTROLE FINANCEIRO E CONCILIAÇÃO DE SAQUE MULTICANAL
=============================================================================
Marketplaces suportados: Shopee, TikTok Shop, Shein
Integração: UpSeller / Relatórios Oficiais de Marketplaces
Tecnologias: Streamlit, Pandas, SQLite3, OpenPyXL
=============================================================================
"""

import base64
import io
import re
import sqlite3
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =============================================================================
# 1. CONFIGURAÇÕES GERAIS E BANCO DE DADOS
# =============================================================================
DB_PATH = Path("ecommerce_financeiro.db")

CHANNELS = ["Shopee", "TikTok Shop", "Shein", "Mercado Livre"]

KNOWN_MARKETPLACES = [
    "Shopee",
    "TikTok Shop",
    "Shein",
    "Mercado Livre",
    "Amazon",
    "Magalu",
    "Kwai Shop",
    "Nuvemshop",
]

DEFAULT_CHANNEL_RATES = {
    "Shopee": {
        "taxa_comissao_pct": 18.0,
        "taxa_fixa_unidade": 5.00,
        "aliquota_imposto_pct": 8.0,
    },
    "TikTok Shop": {
        "taxa_comissao_pct": 22.0,
        "taxa_fixa_unidade": 5.00,
        "aliquota_imposto_pct": 8.0,
    },
    "Shein": {
        "taxa_comissao_pct": 20.0,
        "taxa_fixa_unidade": 4.00,
        "aliquota_imposto_pct": 8.0,
    },
    "Mercado Livre": {
        "taxa_comissao_pct": 19.0,
        "taxa_fixa_unidade": 6.00,
        "aliquota_imposto_pct": 8.0,
    },
}


class DatabaseManager:
    """Gerenciador central do banco de dados relacional SQLite."""

    @classmethod
    def get_all_channels(cls) -> List[str]:
        """Retorna todos os canais cadastrados no banco ou predefinidos, filtrando apenas marketplaces válidos."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT canal FROM configuracoes_canais")
            db_channels = [
                row["canal"] for row in cursor.fetchall()
                if row["canal"] and row["canal"] in KNOWN_MARKETPLACES
            ]
            all_c = list(dict.fromkeys(CHANNELS + db_channels))
            return all_c

    @staticmethod
    def get_connection() -> sqlite3.Connection:
        """Cria e retorna uma conexão com o banco de dados SQLite."""
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn

    @classmethod
    def init_db(cls) -> None:
        """Inicializa as tabelas do banco de dados e popula taxas padrão caso não existam."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()

            # Tabela de configurações de taxas por canal
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS configuracoes_canais (
                    canal TEXT PRIMARY KEY,
                    taxa_comissao_pct REAL NOT NULL,
                    taxa_fixa_unidade REAL NOT NULL,
                    aliquota_imposto_pct REAL NOT NULL
                )
                """
            )

            # Limpeza preventiva de canais espúrios (ex: números de pedidos gravados indevidamente como canais)
            placeholders = ', '.join('?' for _ in KNOWN_MARKETPLACES)
            cursor.execute(f"DELETE FROM configuracoes_canais WHERE canal NOT IN ({placeholders})", KNOWN_MARKETPLACES)

            # Inserir valores padrão para os canais principais se não existirem
            for canal, taxas in DEFAULT_CHANNEL_RATES.items():
                cursor.execute(
                    """
                    INSERT INTO configuracoes_canais 
                    (canal, taxa_comissao_pct, taxa_fixa_unidade, aliquota_imposto_pct)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(canal) DO UPDATE SET
                        taxa_comissao_pct = CASE WHEN configuracoes_canais.taxa_comissao_pct = 20.0 AND excluded.canal = 'Shopee' THEN 18.0 ELSE configuracoes_canais.taxa_comissao_pct END
                    """,
                    (
                        canal,
                        taxas["taxa_comissao_pct"],
                        taxas["taxa_fixa_unidade"],
                        taxas["aliquota_imposto_pct"],
                    ),
                )

            # Tabela cadastral contínua de anúncios/produtos
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS produtos (
                    id_variante TEXT PRIMARY KEY,
                    canal TEXT NOT NULL,
                    id_anuncio TEXT,
                    sku TEXT,
                    nome_anuncio TEXT,
                    variacao TEXT,
                    preco_cheio REAL NOT NULL DEFAULT 0.0,
                    preco_desconto REAL NOT NULL DEFAULT 0.0,
                    preco_base REAL NOT NULL DEFAULT 0.0,
                    custo_cmv REAL NOT NULL DEFAULT 0.0,
                    atualizado_em TIMESTAMP NOT NULL
                )
                """
            )

            # Tabela de vendas processadas e conciliadas
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS vendas_processadas (
                    pedido_id TEXT PRIMARY KEY,
                    numero_venda TEXT,
                    tag_item TEXT,
                    canal TEXT NOT NULL,
                    id_variante TEXT,
                    sku TEXT,
                    nome_anuncio TEXT,
                    quantidade INTEGER NOT NULL,
                    preco_upseller_registrado REAL NOT NULL,
                    preco_base_anuncio REAL NOT NULL,
                    faturamento_bruto REAL NOT NULL,
                    total_taxas_canal REAL NOT NULL,
                    valor_imposto REAL NOT NULL,
                    saque_liquido_previsto REAL NOT NULL,
                    custo_cmv_total REAL NOT NULL,
                    lucro_liquido_real REAL NOT NULL,
                    data_pedido TEXT,
                    data_processamento TIMESTAMP NOT NULL
                )
                """
            )
            # Migrações automáticas caso a tabela já exista
            try:
                cursor.execute("ALTER TABLE vendas_processadas ADD COLUMN data_pedido TEXT")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE vendas_processadas ADD COLUMN numero_venda TEXT")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE vendas_processadas ADD COLUMN tag_item TEXT")
            except sqlite3.OperationalError:
                pass

            # Backfill para registros legados
            cursor.execute("UPDATE vendas_processadas SET numero_venda = pedido_id WHERE numero_venda IS NULL OR numero_venda = ''")
            cursor.execute("UPDATE vendas_processadas SET tag_item = 'Item Único' WHERE tag_item IS NULL OR tag_item = ''")

            # Corrige canais espúrios gravados em vendas_processadas
            cursor.execute(f"SELECT pedido_id, numero_venda, canal FROM vendas_processadas WHERE canal NOT IN ({placeholders})", KNOWN_MARKETPLACES)
            invalid_sales = cursor.fetchall()
            for inv in invalid_sales:
                corr = DataProcessor.detect_channel_from_order_id(inv["numero_venda"] or inv["pedido_id"]) or "Shopee"
                cursor.execute("UPDATE vendas_processadas SET canal = ? WHERE pedido_id = ?", (corr, inv["pedido_id"]))

            conn.commit()

    @classmethod
    def get_channel_configs(cls) -> Dict[str, Dict[str, float]]:
        """Recupera as taxas de todos os canais."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM configuracoes_canais")
            rows = cursor.fetchall()
            configs = {}
            for row in rows:
                configs[row["canal"]] = {
                    "taxa_comissao_pct": float(row["taxa_comissao_pct"]),
                    "taxa_fixa_unidade": float(row["taxa_fixa_unidade"]),
                    "aliquota_imposto_pct": float(row["aliquota_imposto_pct"]),
                }
            return configs

    @classmethod
    def update_channel_config(
        cls,
        canal: str,
        taxa_comissao_pct: float,
        taxa_fixa_unidade: float,
        aliquota_imposto_pct: float,
    ) -> None:
        """Atualiza as taxas de um canal no banco de dados."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO configuracoes_canais (canal, taxa_comissao_pct, taxa_fixa_unidade, aliquota_imposto_pct)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(canal) DO UPDATE SET
                    taxa_comissao_pct = excluded.taxa_comissao_pct,
                    taxa_fixa_unidade = excluded.taxa_fixa_unidade,
                    aliquota_imposto_pct = excluded.aliquota_imposto_pct
                """,
                (canal, taxa_comissao_pct, taxa_fixa_unidade, aliquota_imposto_pct),
            )
            conn.commit()

    @classmethod
    def upsert_products(cls, records: List[Dict[str, Any]]) -> Tuple[int, int]:
        """
        Executa UPSERT idempotente de produtos.
        Se já existir: atualiza preços e metadados, PRESERVANDO o custo_cmv caso o novo seja zero.
        Retorna (novos_inseridos, atualizados).
        """
        if not records:
            return 0, 0

        novos = 0
        atualizados = 0
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with cls.get_connection() as conn:
            cursor = conn.cursor()
            for r in records:
                # Checa se já existe
                cursor.execute(
                    "SELECT id_variante, custo_cmv FROM produtos WHERE id_variante = ?",
                    (str(r["id_variante"]),),
                )
                existing = cursor.fetchone()

                if existing:
                    atualizados += 1
                    # Preserva CMV existente se o novo registro não trouxer CMV definido
                    novo_cmv = float(r.get("custo_cmv", 0.0) or 0.0)
                    cmv_final = novo_cmv if novo_cmv > 0 else float(existing["custo_cmv"])

                    cursor.execute(
                        """
                        UPDATE produtos SET
                            canal = ?,
                            id_anuncio = COALESCE(NULLIF(?, ''), id_anuncio),
                            sku = COALESCE(NULLIF(?, ''), sku),
                            nome_anuncio = COALESCE(NULLIF(?, ''), nome_anuncio),
                            variacao = COALESCE(NULLIF(?, ''), variacao),
                            preco_cheio = ?,
                            preco_desconto = ?,
                            preco_base = ?,
                            custo_cmv = ?,
                            atualizado_em = ?
                        WHERE id_variante = ?
                        """,
                        (
                            r.get("canal", ""),
                            r.get("id_anuncio", ""),
                            r.get("sku", ""),
                            r.get("nome_anuncio", ""),
                            r.get("variacao", ""),
                            float(r.get("preco_cheio", 0.0) or 0.0),
                            float(r.get("preco_desconto", 0.0) or 0.0),
                            float(r.get("preco_base", 0.0) or 0.0),
                            cmv_final,
                            now,
                            str(r["id_variante"]),
                        ),
                    )
                else:
                    novos += 1
                    cursor.execute(
                        """
                        INSERT INTO produtos (
                            id_variante, canal, id_anuncio, sku, nome_anuncio,
                            variacao, preco_cheio, preco_desconto, preco_base,
                            custo_cmv, atualizado_em
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            str(r["id_variante"]),
                            r.get("canal", ""),
                            r.get("id_anuncio", ""),
                            r.get("sku", ""),
                            r.get("nome_anuncio", ""),
                            r.get("variacao", ""),
                            float(r.get("preco_cheio", 0.0) or 0.0),
                            float(r.get("preco_desconto", 0.0) or 0.0),
                            float(r.get("preco_base", 0.0) or 0.0),
                            float(r.get("custo_cmv", 0.0) or 0.0),
                            now,
                        ),
                    )
            conn.commit()

        return novos, atualizados

    @classmethod
    def get_products_dataframe(cls) -> pd.DataFrame:
        """Carrega todos os produtos cadastrados."""
        with cls.get_connection() as conn:
            df = pd.read_sql_query(
                "SELECT * FROM produtos ORDER BY atualizado_em DESC", conn
            )
            return df

    @classmethod
    def update_product_cmv(cls, id_variante: str, novo_cmv: float) -> None:
        """Atualiza individualmente o CMV de um produto por id_variante."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE produtos SET custo_cmv = ? WHERE id_variante = ?",
                (float(novo_cmv), str(id_variante)),
            )
            conn.commit()

    @classmethod
    def batch_update_cmv(cls, updates: List[Tuple[float, str]]) -> None:
        """Atualiza em lote o CMV dos produtos (lista de (novo_cmv, id_variante))."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany(
                "UPDATE produtos SET custo_cmv = ? WHERE id_variante = ?",
                updates,
            )
            conn.commit()

    @classmethod
    def get_existing_pedido_ids(cls) -> set:
        """Retorna o conjunto com todos os pedido_id já gravados no banco de dados."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT pedido_id FROM vendas_processadas")
            return {str(row["pedido_id"]) for row in cursor.fetchall()}

    @classmethod
    def upsert_sales(cls, sales: List[Dict[str, Any]], ignorar_existentes: bool = True) -> Tuple[int, int]:
        """
        Insere ou atualiza vendas processadas de maneira idempotente.
        Se ignorar_existentes=True: pula pedidos que já foram inseridos anteriormente, ignorando-os.
        Se ignorar_existentes=False: atualiza os registros existentes com os novos cálculos.
        Retorna (novas_inseridas, repetidas_ignoradas_ou_atualizadas).
        """
        if not sales:
            return 0, 0

        novas = 0
        repetidas = 0
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with cls.get_connection() as conn:
            cursor = conn.cursor()

            # Garante que qualquer canal detectado na planilha conste em configuracoes_canais (apenas marketplaces válidos)
            for s in sales:
                ch = s.get("canal")
                if ch and ch in KNOWN_MARKETPLACES:
                    cursor.execute("SELECT canal FROM configuracoes_canais WHERE canal = ?", (ch,))
                    if not cursor.fetchone():
                        def_rate = DEFAULT_CHANNEL_RATES.get(
                            ch,
                            {"taxa_comissao_pct": 20.0, "taxa_fixa_unidade": 5.0, "aliquota_imposto_pct": 8.0},
                        )
                        cursor.execute(
                            """
                            INSERT INTO configuracoes_canais (canal, taxa_comissao_pct, taxa_fixa_unidade, aliquota_imposto_pct)
                            VALUES (?, ?, ?, ?)
                            """,
                            (ch, def_rate["taxa_comissao_pct"], def_rate["taxa_fixa_unidade"], def_rate["aliquota_imposto_pct"]),
                        )

            for s in sales:
                pid = str(s["pedido_id"])
                cursor.execute(
                    "SELECT pedido_id FROM vendas_processadas WHERE pedido_id = ?",
                    (pid,),
                )
                exists = cursor.fetchone()

                if exists:
                    repetidas += 1
                    if ignorar_existentes:
                        # Venda repetida ignorada: mantém os dados intactos e não duplica
                        continue
                    else:
                        cursor.execute(
                            """
                            UPDATE vendas_processadas SET
                                numero_venda = ?,
                                tag_item = ?,
                                canal = ?,
                                id_variante = ?,
                                sku = ?,
                                nome_anuncio = ?,
                                quantidade = ?,
                                preco_upseller_registrado = ?,
                                preco_base_anuncio = ?,
                                faturamento_bruto = ?,
                                total_taxas_canal = ?,
                                valor_imposto = ?,
                                saque_liquido_previsto = ?,
                                custo_cmv_total = ?,
                                lucro_liquido_real = ?,
                                data_pedido = COALESCE(NULLIF(?, ''), vendas_processadas.data_pedido),
                                data_processamento = ?
                            WHERE pedido_id = ?
                            """,
                            (
                                str(s.get("numero_venda") or pid),
                                str(s.get("tag_item") or "Item Único"),
                                s.get("canal", ""),
                                s.get("id_variante", ""),
                                s.get("sku", ""),
                                s.get("nome_anuncio", ""),
                                int(s.get("quantidade", 1)),
                                float(s.get("preco_upseller_registrado", 0.0)),
                                float(s.get("preco_base_anuncio", 0.0)),
                                float(s.get("faturamento_bruto", 0.0)),
                                float(s.get("total_taxas_canal", 0.0)),
                                float(s.get("valor_imposto", 0.0)),
                                float(s.get("saque_liquido_previsto", 0.0)),
                                float(s.get("custo_cmv_total", 0.0)),
                                float(s.get("lucro_liquido_real", 0.0)),
                                s.get("data_pedido", now),
                                now,
                                pid,
                            ),
                        )
                else:
                    novas += 1
                    cursor.execute(
                        """
                        INSERT INTO vendas_processadas (
                            pedido_id, numero_venda, tag_item, canal, id_variante, sku, nome_anuncio,
                            quantidade, preco_upseller_registrado, preco_base_anuncio,
                            faturamento_bruto, total_taxas_canal, valor_imposto,
                            saque_liquido_previsto, custo_cmv_total, lucro_liquido_real,
                            data_pedido, data_processamento
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            pid,
                            str(s.get("numero_venda") or pid),
                            str(s.get("tag_item") or "Item Único"),
                            s.get("canal", ""),
                            s.get("id_variante", ""),
                            s.get("sku", ""),
                            s.get("nome_anuncio", ""),
                            int(s.get("quantidade", 1)),
                            float(s.get("preco_upseller_registrado", 0.0)),
                            float(s.get("preco_base_anuncio", 0.0)),
                            float(s.get("faturamento_bruto", 0.0)),
                            float(s.get("total_taxas_canal", 0.0)),
                            float(s.get("valor_imposto", 0.0)),
                            float(s.get("saque_liquido_previsto", 0.0)),
                            float(s.get("custo_cmv_total", 0.0)),
                            float(s.get("lucro_liquido_real", 0.0)),
                            s.get("data_pedido", now),
                            now,
                        ),
                    )
            conn.commit()

        return novas, repetidas

    @classmethod
    def get_sales_dataframe(cls, canal_filtro: Optional[str] = None) -> pd.DataFrame:
        """Recupera o DataFrame com todas as vendas processadas com filtros opcionais."""
        with cls.get_connection() as conn:
            if canal_filtro and canal_filtro != "Todos":
                query = "SELECT * FROM vendas_processadas WHERE canal = ? ORDER BY data_processamento DESC"
                df = pd.read_sql_query(query, conn, params=(canal_filtro,))
            else:
                query = "SELECT * FROM vendas_processadas ORDER BY data_processamento DESC"
                df = pd.read_sql_query(query, conn)
            return df

    @classmethod
    def recalculate_all_sales(cls) -> int:
        """
        Recalcula retroativamente todas as vendas cadastradas aplicando:
        1. Novas taxas de canal do banco.
        2. Preço base e CMV mais atualizados da tabela de produtos.
        Retorna a quantidade de pedidos recalculados.
        """
        with cls.get_connection() as conn:
            cursor = conn.cursor()

            # Pega as vendas atuais
            cursor.execute("SELECT * FROM vendas_processadas")
            sales = [dict(row) for row in cursor.fetchall()]

            if not sales:
                return 0

            # Carrega configs de taxas
            configs = cls.get_channel_configs()

            # Carrega mapa de produtos por id_variante e por sku
            cursor.execute("SELECT * FROM produtos")
            prods = [dict(row) for row in cursor.fetchall()]

            prod_by_variant = {str(p["id_variante"]): p for p in prods if p.get("id_variante")}
            prod_by_sku = {str(p["sku"]).strip().upper(): p for p in prods if p.get("sku")}
            prod_by_title = {DataProcessor.normalize_str(p["nome_anuncio"]): p for p in prods if p.get("nome_anuncio")}

            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            recalculated_count = 0

            for s in sales:
                canal = s["canal"]
                if canal not in KNOWN_MARKETPLACES:
                    canal = DataProcessor.detect_channel_from_order_id(s.get("numero_venda") or s.get("pedido_id")) or "Shopee"

                rates = configs.get(
                    canal,
                    DEFAULT_CHANNEL_RATES.get(
                        canal,
                        {"taxa_comissao_pct": 20.0, "taxa_fixa_unidade": 5.0, "aliquota_imposto_pct": 8.0},
                    ),
                )

                id_var = str(s.get("id_variante") or "")
                sku_val = str(s.get("sku") or "").strip().upper()
                nome_anuncio_s = str(s.get("nome_anuncio") or "").strip()

                # Busca produto correspondente no catálogo (variante -> sku -> título)
                matched_prod = prod_by_variant.get(id_var)
                if not matched_prod and sku_val:
                    matched_prod = prod_by_sku.get(sku_val)
                if not matched_prod and nome_anuncio_s:
                    matched_prod = prod_by_title.get(DataProcessor.normalize_str(nome_anuncio_s))

                if matched_prod:
                    preco_base = float(matched_prod.get("preco_base") or matched_prod.get("preco_cheio") or s["preco_upseller_registrado"])
                    custo_cmv = float(matched_prod.get("custo_cmv") or 0.0)
                    nome_anuncio = matched_prod.get("nome_anuncio") or s["nome_anuncio"]
                else:
                    preco_base = float(s.get("preco_base_anuncio") or s["preco_upseller_registrado"])
                    custo_cmv = 0.0
                    nome_anuncio = s["nome_anuncio"]

                qtd = int(s["quantidade"])
                fat_bruto = round(preco_base * qtd, 2)
                comissao = round(fat_bruto * (rates["taxa_comissao_pct"] / 100.0), 2)
                taxa_fixa = round(rates["taxa_fixa_unidade"] * qtd, 2)
                total_taxas = round(comissao + taxa_fixa, 2)
                imposto = round(fat_bruto * (rates["aliquota_imposto_pct"] / 100.0), 2)
                saque_liq = round(fat_bruto - total_taxas - imposto, 2)
                cmv_total = round(custo_cmv * qtd, 2)
                lucro_real = round(saque_liq - cmv_total, 2)

                cursor.execute(
                    """
                    UPDATE vendas_processadas SET
                        canal = ?,
                        nome_anuncio = ?,
                        preco_base_anuncio = ?,
                        faturamento_bruto = ?,
                        total_taxas_canal = ?,
                        valor_imposto = ?,
                        saque_liquido_previsto = ?,
                        custo_cmv_total = ?,
                        lucro_liquido_real = ?,
                        data_processamento = ?
                    WHERE pedido_id = ?
                    """,
                    (
                        canal,
                        nome_anuncio,
                        preco_base,
                        fat_bruto,
                        total_taxas,
                        imposto,
                        saque_liq,
                        cmv_total,
                        lucro_real,
                        now,
                        s["pedido_id"],
                    ),
                )
                recalculated_count += 1

            conn.commit()
            return recalculated_count

    @classmethod
    def clear_sales_history(cls) -> int:
        """Limpa todo o histórico de vendas processadas, preservando os anúncios e configurações de taxas."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM vendas_processadas")
            count = cursor.fetchone()[0]
            cursor.execute("DELETE FROM vendas_processadas")
            conn.commit()
            return count

    @classmethod
    def get_summary_metrics(cls) -> Dict[str, Any]:
        """Calcula agregados rápidos para os badges do sistema."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) AS total_produtos FROM produtos")
            tot_prod = cursor.fetchone()["total_produtos"]

            cursor.execute("SELECT COUNT(*) AS total_vendas, COALESCE(SUM(faturamento_bruto), 0) AS total_fat FROM vendas_processadas")
            row_vendas = cursor.fetchone()

            return {
                "total_produtos": tot_prod,
                "total_vendas": row_vendas["total_vendas"],
                "total_faturamento": row_vendas["total_fat"],
            }


# =============================================================================
# 2. UTILITÁRIOS DE PROCESSAMENTO, PARSER E NORMALIZAÇÃO
# =============================================================================
class DataProcessor:
    """Funções utilitárias para higienização, mapeamento flexível e cálculos."""

    @staticmethod
    def clean_id_str(val: Any) -> str:
        """Limpa IDs alfanuméricos e numéricos, removendo .0 de inteiros vindos do Excel."""
        if val is None or pd.isna(val):
            return ""
        s = str(val).strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s

    @classmethod
    def detect_channel_from_string(cls, val: Any) -> Optional[str]:
        """Detecta se uma string contém referência a um dos marketplaces conhecidos."""
        if val is None or pd.isna(val):
            return None
        s = str(val).strip().lower()
        if not s:
            return None
        if "shopee" in s:
            return "Shopee"
        elif "tiktok" in s or "tik tok" in s or s == "tts" or "bytedance" in s:
            return "TikTok Shop"
        elif "shein" in s:
            return "Shein"
        elif "mercado" in s or "livre" in s or s == "ml" or "meli" in s:
            return "Mercado Livre"
        elif "amazon" in s:
            return "Amazon"
        elif "magalu" in s or "magazine" in s:
            return "Magalu"
        elif "kwai" in s:
            return "Kwai Shop"
        elif "nuvem" in s:
            return "Nuvemshop"
        return None

    @classmethod
    def detect_channel_from_order_id(cls, order_id: Any) -> Optional[str]:
        """Identifica o marketplace pelo padrão do número do pedido."""
        if order_id is None or pd.isna(order_id):
            return None
        oid = str(order_id).strip().upper()
        if not oid:
            return None
        # Padrão Shopee: ex: 261016Q8536637 ou 241016Q... (data YYMMDD seguida de letra como Q)
        if re.match(r"^\d{6}[A-Z0-9]{6,15}$", oid) and any(c.isalpha() for c in oid[6:]):
            return "Shopee"
        # Padrão TikTok Shop: numérico de 15 a 19 dígitos iniciando com 57 ou 58, ou TTS
        if re.match(r"^5[78]\d{13,17}$", oid) or "TTS" in oid:
            return "TikTok Shop"
        # Padrão Shein: inicia com GSO ou GSON ou GS, ou SHEIN
        if oid.startswith("GSO") or oid.startswith("GS") or "SHEIN" in oid:
            return "Shein"
        # Padrão Mercado Livre: 10 a 16 dígitos numéricos
        if re.match(r"^(20000|4\d{9}|\d{10,16})$", oid):
            return "Mercado Livre"
        return None

    @classmethod
    def normalize_channel(cls, raw_val: Any, default_channel: str = "Shopee") -> str:
        """Normaliza o canal/marketplace para um marketplace válido garantido."""
        detected = cls.detect_channel_from_string(raw_val)
        if detected:
            return detected
        return default_channel if default_channel in KNOWN_MARKETPLACES else "Shopee"

    @staticmethod
    def normalize_str(val: Any) -> str:
        """Remove acentos, espaços extras e converte para minúsculas para matching."""
        if val is None or pd.isna(val):
            return ""
        s = str(val).strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join([c for c in s if not unicodedata.combining(c)])
        return s

    @staticmethod
    def parse_currency(val: Any) -> float:
        """Converte formatos monetários (ex: 'R$ 1.250,50', '49,90', 49.9) para float."""
        if val is None or pd.isna(val):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)

        s = str(val).strip()
        if not s:
            return 0.0

        # Remove 'R$', 'BRL', espaços
        s = re.sub(r"[^\d,\.-]", "", s)

        # Se tiver vírgula e ponto: decide separador decimal
        if "," in s and "." in s:
            # Ex: 1.250,50 -> remove ponto e troca vírgula por ponto
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                # Ex: 1,250.50
                s = s.replace(",", "")
        elif "," in s:
            # Ex: 49,90 -> 49.90
            s = s.replace(",", ".")

        try:
            return float(s)
        except ValueError:
            return 0.0

    @staticmethod
    def parse_date_str(val: Any) -> str:
        """Converte datas em formatos variados para string YYYY-MM-DD HH:MM:SS ou YYYY-MM-DD."""
        if val is None or pd.isna(val):
            return ""
        if isinstance(val, (datetime, pd.Timestamp)):
            return val.strftime("%Y-%m-%d %H:%M:%S")
        s = str(val).strip()
        if not s:
            return ""
        try:
            dt = pd.to_datetime(s, dayfirst=True)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return s[:19]

    @classmethod
    def find_column(cls, df_columns: List[str], synonyms: List[str], exclude_terms: Optional[List[str]] = None) -> Optional[str]:
        """Localiza uma coluna no DataFrame baseando-se em sinônimos normalizados com exclusões preventivas."""
        norm_cols = {col: cls.normalize_str(col) for col in df_columns}
        norm_synonyms = [cls.normalize_str(syn) for syn in synonyms]
        norm_excludes = [cls.normalize_str(ex) for ex in (exclude_terms or [])]

        def is_excluded(col_norm: str) -> bool:
            return any(ex in col_norm for ex in norm_excludes)

        # 1. Match exato
        for orig, norm in norm_cols.items():
            if norm in norm_synonyms and not is_excluded(norm):
                return orig

        # 2. Match parcial (sinônimo contido no nome da coluna)
        for syn in norm_synonyms:
            for orig, norm in norm_cols.items():
                if is_excluded(norm):
                    continue
                if syn in norm:
                    return orig

        return None

    @classmethod
    def read_uploaded_file(cls, uploaded_file) -> pd.DataFrame:
        """Lê arquivo nos formatos .xlsx, .xls ou .csv com fallbacks de encoding."""
        fname = uploaded_file.name.lower()
        if fname.endswith(".csv"):
            for encoding in ["utf-8", "latin1", "cp1252", "utf-8-sig"]:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, encoding=encoding, sep=None, engine="python")
                    return df
                except Exception:
                    continue
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=";")
        else:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file)

    @classmethod
    def map_and_clean_ads_dataframe(cls, df: pd.DataFrame, canal: str) -> pd.DataFrame:
        """Mapeia e normaliza o relatório de anúncios para a estrutura da tabela produtos."""
        cols = list(df.columns)

        col_var_id = cls.find_column(
            cols,
            [
                "id da variante",
                "id da variacao",
                "variation id",
                "codigo da variacao",
                "id do sku",
                "variation_id",
                "id_variante",
                "variante id",
                "id variacao",
            ],
        )

        col_item_id = cls.find_column(
            cols,
            [
                "id do anuncio",
                "id do produto",
                "item id",
                "product id",
                "id_anuncio",
                "id anuncio",
                "id produto",
            ],
        )

        col_sku = cls.find_column(
            cols,
            [
                "sku",
                "sku da variacao",
                "sku do produto",
                "seller sku",
                "codigo sku",
                "código sku",
                "sku id",
            ],
        )

        col_nome = cls.find_column(
            cols,
            [
                "nome do produto",
                "nome do anuncio",
                "titulo",
                "product name",
                "nome",
                "titulo do produto",
            ],
        )

        col_var_nome = cls.find_column(
            cols,
            [
                "variacao",
                "nome da variacao",
                "variacao do produto",
                "variation",
                "opcao",
                "opção",
            ],
        )

        col_preco_cheio = cls.find_column(
            cols,
            [
                "preco original",
                "preco cheio",
                "preco de tabela",
                "preco",
                "original price",
                "price",
                "preço original",
                "preço cheio",
                "preço",
            ],
        )

        col_preco_desc = cls.find_column(
            cols,
            [
                "preco com desconto",
                "preco promocional",
                "preco de promocao",
                "discount price",
                "promotion price",
                "preço com desconto",
                "preço promocional",
                "preco de venda",
                "preço de venda",
            ],
        )

        records = []
        for idx, row in df.iterrows():
            id_variante = cls.clean_id_str(row[col_var_id]) if col_var_id and pd.notna(row[col_var_id]) else ""
            sku_val = cls.clean_id_str(row[col_sku]) if col_sku and pd.notna(row[col_sku]) else ""
            id_anuncio = cls.clean_id_str(row[col_item_id]) if col_item_id and pd.notna(row[col_item_id]) else ""
            nome = str(row[col_nome]).strip() if col_nome and pd.notna(row[col_nome]) else "Anúncio sem Nome"
            variacao = str(row[col_var_nome]).strip() if col_var_nome and pd.notna(row[col_var_nome]) else ""

            # Se não houver id_variante, usa o SKU ou cria identificador baseado no anúncio/índice
            if not id_variante:
                id_variante = sku_val if sku_val else f"VAR_{id_anuncio}_{idx}"

            preco_cheio = cls.parse_currency(row[col_preco_cheio]) if col_preco_cheio else 0.0
            preco_desc = cls.parse_currency(row[col_preco_desc]) if col_preco_desc else 0.0

            # Preço Base: Preço com Desconto (do anúncio) [fallback: Preço Cheio]
            if preco_desc > 0:
                preco_base = preco_desc
            elif preco_cheio > 0:
                preco_base = preco_cheio
            else:
                preco_base = 0.0

            records.append(
                {
                    "id_variante": id_variante,
                    "canal": canal,
                    "id_anuncio": id_anuncio,
                    "sku": sku_val,
                    "nome_anuncio": nome,
                    "variacao": variacao,
                    "preco_cheio": preco_cheio,
                    "preco_desconto": preco_desc,
                    "preco_base": preco_base,
                    "custo_cmv": 0.0,
                }
            )

        return pd.DataFrame(records)

    @classmethod
    def reconcile_sales_batch(
        cls,
        df_sales: pd.DataFrame,
        canal: str,
        channel_configs: Dict[str, Dict[str, float]],
        products_df: pd.DataFrame,
    ) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        """
        Realiza o cruzamento das vendas importadas com a base contínua de produtos/anúncios.
        Aplica a lógica de fallback (id_variante -> sku -> preço de venda).
        Calcula todas as fórmulas financeiras oficiais.
        """
        cols = list(df_sales.columns)

        col_order_id = cls.find_column(
            cols,
            [
                "numero da venda",
                "número da venda",
                "numero de venda",
                "número de venda",
                "id da venda",
                "id de venda",
                "codigo da venda",
                "código da venda",
                "numero do pedido",
                "número do pedido",
                "numero de pedido",
                "número de pedido",
                "id do pedido",
                "id de pedido",
                "codigo do pedido",
                "código do pedido",
                "no. do pedido",
                "no. de pedido",
                "nº do pedido",
                "nº de pedido",
                "n° do pedido",
                "n° de pedido",
                "pedido id",
                "venda id",
                "order id",
                "order no",
                "order sn",
                "order number",
                "encomenda",
                "id da encomenda",
                "número da encomenda",
                "numero da encomenda",
                "pedido_id",
                "pedido",
                "venda",
            ],
        )

        col_var_id = cls.find_column(
            cols,
            [
                "id da variante",
                "id da variacao",
                "variation id",
                "codigo da variacao",
                "id do sku",
                "variation_id",
                "id_variante",
            ],
        )

        col_sku = cls.find_column(
            cols,
            [
                "sku",
                "sku da variacao",
                "sku do produto",
                "seller sku",
                "codigo sku",
                "código sku",
            ],
        )

        col_nome = cls.find_column(
            cols,
            [
                "nome do produto",
                "titulo do produto",
                "product name",
                "nome do anuncio",
                "titulo",
                "nome",
            ],
        )

        col_qtd = cls.find_column(
            cols,
            [
                "quantidade",
                "qtd",
                "quantity",
                "qtd.",
            ],
        )

        col_preco_up = cls.find_column(
            cols,
            [
                "preco de produto",
                "preço de produto",
                "preco unitario",
                "preço unitário",
                "preco de venda",
                "preço de venda",
                "unit price",
                "preco",
                "preço",
            ],
        )

        col_data_pedido = cls.find_column(
            cols,
            [
                "data do pedido",
                "data de criacao",
                "data de criação",
                "order date",
                "data da criacao",
                "data da criação",
                "data pedido",
                "data do pagamento",
                "horario do pedido",
                "horário do pedido",
                "data",
            ],
        )

        col_canal = cls.find_column(
            cols,
            [
                "canal de venda",
                "plataforma de venda",
                "nome da loja",
                "loja de origem",
                "tipo de canal",
                "store name",
                "shop name",
                "plataforma",
                "marketplace",
                "platform",
                "loja",
                "canal",
                "store",
                "channel",
            ],
            exclude_terms=[
                "pedido",
                "order",
                "rastreio",
                "rastreamento",
                "pacote",
                "package",
                "armazem",
                "armazém",
                "numero",
                "número",
                "no.",
                "codigo",
                "código",
                "id do pedido",
                "id pedido",
                "id_pedido",
                "item",
                "produto",
            ],
        )

        # Mapas de produtos para busca rápida O(1)
        prod_map_var = {}
        prod_map_sku = {}
        prod_map_title = {}
        if not products_df.empty:
            for _, p in products_df.iterrows():
                v_id = cls.clean_id_str(p["id_variante"])
                s_val = cls.clean_id_str(p["sku"]).upper() if pd.notna(p["sku"]) else ""
                n_val = cls.normalize_str(p["nome_anuncio"]) if pd.notna(p["nome_anuncio"]) else ""
                if v_id:
                    prod_map_var[v_id] = p
                if s_val:
                    prod_map_sku[s_val] = p
                if n_val and n_val not in prod_map_title:
                    prod_map_title[n_val] = p

        processed_sales = []
        unmapped_items = []
        seen_order_ids: Dict[str, int] = {}
        canais_contagem: Dict[str, int] = {}

        # Contagem prévia de quantas linhas cada venda/pedido possui no lote
        order_counts: Dict[str, int] = {}
        for idx, row in df_sales.iterrows():
            if col_order_id and pd.notna(row[col_order_id]):
                raw_oid = cls.clean_id_str(row[col_order_id])
            else:
                raw_oid = f"VENDA_{idx+1}"
            order_counts[raw_oid] = order_counts.get(raw_oid, 0) + 1

        total_itens = len(df_sales)
        matched_count = 0

        for idx, row in df_sales.iterrows():
            if col_order_id and pd.notna(row[col_order_id]):
                raw_order_id = cls.clean_id_str(row[col_order_id])
            else:
                raw_order_id = f"VENDA_{idx+1}"

            total_itens_venda = order_counts.get(raw_order_id, 1)
            current_item_index = seen_order_ids.get(raw_order_id, 0) + 1
            seen_order_ids[raw_order_id] = current_item_index

            # Taggear o número da venda para que ela não se repita de forma ambígua
            if total_itens_venda > 1:
                tag_item = f"Item {current_item_index} de {total_itens_venda}"
                pedido_id = f"{raw_order_id}-{current_item_index:02d}"
            else:
                tag_item = "Item Único"
                pedido_id = raw_order_id

            # Identificação inteligente do canal de cada linha
            row_canal = None

            # 1. Tenta identificar canal pela coluna dedicada de canal/loja
            if col_canal and pd.notna(row[col_canal]) and str(row[col_canal]).strip():
                row_canal = cls.detect_channel_from_string(row[col_canal])

            # 2. Se não encontrou canal conhecido, varre outras colunas de texto procurando menção aos marketplaces
            if not row_canal:
                for c_col in cols:
                    val_c = str(row[c_col]) if pd.notna(row[c_col]) else ""
                    if val_c:
                        det = cls.detect_channel_from_string(val_c)
                        if det:
                            row_canal = det
                            break

            # 3. Se ainda não encontrou, identifica pelo formato do número do pedido (ex: 261016Q... -> Shopee)
            if not row_canal:
                row_canal = cls.detect_channel_from_order_id(raw_order_id)

            # 4. Fallback final seguro
            if not row_canal:
                if canal and canal not in ["Automático (Detectar pela Planilha Unificada)", "Automático", "Todos", ""]:
                    row_canal = canal
                else:
                    row_canal = "Shopee"

            canais_contagem[row_canal] = canais_contagem.get(row_canal, 0) + 1

            # Busca as taxas específicas do canal desta linha
            rates = channel_configs.get(
                row_canal,
                DEFAULT_CHANNEL_RATES.get(
                    row_canal,
                    {"taxa_comissao_pct": 20.0, "taxa_fixa_unidade": 5.0, "aliquota_imposto_pct": 8.0},
                ),
            )

            id_var = cls.clean_id_str(row[col_var_id]) if col_var_id and pd.notna(row[col_var_id]) else ""
            sku_val = cls.clean_id_str(row[col_sku]) if col_sku and pd.notna(row[col_sku]) else ""
            nome_anuncio = str(row[col_nome]).strip() if col_nome and pd.notna(row[col_nome]) else "Produto sem Nome"

            try:
                qtd = int(row[col_qtd]) if col_qtd and pd.notna(row[col_qtd]) else 1
                qtd = max(1, qtd)
            except Exception:
                qtd = 1

            preco_up = cls.parse_currency(row[col_preco_up]) if col_preco_up else 0.0

            # Algoritmo de Matching no Catálogo e Preço Base
            matched = False
            preco_base = 0.0
            custo_cmv = 0.0

            # 1. Busca por id_variante
            if id_var and id_var in prod_map_var:
                p = prod_map_var[id_var]
                preco_base = float(p["preco_base"]) if float(p["preco_base"]) > 0 else float(p["preco_cheio"])
                custo_cmv = float(p.get("custo_cmv", 0.0) or 0.0)
                if not nome_anuncio or nome_anuncio == "Produto sem Nome":
                    nome_anuncio = str(p["nome_anuncio"])
                matched = True

            # 2. Fallback por SKU
            elif sku_val and sku_val.upper() in prod_map_sku:
                p = prod_map_sku[sku_val.upper()]
                preco_base = float(p["preco_base"]) if float(p["preco_base"]) > 0 else float(p["preco_cheio"])
                custo_cmv = float(p.get("custo_cmv", 0.0) or 0.0)
                if not nome_anuncio or nome_anuncio == "Produto sem Nome":
                    nome_anuncio = str(p["nome_anuncio"])
                matched = True

            # 3. Fallback por Título do Anúncio
            elif nome_anuncio and cls.normalize_str(nome_anuncio) in prod_map_title:
                p = prod_map_title[cls.normalize_str(nome_anuncio)]
                preco_base = float(p["preco_base"]) if float(p["preco_base"]) > 0 else float(p["preco_cheio"])
                custo_cmv = float(p.get("custo_cmv", 0.0) or 0.0)
                matched = True

            # 3. Fallback de Preço se não encontrado
            if not matched:
                preco_base = preco_up
                custo_cmv = 0.0
                unmapped_items.append({"pedido": pedido_id, "numero_venda": raw_order_id, "canal": row_canal, "sku": sku_val, "variante": id_var, "preco_usado": preco_up})
            else:
                matched_count += 1

            # Se mesmo assim preco_base for 0, usa preco_up
            if preco_base <= 0:
                preco_base = preco_up

            # Fórmulas Oficiais de Cálculo com taxas do canal da linha
            faturamento_bruto = round(preco_base * qtd, 2)
            valor_comissao = round(faturamento_bruto * (rates["taxa_comissao_pct"] / 100.0), 2)
            valor_taxa_fixa = round(rates["taxa_fixa_unidade"] * qtd, 2)
            total_taxas_canal = round(valor_comissao + valor_taxa_fixa, 2)
            valor_imposto = round(faturamento_bruto * (rates["aliquota_imposto_pct"] / 100.0), 2)
            saque_liquido_previsto = round(faturamento_bruto - total_taxas_canal - valor_imposto, 2)
            custo_cmv_total = round(custo_cmv * qtd, 2)
            lucro_liquido_real = round(saque_liquido_previsto - custo_cmv_total, 2)

            # Data do Pedido
            data_ped_raw = row[col_data_pedido] if col_data_pedido and pd.notna(row[col_data_pedido]) else None
            data_pedido = cls.parse_date_str(data_ped_raw) if data_ped_raw else datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            processed_sales.append(
                {
                    "pedido_id": pedido_id,
                    "numero_venda": raw_order_id,
                    "tag_item": tag_item,
                    "canal": row_canal,
                    "id_variante": id_var,
                    "sku": sku_val,
                    "nome_anuncio": nome_anuncio,
                    "quantidade": qtd,
                    "preco_upseller_registrado": preco_up,
                    "preco_base_anuncio": preco_base,
                    "faturamento_bruto": faturamento_bruto,
                    "total_taxas_canal": total_taxas_canal,
                    "valor_imposto": valor_imposto,
                    "saque_liquido_previsto": saque_liquido_previsto,
                    "custo_cmv_total": custo_cmv_total,
                    "lucro_liquido_real": lucro_liquido_real,
                    "data_pedido": data_pedido,
                }
            )

        stats = {
            "total_pedidos": total_itens,
            "itens_conciliados": matched_count,
            "itens_nao_cadastrados": len(unmapped_items),
            "taxa_cobertura_pct": round((matched_count / total_itens * 100.0), 1) if total_itens > 0 else 0.0,
            "canais_detectados": canais_contagem,
            "exemplos_nao_cadastrados": unmapped_items[:5],
        }

        return processed_sales, stats


# =============================================================================
# 3. EXPORTADOR EXCEL FORMATADO (.XLSX) COM OPENPYXL
# =============================================================================
class ExcelExporter:
    """Gera planilhas formatadas e estilizadas para download de conciliação."""

    @staticmethod
    def export_sales_to_excel(df: pd.DataFrame) -> bytes:
        """Exporta o DataFrame de vendas com formatação financeira, cabeçalhos e totais."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Conciliação Financeira"
        ws.views.sheetView[0].showGridLines = True

        # Paleta de Cores
        COR_HEADER = "1B365D"  # Azul Marinho Corporativo
        COR_ZEBRA = "F8FAFC"   # Cinza Suave
        COR_TOTAL = "E2E8F0"   # Cinza Médio

        header_fill = PatternFill(start_color=COR_HEADER, end_color=COR_HEADER, fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        border_thin = Side(border_style="thin", color="CBD5E1")
        cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

        # Mapeamento e Ordem de Colunas Amigáveis
        col_labels = [
            ("numero_venda", "Nº da Venda", "@"),
            ("tag_item", "Tag / Linha", "@"),
            ("pedido_id", "ID Linha", "@"),
            ("data_pedido", "Data do Pedido", "@"),
            ("canal", "Canal / Marketplace", "@"),
            ("sku", "SKU", "@"),
            ("id_variante", "ID Variante", "@"),
            ("nome_anuncio", "Nome do Anúncio", "@"),
            ("quantidade", "Qtd", "#,##0"),
            ("preco_upseller_registrado", "Preço UpSeller (Cliente)", "R$ #,##0.00"),
            ("preco_base_anuncio", "Preço Base Anunciado", "R$ #,##0.00"),
            ("faturamento_bruto", "Faturamento Bruto Real", "R$ #,##0.00"),
            ("total_taxas_canal", "Total Taxas Canal", "R$ #,##0.00"),
            ("valor_imposto", "Imposto (8%)", "R$ #,##0.00"),
            ("saque_liquido_previsto", "Saque Líquido Previsto", "R$ #,##0.00"),
            ("custo_cmv_total", "CMV Total", "R$ #,##0.00"),
            ("lucro_liquido_real", "Lucro Líquido Real", "R$ #,##0.00"),
            ("data_processamento", "Data Processamento", "YYYY-MM-DD HH:MM"),
        ]

        # Escrever Cabeçalho
        for col_idx, (_, label, _) in enumerate(col_labels, start=1):
            cell = ws.cell(row=1, column=col_idx, value=label)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
            cell.border = cell_border
        ws.row_dimensions[1].height = 28

        # Escrever Dados
        row_start = 2
        for r_idx, row in df.iterrows():
            current_row = row_start + r_idx
            is_even = r_idx % 2 == 0
            row_fill = PatternFill(start_color="FFFFFF" if is_even else COR_ZEBRA, fill_type="solid")

            for c_idx, (col_name, _, num_fmt) in enumerate(col_labels, start=1):
                val = row.get(col_name, "")
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                cell.fill = row_fill
                cell.border = cell_border
                cell.number_format = num_fmt

                if col_name in ["quantidade", "preco_upseller_registrado", "preco_base_anuncio",
                                "faturamento_bruto", "total_taxas_canal", "valor_imposto",
                                "saque_liquido_previsto", "custo_cmv_total", "lucro_liquido_real"]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

            ws.row_dimensions[current_row].height = 20

        # Linha de Totais
        total_row = row_start + len(df)
        ws.cell(row=total_row, column=1, value="TOTAIS").font = Font(name="Calibri", size=11, bold=True)
        total_fill = PatternFill(start_color=COR_TOTAL, end_color=COR_TOTAL, fill_type="solid")

        sum_cols = [
            "quantidade", "faturamento_bruto", "total_taxas_canal",
            "valor_imposto", "saque_liquido_previsto", "custo_cmv_total", "lucro_liquido_real"
        ]

        for c_idx, (col_name, _, num_fmt) in enumerate(col_labels, start=1):
            cell = ws.cell(row=total_row, column=c_idx)
            cell.fill = total_fill
            cell.border = cell_border
            cell.font = Font(name="Calibri", size=11, bold=True)

            if col_name in sum_cols and len(df) > 0:
                col_letter = get_column_letter(c_idx)
                cell.value = f"=SUM({col_letter}2:{col_letter}{total_row-1})"
                cell.number_format = num_fmt
                cell.alignment = Alignment(horizontal="right", vertical="center")

        ws.row_dimensions[total_row].height = 24

        # Ajuste Automático de Largura de Coluna
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()


# =============================================================================
# 4. INTERFACE GRÁFICA STREAMLIT (app.py)
# =============================================================================
def get_logo_base64() -> Optional[str]:
    """Recupera e codifica a logo em base64 para uso na interface."""
    logo_path = Path("assets/logo_gratitude.png")
    if logo_path.exists():
        try:
            with open(logo_path, "rb") as f:
                return base64.b64encode(f.read()).decode()
        except Exception:
            return None
    return None


def apply_custom_css():
    """Aplica identidade visual sofisticada, moderna e com estética executiva da marca Gratitude Têxtil."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Cinzel:wght@600;700;800&display=swap');

        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        }

        .main {
            background-color: #f8fafc;
        }

        /* Banner de Cabeçalho Executivo */
        .header-banner {
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: linear-gradient(135deg, #0b1120 0%, #172554 50%, #0f172a 100%);
            padding: 22px 28px;
            border-radius: 16px;
            border: 1px solid rgba(212, 175, 55, 0.4);
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2), 0 0 15px rgba(212, 175, 55, 0.1);
            margin-bottom: 22px;
        }

        .brand-title {
            color: #d4af37;
            font-size: 1.85rem;
            font-weight: 800;
            letter-spacing: 0.06em;
            font-family: 'Cinzel', serif, sans-serif;
            background: linear-gradient(135deg, #fef08a 0%, #d4af37 50%, #b45309 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin: 0;
            line-height: 1.1;
        }

        .brand-subtitle {
            color: #94a3b8;
            font-size: 0.9rem;
            font-weight: 500;
            margin-top: 4px;
        }

        /* Cards de Métricas Premium */
        .metric-card {
            background: #ffffff;
            border-radius: 14px;
            padding: 20px 16px;
            box-shadow: 0 4px 15px -2px rgba(15, 23, 42, 0.05), 0 2px 6px -1px rgba(15, 23, 42, 0.03);
            border: 1px solid #e2e8f0;
            text-align: center;
            margin-bottom: 12px;
            position: relative;
            overflow: hidden;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        .metric-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 12px 24px -4px rgba(15, 23, 42, 0.08);
        }

        .metric-card::before {
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
        }

        .card-gold::before { background: linear-gradient(90deg, #ecc94b, #d4af37, #b45309); }
        .card-red::before { background: linear-gradient(90deg, #f87171, #ef4444); }
        .card-amber::before { background: linear-gradient(90deg, #fbbf24, #f59e0b); }
        .card-blue::before { background: linear-gradient(90deg, #60a5fa, #2563eb); }
        .card-green::before { background: linear-gradient(90deg, #34d399, #059669); }

        .metric-title {
            color: #64748b;
            font-size: 0.78rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 6px;
        }

        .metric-value {
            color: #0f172a;
            font-size: 1.6rem;
            font-weight: 800;
            letter-spacing: -0.02em;
        }

        .metric-sub {
            font-size: 0.8rem;
            font-weight: 600;
            margin-top: 4px;
        }

        /* Banner de Diagnóstico Moderno */
        .diagnostic-card {
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            border: 1px solid #cbd5e1;
            border-left: 5px solid #d4af37;
            border-radius: 12px;
            padding: 16px 20px;
            margin: 16px 0;
            box-shadow: 0 4px 12px -2px rgba(0, 0, 0, 0.04);
        }

        /* Custom Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 6px;
            background-color: #f1f5f9;
            padding: 6px;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
        }

        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 8px 18px;
            font-weight: 600;
            color: #475569;
        }

        .stTabs [aria-selected="true"] {
            background-color: #ffffff !important;
            color: #b45309 !important;
            box-shadow: 0 2px 6px rgba(0,0,0,0.06);
        }

        /* Tags de Canais */
        .tag-shopee {
            background-color: #fff1ee;
            color: #ee4d2d;
            padding: 4px 10px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.8rem;
            border: 1px solid #ffedd5;
        }
        .tag-tiktok {
            background-color: #f8fafc;
            color: #0f172a;
            padding: 4px 10px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.8rem;
            border: 1px solid #e2e8f0;
        }
        .tag-shein {
            background-color: #fefce8;
            color: #854d0e;
            padding: 4px 10px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.8rem;
            border: 1px solid #fef08a;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def format_currency_br(val: float) -> str:
    """Formata valor float para formato monetário brasileiro R$ 1.234,56."""
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def render_dashboard_charts(df_sales: pd.DataFrame):
    """Gera e renderiza 4 gráficos modernos e interativos com Plotly."""
    if df_sales.empty:
        return

    st.markdown(
        """
        <div style="margin: 22px 0 14px 0;">
            <h3 style="margin: 0; font-size: 1.25rem; font-weight: 800; color: #0f172a;">
                📈 Inteligência Visual & Desempenho Financeiro
            </h3>
            <span style="color: #64748b; font-size: 0.85rem;">Gráficos gerenciais calculados em tempo real sobre os pedidos filtrados</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_g1, col_g2 = st.columns(2)

    # 1. Donut de Composição do Faturamento Bruto Real
    with col_g1:
        total_saque = df_sales["saque_liquido_previsto"].sum()
        total_taxas = df_sales["total_taxas_canal"].sum()
        total_imposto = df_sales["valor_imposto"].sum()
        total_cmv = df_sales["custo_cmv_total"].sum()
        total_fat = df_sales["faturamento_bruto"].sum()

        labels = ["Saque Líquido Previsto", "Custo Produtos (CMV)", "Taxas Marketplace", "Imposto DAS (8%)"]
        values = [max(0.0, float(total_saque)), max(0.0, float(total_cmv)), max(0.0, float(total_taxas)), max(0.0, float(total_imposto))]
        colors = ["#D4AF37", "#475569", "#EF4444", "#F59E0B"]

        fig_donut = go.Figure(
            data=[
                go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.6,
                    marker=dict(colors=colors, line=dict(color="#ffffff", width=2)),
                    textinfo="percent",
                    hovertemplate="<b>%{label}</b><br>Valor: R$ %{value:,.2f}<br>Participação: %{percent}<extra></extra>",
                )
            ]
        )
        fig_donut.update_layout(
            title=dict(
                text="<b>Destinação do Faturamento Bruto Real</b>",
                font=dict(size=14, color="#1e293b", family="Plus Jakarta Sans"),
            ),
            legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5, font=dict(size=11)),
            margin=dict(l=10, r=10, t=40, b=55),
            height=330,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            annotations=[
                dict(
                    text=f"Total Fat.<br><b>R$ {total_fat:,.0f}</b>",
                    x=0.5,
                    y=0.5,
                    font_size=13,
                    showarrow=False,
                    font_color="#0f172a",
                )
            ],
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    # 2. Comparativo de Marketplaces (Barra Agrupada)
    with col_g2:
        df_canal = df_sales.groupby("canal").agg({
            "faturamento_bruto": "sum",
            "saque_liquido_previsto": "sum",
            "lucro_liquido_real": "sum",
        }).reset_index()

        fig_bar = go.Figure()
        fig_bar.add_trace(
            go.Bar(
                name="Fat. Bruto Real",
                x=df_canal["canal"],
                y=df_canal["faturamento_bruto"],
                marker_color="#1E293B",
                hovertemplate="<b>%{x}</b><br>Fat. Bruto: R$ %{y:,.2f}<extra></extra>",
            )
        )
        fig_bar.add_trace(
            go.Bar(
                name="Saque Previsto",
                x=df_canal["canal"],
                y=df_canal["saque_liquido_previsto"],
                marker_color="#D4AF37",
                hovertemplate="<b>%{x}</b><br>Saque Previsto: R$ %{y:,.2f}<extra></extra>",
            )
        )
        fig_bar.add_trace(
            go.Bar(
                name="Lucro Líquido",
                x=df_canal["canal"],
                y=df_canal["lucro_liquido_real"],
                marker_color="#10B981",
                hovertemplate="<b>%{x}</b><br>Lucro Real: R$ %{y:,.2f}<extra></extra>",
            )
        )
        fig_bar.update_layout(
            barmode="group",
            title=dict(
                text="<b>Comparativo por Marketplace</b>",
                font=dict(size=14, color="#1e293b", family="Plus Jakarta Sans"),
            ),
            legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5, font=dict(size=11)),
            margin=dict(l=10, r=10, t=40, b=55),
            height=330,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#e2e8f0", tickprefix="R$ "),
            xaxis=dict(gridcolor="rgba(0,0,0,0)"),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    col_g3, col_g4 = st.columns(2)

    # 3. Evolução Diária de Faturamento e Saques
    with col_g3:
        df_temp = df_sales.copy()
        col_data_ref = "data_pedido" if ("data_pedido" in df_temp.columns and df_temp["data_pedido"].notna().any()) else "data_processamento"
        df_temp["data_curta"] = pd.to_datetime(df_temp[col_data_ref].fillna(df_temp["data_processamento"]), errors="coerce").dt.date

        df_time = df_temp.groupby("data_curta").agg({
            "faturamento_bruto": "sum",
            "saque_liquido_previsto": "sum",
            "lucro_liquido_real": "sum",
        }).reset_index().sort_values("data_curta")

        fig_time = go.Figure()
        fig_time.add_trace(
            go.Scatter(
                x=df_time["data_curta"].astype(str),
                y=df_time["faturamento_bruto"],
                mode="lines+markers",
                name="Fat. Bruto Real",
                line=dict(color="#1E293B", width=2.5),
                hovertemplate="Data: %{x}<br>Fat. Bruto: R$ %{y:,.2f}<extra></extra>",
            )
        )
        fig_time.add_trace(
            go.Scatter(
                x=df_time["data_curta"].astype(str),
                y=df_time["saque_liquido_previsto"],
                mode="lines+markers",
                name="Saque Previsto",
                line=dict(color="#D4AF37", width=3),
                fill="tozeroy",
                fillcolor="rgba(212, 175, 55, 0.12)",
                hovertemplate="Data: %{x}<br>Saque Previsto: R$ %{y:,.2f}<extra></extra>",
            )
        )
        fig_time.add_trace(
            go.Scatter(
                x=df_time["data_curta"].astype(str),
                y=df_time["lucro_liquido_real"],
                mode="lines+markers",
                name="Lucro Real",
                line=dict(color="#10B981", width=2, dash="dot"),
                hovertemplate="Data: %{x}<br>Lucro Real: R$ %{y:,.2f}<extra></extra>",
            )
        )
        fig_time.update_layout(
            title=dict(
                text="<b>Evolução Diária de Vendas & Saques</b>",
                font=dict(size=14, color="#1e293b", family="Plus Jakarta Sans"),
            ),
            legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5, font=dict(size=11)),
            margin=dict(l=10, r=10, t=40, b=55),
            height=330,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#e2e8f0", tickprefix="R$ "),
            xaxis=dict(gridcolor="rgba(0,0,0,0)"),
        )
        st.plotly_chart(fig_time, use_container_width=True)

    # 4. Top 5 Produtos Mais Lucrativos
    with col_g4:
        df_prod = df_sales.groupby("nome_anuncio").agg({
            "lucro_liquido_real": "sum",
            "faturamento_bruto": "sum",
            "quantidade": "sum",
        }).reset_index().sort_values("lucro_liquido_real", ascending=True).tail(5)

        df_prod["nome_curto"] = df_prod["nome_anuncio"].apply(lambda s: (s[:26] + "...") if len(str(s)) > 26 else str(s))

        fig_top = go.Figure(
            go.Bar(
                x=df_prod["lucro_liquido_real"],
                y=df_prod["nome_curto"],
                orientation="h",
                marker=dict(
                    color=df_prod["lucro_liquido_real"],
                    colorscale=[[0, "#E5C158"], [1, "#B8860B"]],
                    line=dict(color="#ffffff", width=1),
                ),
                text=df_prod["lucro_liquido_real"].apply(lambda v: f"R$ {v:,.2f}"),
                textposition="auto",
                hovertemplate="<b>%{y}</b><br>Lucro Real: R$ %{x:,.2f}<extra></extra>",
            )
        )
        fig_top.update_layout(
            title=dict(
                text="<b>Top 5 Produtos por Lucro Líquido Real</b>",
                font=dict(size=14, color="#1e293b", family="Plus Jakarta Sans"),
            ),
            margin=dict(l=10, r=10, t=40, b=55),
            height=330,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(gridcolor="#e2e8f0", tickprefix="R$ "),
            yaxis=dict(gridcolor="rgba(0,0,0,0)"),
        )
        st.plotly_chart(fig_top, use_container_width=True)


def main():
    st.set_page_config(
        page_title="GRATITUDE TÊXTIL • Hub Financeiro & Conciliação",
        page_icon="assets/logo_gratitude.png" if Path("assets/logo_gratitude.png").exists() else "💰",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_custom_css()

    # Inicializa banco de dados
    DatabaseManager.init_db()

    # Sidebar com Identidade Visual Gratitude Têxtil
    with st.sidebar:
        logo_b64 = get_logo_base64()
        if logo_b64:
            st.markdown(
                f"""
                <div style="text-align: center; padding: 10px 0 16px 0;">
                    <img src="data:image/png;base64,{logo_b64}" style="max-width: 140px; border-radius: 12px; filter: drop-shadow(0 6px 12px rgba(0,0,0,0.35));" />
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown("### GRATITUDE TÊXTIL")

        st.markdown(
            """
            <div style="text-align: center; margin-bottom: 18px;">
                <span style="color: #d4af37; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.06em; font-family: 'Cinzel', serif;">GRATITUDE TÊXTIL</span><br/>
                <span style="color: #64748b; font-size: 0.8rem; font-weight: 500;">Gestão & Conciliação Multicanal</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.divider()

        st.markdown("#### ⚡ Parâmetros Rápidos")
        configs_sb = DatabaseManager.get_channel_configs()
        for canal, rates in configs_sb.items():
            cor_bullet = "#EE4D2D" if canal == "Shopee" else ("#0F172A" if canal == "TikTok Shop" else "#CA8A04")
            st.markdown(
                f"""
                <div style="background: white; border-left: 4px solid {cor_bullet}; padding: 8px 12px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #e2e8f0;">
                    <b style="color: #0f172a; font-size: 0.85rem;">{canal}</b><br/>
                    <span style="font-size: 0.76rem; color: #64748b;">Comissão: <b>{rates['taxa_comissao_pct']}%</b> | Fixa: <b>R$ {rates['taxa_fixa_unidade']:.2f}</b></span>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.divider()
        st.caption("Gratitude Têxtil • V2.0 • UpSeller Hub")

    # Cabeçalho Principal com Banner Executivo
    metrics_db = DatabaseManager.get_summary_metrics()
    logo_b64 = get_logo_base64()
    logo_img_tag = (
        f'<img src="data:image/png;base64,{logo_b64}" style="height: 72px; border-radius: 10px; filter: drop-shadow(0 4px 10px rgba(0,0,0,0.35));" />'
        if logo_b64
        else ""
    )

    st.markdown(
        f"""
        <div class="header-banner">
            <div style="display: flex; align-items: center; gap: 20px;">
                {logo_img_tag}
                <div>
                    <h1 class="brand-title">GRATITUDE TÊXTIL</h1>
                    <div class="brand-subtitle">Controle Financeiro & Conciliação de Saque Multicanal • Shopee • TikTok Shop • Shein</div>
                </div>
            </div>
            <div style="text-align: right; background: rgba(255,255,255,0.06); padding: 12px 20px; border-radius: 12px; border: 1px solid rgba(212,175,55,0.25);">
                <span style="color: #fef08a; font-weight: 700; font-size: 0.85rem;">● Sistema Conectado</span><br/>
                <span style="color: #f1f5f9; font-size: 0.88rem;"><b>{metrics_db['total_produtos']}</b> Produtos &nbsp;|&nbsp; <b>{metrics_db['total_vendas']}</b> Vendas</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Criação das 4 Abas Principais
    tab_dashboard, tab_anuncios, tab_vendas, tab_config = st.tabs(
        [
            "📊 Dashboard & Indicadores",
            "📦 Importação de Anúncios",
            "🛒 Importação & Conciliação de Vendas",
            "⚙️ Gerenciador de Taxas e CMV",
        ]
    )

    # =========================================================================
    # ABA 1: DASHBOARD & INDICADORES
    # =========================================================================
    with tab_dashboard:
        st.subheader("Painel de Desempenho e Conciliação Financeira")

        # Barra de Filtros Avançada
        with st.container():
            col_f1, col_f2, col_f3 = st.columns([1.3, 1.4, 2.3])
            with col_f1:
                filtro_canal = st.selectbox("Canal / Marketplace:", ["Todos"] + DatabaseManager.get_all_channels(), key="dash_canal")
            with col_f2:
                preset_data = st.selectbox(
                    "Período Rápido:",
                    [
                        "Todo o Histórico",
                        "Hoje",
                        "Ontem",
                        "Últimos 7 dias",
                        "Últimos 30 dias",
                        "Mês Atual",
                        "Personalizado",
                    ],
                    key="dash_preset_data",
                )

            # Cálculo de datas baseado no preset
            hoje = datetime.now().date()
            if preset_data == "Hoje":
                d_ini_def = hoje
                d_fim_def = hoje
            elif preset_data == "Ontem":
                d_ini_def = hoje - timedelta(days=1)
                d_fim_def = hoje - timedelta(days=1)
            elif preset_data == "Últimos 7 dias":
                d_ini_def = hoje - timedelta(days=6)
                d_fim_def = hoje
            elif preset_data == "Últimos 30 dias":
                d_ini_def = hoje - timedelta(days=29)
                d_fim_def = hoje
            elif preset_data == "Mês Atual":
                d_ini_def = hoje.replace(day=1)
                d_fim_def = hoje
            else:
                d_ini_def = None
                d_fim_def = None

            with col_f3:
                if preset_data != "Todo o Histórico":
                    intervalo_data = st.date_input(
                        "Intervalo de Datas:",
                        value=(d_ini_def, d_fim_def) if d_ini_def else (),
                        format="DD/MM/YYYY",
                        key="dash_intervalo_data",
                    )
                else:
                    st.text_input("Intervalo de Datas:", value="Todo o histórico selecionado", disabled=True)
                    intervalo_data = None

            col_s1, col_s2, col_s3 = st.columns([2.5, 1.3, 1.0])
            with col_s1:
                filtro_busca = st.text_input("Buscar por Pedido, SKU ou Nome:", placeholder="Ex: PED-1234, CAMISA-P...", key="dash_busca")
            with col_s2:
                campo_data_filtro = st.selectbox("Filtrar data por:", ["Data do Pedido", "Data de Processamento"], key="dash_campo_data")
            with col_s3:
                st.write("")
                st.write("")
                btn_refresh = st.button("🔄 Atualizar", use_container_width=True)

        # Carregar Vendas do Banco
        df_sales = DatabaseManager.get_sales_dataframe(filtro_canal)

        if not df_sales.empty:
            # Seleciona a coluna para comparação de data
            col_dt_usada = "data_pedido" if (campo_data_filtro == "Data do Pedido" and "data_pedido" in df_sales.columns and df_sales["data_pedido"].notna().any()) else "data_processamento"
            dt_series = pd.to_datetime(df_sales[col_dt_usada].fillna(df_sales["data_processamento"]), errors="coerce")

            # Aplicação do filtro de data
            if intervalo_data:
                if isinstance(intervalo_data, (list, tuple)):
                    if len(intervalo_data) == 2:
                        d_inicio, d_fim = intervalo_data
                        mask_data = (dt_series.dt.date >= d_inicio) & (dt_series.dt.date <= d_fim)
                        df_sales = df_sales[mask_data]
                    elif len(intervalo_data) == 1:
                        d_inicio = intervalo_data[0]
                        mask_data = (dt_series.dt.date == d_inicio)
                        df_sales = df_sales[mask_data]

            # Aplicação do filtro de busca textual
            if filtro_busca:
                b_norm = DataProcessor.normalize_str(filtro_busca)
                num_venda_col = df_sales["numero_venda"].astype(str).str.lower() if "numero_venda" in df_sales.columns else pd.Series(dtype=str)
                tag_item_col = df_sales["tag_item"].astype(str).str.lower() if "tag_item" in df_sales.columns else pd.Series(dtype=str)
                df_sales = df_sales[
                    df_sales["pedido_id"].astype(str).str.lower().str.contains(b_norm)
                    | num_venda_col.str.contains(b_norm)
                    | tag_item_col.str.contains(b_norm)
                    | df_sales["sku"].astype(str).str.lower().str.contains(b_norm)
                    | df_sales["nome_anuncio"].astype(str).str.lower().str.contains(b_norm)
                ]

        if df_sales.empty:
            st.warning("⚠️ Nenhuma venda processada encontrada para o filtro selecionado (canal, data ou busca). Tente ampliar o período ou alterar os filtros.")
        else:
            if intervalo_data and len(intervalo_data) >= 1:
                d_ini_txt = intervalo_data[0].strftime("%d/%m/%Y")
                d_fim_txt = intervalo_data[1].strftime("%d/%m/%Y") if len(intervalo_data) == 2 else d_ini_txt
                st.info(f"📅 **Período Selecionado ({campo_data_filtro}):** de `{d_ini_txt}` até `{d_fim_txt}` — **{len(df_sales)} pedidos** conciliados.")
            # Cálculos Agregados de KPIs
            total_fat_real = df_sales["faturamento_bruto"].sum()
            total_fat_upseller = (df_sales["preco_upseller_registrado"] * df_sales["quantidade"]).sum()
            total_taxas = df_sales["total_taxas_canal"].sum()
            total_imposto = df_sales["valor_imposto"].sum()
            total_saque_liq = df_sales["saque_liquido_previsto"].sum()
            total_cmv = df_sales["custo_cmv_total"].sum()
            total_lucro_real = df_sales["lucro_liquido_real"].sum()
            diferenca_cupons = total_fat_real - total_fat_upseller
            margem_liq_pct = (total_lucro_real / total_fat_real * 100.0) if total_fat_real > 0 else 0.0

            # Cards de Indicadores (KPIs)
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                st.markdown(
                    f"""
                    <div class="metric-card card-gold">
                        <div class="metric-title">Faturamento Bruto Real</div>
                        <div class="metric-value">{format_currency_br(total_fat_real)}</div>
                        <div class="metric-sub" style="color: #64748b;">Base Anunciada</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c2:
                st.markdown(
                    f"""
                    <div class="metric-card card-red">
                        <div class="metric-title">Taxas Marketplace</div>
                        <div class="metric-value" style="color: #ef4444;">{format_currency_br(total_taxas)}</div>
                        <div class="metric-sub" style="color: #ef4444;">Comissão + Fixa</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c3:
                st.markdown(
                    f"""
                    <div class="metric-card card-amber">
                        <div class="metric-title">Imposto Simples (8%)</div>
                        <div class="metric-value" style="color: #f59e0b;">{format_currency_br(total_imposto)}</div>
                        <div class="metric-sub" style="color: #f59e0b;">Provisão DAS</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c4:
                st.markdown(
                    f"""
                    <div class="metric-card card-blue">
                        <div class="metric-title" style="color: #2563eb;">Saque Líquido Previsto</div>
                        <div class="metric-value" style="color: #2563eb;">{format_currency_br(total_saque_liq)}</div>
                        <div class="metric-sub" style="color: #2563eb;">Repasse na Conta</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c5:
                cor_lucro = "#059669" if total_lucro_real >= 0 else "#ef4444"
                st.markdown(
                    f"""
                    <div class="metric-card card-green">
                        <div class="metric-title" style="color: {cor_lucro};">Lucro Líquido Real</div>
                        <div class="metric-value" style="color: {cor_lucro};">{format_currency_br(total_lucro_real)}</div>
                        <div class="metric-sub" style="color: {cor_lucro};">Margem: {margem_liq_pct:.1f}%</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            # Barra de Destaque da Distorção de Cupons
            st.markdown(
                f"""
                <div class="diagnostic-card">
                    <b style="color: #b45309; font-size: 0.95rem;">💡 Diagnóstico da Distorção de Cupons:</b><br/>
                    O valor pago pelo cliente (registrado no UpSeller) foi de <b>{format_currency_br(total_fat_upseller)}</b>.
                    Porém, o Faturamento Base Anunciado real é de <b>{format_currency_br(total_fat_real)}</b> 
                    (Diferença de <b style="color: #059669;">+{format_currency_br(diferenca_cupons)}</b> subsidiada por cupons da plataforma). 
                    As comissões e o saque foram calculados sobre o preço anunciado, garantindo precisão financeira absoluta para a <b>Gratitude Têxtil</b>.
                </div>
                """,
                unsafe_allow_html=True,
            )

            # Renderização dos 4 Gráficos Interativos Plotly
            render_dashboard_charts(df_sales)
            st.divider()

            # Botão de Exportação Excel
            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                st.caption(f"Exibindo {len(df_sales)} pedidos processados.")
            with col_b2:
                excel_bytes = ExcelExporter.export_sales_to_excel(df_sales)
                st.download_button(
                    label="📥 Baixar Relatório Formatado (.xlsx)",
                    data=excel_bytes,
                    file_name=f"conciliacao_financeira_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            # Tabela Detalhada com Formatação
            df_display = df_sales.copy()
            col_monetarias = [
                "preco_upseller_registrado",
                "preco_base_anuncio",
                "faturamento_bruto",
                "total_taxas_canal",
                "valor_imposto",
                "saque_liquido_previsto",
                "custo_cmv_total",
                "lucro_liquido_real",
            ]
            for c in col_monetarias:
                df_display[c] = df_display[c].apply(format_currency_br)

            col_renames = {
                "numero_venda": "Nº da Venda",
                "tag_item": "Tag / Linha",
                "pedido_id": "ID Linha",
                "data_pedido": "Data do Pedido",
                "canal": "Canal",
                "sku": "SKU",
                "id_variante": "ID Variante",
                "nome_anuncio": "Nome do Anúncio",
                "quantidade": "Qtd",
                "preco_upseller_registrado": "Preço UpSeller",
                "preco_base_anuncio": "Preço Base Anúncio",
                "faturamento_bruto": "Fat. Bruto Real",
                "total_taxas_canal": "Taxas Canal",
                "valor_imposto": "Imposto DAS",
                "saque_liquido_previsto": "Saque Previsto",
                "custo_cmv_total": "CMV Total",
                "lucro_liquido_real": "Lucro Real",
                "data_processamento": "Processado Em",
            }
            cols_to_show = [c for c in col_renames.keys() if c in df_display.columns]
            df_display = df_display[cols_to_show].rename(columns=col_renames)

            st.dataframe(
                df_display,
                use_container_width=True,
                hide_index=True,
                height=450,
            )

    # =========================================================================
    # ABA 2: IMPORTAÇÃO DE ANÚNCIOS (PRODUTOS)
    # =========================================================================
    with tab_anuncios:
        st.subheader("Catálogo Contínuo de Anúncios e Variações")
        st.write(
            "Importe a planilha de anúncios exportada do UpSeller ou marketplaces (.xlsx / .csv). "
            "Se o anúncio não existir no banco, ele será cadastrado automaticamente. "
            "Se já existir, seus preços e títulos serão atualizados via UPSERT sem sobrescrever o CMV."
        )

        col_up1, col_up2 = st.columns([1, 2])
        with col_up1:
            canal_anuncio = st.selectbox("Selecione o Marketplace de Origem:", DatabaseManager.get_all_channels(), key="sel_canal_anuncio")
            file_anuncios = st.file_uploader(
                "Carregar Planilha de Anúncios (.xlsx, .xls, .csv):",
                type=["xlsx", "xls", "csv"],
                key="uploader_anuncios",
            )

        with col_up2:
            if file_anuncios:
                try:
                    df_raw_anuncios = DataProcessor.read_uploaded_file(file_anuncios)
                    st.success(f"Arquivo lido com sucesso: **{len(df_raw_anuncios)} linhas** encontradas.")

                    with st.expander("🔍 Prévia das colunas detectadas", expanded=False):
                        st.dataframe(df_raw_anuncios.head(5), use_container_width=True)

                    df_clean_anuncios = DataProcessor.map_and_clean_ads_dataframe(df_raw_anuncios, canal_anuncio)

                    st.write(f"**Pré-visualização da Normalização ({len(df_clean_anuncios)} itens prontos para UPSERT):**")
                    st.dataframe(df_clean_anuncios.head(5), use_container_width=True, hide_index=True)

                    if st.button("🚀 Executar UPSERT de Anúncios no Banco", type="primary", key="btn_save_anuncios"):
                        records = df_clean_anuncios.to_dict(orient="records")
                        novos, atualizados = DatabaseManager.upsert_products(records)
                        st.balloons()
                        st.success(
                            f"✅ Operação concluída com sucesso! **{novos} novos produtos cadastrados** e **{atualizados} atualizados**."
                        )
                except Exception as e:
                    st.error(f"Erro ao processar arquivo de anúncios: {str(e)}")

        st.divider()
        st.subheader("Produtos Cadastrados no Banco de Dados")

        df_prods = DatabaseManager.get_products_dataframe()
        if df_prods.empty:
            st.info("Nenhum produto cadastrado no catálogo ainda.")
        else:
            col_search_prod, col_filter_prod = st.columns([3, 1])
            with col_search_prod:
                busca_p = st.text_input("Buscar produto por SKU, ID Variante ou Nome:", key="busca_prod_catalogo")
            with col_filter_prod:
                canal_p_filtro = st.selectbox("Filtrar Canal:", ["Todos"] + DatabaseManager.get_all_channels(), key="canal_prod_catalogo")

            df_prods_filtered = df_prods.copy()
            if canal_p_filtro != "Todos":
                df_prods_filtered = df_prods_filtered[df_prods_filtered["canal"] == canal_p_filtro]
            if busca_p:
                b_norm = DataProcessor.normalize_str(busca_p)
                df_prods_filtered = df_prods_filtered[
                    df_prods_filtered["id_variante"].astype(str).str.lower().str.contains(b_norm)
                    | df_prods_filtered["sku"].astype(str).str.lower().str.contains(b_norm)
                    | df_prods_filtered["nome_anuncio"].astype(str).str.lower().str.contains(b_norm)
                ]

            df_prods_filtered["id_variante"] = df_prods_filtered["id_variante"].apply(DataProcessor.clean_id_str)
            st.write(f"Exibindo **{len(df_prods_filtered)}** produtos.")
            st.dataframe(
                df_prods_filtered[[
                    "id_variante", "canal", "sku", "nome_anuncio", "variacao",
                    "preco_cheio", "preco_desconto", "preco_base", "custo_cmv", "atualizado_em"
                ]],
                use_container_width=True,
                hide_index=True,
                height=400,
            )

    # =========================================================================
    # ABA 3: IMPORTAÇÃO E CONCILIAÇÃO DE VENDAS
    # =========================================================================
    with tab_vendas:
        st.subheader("Importação e Conciliação Automática de Pedidos")
        st.write(
            "Faça o upload do relatório de vendas exportado (ex: `Export_Order...`). "
            "O sistema fará o cruzamento automático com a tabela de anúncios através do **ID da Variante** (ou **SKU**), "
            "substituindo o valor pago pelo cliente pelo **Preço Base Anunciado** e deduzindo as taxas oficiais da plataforma."
        )

        col_v1, col_v2 = st.columns([1, 2])
        with col_v1:
            opcoes_canais = ["Automático (Detectar pela Planilha Unificada)"] + DatabaseManager.get_all_channels()
            canal_vendas = st.selectbox(
                "Canal do Relatório de Vendas:",
                opcoes_canais,
                index=0,
                key="sel_canal_vendas",
                help="Se sua planilha for unificada (com vendas de Shopee, TikTok, Shein, Mercado Livre misturadas), selecione 'Automático'. O sistema identificará cada linha individualmente pela coluna da plataforma.",
            )
            file_vendas = st.file_uploader(
                "Carregar Relatório de Vendas (.xlsx, .xls, .csv):",
                type=["xlsx", "xls", "csv"],
                key="uploader_vendas",
            )

        with col_v2:
            if file_vendas:
                try:
                    df_raw_vendas = DataProcessor.read_uploaded_file(file_vendas)
                    st.success(f"Arquivo de vendas carregado: **{len(df_raw_vendas)} linhas**.")

                    # Carregar configurações, catálogo e IDs existentes do banco
                    configs = DatabaseManager.get_channel_configs()
                    df_prods_catalog = DatabaseManager.get_products_dataframe()
                    existing_pedido_ids = DatabaseManager.get_existing_pedido_ids()

                    # Executa Conciliação em Memória
                    processed_sales, stats = DataProcessor.reconcile_sales_batch(
                        df_raw_vendas,
                        canal_vendas,
                        configs,
                        df_prods_catalog,
                    )

                    # Contagem de novos vs repetidos no lote
                    repetidos_lote = sum(1 for s in processed_sales if s["pedido_id"] in existing_pedido_ids)
                    novos_lote = len(processed_sales) - repetidos_lote

                    # Badges visuais dos canais identificados no arquivo
                    canais_detectados = stats.get("canais_detectados", {})
                    cores_map = {
                        "Shopee": ("#EE4D2D", "#FFF1EE", "🟠"),
                        "TikTok Shop": ("#0F172A", "#F1F5F9", "⚫"),
                        "Shein": ("#B45309", "#FEF3C7", "🟡"),
                        "Mercado Livre": ("#854D0E", "#FEF9C3", "🟡"),
                    }
                    badges_canais_html = []
                    for ch_nome, ch_count in canais_detectados.items():
                        c_text, c_bg, c_icon = cores_map.get(ch_nome, ("#1E293B", "#F8FAFC", "🔵"))
                        badges_canais_html.append(
                            f'<span style="background: {c_bg}; color: {c_text}; border: 1px solid {c_text}40; padding: 2px 7px; border-radius: 6px; font-weight: 600; font-size: 0.85rem; margin-right: 6px; display: inline-block; margin-top: 3px;">{c_icon} {ch_nome}: {ch_count}</span>'
                        )
                    badges_str = " ".join(badges_canais_html) if badges_canais_html else "<i>Não especificado</i>"

                    # Exibe Estatísticas do Lote
                    cor_badge_rep = "#EF4444" if repetidos_lote > 0 else "#10B981"
                    st.markdown(
                        f"""
                        <div style="background: white; border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; margin-bottom: 12px;">
                            <b style="color: #0f172a; font-size: 1rem;">Resumo da Conciliação do Lote:</b><br/>
                            • Total de Linhas/Itens no Arquivo: <b>{stats['total_pedidos']}</b><br/>
                            • Novos Pedidos a Cadastrar: <b style="color: #10B981;">{novos_lote}</b><br/>
                            • Pedidos Repetidos (já existentes no banco): <b style="color: {cor_badge_rep};">{repetidos_lote}</b> (serão ignorados para não duplicar)<br/>
                            • Separação por Canal/Marketplace: {badges_str}<br/>
                            • Itens Vinculados ao Catálogo de Anúncios: <b>{stats['itens_conciliados']}</b> ({stats['taxa_cobertura_pct']}% de cobertura)<br/>
                            • Itens sem Anúncio Pré-Cadastrado: <b>{stats['itens_nao_cadastrados']}</b> (utilizado preço de venda como fallback)<br/>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    if repetidos_lote > 0:
                        st.info(
                            f"ℹ️ **Proteção Antiduplicação Ativa:** Identificamos **{repetidos_lote} pedidos repetidos** que já foram importados anteriormente. "
                            "O sistema irá ignorá-los automaticamente para que seu faturamento e taxas não sejam duplicados."
                        )

                    if stats["itens_nao_cadastrados"] > 0:
                        st.warning(
                            f"⚠️ Atenção: {stats['itens_nao_cadastrados']} itens não tinham anúncio previamente cadastrado na aba 'Importação de Anúncios'. "
                            "Para esses itens, o sistema utilizou o preço do relatório como fallback. "
                            "Recomenda-se importar a planilha de anúncios do marketplace para garantir a recuperação do preço com desconto real."
                        )

                    # Prévia dos dados conciliados
                    df_preview_conciliado = pd.DataFrame(processed_sales)
                    cols_preview = [
                        c for c in [
                            "numero_venda", "tag_item", "canal", "pedido_id", "sku", "nome_anuncio", "quantidade",
                            "preco_upseller_registrado", "preco_base_anuncio",
                            "faturamento_bruto", "total_taxas_canal", "valor_imposto",
                            "saque_liquido_previsto", "lucro_liquido_real"
                        ] if c in df_preview_conciliado.columns
                    ]
                    with st.expander("👀 Ver prévia detalhada dos cálculos do lote", expanded=True):
                        st.dataframe(
                            df_preview_conciliado[cols_preview].rename(columns={
                                "numero_venda": "Nº da Venda",
                                "tag_item": "Tag / Linha",
                                "canal": "Canal / Marketplace",
                                "pedido_id": "ID Linha",
                                "sku": "SKU",
                                "nome_anuncio": "Nome do Anúncio",
                                "quantidade": "Qtd",
                                "preco_upseller_registrado": "Preço UpSeller",
                                "preco_base_anuncio": "Preço Base",
                                "faturamento_bruto": "Fat. Bruto",
                                "total_taxas_canal": "Taxas Canal",
                                "valor_imposto": "Imposto DAS",
                                "saque_liquido_previsto": "Saque Previsto",
                                "lucro_liquido_real": "Lucro Real",
                            }).head(15),
                            use_container_width=True,
                            hide_index=True,
                        )

                    chk_ignorar = st.checkbox(
                        "🛡️ Ignorar vendas repetidas já existentes no banco (recomendado para não duplicar)",
                        value=True,
                        key="chk_ignorar_repetidos",
                    )

                    if st.button("💾 Confirmar e Gravar Vendas no Banco de Dados", type="primary", key="btn_save_vendas"):
                        novas, repetidas = DatabaseManager.upsert_sales(processed_sales, ignorar_existentes=chk_ignorar)
                        if novas > 0 and repetidas > 0:
                            st.balloons()
                            st.success(
                                f"🎉 Conciliação gravada com sucesso! **{novas} novos pedidos inseridos** e **{repetidas} pedidos repetidos foram ignorados** (nenhuma duplicação)."
                            )
                        elif novas > 0:
                            st.balloons()
                            st.success(f"🎉 **{novas} novos pedidos** inseridos com sucesso no banco de dados!")
                        else:
                            st.info(
                                f"ℹ️ Todos os **{repetidas} pedidos** desta planilha já haviam sido importados e foram **ignorados**. O faturamento permanece 100% íntegro sem duplicações!"
                            )
                except Exception as e:
                    st.error(f"Erro ao conciliar vendas: {str(e)}")

    # =========================================================================
    # ABA 4: GERENCIADOR DE TAXAS E CMV
    # =========================================================================
    with tab_config:
        st.subheader("Gerenciador de Taxas de Marketplaces & Custos (CMV)")

        # PARTE 1: Painel de Edição de Taxas por Canal
        st.markdown("#### 1. Configuração de Taxas e Impostos por Canal")
        st.caption("Altere as comissões, taxa fixa por peça e alíquota de imposto (Simples Nacional) para cada canal.")

        configs_atuais = DatabaseManager.get_channel_configs()
        all_channels = DatabaseManager.get_all_channels()

        cores_map = {
            "Shopee": ("#EE4D2D", "#FFF1EE", "🟠"),
            "TikTok Shop": ("#0F172A", "#F1F5F9", "⚫"),
            "Shein": ("#92400E", "#FEF3C7", "🟡"),
            "Mercado Livre": ("#854D0E", "#FEF9C3", "🟡"),
        }

        novas_taxas = {}

        # Renderização dinâmica em colunas balanceadas para todos os canais
        num_cols = min(len(all_channels), 4) if all_channels else 1
        cols_grid = st.columns(num_cols)
        for idx, canal_nome in enumerate(all_channels):
            with cols_grid[idx % num_cols]:
                c_text, c_bg, c_icon = cores_map.get(canal_nome, ("#1E293B", "#F8FAFC", "🔵"))
                st.markdown(
                    f"""
                    <div style="background: {c_bg}; border-radius: 8px; padding: 12px; border: 1px solid {c_text}40; margin-bottom: 10px;">
                        <b style="color: {c_text}; font-size: 1.05rem;">{c_icon} {canal_nome}</b>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                tax_info = configs_atuais.get(
                    canal_nome,
                    DEFAULT_CHANNEL_RATES.get(
                        canal_nome,
                        {"taxa_comissao_pct": 20.0, "taxa_fixa_unidade": 5.0, "aliquota_imposto_pct": 8.0},
                    ),
                )
                safe_key = re.sub(r'[^a-zA-Z0-9_]', '_', canal_nome).lower()
                com_val = st.number_input(f"Comissão ({canal_nome}) %:", value=float(tax_info["taxa_comissao_pct"]), step=0.5, key=f"cfg_com_{safe_key}")
                fix_val = st.number_input(f"Taxa Fixa ({canal_nome}) R$:", value=float(tax_info["taxa_fixa_unidade"]), step=0.5, key=f"cfg_fix_{safe_key}")
                imp_val = st.number_input(f"Imposto DAS ({canal_nome}) %:", value=float(tax_info["aliquota_imposto_pct"]), step=0.5, key=f"cfg_imp_{safe_key}")
                novas_taxas[canal_nome] = {"taxa_comissao_pct": com_val, "taxa_fixa_unidade": fix_val, "aliquota_imposto_pct": imp_val}

        if st.button("💾 Salvar Novas Taxas no Banco", type="primary", key="btn_save_taxas"):
            for canal, t in novas_taxas.items():
                DatabaseManager.update_channel_config(
                    canal,
                    t["taxa_comissao_pct"],
                    t["taxa_fixa_unidade"],
                    t["aliquota_imposto_pct"],
                )
            st.success("✅ Configurações de taxas atualizadas com sucesso!")

        st.divider()

        # PARTE 2: Gerenciador de Custo da Mercadoria Vendida (CMV)
        st.markdown("#### 2. Tabela de Gestão de Custo Unitário (CMV) por Produto")
        st.caption("Edite os custos diretamente na tabela abaixo e clique em 'Salvar Alterações de CMV'.")

        df_prods_cmv = DatabaseManager.get_products_dataframe()

        if df_prods_cmv.empty:
            st.info("Nenhum produto cadastrado no catálogo para gerenciar CMV. Cadastre anúncios na aba 'Importação de Anúncios'.")
        else:
            df_editor = df_prods_cmv[["id_variante", "sku", "nome_anuncio", "canal", "preco_base", "custo_cmv"]].copy()
            df_editor["id_variante"] = df_editor["id_variante"].apply(DataProcessor.clean_id_str)

            edited_df = st.data_editor(
                df_editor,
                column_config={
                    "id_variante": st.column_config.TextColumn("ID Variante", disabled=True),
                    "sku": st.column_config.TextColumn("SKU", disabled=True),
                    "nome_anuncio": st.column_config.TextColumn("Nome do Anúncio", disabled=True),
                    "canal": st.column_config.TextColumn("Canal", disabled=True),
                    "preco_base": st.column_config.NumberColumn("Preço Base (R$)", format="R$ %.2f", disabled=True),
                    "custo_cmv": st.column_config.NumberColumn(
                        "Custo CMV (R$)",
                        help="Insira o custo unitário deste produto",
                        min_value=0.0,
                        step=0.5,
                        format="R$ %.2f",
                    ),
                },
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                height=400,
                key="editor_cmv",
            )

            col_salvar_cmv, col_recalc = st.columns([1, 1])
            with col_salvar_cmv:
                if st.button("💾 Salvar Alterações de CMV", type="primary", key="btn_save_cmv"):
                    # Detecta linhas alteradas
                    updates = []
                    for _, row in edited_df.iterrows():
                        updates.append((float(row["custo_cmv"]), str(row["id_variante"])))
                    DatabaseManager.batch_update_cmv(updates)
                    st.success("✅ Custos CMV atualizados com sucesso no banco de dados!")

            with col_recalc:
                st.write("")
                if st.button("🔄 Recalcular Todas as Vendas com Novas Taxas e CMV", key="btn_recalc_vendas"):
                    recalc_count = DatabaseManager.recalculate_all_sales()
                    st.balloons()
                    st.success(f"🎉 Recálculo concluído! **{recalc_count} vendas** foram recalculadas com as taxas e CMVs atualizados.")

        st.divider()

        # PARTE 3: Manutenção do Sistema & Reinício do Histórico
        st.markdown("#### 3. Manutenção e Reinício do Histórico de Vendas")
        st.caption("Esta opção zera exclusivamente os pedidos e faturamento da aba de vendas. Seus anúncios cadastrados no catálogo e as taxas configuradas permanecem 100% intactos.")

        with st.expander("🗑️ Zona de Limpeza: Zerar Vendas Processadas", expanded=False):
            st.warning("⚠️ **Atenção:** Esta ação apagará todas as vendas gravadas no banco de dados para que você possa iniciar o controle financeiro do zero com as novas planilhas.")
            chk_confirma_zerar = st.checkbox(
                "Sim, confirmo que desejo apagar todo o histórico de vendas para começar do zero.",
                key="chk_confirma_zerar",
            )
            if st.button("🚨 Confirmar e Apagar Todas as Vendas", type="primary", key="btn_zerar_vendas", disabled=not chk_confirma_zerar):
                qtd_removidas = DatabaseManager.clear_sales_history()
                st.success(f"✅ Histórico financeiro zerado com sucesso! **{qtd_removidas} vendas** foram removidas. O catálogo de anúncios continua 100% preservado.")
                st.rerun()


if __name__ == "__main__":
    main()

# Sistema de Controle Financeiro e Conciliação de Saque Multicanal
**Shopee • TikTok Shop • Shein • UpSeller**

Sistema completo e autônomo para conciliação financeira de vendas em marketplaces com correção automática da distorção de cupons da plataforma, controle de estoque/anúncios cadastrais, dedução de taxas parametrizáveis e apuração de saque líquido e lucro real com base no CMV.

---

## 🚀 Como Executar o Sistema

O ambiente virtual local já está configurado na pasta `.venv`. Para iniciar o painel interativo, execute no terminal PowerShell:

```powershell
.\.venv\Scripts\streamlit.exe run app.py
```

O aplicativo abrirá automaticamente no seu navegador padrão no endereço:
`http://localhost:8501`

---

## 🗄️ Estrutura do Banco de Dados SQLite (`ecommerce_financeiro.db`)

1. **`configuracoes_canais`**:
   - `canal`: Shopee, TikTok Shop, Shein
   - `taxa_comissao_pct`: Comissão da plataforma (ex: 20%, 22%)
   - `taxa_fixa_unidade`: Custo fixo por peça vendida (ex: R$ 5,00, R$ 4,00)
   - `aliquota_imposto_pct`: Alíquota do Simples Nacional (padrão: 8,0%)

2. **`produtos`**:
   - Catálogo contínuo com `id_variante` como chave primária, `sku`, `id_anuncio`, `nome_anuncio`, `variacao`, `preco_cheio`, `preco_desconto`, `preco_base` e `custo_cmv`.
   - Suporta UPSERT idempotente: atualiza preços sem perder custos CMV já cadastrados.

3. **`vendas_processadas`**:
   - Pedidos conciliados com identificador único `pedido_id`, quantidade, preço do cliente, preço base anunciado, faturamento bruto real, comissões, impostos, saque previsto, CMV total e lucro líquido real.

---

## 📐 Fórmulas Financeiras Aplicadas

Para cada item vendido:

$$\text{Preço Base} = \text{Preço com Desconto (do Anúncio)} \quad [\text{fallback: Preço Cheio ou Preço Venda}]$$

$$\text{Faturamento Bruto} = \text{Preço Base} \times \text{Quantidade}$$

$$\text{Valor Comissão} = \text{Faturamento Bruto} \times \text{Taxa Comissão (\% canal)}$$

$$\text{Valor Taxa Fixa} = \text{Taxa Fixa por Peça} \times \text{Quantidade}$$

$$\text{Total Taxas Canal} = \text{Valor Comissão} + \text{Valor Taxa Fixa}$$

$$\text{Imposto} = \text{Faturamento Bruto} \times \text{Alíquota Imposto (\%)}$$

$$\text{Saque Líquido Previsto} = \text{Faturamento Bruto} - \text{Total Taxas Canal} - \text{Imposto}$$

$$\text{Lucro Líquido Real} = \text{Saque Líquido Previsto} - (\text{Custo Unitário CMV} \times \text{Quantidade})$$

---

## 📂 Arquivos de Exemplo para Teste
Na pasta `exemplos/`, foram criadas duas planilhas para validação imediata:
- `exemplos/exemplo_anuncios_upseller.xlsx`: Exemplo de anúncios cadastrais.
- `exemplos/exemplo_vendas_upseller.xlsx`: Exemplo de vendas com cupons aplicados pelo cliente.

---

## 🌐 Publicação Gratuita na Nuvem (Deploy)

### 1. Subir para o GitHub
Execute o arquivo [`subir_github.bat`](subir_github.bat) ou rode no terminal:
```bash
git push -u origin main
```

### 2. Por que usar Streamlit Community Cloud em vez da Vercel?
- A **Vercel** foi projetada para sites estáticos e funções *serverless* efêmeras (cada requisição apaga o disco local, o que destruiria o banco SQLite e cortaria as conexões WebSocket do Streamlit).
- O **Streamlit Community Cloud (share.streamlit.io)** é a plataforma oficial, **100% gratuita para sempre**, com servidor persistente, certificado SSL e suporte nativo a WebSockets e SQLite.

### 3. Como colocar no ar em 2 minutos (Streamlit Cloud):
1. Acesse [share.streamlit.io](https://share.streamlit.io/) e faça login com sua conta do GitHub.
2. Clique em **"New app"**.
3. Selecione o repositório: `jvitoria775-cyber/FINANCEIROMTKPLACEGT`.
4. Branch: `main` | Main file path: `app.py`.
5. Clique em **"Deploy!"**.
6. Seu painel estará publicado e acessível publicamente via link HTTPS gratuito!

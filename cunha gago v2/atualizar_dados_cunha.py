"""
atualizar_dados_cunha.py
Script para extrair e atualizar os dados da unidade Cunha Gago a partir de vendtef_local.js
Gera tanto relatorio_vendas_geral.js quanto relatorio_vendas_geral.csv
"""
import os
import sys
from pathlib import Path

def atualizar():
    script_dir = Path(__file__).parent.resolve()
    vendtef_path = script_dir.parent / "vendtef_local.js"

    if not vendtef_path.exists():
        print(f"ERRO: {vendtef_path} não encontrado!")
        sys.exit(1)

    with open(vendtef_path, "r", encoding="utf-8", errors="ignore") as f:
        raw = f.read()

    lines = raw.split("\n")
    header = "Cliente;Máquina;Modelo;Fabricante;Pagamento;Produtos;Mola;Venda (R$);Preço (R$);Total;Código Promocional;Data;Hora;Nº Logico;NSU;Autorização;Tipo Cartão;Rede;Bandeira;Usuário;Nº Cartão;Matricula"

    cunha_rows = []
    for line in lines:
        l = line.strip()
        if not l:
            continue
        if "40# AHI - Cunha Gago" in l:
            cunha_rows.append(l)

    print(f"Total de transações encontradas para Cunha Gago: {len(cunha_rows)}")

    csv_content = header + "\n" + "\n".join(cunha_rows)

    # 1. Salva relatorio_vendas_geral.csv
    csv_file = script_dir / "relatorio_vendas_geral.csv"
    with open(csv_file, "w", encoding="utf-8") as f:
        f.write(csv_content)
    print(f"Salvo com sucesso: {csv_file} ({os.path.getsize(csv_file)} bytes)")

    # 2. Salva relatorio_vendas_geral.js
    escaped = csv_content.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
    js_content = "window.LAVAI_CSV_DATA = `\n" + escaped + "`;\n"
    js_file = script_dir / "relatorio_vendas_geral.js"
    with open(js_file, "w", encoding="utf-8") as f:
        f.write(js_content)
    print(f"Salvo com sucesso: {js_file} ({os.path.getsize(js_file)} bytes)")

    # Mostra resumo das transações
    if cunha_rows:
        primeira_dt = cunha_rows[0].split(";")[11] if len(cunha_rows[0].split(";")) > 11 else "?"
        ultima_dt = cunha_rows[-1].split(";")[11] if len(cunha_rows[-1].split(";")) > 11 else "?"
        ultima_hr = cunha_rows[-1].split(";")[12] if len(cunha_rows[-1].split(";")) > 12 else "?"
        print(f"Período: {primeira_dt} até {ultima_dt} às {ultima_hr}")

if __name__ == "__main__":
    atualizar()

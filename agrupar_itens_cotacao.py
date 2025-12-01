import pandas as pd
from collections import Counter
from datetime import datetime
import html
import os
import glob
import re
import json

def processar_planilha_para_cotacao():
    """Processa a planilha e agrupa itens repetidos"""
    
    # Ler planilha
    df = pd.read_excel('Dartagnan.xlsx', header=None)
    
    # Procurar linha de cabeçalho
    header_row = None
    for i in range(min(10, len(df))):
        linha = df.iloc[i].astype(str).tolist()
        if 'Item' in linha and 'Descrição' in linha:
            header_row = i
            break
    
    if header_row is None:
        header_row = 3
    
    # Ler com cabeçalho
    df_header = pd.read_excel('Dartagnan.xlsx', header=header_row)
    
    # Identificar colunas
    col_descricao = None
    col_total = None
    col_unidade = None
    col_quantidade = None
    
    for col in df_header.columns:
        col_str = str(col).lower()
        if 'descrição' in col_str or 'descricao' in col_str:
            col_descricao = col
        elif 'total' in col_str:
            col_total = col
        elif 'und' in col_str or 'unidade' in col_str:
            col_unidade = col
        elif 'quant' in col_str:
            col_quantidade = col
    
    if col_descricao is None:
        # Procurar coluna com mais texto
        for col in df_header.columns:
            if df_header[col].astype(str).str.len().mean() > 10:
                col_descricao = col
                break
    
    if col_total is None:
        col_total = df_header.columns[-1]
    
    print(f"Colunas identificadas:")
    print(f"  Descrição: {col_descricao}")
    print(f"  Unidade: {col_unidade}")
    print(f"  Quantidade: {col_quantidade}")
    print(f"  Total: {col_total}")
    
    # Coletar TODOS os itens primeiro (sem filtros)
    todos_itens_raw = []
    
    for idx, row in df_header.iterrows():
        desc_val = row[col_descricao] if pd.notna(row[col_descricao]) else None
        total_val = row[col_total] if pd.notna(row[col_total]) else None
        unidade_val = row[col_unidade] if col_unidade and pd.notna(row[col_unidade]) else None
        quantidade_val = row[col_quantidade] if col_quantidade and pd.notna(row[col_quantidade]) else None
        
        if pd.isna(desc_val):
            continue
        
        desc_str = str(desc_val).strip()
        
        # Pular cabeçalhos e valores inválidos
        if desc_str.lower() in ['descrição', 'descricao', 'obra', 'nan', '']:
            continue
        
        # Pular linhas de total
        if 'total' in desc_str.lower() or 'geral' in desc_str.lower():
            continue
        
        # Obter valor (tratar NaN como 0)
        valor = 0
        if pd.notna(total_val):
            try:
                valor = float(total_val)
                if pd.isna(valor):
                    valor = 0
            except:
                valor = 0
        
        # Obter unidade e quantidade
        unidade = str(unidade_val).strip() if pd.notna(unidade_val) else None
        quantidade = None
        if pd.notna(quantidade_val):
            try:
                quantidade = float(quantidade_val)
                if pd.isna(quantidade):
                    quantidade = None
            except:
                quantidade = None
        
        # Verificar se unidade ou quantidade estão em branco
        unidade_valida = unidade and unidade.lower() not in ['nan', 'none', '', 'undefined']
        quantidade_valida = quantidade is not None and not pd.isna(quantidade)
        
        # Adicionar apenas itens com unidade E quantidade válidas
        if unidade_valida and quantidade_valida:
            todos_itens_raw.append({
                'descricao': desc_str,
                'valor': valor,
                'unidade': unidade,
                'quantidade': quantidade
            })
    
    # Contar TODAS as ocorrências na planilha (incluindo as com valor zero ou NaN)
    # Para isso, precisamos contar antes de filtrar por valor
    todas_descricoes_planilha = []
    for idx, row in df_header.iterrows():
        desc_val = row[col_descricao] if pd.notna(row[col_descricao]) else None
        if pd.isna(desc_val):
            continue
        desc_str = str(desc_val).strip()
        if desc_str.lower() not in ['descrição', 'descricao', 'obra', 'nan', '']:
            if 'total' not in desc_str.lower() and 'geral' not in desc_str.lower():
                todas_descricoes_planilha.append(desc_str)
    
    contador_todos = Counter(todas_descricoes_planilha)
    
    # Filtrar: manter itens que se repetem OU são itens finais detalhados
    itens = []
    categorias_genericas = [
        'esquadrias', 'piso', 'revestimento', 'louças', 'acessórios', 
        'metais', 'vidro', 'diversos', 'área', 'reforma', 'sala',
        'banheiro', 'depósito', 'hall', 'barrilete', 'bombas',
        'quadro', 'comando', 'escada', 'acesso', 'execução',
        'elevatória', 'água', 'bruta', 'bate', 'estaca', 
        'sistema', 'cloração', 'eta', 'nova', 'oficina', 
        'hidrômetros', 'pitometria', 'almoxarifado', 'estação', 
        'tratamento', 'casa', 'química', 'laboratório', 'guarita', 
        'administração', 'local', 'serviços', 'preliminares'
    ]
    
    for item in todos_itens_raw:
        desc = item['descricao']
        desc_lower = desc.lower()
        palavras_desc = desc_lower.split()
        
        # Verificar se é categoria genérica de UMA palavra (pular apenas essas)
        # Exemplos: "Piso", "Vidro", "Esquadrias" (uma palavra só)
        if len(palavras_desc) == 1 and len(desc) < 20:
            if desc_lower in categorias_genericas:
                continue  # Pular apenas categorias de uma palavra
        
        # Verificar se é apenas lista de categorias separadas por vírgula (sem especificações)
        if ',' in desc and len(palavras_desc) <= 5:
            palavras_separadas = [p.strip() for p in desc.split(',')]
            # Se todas as palavras são categorias genéricas E não tem especificações técnicas
            todas_categorias = all(p.lower() in categorias_genericas for p in palavras_separadas if len(p) > 2)
            tem_especificacao = any(marker in desc.upper() for marker in ['AF_', 'NBR', 'CM', 'MM', 'X', 'DE', 'PARA'])
            if todas_categorias and not tem_especificacao:
                continue  # Pular listas de categorias sem especificações
        
        # Se tem valor > 0, incluir (tanto repetidos quanto únicos)
        if item['valor'] > 0:
            # Verificar se é item final (tem código técnico, descrição detalhada, ou é item principal)
            tem_codigo_tecnico = any(marker in desc.upper() for marker in ['AF_', 'NBR'])
            tem_descricao_detalhada = len(desc) > 50
            tem_especificacoes = any(marker in desc.upper() for marker in ['CM', 'MM', 'X', 'M²', 'M2'])
            
            # Itens principais (não são categorias genéricas de uma palavra)
            # Se tem mais de 2 palavras OU mais de 30 caracteres, é item principal
            e_item_principal = len(palavras_desc) > 2 or len(desc) > 30
            
            # Incluir se:
            # - Tem código técnico OU
            # - Tem descrição detalhada OU
            # - Tem especificações (dimensões) OU
            # - É item principal (mais de 2 palavras ou mais de 30 caracteres)
            if tem_codigo_tecnico or tem_descricao_detalhada or tem_especificacoes or e_item_principal:
                itens.append(item)
    
    # Agrupar itens iguais
    # Usar contador_todos para quantidade real de repetições na planilha
    contador = Counter([item['descricao'] for item in itens])
    
    # Criar lista agrupada
    itens_agrupados = []
    valores_por_item = {}
    unidades_por_item = {}
    quantidades_por_item = {}
    
    for item in itens:
        desc = item['descricao']
        if desc not in valores_por_item:
            valores_por_item[desc] = []
            unidades_por_item[desc] = []
            quantidades_por_item[desc] = []
        valores_por_item[desc].append(item['valor'])
        unidades_por_item[desc].append(item['unidade'])
        quantidades_por_item[desc].append(item['quantidade'])
    
    for descricao, qtd_ocorrencias in contador.items():
        # Usar contador_todos para quantidade real na planilha (incluindo as com valor 0)
        qtd_real_planilha = contador_todos.get(descricao, qtd_ocorrencias)
        valores = valores_por_item[descricao]
        unidades = unidades_por_item[descricao]
        quantidades = quantidades_por_item[descricao]
        
        # Filtrar apenas valores > 0 (remover os com valor 0/NaN que foram incluídos)
        valores_filtrados = [v for v in valores if v > 0]
        if not valores_filtrados:
            continue  # Pular se não há valores > 0
        
        valor_total = sum(valores_filtrados)
        valor_medio = valor_total / len(valores_filtrados) if valores_filtrados else 0
        quantidade_total = sum(quantidades) if quantidades else qtd_ocorrencias
        
        # Usar unidade mais comum ou primeira disponível
        unidade_mais_comum = max(set(unidades), key=unidades.count) if unidades else 'UN'
        
        itens_agrupados.append({
            'descricao': descricao,
            'quantidade': qtd_real_planilha,  # Número real de vezes que aparece na planilha (incluindo as com valor 0)
            'quantidade_total': quantidade_total,  # Soma das quantidades
            'unidade': unidade_mais_comum,
            'valor_total': valor_total,
            'valor_unitario': valor_medio,
            'valores': valores_filtrados
        })
    
    # Ordenar por quantidade (mais repetidos primeiro)
    itens_agrupados.sort(key=lambda x: x['quantidade'], reverse=True)
    
    return itens_agrupados

def buscar_imagem_item(numero_item):
    """Busca imagem para um item baseado no número sequencial (1 a 39)"""
    # Procurar pasta de imagens
    pastas_imagens = ['imagens', 'fotos', 'images', 'photos', '.']
    extensoes = ['jpg', 'jpeg', 'png', 'gif', 'webp']
    
    # Garantir que numero_item é inteiro
    numero_item = int(numero_item)
    
    # Buscar arquivo com o número do item
    for pasta in pastas_imagens:
        if not os.path.exists(pasta):
            continue
        
        # Buscar arquivos com o número do item
        for ext in extensoes:
            # Tentar diferentes formatos: 1.jpg, 01.jpg, item1.jpg, etc.
            possiveis_nomes = [
                f'{numero_item}.{ext}',
                f'{numero_item:02d}.{ext}',  # 01.jpg, 02.jpg, etc.
                f'item{numero_item}.{ext}',
                f'item{numero_item:02d}.{ext}',
                f'#{numero_item}.{ext}',
                f'#{numero_item:02d}.{ext}'
            ]
            
            for nome in possiveis_nomes:
                caminho = os.path.join(pasta, nome)
                if os.path.exists(caminho):
                    return caminho.replace('\\', '/')
    
    return None

def processar_checklist():
    """Processa arquivo Excel de checklist"""
    arquivos_checklist = ['checklist.xlsx', 'Checklist.xlsx', 'CHECKLIST.xlsx']
    arquivo_encontrado = None
    
    for arquivo in arquivos_checklist:
        if os.path.exists(arquivo):
            arquivo_encontrado = arquivo
            break
    
    if not arquivo_encontrado:
        return None
    
    try:
        df = pd.read_excel(arquivo_encontrado, header=None)
        checklist_data = []
        
        # Processar dados do checklist
        for idx, row in df.iterrows():
            item = None
            status = False
            
            if len(df.columns) == 1:
                if pd.notna(row[0]):
                    item = str(row[0]).strip()
            else:
                if pd.notna(row[0]):
                    item = str(row[0]).strip()
                # Verificar outras colunas para status
                for col_idx in range(1, min(3, len(df.columns))):
                    if pd.notna(row[col_idx]):
                        valor = str(row[col_idx]).strip().lower()
                        if valor in ['x', 'sim', 'ok', 'concluído', 'concluido', 'feito', '1']:
                            status = True
            
            # Ignorar primeira linha se for cabeçalho comum
            if idx == 0 and item:
                item_lower = item.lower()
                if item_lower in ['item', 'descrição', 'descricao', 'status', 'concluído', 'concluido']:
                    continue
            
            if item and item.lower() not in ['nan', 'none', '']:
                checklist_data.append({
                    'item': item,
                    'concluido': status
                })
        
        return checklist_data
    except Exception as e:
        print(f"Erro ao processar checklist: {e}")
        return None

def criar_html_checklist(checklist_data):
    """Cria HTML para o checklist"""
    if not checklist_data:
        return """
        <div style="text-align: center; padding: 40px;">
            <p style="color: #666; margin-bottom: 20px;">
                📋 Nenhum arquivo de checklist encontrado.
            </p>
            <p style="color: #999; font-size: 0.9em;">
                Crie um arquivo <strong>checklist.xlsx</strong> na pasta do projeto com os itens do checklist.
            </p>
        </div>
        """
    
    total = len(checklist_data)
    concluidos = sum(1 for item in checklist_data if item.get('concluido', False))
    pendentes = total - concluidos
    
    html_content = f"""
    <div style="max-width: 800px; margin: 0 auto;">
        <div style="background: #f8f9fa; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
            <p style="margin: 0; color: #666;">
                <strong>Total de itens:</strong> {total} | 
                <strong>Concluídos:</strong> <span style="color: #2e7d32;">{concluidos}</span> | 
                <strong>Pendentes:</strong> <span style="color: #d32f2f;">{pendentes}</span>
            </p>
        </div>
        
        <div style="background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
            <table style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr style="background: #667eea; color: white;">
                        <th style="padding: 15px; text-align: left; width: 50px;">✓</th>
                        <th style="padding: 15px; text-align: left;">Item do Checklist</th>
                    </tr>
                </thead>
                <tbody>
"""
    
    for i, item_data in enumerate(checklist_data, 1):
        item_text = html.escape(str(item_data['item']))
        concluido = item_data.get('concluido', False)
        checked = 'checked' if concluido else ''
        linha_style = 'background: #e8f5e9;' if concluido else ''
        
        html_content += f"""
                    <tr style="{linha_style}">
                        <td style="padding: 12px 15px; text-align: center;">
                            <input type="checkbox" id="check_{i-1}" {checked} onchange="atualizarChecklist({i-1}, this.checked)" 
                                   style="width: 20px; height: 20px; cursor: pointer;">
                        </td>
                        <td style="padding: 12px 15px;">
                            <label for="check_{i-1}" style="cursor: pointer; display: block; margin: 0;">
                                {item_text}
                            </label>
                        </td>
                    </tr>
"""
    
    checklist_json = json.dumps(checklist_data, ensure_ascii=False)
    
    html_content += f"""
                </tbody>
            </table>
        </div>
    </div>
    
    <script>
        let checklistData = {checklist_json};
        
        function atualizarChecklist(index, concluido) {{
            checklistData[index].concluido = concluido;
            atualizarEstatisticas();
            salvarChecklist();
        }}
        
        function atualizarEstatisticas() {{
            const total = checklistData.length;
            const concluidos = checklistData.filter(item => item.concluido).length;
            const pendentes = total - concluidos;
            
            // Atualizar visualmente as linhas
            const rows = document.querySelectorAll('#checklist-content tbody tr');
            rows.forEach((row, index) => {{
                if (checklistData[index].concluido) {{
                    row.style.background = '#e8f5e9';
                }} else {{
                    row.style.background = '';
                }}
            }});
        }}
        
        function salvarChecklist() {{
            localStorage.setItem('checklist_caerd', JSON.stringify(checklistData));
        }}
        
        function carregarChecklist() {{
            const saved = localStorage.getItem('checklist_caerd');
            if (saved) {{
                const savedData = JSON.parse(saved);
                savedData.forEach((item, index) => {{
                    if (index < checklistData.length) {{
                        checklistData[index].concluido = item.concluido;
                    }}
                }});
                
                const checkboxes = document.querySelectorAll('#checklist-content input[type="checkbox"]');
                checkboxes.forEach((checkbox, index) => {{
                    if (index < checklistData.length) {{
                        checkbox.checked = checklistData[index].concluido;
                    }}
                }});
                
                atualizarEstatisticas();
            }}
        }}
        
        document.addEventListener('DOMContentLoaded', function() {{
            carregarChecklist();
        }});
    </script>
"""
    
    return html_content

def criar_html_cotacao(itens_agrupados):
    """Cria página HTML focada em cotação"""
    
    # Incluir TODOS os itens (repetidos e únicos)
    itens_repetidos = itens_agrupados
    
    # Buscar imagens para cada item pelo número sequencial
    for i, item in enumerate(itens_repetidos, 1):
        item['imagem'] = buscar_imagem_item(i)
        item['numero_item'] = i  # Adicionar número do item
    
    # Processar checklist
    checklist_data = processar_checklist()
    html_checklist = criar_html_checklist(checklist_data)
    
    html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Itens para Cotação - Dartagnan</title>
    <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>📋</text></svg>">
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
        }}
        
        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .content {{
            padding: 30px;
        }}
        
        .tabs-container {{
            margin-bottom: 20px;
        }}
        
        .tabs {{
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            border-bottom: 2px solid #e0e0e0;
        }}
        
        .tab-button {{
            padding: 12px 24px;
            background: transparent;
            border: none;
            border-bottom: 3px solid transparent;
            cursor: pointer;
            font-size: 1.1em;
            font-weight: 500;
            color: #666;
            transition: all 0.3s;
        }}
        
        .tab-button:hover {{
            color: #667eea;
            background: #f5f5f5;
        }}
        
        .tab-button.active {{
            color: #667eea;
            border-bottom-color: #667eea;
            font-weight: 600;
        }}
        
        .tab-content {{
            display: none;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        .section-title {{
            font-size: 1.8em;
            color: #333;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #667eea;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            background: white;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        
        th {{
            background: #667eea;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }}
        
        th.number {{
            text-align: right;
        }}
        
        td {{
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }}
        
        tr:hover {{
            background: #f5f5f5;
        }}
        
        .number {{
            text-align: right;
            font-family: 'Courier New', monospace;
        }}
        
        .descricao {{
            max-width: 600px;
            word-wrap: break-word;
        }}
        
        .quantidade-badge {{
            background: #667eea;
            color: white;
            padding: 5px 12px;
            border-radius: 20px;
            font-weight: bold;
            display: inline-block;
        }}
        
        .item-com-imagem {{
            position: relative;
            cursor: help;
        }}
        
        .tooltip {{
            position: absolute;
            background: white;
            border: 2px solid #667eea;
            border-radius: 10px;
            padding: 10px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            z-index: 1000;
            display: none;
            max-width: 400px;
            pointer-events: none;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            margin-bottom: 10px;
        }}
        
        .tooltip::after {{
            content: '';
            position: absolute;
            top: 100%;
            left: 50%;
            transform: translateX(-50%);
            border: 10px solid transparent;
            border-top-color: #667eea;
        }}
        
        .tooltip img {{
            max-width: 350px;
            max-height: 300px;
            width: auto;
            height: auto;
            border-radius: 5px;
            display: block;
            object-fit: contain;
        }}
        
        .tooltip .tooltip-text {{
            margin-top: 8px;
            font-size: 0.85em;
            color: #666;
            text-align: center;
            padding: 5px;
        }}
        
        .item-com-imagem:hover .tooltip {{
            display: block;
        }}
        
        .icon-imagem {{
            display: inline-block;
            margin-left: 5px;
            color: #667eea;
            font-size: 0.9em;
            opacity: 0.7;
            transition: opacity 0.2s;
        }}
        
        .item-com-imagem:hover .icon-imagem {{
            opacity: 1;
        }}
        
        /* Responsividade Mobile */
        .table-wrapper {{
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            margin: 0 -15px;
            padding: 0 15px;
        }}
        
        .table-wrapper table {{
            min-width: 800px;
        }}
        
        /* Cards para mobile - alternativa à tabela */
        .mobile-card {{
            display: none;
            background: white;
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 15px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        
        .mobile-card-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
            padding-bottom: 10px;
            border-bottom: 2px solid #667eea;
        }}
        
        .mobile-card-title {{
            font-weight: bold;
            color: #667eea;
            font-size: 1.1em;
        }}
        
        .mobile-card-content {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            font-size: 0.9em;
        }}
        
        .mobile-card-label {{
            font-weight: 600;
            color: #666;
        }}
        
        .mobile-card-value {{
            text-align: right;
            color: #333;
        }}
        
        .mobile-card-desc {{
            grid-column: 1 / -1;
            margin-top: 10px;
            padding-top: 10px;
            border-top: 1px solid #e0e0e0;
            font-size: 0.85em;
            line-height: 1.4;
        }}
        
        @media (max-width: 768px) {{
            body {{
                padding: 10px;
            }}
            
            .container {{
                border-radius: 10px;
            }}
            
            .header {{
                padding: 20px 15px;
            }}
            
            .header h1 {{
                font-size: 1.8em;
            }}
            
            .content {{
                padding: 15px;
            }}
            
            .tabs {{
                gap: 5px;
                flex-wrap: wrap;
            }}
            
            .tab-button {{
                padding: 14px 20px;
                font-size: 1em;
                flex: 1;
                min-width: 120px;
                -webkit-tap-highlight-color: transparent;
            }}
            
            .section-title {{
                font-size: 1.4em;
                margin-bottom: 15px;
            }}
            
            /* Esconder tabela em mobile, mostrar cards */
            .table-wrapper {{
                display: none;
            }}
            
            .mobile-cards-container {{
                display: block;
            }}
            
            .mobile-card {{
                display: block;
            }}
            
            /* Tooltip mobile - usar touch */
            .item-com-imagem {{
                cursor: pointer;
            }}
            
            .item-com-imagem.active .tooltip {{
                display: block;
                position: fixed;
                bottom: 20px;
                left: 50%;
                transform: translateX(-50%);
                max-width: 90vw;
                max-height: 60vh;
                z-index: 10000;
            }}
            
            .item-com-imagem.active .tooltip::after {{
                display: none;
            }}
            
            .tooltip img {{
                max-width: 100%;
                max-height: 50vh;
            }}
            
            /* Checklist mobile */
            #checklist-content table {{
                font-size: 0.9em;
            }}
            
            #checklist-content th,
            #checklist-content td {{
                padding: 10px 8px;
            }}
            
            #checklist-content input[type="checkbox"] {{
                width: 24px;
                height: 24px;
            }}
        }}
        
        @media (min-width: 769px) {{
            .mobile-cards-container {{
                display: none;
            }}
        }}
        
        @media print {{
            body {{
                background: white;
            }}
            .tooltip {{
                display: none !important;
            }}
            .mobile-cards-container {{
                display: none;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📋 Itens para Cotação</h1>
        </div>
        
        <div class="content">
            <div class="tabs-container">
                <div class="tabs">
                    <button class="tab-button active" onclick="showTab('insumos')">Insumos</button>
                    <button class="tab-button" onclick="showTab('checklist')">Checklist</button>
                </div>
            </div>
            
            <div id="insumos" class="tab-content active">
                <h2 class="section-title">📊 Insumos Medição Dezembro</h2>
                
                <div class="table-wrapper">
                    <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Descrição do Item</th>
                            <th class="number">Quantidade</th>
                            <th class="number">Unidade</th>
                            <th class="number">Valor Unitário (R$)</th>
                            <th class="number">Valor Total (R$)</th>
                        </tr>
                    </thead>
                    <tbody>
"""
    
    # Adicionar itens repetidos
    for i, item in enumerate(itens_repetidos, 1):
        desc_escaped = html.escape(item['descricao'])
        unidade_escaped = html.escape(str(item.get('unidade', 'UN')))
        imagem = item.get('imagem', None)
        numero_item = item.get('numero_item', i)
        
        # Se tiver imagem, adicionar tooltip
        if imagem:
            imagem_escaped = html.escape(imagem)
            html_content += f"""
                    <tr>
                        <td><strong>#{numero_item}</strong></td>
                        <td class="descricao item-com-imagem">
                            {desc_escaped}
                            <span class="icon-imagem">📷</span>
                            <div class="tooltip">
                                <img src="{imagem_escaped}" alt="Imagem do item #{numero_item}" onerror="this.style.display='none'; this.parentElement.querySelector('.tooltip-text').textContent='Imagem #{numero_item} não encontrada';">
                                <div class="tooltip-text">Item #{numero_item}: {desc_escaped[:50]}...</div>
                            </div>
                        </td>
                        <td class="number">
                            <span class="quantidade-badge">{item.get('quantidade_total', item['quantidade'])}</span>
                        </td>
                        <td class="number">{unidade_escaped}</td>
                        <td class="number">R$ {item['valor_unitario']:,.2f}</td>
                        <td class="number"><strong>R$ {item['valor_total']:,.2f}</strong></td>
                    </tr>
"""
        else:
            html_content += f"""
                    <tr>
                        <td><strong>#{numero_item}</strong></td>
                        <td class="descricao">{desc_escaped}</td>
                        <td class="number">
                            <span class="quantidade-badge">{item.get('quantidade_total', item['quantidade'])}</span>
                        </td>
                        <td class="number">{unidade_escaped}</td>
                        <td class="number">R$ {item['valor_unitario']:,.2f}</td>
                        <td class="number"><strong>R$ {item['valor_total']:,.2f}</strong></td>
                    </tr>
"""
    
    html_content += f"""
                    </tbody>
                </table>
                </div>
                
                <!-- Cards para Mobile -->
                <div class="mobile-cards-container">
"""
    
    # Adicionar cards mobile
    for i, item in enumerate(itens_repetidos, 1):
        desc_escaped = html.escape(item['descricao'])
        unidade_escaped = html.escape(str(item.get('unidade', 'UN')))
        imagem = item.get('imagem', None)
        numero_item = item.get('numero_item', i)
        
        # Se tiver imagem, adicionar tooltip
        if imagem:
            imagem_escaped = html.escape(imagem)
            html_content += f"""
                    <div class="mobile-card">
                        <div class="mobile-card-header">
                            <span class="mobile-card-title">#{numero_item}</span>
                            <span class="quantidade-badge">{item.get('quantidade_total', item['quantidade'])} {unidade_escaped}</span>
                        </div>
                        <div class="mobile-card-content">
                            <span class="mobile-card-label">Valor Unitário:</span>
                            <span class="mobile-card-value">R$ {item['valor_unitario']:,.2f}</span>
                            <span class="mobile-card-label">Valor Total:</span>
                            <span class="mobile-card-value"><strong>R$ {item['valor_total']:,.2f}</strong></span>
                        </div>
                        <div class="mobile-card-desc item-com-imagem">
                            {desc_escaped}
                            <span class="icon-imagem">📷</span>
                            <div class="tooltip">
                                <img src="{imagem_escaped}" alt="Imagem do item #{numero_item}" onerror="this.style.display='none'; this.parentElement.querySelector('.tooltip-text').textContent='Imagem #{numero_item} não encontrada';">
                                <div class="tooltip-text">Item #{numero_item}: {desc_escaped[:50]}...</div>
                            </div>
                        </div>
                    </div>
"""
        else:
            html_content += f"""
                    <div class="mobile-card">
                        <div class="mobile-card-header">
                            <span class="mobile-card-title">#{numero_item}</span>
                            <span class="quantidade-badge">{item.get('quantidade_total', item['quantidade'])} {unidade_escaped}</span>
                        </div>
                        <div class="mobile-card-content">
                            <span class="mobile-card-label">Valor Unitário:</span>
                            <span class="mobile-card-value">R$ {item['valor_unitario']:,.2f}</span>
                            <span class="mobile-card-label">Valor Total:</span>
                            <span class="mobile-card-value"><strong>R$ {item['valor_total']:,.2f}</strong></span>
                        </div>
                        <div class="mobile-card-desc">
                            {desc_escaped}
                        </div>
                    </div>
"""
    
    html_content += """
                </div>
            </div>
            
            <div id="checklist" class="tab-content">
                <h2 class="section-title">✅ Checklist da Obra</h2>
                <div id="checklist-content">
                    PLACEHOLDER_CHECKLIST_HTML
                </div>
            </div>
        </div>
    </div>
    
    <script>
        function showTab(tabName) {{
            // Esconder todos os conteúdos
            document.querySelectorAll('.tab-content').forEach(content => {{
                content.classList.remove('active');
            }});
            
            // Remover active de todos os botões
            document.querySelectorAll('.tab-button').forEach(button => {{
                button.classList.remove('active');
            }});
            
            // Mostrar conteúdo selecionado
            document.getElementById(tabName).classList.add('active');
            
            // Ativar botão correspondente
            event.target.classList.add('active');
        }}
        
        // Melhorar posicionamento dos tooltips e suporte mobile
        document.addEventListener('DOMContentLoaded', function() {{
            const itemsComImagem = document.querySelectorAll('.item-com-imagem');
            let activeTooltip = null;
            
            function fecharTooltip() {{
                if (activeTooltip) {{
                    activeTooltip.classList.remove('active');
                    activeTooltip = null;
                }}
            }}
            
            itemsComImagem.forEach(item => {{
                const tooltip = item.querySelector('.tooltip');
                if (!tooltip) return;
                
                // Desktop: hover
                item.addEventListener('mouseenter', function(e) {{
                    if (window.innerWidth > 768) {{
                        // Ajustar posicionamento baseado na posição na tela
                        const rect = item.getBoundingClientRect();
                        const tooltipRect = tooltip.getBoundingClientRect();
                        
                        // Se tooltip sair da tela à direita, alinhar à direita
                        if (rect.left + tooltipRect.width > window.innerWidth) {{
                            tooltip.style.left = 'auto';
                            tooltip.style.right = '0';
                            tooltip.style.transform = 'none';
                        }} else {{
                            tooltip.style.left = '50%';
                            tooltip.style.right = 'auto';
                            tooltip.style.transform = 'translateX(-50%)';
                        }}
                        
                        // Se tooltip sair da tela acima, mostrar abaixo
                        if (rect.top - tooltipRect.height < 0) {{
                            tooltip.style.bottom = 'auto';
                            tooltip.style.top = '100%';
                            tooltip.style.marginBottom = '0';
                            tooltip.style.marginTop = '10px';
                        }} else {{
                            tooltip.style.bottom = '100%';
                            tooltip.style.top = 'auto';
                            tooltip.style.marginBottom = '10px';
                            tooltip.style.marginTop = '0';
                        }}
                    }}
                }});
                
                // Mobile: touch
                item.addEventListener('touchstart', function(e) {{
                    if (window.innerWidth <= 768) {{
                        e.preventDefault();
                        fecharTooltip();
                        item.classList.add('active');
                        activeTooltip = item;
                    }}
                }});
                
                // Fechar tooltip ao tocar fora (mobile)
                item.addEventListener('mouseleave', function(e) {{
                    if (window.innerWidth > 768) {{
                        // Desktop: fechar no mouseleave
                    }}
                }});
            }});
            
            // Fechar tooltip ao tocar em qualquer lugar (mobile)
            document.addEventListener('touchstart', function(e) {{
                if (window.innerWidth <= 768 && activeTooltip && !activeTooltip.contains(e.target)) {{
                    fecharTooltip();
                }}
            }});
            
            // Fechar tooltip ao clicar fora (mobile)
            document.addEventListener('click', function(e) {{
                if (window.innerWidth <= 768 && activeTooltip && !activeTooltip.contains(e.target)) {{
                    fecharTooltip();
                }}
            }});
        }});
    </script>
</body>
</html>
"""
    
    # Substituir placeholder pelo HTML do checklist
    html_content = html_content.replace('PLACEHOLDER_CHECKLIST_HTML', html_checklist)
    
    return html_content

# Processar
print("Processando planilha para identificar itens repetidos...")
print("="*60)

itens_agrupados = processar_planilha_para_cotacao()

# Incluir TODOS os itens (repetidos e únicos)
itens_repetidos = itens_agrupados

print(f"\n✅ Itens processados: {len(itens_agrupados)}")
print(f"✅ Total de itens (repetidos e únicos): {len(itens_repetidos)}")

if itens_repetidos:
    # Separar repetidos e únicos para estatísticas
    repetidos = [item for item in itens_repetidos if item['quantidade'] > 1]
    unicos = [item for item in itens_repetidos if item['quantidade'] == 1]
    print(f"\n📊 Estatísticas:")
    print(f"   - Itens repetidos: {len(repetidos)}")
    print(f"   - Itens únicos: {len(unicos)}")
    print(f"\n📊 Top 10 itens mais repetidos:")
    for i, item in enumerate(sorted(repetidos, key=lambda x: x['quantidade'], reverse=True)[:10], 1):
        print(f"   {i}. {item['descricao'][:60]}... - {item['quantidade']}x")

# Criar HTML
html_content = criar_html_cotacao(itens_agrupados)

# Salvar
with open('itens_cotacao_dartagnan.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\n✅ Página HTML criada: itens_cotacao_dartagnan.html")

# Gerar CSV também
df_csv = pd.DataFrame(itens_repetidos)
df_csv = df_csv[['descricao', 'quantidade_total', 'unidade', 'valor_unitario', 'valor_total']]
df_csv.columns = ['Descrição', 'Quantidade', 'Unidade', 'Valor Unitário (R$)', 'Valor Total (R$)']
df_csv.to_csv('itens_cotacao_dartagnan.csv', index=False, encoding='utf-8-sig')
print(f"✅ Arquivo CSV criado: itens_cotacao_dartagnan.csv")


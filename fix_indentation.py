
import os

def fix_file():
    source_file = 'app_test.py'
    target_file = 'app_test_2.py'
    
    if not os.path.exists(source_file):
        print(f"Error: {source_file} not found.")
        return

    with open(source_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Split point: after line 782 (index 782)
    # Verify content at split point
    # Line 782 (index 781) should be: "            row_e = best_match['row']\n"
    # Line 783 (index 782) should be: "            \n" or similar
    
    # Let's check a few lines around 782 to be sure
    # print(f"Line 782: {lines[781]}")
    # print(f"Line 783: {lines[782]}")
    
    part1 = lines[:782]
    part2 = lines[782:]

    new_rule = """            # >>> NOVA REGRA DE SEGURANÇA <<<
            # Se NÃO for Casa de Portugal e HÁ documento, só aceitamos o match se:
            #  - o documento da entrada corresponde ao da saída
            #  - OU se não há documento na entrada mas a similaridade é muito alta (>85) e quantidade próxima
            if doc_num and doc_num != '' and not destino_eh_cp:
                doc_num_e = row_e['doc_num']
                qtd_e_temp = float(row_e.get('qt_entrada', 0))
                diferenca_qtd_abs = abs(qtd_s - qtd_e_temp)
                tolerancia_qtd = max(1, 0.1 * max(qtd_s, 1))  # 10% ou 1 unidade
                
                # Se o documento da entrada não corresponde ao da saída
                if doc_num_e != doc_num:
                    # Rejeita o match - considera como "não encontrado"
                    stats['nao_encontrados'] += 1
                    stats['nao_conformes'] += 1
                    
                    motivo = f"Documento {doc_num} não encontrado na entrada"
                    
                    analise.append([
                        data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num,
                        produto_s, "-",  # Produto (Entrada)
                        row_s.get('especie', ''), valor_s, None, None,
                        qtd_s, None, None,
                        "❌ Não Conforme", motivo, "-",
                        "Sem correspondência encontrada", "-"
                    ])
                    continue
            # <<< FIM DA NOVA REGRA >>>
"""

    # Ensure new_rule ends with newline if needed
    if not new_rule.endswith('\n'):
        new_rule += '\n'

    with open(target_file, 'w', encoding='utf-8') as f:
        f.writelines(part1)
        f.write(new_rule)
        f.writelines(part2)
    
    print(f"Successfully created {target_file} with corrected indentation.")

if __name__ == "__main__":
    fix_file()

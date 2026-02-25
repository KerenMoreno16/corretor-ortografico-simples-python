import pandas as pd
import sys
import os
import language_tool_python
from tqdm import tqdm

def analisar_ortografia_gramatica_excel(caminho_arquivo, coluna_texto):
    # Verifica se arquivo existe
    if not os.path.isfile(caminho_arquivo):
        print(f"Erro: Arquivo '{caminho_arquivo}' não encontrado.")
        return

    # Leitura do arquivo
    try:
        df = pd.read_excel(caminho_arquivo)
        print(f"Arquivo carregado: {len(df)} registros")
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return

    # Verifica se coluna existe
    if coluna_texto not in df.columns:
        print(f"Erro: Coluna '{coluna_texto}' não encontrada.")
        print(f"Colunas disponíveis: {df.columns.tolist()}")
        return

    # Inicializa o LanguageTool
    print("Inicializando LanguageTool...")
    tool = language_tool_python.LanguageTool('pt-BR')

    resultados = []
    total_erros = 0
    textos_com_erros = 0

    print(f"Analisando {len(df)} textos...\n")

    # Processa cada texto com barra de progresso
    for idx, texto in enumerate(tqdm(df[coluna_texto], desc="Analisando", unit="linha")):
        if not isinstance(texto, str) or not texto.strip():
            continue

        try:
            # Verifica erros
            matches = tool.check(texto)

            if matches:
                textos_com_erros += 1
                total_erros += len(matches)

                erros_detalhes = []
                for match in matches:
                    erros_detalhes.append({
                        'erro': match.matched_text,           
                        'mensagem': match.message,
                        'sugestoes': ', '.join(match.replacements[:3]) if match.replacements else 'N/A',
                        'tipo': match.rule_issue_type,
                        'categoria': match.category
                    })

                # Corrige o texto
                texto_corrigido = tool.correct(texto)

                resultados.append({
                    'id_registro': idx + 1,
                    'Texto Original': texto,
                    'Texto Corrigido': texto_corrigido,
                    'Quantidade Erros': len(matches),
                    'Erros Encontrados': '; '.join([f"{e['erro']} ({e['mensagem']})" for e in erros_detalhes[:5]]),
                    'Detalhes Erros': str(erros_detalhes),  # Salva todos os detalhes
                    'Tipo Principal': erros_detalhes[0]['tipo'] if erros_detalhes else 'N/A'
                })

        except Exception as e:
            print(f"Erro no registro {idx + 1}: {str(e)}")
            continue

    # Fecha o LanguageTool
    tool.close()

    # Salva resultados
    if resultados:
        df_resultados = pd.DataFrame(resultados)
        saida = "analise_ortografica_gramatical.xlsx"
        df_resultados.to_excel(saida, index=False)

        print(f"\n" + "="*60)
        print("ANÁLISE CONCLUÍDA!")
        print("="*60)
        print(f"Total de textos analisados: {len(df)}")
        print(f"Textos com erros: {textos_com_erros}")
        print(f"Total de erros encontrados: {total_erros}")
        print(f"Resultados salvos em: {saida}")
        print("="*60)

        # Analisando os Erros
        if total_erros > 0:
            print("\nTOP 5 ERROS MAIS FREQUENTES:")
            todos_erros = []
            for r in resultados:
                if r['Detalhes Erros'] != '[]':
                    import ast
                    try:
                        erros = ast.literal_eval(r['Detalhes Erros'])
                        todos_erros.extend([e['erro'] for e in erros])
                    except:
                        pass
            from collections import Counter
            top_erros = Counter(todos_erros).most_common(5)
            for i, (erro, qtd) in enumerate(top_erros, 1):
                print(f"  {i}. '{erro}' - {qtd} ocorrências")
    else:
        print("\nNenhum erro encontrado!")

    return {
        'total_textos': len(df),
        'textos_com_erros': textos_com_erros,
        'total_erros': total_erros,
        'arquivo_saida': saida if resultados else None
    }


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("="*60)
        print("USO DO SCRIPT")
        print("="*60)
        print("python script.py <arquivo_excel> <coluna_texto>")
        print("\nExemplo:")
        print('  python analise.py "dialogo.xlsx" "Falas"')
        print("="*60)
    else:
        try:
            analisar_ortografia_gramatica_excel(sys.argv[1], sys.argv[2])
        except KeyboardInterrupt:
            print("\n Processo interrompido pelo usuário.")
        except Exception as e:
            print(f"\nErro crítico: {e}")
            import traceback
            traceback.print_exc()
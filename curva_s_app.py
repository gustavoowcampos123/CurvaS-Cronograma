# Função para calcular o caminho crítico com verificações e exibir número da linha das tarefas sem predecessoras
def calculate_critical_path(df):
    G = nx.DiGraph()
    
    # Verificar se a coluna de predecessoras existe
    if 'Predecessoras' in df.columns:
        for i, row in df.iterrows():
            # Verificar se há predecessoras
            if pd.notna(row['Predecessoras']):
                predecessoras = str(row['Predecessoras']).split(';')
                for pred in predecessoras:
                    pred_clean = remove_prefix(pred.split('-')[0].strip())
                    try:
                        # Verificar se a duração não é nula e está no formato correto
                        duration = int(row['Duração'].split()[0])
                        if pred_clean:
                            G.add_edge(pred_clean, row['Nome da tarefa'], weight=duration)
                    except ValueError:
                        st.error(f"Duração inválida para a tarefa {row['Nome da tarefa']}: {row['Duração']} (linha {i+1})")
            else:
                # Exibir número da linha quando não houver predecessoras
                st.warning(f"A tarefa {row['Nome da tarefa']} (linha {i+1}) não tem predecessoras.")
    else:
        st.error("A coluna 'Predecessoras' não foi encontrada no arquivo.")
    
    if len(G.nodes) == 0:
        st.error("O grafo de atividades está vazio. Verifique as predecessoras e a duração das atividades.")
        return []
    
    try:
        critical_path = nx.dag_longest_path(G, weight='weight')
        return critical_path
    except Exception as e:
        st.error(f"Erro ao calcular o caminho crítico: {e}")
        return []

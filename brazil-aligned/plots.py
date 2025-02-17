import plotly.subplots as sp
import pandas as pd
import plotly.graph_objects as go

def plot_grid_all_agencies(final_df: pd.DataFrame, cols: int = 2, height: int = 300, width: int = 800, **kwargs) -> go.Figure:
    # Filtrar apenas as categorias relevantes
    valid_categories = ['contra', 'alinhada']
    agencies = final_df[final_df['category'].isin(valid_categories)]['agency'].unique()
    agencies = [agency.split()[0] for agency in agencies]  # Manter apenas a primeira palavra do nome da agência
    num_agencies = len(agencies)
    rows = (num_agencies // cols) + (num_agencies % cols > 0)  # Definir número de linhas

    fig = sp.make_subplots(rows=rows, cols=cols, subplot_titles=[f'Agência {agency}' for agency in agencies])

    row, col = 1, 1
    line_colors = {'alinhada': 'blue', 'contra': 'red'}
    marker_colors = {'alinhada': 'blue', 'contra': 'red'}

    for agency in agencies:
        df_agency = final_df[final_df['agency'].str.startswith(agency)]
        
        # Criar gráfico de linha
        for category in df_agency['category'].unique():
            df_cat = df_agency[df_agency['category'] == category]
            color = line_colors.get(category, 'black')  # Definir cor padrão se categoria não for encontrada
            fig.add_trace(
                go.Scatter(
                    x=df_cat['year'], y=df_cat['conc_parc'],
                    mode='lines+markers',
                    line=dict(dash='solid' if category == 'alinhada' else 'dash', color=color),
                    marker=dict(color=color),
                    name=f'{category}',
                    legendgroup=category,
                    showlegend=row == 1 and col == 1  # Exibir legenda apenas uma vez
                ), row=row, col=col
            )
        
        # Adicionar pontos específicos para os presidentes
        df_presidents = final_df[['year', 'president']].drop_duplicates()
        for president in df_presidents['president'].unique():
            df_pres = df_agency[df_agency['president'] == president]
            fig.add_trace(
                go.Scatter(
                    x=df_pres['year'], y=df_pres['conc_parc'],
                    mode='markers',
                    marker=dict(size=10, symbol='circle', color='black'),
                    name=president,
                    legendgroup=president,
                    showlegend=row == 1 and col == 1  # Exibir legenda apenas uma vez
                ), row=row, col=col
            )
        
        col += 1
        if col > cols:
            col = 1
            row += 1

    fig.update_layout(
        title='Média de Conformidade das Agências ao Longo dos Anos',
        height = height * rows, width = width,  # Ajustar tamanho do grid
        showlegend=True,
        legend=dict(x=0.5, y=-0.2, orientation='h')  # Mover legenda para a parte inferior
    )

    return fig

if __name__ == "__main__":
    final_df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
    plot_grid_all_agencies(final_df, width=300)
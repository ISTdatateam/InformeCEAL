'''
    # Procesamiento de df_rec para crear df_reco
    df_reco = pd.DataFrame(columns=['Dimensión', 'Rubro', 'Recomendación'])

    for index, row in df_rec.iterrows():
        dimension = row['Dimensión']
        for i in range(0, len(row) - 1, 2):
            rubro_col = f'Rubro.{i // 2}' if i // 2 > 0 else 'Rubro'
            recomendacion_col = f'Recomendación.{i // 2}' if i // 2 > 0 else 'Recomendación'

            if rubro_col in row and recomendacion_col in row:
                rubro = row[rubro_col]
                recomendacion = row[recomendacion_col]

                if pd.notna(rubro) and pd.notna(recomendacion):
                    df_reco = pd.concat(
                        [df_reco, pd.DataFrame([[dimension, rubro, recomendacion]], columns=df_reco.columns)],
                        ignore_index=True)

    # Realizar el merge entre df_ciiu y df_reco
    df_recomendaciones = pd.merge(df_ciiu, df_reco, left_on='Sección', right_on='Rubro', how='left')


    with pd.ExcelWriter('recomendaciones.xlsx', engine='xlsxwriter') as writer:
        df_recomendaciones.to_excel(writer, sheet_name='df_recomendaciones', index=False)

'''
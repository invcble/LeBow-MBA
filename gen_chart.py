import pandas as pd
from chart_single import plot_and_save_single
from chart_multi import plot_and_save_multi

def save_chart(row, df):
    scale_5_columns = {'CHAL', 'CONN', 'DOER', 'INNOV', 'ORG', 'PC', 'TB'}
    demographic_columns = ['Name', 'MBA program', 'Email', 'Academic years', 'Academic term', 'DOB', 'Gender', 'Profession', 'Work Ex. years']
    strength_columns = ['Total_Size', 'Strong', 'Weak']
    breadth_columns = ['High_Lvl', 'Ext', 'Cross_Func']

    exclude_cols = demographic_columns + strength_columns + breadth_columns
    unique_columns = [col for col in df.columns if col not in exclude_cols and not col.endswith('_MBAavg')]

    # Iterate over the unique columns and plot using the appropriate scale
    for col in unique_columns:
        # Determine if the current column should use a scale of 5 or 7
        switch_scale = col in scale_5_columns
        
        # Get the corresponding _MBAavg and _SD values
        your_score = row[col] if not pd.isna(row[col]) else 0
        mba_avg_col = f'{col}_MBAavg'
        sd_col = f'{col}_SD'
        
        if mba_avg_col in row and sd_col in row:
            drexel_mba_avg = row[mba_avg_col]
            standard_deviation = row[sd_col]
            
            plot_and_save_single(col, your_score, drexel_mba_avg, standard_deviation, switch_scale=switch_scale)

    your_scores = [row[col] if not pd.isna(row[col]) else 0 for col in strength_columns]
    drexel_mba_scores = [row[f'{col}_MBAavg'] for col in strength_columns if f'{col}_MBAavg' in row]

    plot_and_save_multi(your_scores, drexel_mba_scores, switch_categories=False)

    your_scores_ = [row[col] if not pd.isna(row[col]) else 0 for col in breadth_columns]
    drexel_mba_scores_ = [row[f'{col}_MBAavg'] for col in breadth_columns if f'{col}_MBAavg' in row]

    plot_and_save_multi(your_scores_, drexel_mba_scores_, switch_categories=True)
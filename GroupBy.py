import pandas as pd

# This script is used AFTER executing the GroupByTotal.py script
# Using this Script to group by name of clients, this will allow me to see how much each customer is spending. 

df = pd.read_excel('Modified Client Transaction February 2023.xlsx')

grouped_df = df.groupby('Persona')


result_df = pd.DataFrame()


for persona, group in grouped_df:
    group.loc[len(group.index)] = [''] * len(group.columns) # adding an empty row at the end of each group by creating a new row with empty strings as values for each column
    result_df = result_df.append(group)


writer = pd.ExcelWriter('Modified NEW Client Transaction February 2023.xlsx')

result_df.to_excel(writer, sheet_name='Combined', index=False)

writer.save()


df  = pd.read_excel('Modified NEW Client Transaction February 2023.xlsx')

# names_to_drop = ['Consumidor Final']

# df = df[(df['Descuento'] > 0) | (~df['Persona'].isin(names_to_drop))]


grouped_df = df.groupby('Persona')

product_lists = grouped_df['Nombre Manual'].apply(list)

total_spent = grouped_df['Total'].sum()

result_df = pd.DataFrame({'Products Bought' : product_lists, 'Total Amount Spent' : total_spent})

result_df = result_df.sort_values('Total Amount Spent', ascending=False).reset_index()

result_df = result_df.drop(result_df.loc[result_df['Persona'] == 'MIRANDA ARTEAGA PIERO PABLO'].index)


result_df.to_excel('Customer Spending February 2023.xlsx', index=False)



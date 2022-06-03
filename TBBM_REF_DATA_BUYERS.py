import pandas as pd


# Read the Excel file and make transformations to dates and null values
df = pd.read_excel('TBBM_REF_DATA_BUYERS.xlsx')
df = df.fillna('')
df['EFFECT_TO_DAT'] = df['EFFECT_TO_DAT'].apply(lambda x: x.strftime("%d-%b-%y"))
df['EFFECT_FROM_DAT'] = df['EFFECT_FROM_DAT'].apply(lambda x: x.strftime("%d-%b-%y"))


def add_comas(str):
    return f"'{str}'"


# Create one unique SQL file per DOMAIN_NAME in the input excel and create one SQL file for all the DOMAIN_NAME
CompleteFile = ''

for domain in range(len(df['DOMAIN_NAME'].unique())):
    subDF = df.loc[df['DOMAIN_NAME'] == df['DOMAIN_NAME'].unique()[domain]]
    insert = ''
    columns = ', '.join(i for i in list(df.columns))
    for i in range(len(subDF)):
        if i != len(subDF)-1:
            insert += f"INTO TBBM_REF_DATA_BUYERS ({columns}) VALUES ({', '.join(add_comas(i) for i in list(subDF.iloc[i]))})\n\t"
        else:
            insert += f"INTO TBBM_REF_DATA_BUYERS ({columns}) VALUES ({', '.join(add_comas(i) for i in list(subDF.iloc[i]))})"
    code = f"""INSERT ALL
    {insert}
SELECT 1 FROM DUAL;"""

    CompleteFile += code + '\n\n\n\n'

    filename = f"./BUYER_MANAGEMENT_REFERENCE_DATA/{df['DOMAIN_NAME'].unique()[domain]}.sql"
    with open(filename, 'w') as f:
        f.write(code)

filename = "./COMPLETE_REFERENCE_DATA/BUYER_MANAGEMENT_REFERENCE_DATA.sql"
with open(filename, 'w') as f:
    f.write(CompleteFile)
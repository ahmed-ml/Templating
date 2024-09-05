import pandas as pd
import numpy as np
import os
import glob
from docxtpl import DocxTemplate
from docx import Document

def input(source_file, nom_mission):
    if nom_mission == 'AG4':
        df = pd.read_excel(source_file, header=1,
                           dtype={'COEFF D_ECHANGE ACANTHE': str, 'COEFF D_ECHANGE CIMOFAT': str,
                                  'COEFF D_ECHANGE VALOREST': str
                               , 'COEFF D_ECHANGE SOLIANCE': str, 'COEFF D_ECHANGE NEOCLAR': str},
                           usecols=list(range(0, 44)))
        df.columns = df.columns.str.replace(' ', '_').str.replace("’", "_").str.replace("'", "_")
        print(df.columns)
        df['ADRESSE_SC'] = df['ADRESSE_SC'].str.title()
        df = df[(df['NOM_SC'] != 'JOSSY') & (df['NOM_SC'] != 'SEPT SEPT')]
        df = df[df['NOM_SC'].notna()]

        for col in ['VILLE_SC', 'VILLE_RCS_SC', 'PRENOM_NOM_GERANT_SC', 'CODE_ASSOCIE_GERANT',
                    'DUREE_SC']:
            df[col] = df[col].astype(str)

        # Préparation des champs numériques (Séparateur de milliers et 2 décimales) :
        columns_float = df.columns[(df.columns.str.contains('MONTANT')) | (df.columns.str.contains('VALEUR'))
                                      | (df.columns.str.contains('PRIME')) | (df.columns.str.contains('NOUVEAU'))
                                      | (df.columns.str.contains('CAPITAL')) | (df.columns.str.contains('ACTIF'))]
        for col in df[columns_float]:
            df[col] = df[col].apply(lambda x: '{:,.2f}'.format(x).replace(',', ' ').replace('.', ',').replace(' ', '.'))

        df['NUMERO_RCS_SC'] = df['NUMERO_RCS_SC'].replace(' ', '')

        df['NUMERO_RCS_SC'] = pd.to_numeric(df['NUMERO_RCS_SC']).apply(
            lambda x: '{:,.0F}'.format(x).replace(',', ';').
            replace('.', ' ').replace(';', ' '))



        for col in ['DUREE_SC', 'CODE_POSTAL_SC']:
            df[col] = df[col].astype(str).str.split('.').str[0].astype(int)

        df['NUMERO_ANNEXE'] = df['NUMERO_ANNEXE'].astype(int)

        df = df.sort_values(by='NUMERO_ANNEXE')
        return df
    elif nom_mission == 'AG4_Annexe':
        df = pd.read_excel(source_file,
                           sheet_name="Feuil1", header=1,
                           dtype={'COEFF D_ECHANGE ACANTHE': str, 'COEFF D_ECHANGE CIMOFAT': str,
                                  'COEFF D_ECHANGE VALOREST': str
                               , 'COEFF D_ECHANGE SOLIANCE': str, 'COEFF D_ECHANGE NEOCLAR': str})

        df.columns = df.columns.str.replace(' ', '_').str.replace("’", "_").str.replace("'", "_")

        print(df.columns)
        # Formater le champ Adresse_SC en miniscule :
        df['ADRESSE_SC'] = df['ADRESSE_SC'].str.title()
        df = df[(df['NOM_SC'] != 'JOSSY') & (df['NOM_SC'] != 'SEPT SEPT')]
        df = df[df['NOM_SC'].notna()]

        # Préparation des champs numériques (Séparateur de milliers et 2 décimales) :
        columns_float = df.columns[(df.columns.str.contains('MONTANT')) | (df.columns.str.contains('VALEUR'))
                                   | (df.columns.str.contains('PRIME')) | (df.columns.str.contains('NOUVEAU'))
                                   | (df.columns.str.contains('ACTIF'))]

        # def format_number(number):
        # return f'{number:,.2f}'.replace(',', ' ').replace('.', ',')

        for col in df[columns_float]:
            df[col] = df[col].apply(lambda x: '{:,.2f}'.format(x).replace(',', ' ').replace('.', ',').replace(' ', '.'))

        # df['CAPITAL_SC'] = df['CAPITAL_SC'].str.replace('\xa0', '').str.replace(' ', '').str.replace(',', '.')  # removWDWDe non-breaking space character
        df['CAPITAL_SC'] = pd.to_numeric(df['CAPITAL_SC']).apply(
            lambda x: '{:,.2f}'.format(x).replace(',', ';').replace('.', ',').replace(';', '.'))

        # def transform_value(value):
        # Remove commas as thousands separators
        # value = value.replace(",", "")
        # Convert the value to a float to remove trailing zeros
        #  value = float(value)
        # Format the value as an integer with periods as thousands separators
        #   value = '{:,.0f}'.format(value).replace(",", ".")
        # return value

        # for col in columns_int:
        # df[col] = df[col].apply(transform_value)

        # print(df.iloc[:3])
        # Formater le champ sur 6 decimale:
        # Séparateur de milliers pour le champ NUMERO_RCS_SC :

        df['NUMERO_RCS_SC'] = df['NUMERO_RCS_SC'].replace(' ', '')

        df['NUMERO_RCS_SC'] = pd.to_numeric(df['NUMERO_RCS_SC']).apply(lambda x: '{:,.0F}'.format(x).replace(',', ';').
                                                                       replace('.', ' ').replace(';', ' '))

        numbers = df.columns[df.columns.str.contains('NOMBRE')]
        for column in numbers:
            df[column] = df[column].apply(lambda x: '{:,.0f}'.format(x).replace(',', '.'))

        numbers = df.columns[df.columns.str.contains('COEFF')]
        for column in numbers:
            df[column] = df[column].astype(str).str.replace('o', ',')
            # for col in df[numbers]:
            # df[col] = df[col].str.replace('.', '').str.replace(',', '.')  # Replace comma with period
        # df[col] = pd.to_numeric(df[col]).apply(lambda x: '{:,.0f}'.format(x).replace(',', ';').replace('.', ' ').replace(';', ' '))

        for col in ['VILLE_SC', 'VILLE_RCS_SC', 'CIVILITE_GERANT', 'PRENOM_NOM_GERANT_SC', 'CODE_ASSOCIE_GERANT',
                    'DUREE_SC']:
            df[col] = df[col].astype(str)

        for col in ['DUREE_SC', 'CODE_POSTAL_SC']:
            df[col] = df[col].astype(str).str.split('.').str[0].astype(int)

        df['NUMERO_ANNEXE'] = df['NUMERO_ANNEXE'].astype(int)

        df = df.sort_values(by='NUMERO_ANNEXE')

        return df

    elif nom_mission == 'AG3':
        df = pd.read_excel(source_file,
                              dtype=str,
                              sheet_name='AG 3', header=0)

        df.columns = df.columns.str.replace(' ', '_').str.replace("’", "_")

        print(df.columns)
        # Formater le champ Adresse_SC en miniscule :
        df['ADRESSE_SC'] = df['ADRESSE_SC'].str.title()
        df = df[(df['NOM_SC'] != 'JOSSY') & (df['NOM_SC'] != 'SEPT SEPT')]

        columns_float = df.columns[(df.columns.str.contains('MONTANT')) | (
                    df.columns == 'VALEUR_NOMINALE_PART_SC_AVANT_AUGMENTATION_DE_CAPITAL')
                                   | (df.columns.str.contains('NOMBRE')) | (df.columns == 'PRIME_D_EMISSION') | (
                                       df.columns.str.contains('NOUVEAU'))]

        for col in df[columns_float]:
            df[col] = df[col].str.replace('\xa0', '').str.replace(' ', '').str.replace(',',
                                                                                       '.')  # remove non-breaking space character
            df[col] = pd.to_numeric(df[col], errors="coerce").apply(
                lambda x: '{:,.2f}'.format(x).replace(',', ';').replace('.', ',').replace(';', '.'))

        df['CAPITAL_SC'] = df['CAPITAL_SC'].str.replace('\xa0', '').str.replace(' ', '').str.replace(',',
                                                                                                     '.')  # remove non-breaking space character
        df['CAPITAL_SC'] = pd.to_numeric(df['CAPITAL_SC']).apply(
            lambda x: '{:,.2f}'.format(x).replace(',', ';').replace('.', ',').replace(';', '.'))

        columns_int = df.columns[(df.columns.str.contains('SOCIALES')) |
                                 (df.columns.str.contains('NOMBRE_DE_PARTS_DE_LA_SCF_AVANT_REDUCTION'))]
        print(df[columns_int])
        for col in columns_int:
            df[col] = pd.to_numeric(df[col].str.replace('.', '').str.replace(',', '.'), errors='coerce')
            df[col] = df[col].apply(lambda x: '{:,.0f}'.format(x).replace(',', ';').replace('.', ',')).str.replace(';',
                                                                                                                   '.')

        print(df.iloc[:3])
        # Formater le champ sur 6 decimale:
        df['NOUVELLE_VALEUR_NOMINALE_PART_SC'] = df['NOUVELLE_VALEUR_NOMINALE_PART_SC'].apply(
            lambda x: '{:,.6f}'.format(float(x)).replace('.', ','))
        df['PRIME_D_EMISSION_PAR_PART_SOCIALE'] = df['PRIME_D_EMISSION_PAR_PART_SOCIALE'].replace('-', np.nan)
        df['PRIME_D_EMISSION_PAR_PART_SOCIALE'] = df['PRIME_D_EMISSION_PAR_PART_SOCIALE'].apply(
            lambda x: '{:,.6f}'.format(float(x)).replace('.', ','))
        # Séparateur de milliers pour le champ NUMERO_RCS_SC :

        df['NUMERO_RCS_SC'] = df['NUMERO_RCS_SC'].replace(' ', '')

        df['NUMERO_RCS_SC'] = pd.to_numeric(df['NUMERO_RCS_SC']).apply(lambda x: '{:,.0F}'.format(x).replace(',', ';').
                                                                       replace('.', ' ').replace(';', ' '))

        numbers = df.columns[df.columns.str.contains('NUMERO_D')]
        for column in numbers:
            df[column] = pd.to_numeric(df[column], errors='coerce').astype('Int64')

        return df
    elif nom_mission == 'AG2':
        df = pd.read_excel(source_file, dtype = str, sheet_name = 'AG 2', header = 0)
        df.columns = df.columns.str.replace(' ', '_').str.replace("'", "_")

        print(df.columns)

        # Formater le champ Adresse_SC en miniscule :
        df['ADRESSE_SC'] = df['ADRESSE_SC'].str.title()
        df = df[(df['NOM_SC'] != 'JOSSY') & (df['NOM_SC'] != 'SEPT SEPT')]

        # df['MONTANT_DISTRIBUTION_RESULTAT_2023'] = df['MONTANT_DISTRIBUTION_RESULTAT_2023'].replace("0",np.nan)
        # Préparation des champs numériques (Séparateur de milliers et 2 décimales) :
        columns_float = df.columns[(df.columns.str.contains('MONTANT')) | (df.columns.str.contains('CAPITAL'))]

        for col in df[columns_float]:
            df[col] = df[col].str.replace('\xa0', '').str.replace(' ', '').str.replace(',',
                                                                                       '.')  # remove non-breaking space character
            df[col] = pd.to_numeric(df[col], errors='coerce').apply(
                lambda x: '{:,.2f}'.format(x).replace(',', ';').replace('.', ',').replace(';', '.'))

        # Séparateur de milliers pour le champ NUMERO_RCS_SC :
        df['NUMERO_RCS_SC'] = df['NUMERO_RCS_SC'].replace(' ', '')
        df['NUMERO_RCS_SC'] = pd.to_numeric(df['NUMERO_RCS_SC']).apply(lambda x: '{:,.0F}'.format(x).replace(',', ';').
                                                                       replace('.', ' ').replace(';', ' '))
        return df
    elif nom_mission == 'AG1':
        df = pd.read_excel("Templates_AG1/Tableau d'automatisation AG 1 v21.04.2023.xlsx", dtype=str,
                           sheet_name='Feuil1', header=1)

        df.columns = df.columns.str.replace(' ', '_').str.replace("'", "_")

        print(df.columns)

        # Formater le champ Adresse_SC en miniscule :
        df['ADRESSE_SC'] = df['ADRESSE_SC'].str.title()

        # Filtrer sur les lignes totalement remplies et 2 entités « SEPT SEPT » et « SMC » sont des sociétés commerciales => pas génération automatique.:
        df = df[(df['MONTANT_DU_RAN'].notna()) & (df['NOM_SC'] != 'SEPT SEPT') & (df['NOM_SC'] != 'SMC')]

        # Préparation des champs numériques (Séparateur de milliers et 2 décimales) :
        columns_float = df.columns[(df.columns.str.contains('MONTANT')) | (df.columns.str.contains('SOMME'))]
        print(df)
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        # Entité JOSSY est une société à capial variable, donc SOCIAL_SC a des textes dedans.
        if (df['NOM_SC'] == 'JOSSY').any():
            df['CAPITAL_SC'] = df['CAPITAL_SC'].astype(str)
        else:
            df['CAPITAL_SC'] = pd.to_numeric(df['CAPITAL_SC']).apply(
                lambda x: '{:,.2f}'.format(x).replace(',', ';').replace('.', ',').replace(';', '.'))
        df['MONTANT_DISTRIBUE_2021'] = df['MONTANT_DISTRIBUE_2021'].str.replace(" ", "").str.replace(",", ".")
        df['MONTANT_DISTRIBUE_2020'] = df['MONTANT_DISTRIBUE_2020'].str.replace(" ", "").str.replace(",", ".")

        montant_distribue = df.columns[df.columns.str.contains('MONTANT_DISTRIBUE_20')]

        for col in df[columns_float]:
            df[col] = df[col].str.replace('\xa0', '').str.replace(' ', '').str.replace(',',
                                                                                       '.')  # remove non-breaking space character
            df[col] = pd.to_numeric(df[col]).apply(
                lambda x: '{:,.2f}'.format(x).replace(',', ';').replace('.', ',').replace(';', '.'))

        # Séparateur de milliers pour le champ NUMERO_RCS_SC :
        df['NUMERO_RCS_SC'] = df['NUMERO_RCS_SC'].replace(' ', '')

        df['NUMERO_RCS_SC'] = pd.to_numeric(df['NUMERO_RCS_SC']).apply(
            lambda x: '{:,.0F}'.format(x).replace(',', ';').
            replace('.', ' ').replace(';', ' '))
        return df

def output(df, word_template_paths,nom_mission,output_folder, word_template_sans, word_template_avec):
        """

        Handle the output logic for rendering templates and saving them.
        """
        # Specify your code related to rendering the templates and saving them based on the DataFrame
        # ...

        if nom_mission == 'AG2':
            for record in df.to_dict(orient='records'):
                if record['MONTANT_DISTRIBUTION_DE_RESERVES'] == "0,00":
                    doc = DocxTemplate(word_template_sans)
                    doc.render(record)
                    filename = os.path.splitext(os.path.basename(word_template_sans))[0]
                    # filename = filename.split('SCF')[0].split('vEY')[0].split('trame')[0].split('1')[0]
                    output_path = f'{output_folder}/{record["NOM_SC"]}_PV AG 2_4 juillet 2023.docx'
                    doc.save(output_path)
                else:
                    doc = DocxTemplate(word_template_avec)
                    doc.render(record)
                    filename2 = os.path.splitext(os.path.basename(word_template_avec))[0]
                    # filename = filename.split('SCF')[0].split('vEY')[0].split('trame')[0].split('1')[0]
                    output_path = f'{output_folder}/{record["NOM_SC"]}_PV AG 2_4 juillet 2023.docx'
                    doc.save(output_path)
        elif nom_mission == 'AG4_Annexe':
            sté = ['SOLIANCE', 'ACANTHE', 'VALOREST', 'NEOCLAR']

            for company in sté:
                word_template_paths = glob.glob(f'Templates_AG4/Annexes {company} - AGE 14 octobre 2023.docx')

                if os.path.exists(f"Templates_AG4/Output_{company}"):
                    files = glob.glob(f'Templates_AG4/Output_{company}/*')
                    for f in files:
                        os.remove(f)
                else:
                    os.makedirs(f"Templates_AG4/Output_{company}")

                merged_doc = Document()
                for record in df.to_dict(orient='records'):
                    for word_template_path in word_template_paths:
                        doc = DocxTemplate(word_template_path)
                        print(f"Processing template: {word_template_path}")
                        doc.render(record)
                        filename = os.path.splitext(os.path.basename(word_template_path))[0]
                        # filename = filename.split('SCF')[0].split('vEY')[0].split('trame')[0].split('1')[0]
                        output_path = f'Templates_AG4/Output_{company}/{record["NOM_SC"]}_{company}_Annexes_AG 14 octobre 2023.docx'
                        print(f"Output path: {output_path}")
                        doc.save(output_path)
                        for element in doc.element.body:
                            merged_doc.element.body.append(element)
                        merged_doc.add_page_break()

                merged_doc.save(f'Templates_AG4/{company}_Annexes_AG 14 octobre 2023.docx')

        else:
            for record in df.to_dict(orient='records'):
                for word_template_path in word_template_paths:
                    doc = DocxTemplate(word_template_path)
                    print(f"Processing template: {word_template_path}")
                    doc.render(record)
                    if  nom_mission == 'AG4':
                        output_path = f'{output_folder}/{record["NOM_SC"]}_AG de dissolution_10 octobre 2023.docx'
                    elif nom_mission == 'AG3':
                        output_path = f'Templates_AG2_AG3/AG3/Output/{record["NOM_SC"]}_PV AG 3_13 juillet 2023.docx'
                    elif nom_mission == 'AG1':
                        filename = os.path.splitext(os.path.basename(word_template_path))[0]
                        filename = filename.split('SCF')[0].split('vEY')[0].split('trame')[0].split('1')[0]
                        output_path = f'Templates_AG1/AG1/Output/{record["NOM_SC"]} {filename.replace("v07.04.2023 clean", "")}du 19.05.2023 .docx'
                    print(f"Output path: {output_path}")
                    doc.save(output_path)


def clear_output_folder(output_folder):
        """
        Clear the output folder or create it if it doesn't exist.
        """
        # Supprimer le répertoire et non par fichier
        try:
            if os.path.exists(output_folder):
                files = glob.glob(os.path.join(output_folder, '*'))
                for f in files:
                    try:
                        os.remove(f)
                    except PermissionError as e:
                        print(f"PermissionError: {e}. Unable to delete {f}.")
                        # Log the error or take appropriate action
            else:
                os.makedirs(output_folder)

        except Exception as e:
            print(f"An error occurred while clearing the output folder: {e}")
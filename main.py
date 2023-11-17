import os
from io import BytesIO

import pandas as pd
import streamlit as st
from langchain.chat_models import ChatOpenAI
from langchain.prompts import HumanMessagePromptTemplate, ChatPromptTemplate
from langchain.schema import SystemMessage

#os.environ['OPENAI_API_KEY'] = "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"


def calculate_square(description):
    chat_template = ChatPromptTemplate.from_messages(
        [
            SystemMessage(
                content=(
                    """
                    Act as a agent who Standardize Fabric description given to standardized format  examples for Standard Formats are '''91% Cotton , 9% Spandex
                    57% Cotton , 38% Modal , 5% Spandex
                    100% Cotton 
                    55% Cotton , 37% Polyester, 8% Elastain
                    53% Cotton , 17.5% Modal , 17.5% Recycle Polyester, 12% Spandex
                    83% Cotton , 17% Elastain
                    95% Cotton , 5% Spandex
                    92% Organic Cotton , 8% Spandex
                    55% S.Cotton , 37% M.Modal , 8% Spandex
                    53% S.Cotton , 35% M.Modal , 12% Spandex
                    95% BCI Cotton , 5% Elastain
                    96% Cotton , 4% Elastain
                    96% M.Modal , 4% Elastain
                    57% Cotton , 38% Polyester , 5% Elastain
                    88% Recycle Polyester , 12% Spandex
                    90% BCI Cotton , 10% Lycra
                    85% Nylon , 15% Spandex
                    54% Cotton , 21% Polyester , 20% Viscose , 5% Spandex
                    97% Cotton , 3% Spandex
                     92% Polyester , 8% Spandex
                    94% Lenzing Modal , 6% Spandex
                    95% Ecovera Lenzing Viscose , 5% Spandex
                    85% Cooling Nylon , 15% Spandex
                    85% Polyester , 15% Elastane
                    87% Cotton , 13% Elastane
                    80% Cotton , 20% Polyester 
                    79% Polyester , 21% Elastane
                    100% Cotton
                    95% BCI Cotton , 5% Spandex
                    92% Recycle Polyester , 8% Spandex
                    90% Recycle Polyester , 10% Spandex
                    55% Cotton , 6% Organic Cotton , 33% Recycle Polyester , 6% Elastane
                    90% BCI Compact Cotton , 10% Lycra
                    53% Supima Cotton , 35% Micro Modal , 12% Lycra
                    74% BCI Cotton , 21% Recycle Cotton , 5% Spandex
                    90% BCI Cotton , 10% Elastane
                    90% BCI Cotton , 10% Elastane 
                    47% BCI Cotton , 41% Recycle Polyester, 12 % Dyeable Spandex
                    53% Supima Cotton , 35% Micro Modal , 12% Lycra 
                    61% BCI Cotton Marl , 33% Recycle Polyester , 6% Spandex 
                    53% Cotton , 36% Modal , 11% Spandex 
                    95% BCI Cotton , 5% Spandex 
                    90% Cotton , 10% Organic Cotton 
                    100% Cotton Combed Compact
                    96% Cotton , 4% Elastane Combed Compact
                    100% Cotton  
                    95% Cotton , 5% Lycra Combed Compact
                    90% BCI Cotton , 10% Organic Cotton Combed Compact
                    97% Cotton , 3% Elastane
                    97% Pima Cotton , 3% Spandex
                    100% Pima Cotton
                    100% Organic Cotton
                    95% Organic Cotton , 5% Lycra
                    '''  these Standerd descpertions are made from core fabrics listed '''Cotton
                    Polyester
                    Silk
                    Lurex
                    Viscose
                    Polyamide
                    Organic Cotton
                    Spandex
                    Modal
                    BCI Cotton (Better Cotton Initiative)
                    Recycled Polyester
                    Elastane
                    Micro Modal
                    Lycra
                    Ecovera Lenzing Viscose
                    Cooling Nylon
                    Recycled Cotton
                    Nylon (Recycled)
                    BCI Compact Cotton
                    Supima Cotton
                    Pima Cotton
                    Easy Set Lycra
                    Combed Compact Cotton
                    RECPOL (Recycled Polyester)
                    100% Recycled Polyamide
                    BCI Compact Supima Cotton
                    Dyeable Spandex''' people can create descriptions from above Original Fabric names with shirt formats, text errors, typing mistakes with percentage incaution.
                    
                    refering above Convert this {text} 
                    to Standard Format
                    """
                )
            ),
            HumanMessagePromptTemplate.from_template("{text}"),
        ]
    )

    llm = ChatOpenAI()
    result = llm(chat_template.format_messages(text=description))
    return result.content


st.title('RM Synthesis')

uploaded_file = st.file_uploader('Upload Excel file', type=['xlsx', 'xls'])

if uploaded_file is not None:
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(uploaded_file)

    st.write('Original DataFrame:')
    st.write(df)

    # Perform calculation on each row
    calculated_values = []
    for index, row in df.iterrows():
        square_value = calculate_square(row['Description'])  # Assuming 'Number' is the column name
        calculated_values.append(square_value)

    # Add a new column with the calculated values to the DataFrame
    df['Reformatted Values'] = calculated_values

    st.write('Updated Excel with reformatted values:')
    st.write(df)

    # Allow users to download the modified Excel file
    output_file = BytesIO()
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()

    # Create a download link for the modified Excel file
    output_file.seek(0)
    st.download_button('Download Updated Excel File', data=output_file, file_name='updated_file.xlsx')

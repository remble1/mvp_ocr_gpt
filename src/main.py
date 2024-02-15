import fitz
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from pandas import DataFrame
import json
import os
from openpyxl import load_workbook


def pdf_to_str(doc: str) -> str:
    text = ""
    with fitz.open(doc) as doc:
        for page in doc:
            text += page.get_text()
    return text


def res_to_df(value: str) -> DataFrame:
    json_data = json.loads(value)
    return DataFrame([json_data])


def save_df(df: DataFrame) -> None:
    excel_filepath = "result_data/medical_data.xlsx"

    if not os.listdir("result_data"):
        df.to_excel(excel_filepath, index=False)
    else:
        wb = load_workbook(excel_filepath)
        sheet = wb.active
        pierwszy_wiersz_jako_lista = df.iloc[0].tolist()
        sheet.append(pierwszy_wiersz_jako_lista)
        wb.save(filename=excel_filepath)


llm = ChatOpenAI(openai_api_key="<openAi_key>", temperature=0)

template = ChatPromptTemplate.from_messages(
    [
        (
            "system",
            """1. you are an OCR system to retrieve text from a document in the form of a large string. Which you will get below. \
        2. you need to find the values for five keywords that are in Polish:\
        'trzpień', 'panewka', 'wkładka', 'głowa', 'PESEL'\
        3. the values are given just after the keyword example:\
        trzpień ***interesting content*** \
        4. you need to capture the names after the keyword and save in JSON form as below:\
            'trzpień': ***contents found***,\
            'panewka': ***contents found***,\
            'wkładka': ***contents found***,\
            'głowa': ***contents found***\
            'PESEL': ***contents found***\
        5. If you can't find information on one of these four items then just type 'None' as in the example: 'głowa': None \
     """,
        ),
        ("human", "{input}"),
    ]
)


chain = template | llm

pdf_dir = "data/"

pdf_folder = [
    os.path.join(pdf_dir, nazwa)
    for nazwa in os.listdir(pdf_dir)
    if os.path.isfile(os.path.join(pdf_dir, nazwa))
]

for pdf in pdf_folder:
    print(pdf)
    res = chain.invoke({"input": pdf_to_str(str(pdf))})
    df = res_to_df(res.content)
    save_df(df)

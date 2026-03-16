import os
import pandas as pd
import datetime as dt
import json
import sqlalchemy
from sqlalchemy import MetaData, Table, Column, Integer, create_engine, String, INTEGER, VARCHAR, DateTime

from dotenv import load_dotenv
from langchain_ollama import OllamaLLM
import pyodbc


load_dotenv()

class ai_ollama():
    def run_ai(df):
        OLLAMA_MODEL_32b: str = os.getenv('OLLAMA_MODEL_Q30b_q8')
        OLLAMA_MODEL_70b: str = os.getenv('OLLAMA_MODEL_llama3.1_70b')
        ollama_url: str = f"http://{os.getenv('OLLAMA_IP_ADDRESS')}:{os.getenv('OLLAMA_PORT')}"
        filename1:str ="12.2.13 Deed of Covenant - 20.07.2006 (RBSII).pdf"
        filename2: str = "12.2.6 HISTORIC - Charge - 17.07.2006.pdf"
        filename3: str = "12.2.1 (A5, B7, C11) - Transfer - .pdf"
        filename4: str = "01-07-2023 - 30-09-2023 Evelyn Investment Report.pdf "
        result_data: list = []
        id: int = 1
        context_header: str = f""" 
        
        You are provided with a filename from windows filesystem.
        Your job is to extract detail from a provided input and return them as a structured JSON object using the following fields.
        
        Values to extract:
        -Date - the date provided in filename 
        
        Extraction Rules:

        Extract the date in the format DD_MM_YYYY.
        The date may appear anywhere in the filename.
        
        If no valid date is found, return an empty string for the "Date" field example: "".

        Output Examples:
        
        Input: {filename1} output: "21_04_2006"
        Input: {filename2} output: "04_05_2021"
        Input: {filename3} output: ''
        
        You must provide a string output at all times nothing additional. 
        
        """

        chat_model: OllamaLLM = OllamaLLM(model=OLLAMA_MODEL_32b, base_url=ollama_url)

        start = dt.datetime.now()
        for i in df:
            prompt = [context_header, i]
            result = chat_model.invoke(prompt)
            #Create new list with results, ensure id is the same as data frame 
            result_data.append({"ID": id, 
                         "Date": result})
            id += 1
        end = dt.datetime.now()
        seconds = end - start
        print(f"took - {seconds.seconds}s")
        return result_data
        # for data in result_data:
        #     print(data['Date'])

# if __name__ == "__main__":
#     main()

#     #OLLAMA_MODEL_Q30b_q8 - 18s
#     #OLLAMA_MODEL_70b - 213s
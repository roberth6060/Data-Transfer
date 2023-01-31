import psycopg2
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy import URL
import os

# Environment variables
db_username = os.environ.get('DB_USERNAME')
db_password = os.environ.get('DB_PASSWORD')
db_host = os.environ.get('DB_HOST')

url_object = URL.create(
    "postgresql+psycopg2",
    username=db_username,
    password=db_password,
    host=db_host,
    database='SpaceNK',
)

# Method - for extracting data from first sheet in excel file and load it to postgres database table


def extract_transform_load(file_path, table_name):

    engine = create_engine(url_object)

    # Load data from the first sheet of the Excel report
    with pd.ExcelFile(file_path) as xls:
        df = pd.read_excel(xls, sheet_name=0, header=5,
                           usecols="C,D,F,G,G,I,J,K,L,M,N,O")
    # Remove (sub)totals
    df = df[df['Store No'] != 'Total']
    # Create the table in the database
    df.to_sql(name=table_name, con=engine, if_exists='replace', index=False)


# Run function with params
extract_transform_load("assets/docs/SpaceNK_2.0.xlsx",
                       "Last Week Report by Store")

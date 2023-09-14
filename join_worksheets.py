import os
import asyncio
import aiofiles
import numpy as np
import pandas as pd
from itertools import product
from time import perf_counter


EXCEL_FILES = {
    'data/PlantData1_2.xlsx': ['First', 'Second'],
    'data/PlantData3.xlsx': ['Third']
}


def get_file() -> list:
    result = []

    for filename, sheetnames in EXCEL_FILES.items():
        result.extend(product([filename], sheetnames))

    return result
        

def test():
    df_first_shift = pd.read_excel('data/PlantData1_2.xlsx', sheet_name='First')
    print(df_first_shift.head())

    df_second_shift = pd.read_excel('data/PlantData1_2.xlsx', sheet_name='Second')
    df_third_shift = pd.read_excel('data/PlantData3.xlsx')

    frames = [df_first_shift, df_second_shift, df_third_shift]

    # append new columns
    total_information_df = pd.concat(frames, axis=1)
    print(total_information_df)

    # append by column
    total_information_df = pd.concat(frames, axis=0)
    print(total_information_df)

    print(total_information_df.columns)

    total_information_df['Time Per Unit'] = total_information_df['Amount Produced'] / total_information_df['Working Time (Minutes)']
    print(total_information_df.columns)
    print(total_information_df['Time Per Unit'])

    hard_drive_df = total_information_df[total_information_df['Product Produced'] == 'Hard Drive']
    print(hard_drive_df)
    print(hard_drive_df.columns)

    hard_drive_df = hard_drive_df.drop(columns=['Product Produced'])
    print(hard_drive_df.columns)
    print(hard_drive_df.groupby(by=['Shift'], dropna=True).mean()['Time Per Unit'])


def automate_report(employee_number: int) -> None:
    t1 = perf_counter()

    frames = []
    for filename, sheetname in get_file():
        frames.append(pd.read_excel(filename, sheet_name=sheetname))

    df = pd.concat(frames, axis=0)
    df['Time Per Unit'] = df['Amount Produced'] / df['Working Time (Minutes)']
    employee_df = df[df['Worker'] == employee_number]

    outpath = os.path.join('./output', str(employee_number) + '-Report.xlsx')
    employee_df.to_excel(outpath, index=False)

    print(perf_counter() - t1)
    return None


async def load_excel(*file):
    filename, sheetname = file
    async with aiofiles.open(filename, mode='rb') as f:
        data = await f.read()
        df = pd.read_excel(data, sheet_name=sheetname, engine='openpyxl')
        return df


async def automate_report_async(employee_number: int) -> None:
    t1 = perf_counter()

    tasks = [load_excel(*file) for file in get_file()]

    loaded_dataframes = await asyncio.gather(*tasks)

    df = pd.concat(loaded_dataframes, axis=0)
    df['Time Per Unit'] = df['Amount Produced'] / df['Working Time (Minutes)']
    employee_df = df[df['Worker'] == employee_number]

    outpath = os.path.join('./output', str(employee_number) + '-Report.xlsx')
    employee_df.to_excel(outpath, index=False)

    print(perf_counter() - t1)
    return None


if __name__ == "__main__":
    employee_number = 110132
    asyncio.run(automate_report_async(employee_number))
    # automate_report(employee_number)

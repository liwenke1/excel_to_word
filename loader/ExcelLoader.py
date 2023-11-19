import os
from typing import List, Dict, Tuple
import pandas as pd


class ExcelLoader:

    def __init__(self, excel_path) -> None:
        self.header, self.data = self._load(excel_path)

    def _load(self, excel_path) -> Tuple[List[str], Dict[str, str]]:
        if not os.path.exists(excel_path):
            raise ValueError("The path of excel file is not exist!")

        _, suffix = os.path.splitext(excel_path)
        if suffix == ".xlsx":
            df = pd.read_excel(excel_path)
        elif suffix == ".csv":
            df = pd.read_csv(excel_path)
        else:
            raise ValueError(
                "The file format is not excel and suffix not in [`xlsx`, `csv`]!"
            )

        header = list(df.columns)
        data = []
        for _, row in df.iterrows():
            column = {}
            for title in header:
                column[title] = row[title] if not pd.isnull(row[title]) else ""
            data.append(column)

        return (header, data)


if __name__ == "__main__":
    path = "/Users/liwenke/Documents/school/项目/CVE和CNVD收集信息/CNVD和CVE收集表.xlsx"
    excel_loader = ExcelLoader(excel_path=path)
    print(excel_loader.header)
    print(excel_loader.data[0])

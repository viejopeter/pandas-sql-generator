import pandas as pd

class AnalyzerFile(object):

    def __init__(self,filepath,start,end):
         self.filepath = filepath
         self.start = start
         self.end = end

    def read_data(self):
         try:
              df = pd.read_excel(self.filepath, engine='openpyxl')
              return df
         except Exception as e:
              print(f"Error reading file: {e}")
              return None

    def filter_data(self):
         df = self.read_data()
         if df is not None:
            df_filtered = df.iloc[:, self.start:self.end].copy()
            df_filtered.loc[:,'Id'] = range(1, len(df_filtered)+1)
            df_filtered = df_filtered[['Id'] + [col for col in df_filtered.columns if col != 'Id']]
            return df_filtered
         else:
             print("No data found")
             return None

    def buildQuery(self):
         try:
             df = self.filter_data()
             string_columns = ", ".join(df.columns)
             query = []
             for index, row in df.iterrows():
                 values = ""
                 for col in df.columns:
                     values += f'"{row[col]}",'
                 values = values.rstrip(",")
                 query.append(f"INSERT INTO ({string_columns}) VALUES ({values});\n")

             with open("query.txt", "w") as query_file:
                 for line in query:
                     query_file.writelines(line)

         except Exception as e:
             print(f"Error building query: {e}")


data = AnalyzerFile("db.xlsx",0,3)
data.buildQuery()


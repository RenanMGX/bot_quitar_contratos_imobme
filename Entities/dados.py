from pandas import read_excel, DataFrame, Series
import xlwings as xw
from shutil import copy2
from os import unlink


class Dados:
    
    @staticmethod
    def carregar_arquivo(path:str) -> DataFrame:
        file_path = path
        if (not file_path.endswith('.xlsx')) and (not file_path.endswith('.xlsm')):
            raise Exception("apenas arquivos excel é permitido")
        
        final_df:DataFrame = DataFrame()
        df:DataFrame
        for num in range(5):
            try:
                df = read_excel(file_path, sheet_name=num)
            
            except TypeError:
                file_path_temp = file_path[0:-5] + "_temp_" + file_path[-5:]
                copy2(file_path, file_path_temp)
                
                app = xw.App(visible=False)
                with app.books.open(file_path_temp) as wb:
                    wb.sheets[0].delete()
                    wb.save(file_path_temp)
                    
                for apps in xw.apps:
                    for open_app in apps.books:
                        if (open_app.name in file_path_temp) or (open_app.name == "Pasta1"):
                            open_app.close()
                
                df = read_excel(file_path_temp, sheet_name=num)
                unlink(file_path_temp)
                
            try:
                df[['Empreendimento', 'Bloco', 'Unidade']]
                final_df = df[['Empreendimento', 'Bloco', 'Unidade']]
                break
            except KeyError:
                continue
        
        if not final_df.empty:
            return final_df
        else:
            raise FileNotFoundError("não foi possivel encontrar os dados nesta planilha")
        

if __name__ == "__main__":
    bot = Dados()
    
    from tkinter.filedialog import askopenfilename
    
    print(Dados.carregar_arquivo(askopenfilename()))
    
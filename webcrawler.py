import requests
import xlwings as xw
from bs4 import BeautifulSoup

class WebCrawler():

    def __init__(self, url: str):
        self.webpage = requests.get(url)
        self.html = self.webpage.text
    
    def parse_table(self):
        self.soup_obj = BeautifulSoup(self.html, features="html.parser")
        table_left, table_right = self.soup_obj.find_all("table")
        self.table_headings = []

        table_left_hd = table_left.thead.find_all("tr")[1:3]
        title = table_left_hd[0].find_all("td")
        for count, subt in enumerate(table_left_hd[1].find_all("td")):
            if count >= 2:
                self.table_headings.append('{} {}'.format(title[count-1].text.strip(), subt.text.strip()).strip())
            else:
                self.table_headings.append('{} {}'.format(title[count].text.strip(), subt.text.strip()).strip())
        
        table_right_hd = table_right.thead.find_all("tr")[1:3]
        title = table_right_hd[0].find_all("td")
        for count, subt in enumerate(table_right_hd[1].find_all("td")):
                self.table_headings.append('{} {}'.format(title[count].text.strip(), subt.text.strip()).strip())
        
        self.table_data = []
        table_left_data = table_left.tbody.find_all("tr")
        table_right_data = table_right.tbody.find_all("tr")
        for line_num in range(len(table_left_data)):
            line = []
            for element in table_left_data[line_num].find_all("td"):
                line.append(element.text.strip())
            for element in table_right_data[line_num].find_all("td"):
                line.append(element.text.strip())
            self.table_data.append(line)
        
        self.table = [self.table_headings] + self.table_data

        return self.table
    
def grava_tabela(tabela: list):

    workbook = xw.Book()
    worksheet = workbook.sheets[0]

    worksheet.range('B2').value = tabela

    workbook.save('./tabela.xlsx')
    workbook.close()



if __name__ == "__main__":
    
    crwl = WebCrawler("https://especial.valor.com.br/valor1000/2020/ranking1000maiores?colunas=id_coluna_1-id_coluna_2-id_coluna_3-id_coluna_4-id_coluna_5-id_coluna_6-id_coluna_7-id_coluna_8-id_coluna_9-id_coluna_10-id_coluna_11-id_coluna_13-id_coluna_14-id_coluna_15-id_coluna_16-id_coluna_17-id_coluna_18-id_coluna_19-id_coluna_20-id_coluna_21-")
    tabela = crwl.parse_table()
    grava_tabela(tabela)
import xlsxwriter

class ExportExcel:

    def __init__(self, cabecalho, dados, nome_arquivo):
        self.__cabecalho = cabecalho
        self.__dados = dados
        self.__nome_arquivo = nome_arquivo

    def __inicia_workbook(self, nome):
        """Faz as configurações iniciais do workbook que será usado"""

        workbook = xlsxwriter.Workbook(nome, {'constant_memory':True})
        return workbook

    def __cria_planilha_tabelas(self, dados, cabecalho, worksheet):
        """Panilha que irá conter as tabelas usadas para o gráfico"""

        coordenadas_grafico = []
        coordenadas = []
        row = 0

        for lista in dados:
            worksheet.write_row(row, 0, cabecalho)
            coordenadas.append(row)
            row += 1

            for linha in lista:
                worksheet.write_row(row, 0, linha)
                coordenadas.append(row)
                row += 1                
            coordenadas_grafico.append(coordenadas)
            coordenadas = []


        return coordenadas_grafico

    def __cria_planilha_grafico(self, coordenadas_x, worksheet_grafico, workbook):
        """Planilha que almazenará os gráficos das tabelas contidas na outra planilha"""

        row = 0
        col = 0

        for x in coordenadas_x:
            chart = workbook.add_chart({'type': 'line'})

            # Última linha do eixo x de coordenadas
            ul = len(x) - 1
            chart.add_series({
                'categories': ['Tabelas', x[1], 0, x[ul], 0],
                'values': ['Tabelas', x[1], 1, x[ul], 1],
                'line': {'color': 'red'},
            })

            chart.add_series({
                'categories': ['Tabelas', x[1], 0, x[ul], 0],
                'values': ['Tabelas', x[1], 2, x[ul], 2],
                'line': {'color': 'blue'},
            })

            worksheet_grafico.insert_chart(row, col, chart)
            row += 15

    def gera_excel(self):
        workbook = self.__inicia_workbook(self.__nome_arquivo)
        worksheet_grafico = workbook.add_worksheet("Graficos")
        worksheet_tabela = workbook.add_worksheet("Tabelas")
        tabelas = self.__cria_planilha_tabelas(self.__dados, self.__cabecalho, worksheet_tabela)
        grafico = self.__cria_planilha_grafico(tabelas, worksheet_grafico, workbook)
        return workbook.close()


cabecalho = ['data', 'max', 'min']
dados = [
    [('01/05/2013', 100, 10), ('09/02/2013', 200, 20), ('07/05/2020', 150, 15), ('08/09/2020', 80, 8)],
    [('02/06/2014', 110, 11), ('06/03/2013', 220, 22), ('03/06/2018', 300, 20), ('15/12/2021', 180, 15)],
    [('03/07/2015', 185, 19), ('07/04/2013', 167, 35), ('06/11/2019', 273, 50), ('17/11/2017', 280, 20)]
]
nome = 'teste.xlsx'

obj = ExportExcel(cabecalho, dados, nome)

obj.gera_excel()
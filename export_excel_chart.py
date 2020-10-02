import xlsxwriter

class ExportExcel:

    def __init__(self, cabecalho, dados, nome_arquivo):
        self.__cabecalho = cabecalho
        self.__dados = dados
        self.__nome_arquivo = nome_arquivo

    def __formatacoes(self, workbook, cabecalho_planilha=False, cabecalho_tabela=False, conteudo_tabela=False):
    
        if cabecalho_planilha:

            formatacao = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color':'#FF8C00',
                                              'align':'center', 'font_size':25})
            return formatacao

        elif cabecalho_tabela:

            formatacao = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color':'#FFA500',
                                              'align':'center', 'font_size':12})
            return formatacao

        else:

            formatacao = workbook.add_format({'align':'center', 'font_size':10})
            return formatacao



    def __inicia_workbook(self, nome):
        """Faz as configurações iniciais do workbook que será usado"""

        workbook = xlsxwriter.Workbook(nome, {'constant_memory':True})
        return workbook

    def __cria_planilha_tabelas(self, dados, cabecalho, worksheet, workbook):
        """Panilha que irá conter as tabelas usadas para o gráfico"""

        # Adiciona cabeçalho na planilha
        formatacao_cabecalho = self.__formatacoes(workbook, cabecalho_planilha=True)
        worksheet.merge_range('A1:E2', 'Relatório Excel', formatacao_cabecalho)

        coordenadas_grafico = []
        coordenadas = []
        formatacao_cabecalho_tabelas = self.__formatacoes(workbook, cabecalho_tabela=True)
        formatacao_conteudo_tabelas = self.__formatacoes(workbook, conteudo_tabela=True)
        row = 4


        for lista in dados:
            for linha in lista:
                if lista[0] == linha:
                    worksheet.merge_range(row, 0, row, 2, linha[0], formatacao_cabecalho_tabelas)
                    coordenadas.append(row)
                    row += 1

                    worksheet.write_row(row, 0, cabecalho, formatacao_cabecalho_tabelas)
                    coordenadas.append(row)
                    row += 1  
                else:
                    worksheet.write_row(row, 0, linha, formatacao_conteudo_tabelas)
                    coordenadas.append(row)
                    row += 1
            coordenadas_grafico.append(coordenadas)
            coordenadas = []
            row += 2

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
        tabelas = self.__cria_planilha_tabelas(self.__dados, self.__cabecalho, worksheet_tabela, workbook)
        grafico = self.__cria_planilha_grafico(tabelas, worksheet_grafico, workbook)
        return workbook.close()


cabecalho = ['data', 'max', 'min']
dados = [
    (('tabela1',), ('01/05/2013', 100, 10), ('09/02/2013', 200, 20), ('07/05/2020', 150, 15), ('08/09/2020', 80, 8)),
    (('tabela2',), ('02/06/2014', 110, 11), ('06/03/2013', 220, 22), ('03/06/2018', 300, 20), ('15/12/2021', 180, 15)),
    (('tabela3',), ('03/07/2015', 185, 19), ('07/04/2013', 167, 35), ('06/11/2019', 273, 50), ('17/11/2017', 280, 20))
]
nome = 'teste.xlsx'

obj = ExportExcel(cabecalho, dados, nome)

obj.gera_excel()
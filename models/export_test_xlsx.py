# -*- coding: utf-8 -*-
from odoo import models

class ReporteXlsx(models.AbstractModel):
    _name='report.tvp_hr_survey_zavic.tvp_hr_survey_zavic'
    _inherit = 'report.report_xlsx.abstract'


    def generate_xlsx_report(self, workbook, data, survey):
        for obj in survey:
            report_name=obj.email
            # One sheet by partner
            sheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            sheet.write(0, 0,report_name, bold)

            sheet.write(2, 0, u"Pregunta", bold)
            sheet.write(2, 1, u"Opciones", bold)
            sheet.write(2, 2, u"Valor Asignado", bold)
            line = 3

            for mov in obj.user_input_line:
                sheet.write(line, 0, mov.question_id.question)
                sheet.write(line, 1, mov.value_suggested_row.value)
                sheet.write(line, 2, mov.value_suggested.quizz_mark)
                line += 1

            for row in range(1, 88):
                sheet.set_row(row, options= {'hidden': True})

            # Formatos que aplicare a las celdas

            cell_format = workbook.add_format({'font_size': 24,
                                               'color': 'white',
                                               'align': 'center',
                                               'valign': 'vcenter',
                                               'bold': 1,
                                               'fg_color': 'blue',
                                               'border': 2})

            cell_format2 = workbook.add_format({'font_size': 24,
                                                'color': 'white',
                                                'align': 'center',
                                                'valign': 'vcenter',
                                                'bold': 1,
                                                'fg_color': 'green',
                                                'border': 2})

            # Combinacion de celdas para valores e intereses
            sheet.merge_range('A90:I90', "VALORES",cell_format)
            sheet.merge_range('J90:R90', "INTERESES",cell_format2)

            #
            formato3 = workbook.add_format({'font_size': 12,
                                            'align': 'center',
                                            'valign': 'vcenter',
                                            'bold': 1,
                                            'fg_color': 'gray',
                                            'border': 1})

            sheet.write('A91', "ITEM", formato3)
            sheet.write('A92', "3", formato3)
            sheet.write('A93', "4", formato3)
            sheet.write('A94', "6", formato3)
            sheet.write('A95', "8", formato3)
            sheet.write('A96', "9", formato3)
            sheet.write('A97', "12", formato3)
            sheet.write('A98', "13", formato3)
            sheet.write('A99', "15", formato3)
            sheet.write('A100', "17", formato3)
            sheet.write('A101', "19", formato3)
            sheet.write('A102', "SUMAS", formato3)

            sheet.write('J91', "ITEM", formato3)
            sheet.write('J92', "1", formato3)
            sheet.write('J93', "2", formato3)
            sheet.write('J94', "5", formato3)
            sheet.write('J95', "7", formato3)
            sheet.write('J96', "10", formato3)
            sheet.write('J97', "11", formato3)
            sheet.write('J98', "14", formato3)
            sheet.write('J99', "16", formato3)
            sheet.write('J100', "18", formato3)
            sheet.write('J101', "20", formato3)
            sheet.write('J102', "SUMAS", formato3)

            # Subtitulos
            subtitulos = workbook.add_format({'font_size': 12,
                                              'align': 'center',
                                              'valign': 'vcenter',
                                              'bold': 1,
                                              'fg_color': 'gray',
                                              'border': 1})

            sheet.merge_range('B91:C91', "MORAL", subtitulos)
            sheet.merge_range('D91:E91', "LEGALIDAD", subtitulos)
            sheet.merge_range('F91:G91', "INDIFERENCIA", subtitulos)
            sheet.merge_range('H91:I91', "CORRUPTO", subtitulos)
            sheet.merge_range('K91:L91', "ECONOMICO", subtitulos)
            sheet.merge_range('M91:N91', "POLITICO", subtitulos)
            sheet.merge_range('O91:P91', "SOCIAL", subtitulos)
            sheet.merge_range('Q91:R91', "RELIGIOSO", subtitulos)

            # OPCIONES
            opciones = workbook.add_format({'font_size': 12,
                                            'align': 'center',
                                            'valign': 'vcenter',
                                            'border': 1})

            sheet.write('B92', "A)", opciones)
            sheet.write('B93', "D)", opciones)
            sheet.write('B94', "A)", opciones)
            sheet.write('B95', "B)", opciones)
            sheet.write('B96', "B)", opciones)
            sheet.write('B97', "B)", opciones)
            sheet.write('B98', "A)", opciones)
            sheet.write('B99', "D)", opciones)
            sheet.write('B100', "D)", opciones)
            sheet.write('B101', "A)", opciones)

            sheet.write('D92', "B)", opciones)
            sheet.write('D93', "C)", opciones)
            sheet.write('D94', "B)", opciones)
            sheet.write('D95', "A)", opciones)
            sheet.write('D96', "A)", opciones)
            sheet.write('D97', "D)", opciones)
            sheet.write('D98', "B)", opciones)
            sheet.write('D99', "C)", opciones)
            sheet.write('D100', "B)", opciones)
            sheet.write('D101', "D)", opciones)

            sheet.write('F92', "C)", opciones)
            sheet.write('F93', "A)", opciones)
            sheet.write('F94', "D)", opciones)
            sheet.write('F95', "C)", opciones)
            sheet.write('F96', "D)", opciones)
            sheet.write('F97', "A)", opciones)
            sheet.write('F98', "C)", opciones)
            sheet.write('F99', "B)", opciones)
            sheet.write('F100', "A)", opciones)
            sheet.write('F101', "B)", opciones)

            sheet.write('H92', "D)", opciones)
            sheet.write('H93', "B)", opciones)
            sheet.write('H94', "C)", opciones)
            sheet.write('H95', "D)", opciones)
            sheet.write('H96', "C)", opciones)
            sheet.write('H97', "C)", opciones)
            sheet.write('H98', "D)", opciones)
            sheet.write('H99', "A)", opciones)
            sheet.write('H100', "C)", opciones)
            sheet.write('H101', "C)", opciones)

            sheet.write('K92', "C)", opciones)
            sheet.write('K93', "C)", opciones)
            sheet.write('K94', "D)", opciones)
            sheet.write('K95', "B)", opciones)
            sheet.write('K96', "A)", opciones)
            sheet.write('K97', "A)", opciones)
            sheet.write('K98', "A)", opciones)
            sheet.write('K99', "A)", opciones)
            sheet.write('K100', "B)", opciones)
            sheet.write('K101', "A)", opciones)

            sheet.write('M92', "B)", opciones)
            sheet.write('M93', "D)", opciones)
            sheet.write('M94', "B)", opciones)
            sheet.write('M95', "C)", opciones)
            sheet.write('M96', "B)", opciones)
            sheet.write('M97', "D)", opciones)
            sheet.write('M98', "D)", opciones)
            sheet.write('M99', "B)", opciones)
            sheet.write('M100', "C)", opciones)
            sheet.write('M101', "B)", opciones)

            sheet.write('O92', "A)", opciones)
            sheet.write('O93', "B)", opciones)
            sheet.write('O94', "A)", opciones)
            sheet.write('O95', "A)", opciones)
            sheet.write('O96', "D)", opciones)
            sheet.write('O97', "B)", opciones)
            sheet.write('O98', "C)", opciones)
            sheet.write('O99', "D)", opciones)
            sheet.write('O100', "D)", opciones)
            sheet.write('O101', "D)", opciones)

            sheet.write('Q92', "D)", opciones)
            sheet.write('Q93', "A)", opciones)
            sheet.write('Q94', "C)", opciones)
            sheet.write('Q95', "D)", opciones)
            sheet.write('Q96', "C)", opciones)
            sheet.write('Q97', "C)", opciones)
            sheet.write('Q98', "B)", opciones)
            sheet.write('Q99', "C)", opciones)
            sheet.write('Q100', "A)", opciones)
            sheet.write('Q101', "C)", opciones)

            # PUNTUACION DEL USUARIO

            sheet.write('C92', "=C12", opciones)
            sheet.write('C93', "=C19", opciones)
            sheet.write('C94', "=C24", opciones)
            sheet.write('C95', "=C33", opciones)
            sheet.write('C96', "=C37", opciones)
            sheet.write('C97', "=C49", opciones)
            sheet.write('C98', "=C52", opciones)
            sheet.write('C99', "=C63", opciones)
            sheet.write('C100', "=C71", opciones)
            sheet.write('C101', "=C76", opciones)
            sheet.write('C102', '=SUM(C92:C101)', subtitulos)

            sheet.write('E92', "=C13", opciones)
            sheet.write('E93', "=C18", opciones)
            sheet.write('E94', "=C25", opciones)
            sheet.write('E95', "=C32", opciones)
            sheet.write('E96', "=C36", opciones)
            sheet.write('E97', "=C51", opciones)
            sheet.write('E98', "=C53", opciones)
            sheet.write('E99', "=C62", opciones)
            sheet.write('E100', "=C69", opciones)
            sheet.write('E101', "=C79", opciones)
            sheet.write('E102', '=SUM(E92:E101)', subtitulos)

            sheet.write('G92', "=C14", opciones)
            sheet.write('G93', "=C16", opciones)
            sheet.write('G94', "=C27", opciones)
            sheet.write('G95', "=C34", opciones)
            sheet.write('G96', "=C39", opciones)
            sheet.write('G97', "=C48", opciones)
            sheet.write('G98', "=C54", opciones)
            sheet.write('G99', "=C61", opciones)
            sheet.write('G100', "=C68", opciones)
            sheet.write('G101', "=C77", opciones)
            sheet.write('G102', '=SUM(G92:G101)', subtitulos)

            sheet.write('I92', "=C15", opciones)
            sheet.write('I93', "=C17", opciones)
            sheet.write('I94', "=C26", opciones)
            sheet.write('I95', "=C35", opciones)
            sheet.write('I96', "=C38", opciones)
            sheet.write('I97', "=C50", opciones)
            sheet.write('I98', "=C55", opciones)
            sheet.write('I99', "=C60", opciones)
            sheet.write('I100', "=C70", opciones)
            sheet.write('I101', "=C78", opciones)
            sheet.write('I102', '=SUM(I92:I101)', subtitulos)

            sheet.write('L92', "=C6", opciones)
            sheet.write('L93', "=C10", opciones)
            sheet.write('L94', "=C23", opciones)
            sheet.write('L95', "=C29", opciones)
            sheet.write('L96', "=C40", opciones)
            sheet.write('L97', "=C44", opciones)
            sheet.write('L98', "=C56", opciones)
            sheet.write('L99', "=C64", opciones)
            sheet.write('L100', "=C73", opciones)
            sheet.write('L101', "=C70", opciones)
            sheet.write('L102', '=SUM(L92:L101)', subtitulos)

            sheet.write('N92', "=C5", opciones)
            sheet.write('N93', "=C11", opciones)
            sheet.write('N94', "=C21", opciones)
            sheet.write('N95', "=C30", opciones)
            sheet.write('N96', "=C41", opciones)
            sheet.write('N97', "=C47", opciones)
            sheet.write('N98', "=C59", opciones)
            sheet.write('N99', "=C65", opciones)
            sheet.write('N100', "=C74", opciones)
            sheet.write('N101', "=C81", opciones)
            sheet.write('N102', '=SUM(N92:N101)', subtitulos)

            sheet.write('P92', "=C4", opciones)
            sheet.write('P93', "=C9", opciones)
            sheet.write('P94', "=C20", opciones)
            sheet.write('P95', "=C28", opciones)
            sheet.write('P96', "=C43", opciones)
            sheet.write('P97', "=C45", opciones)
            sheet.write('P98', "=C58", opciones)
            sheet.write('P99', "=C67", opciones)
            sheet.write('P100', "=C75", opciones)
            sheet.write('P101', "=C83", opciones)
            sheet.write('P102', '=SUM(P92:P101)', subtitulos)

            sheet.write('R92', "=C7", opciones)
            sheet.write('R93', "=C8", opciones)
            sheet.write('R94', "=C22", opciones)
            sheet.write('R95', "=C31", opciones)
            sheet.write('R96', "=C42", opciones)
            sheet.write('R97', "=C46", opciones)
            sheet.write('R98', "=C57", opciones)
            sheet.write('R99', "=C66", opciones)
            sheet.write('R100', "=C72", opciones)
            sheet.write('R101', "=C82", opciones)
            sheet.write('R102', '=SUM(R92:R101)', subtitulos)

            grafico1 = workbook.add_chart({'type': 'area'})
            grafico1.add_series({
                'name': '=Sheet1!$B$1',
                'categories': '=(Sheet1!$B$91,Sheet1!$D$91,Sheet1!$F$91,Sheet1!$H$91,Sheet1!$K$91,Sheet1!$M$91,Sheet1!$O$91,Sheet1!$Q$91)',
                'values': '=(Sheet1!$C$102,Sheet1!$E$102,Sheet1!$G$102,Sheet1!$I$102,Sheet1!$L$102,Sheet1!$N$102,Sheet1!$P$102,Sheet1!$R$102)',
            })

            grafico1.set_style(26)

            sheet.insert_chart('A104', grafico1, {'x_scale': 2.4, 'y_scale': 1.5})


	    #GRAFICO CIRCULAR VALORES

            ftgraficov = workbook.add_format({'font_size': 24,
                                              'color': 'green',
                                              'align': 'center',
                                              'valign': 'vcenter',
                                             })

            sheet.merge_range('B127:C127', "VALORES", ftgraficov)
            sheet.merge_range('B128:C128', "PUNTAJE TOTAL", bold)
            sheet.merge_range('B129:C129', "   -MORAL", bold)
            sheet.merge_range('B130:C130', "   -LEGALIDAD", bold)
            sheet.merge_range('B131:C131', "   -INDIFERENCIA", bold)
            sheet.merge_range('B132:C132', "   -CORRUPTO", bold)

            sheet.write('D128', "=SUM(D129:D132)", bold)
            sheet.write('D129', "=C102", bold)
            sheet.write('D130', "=E102", bold)
            sheet.write('D131', "=G102", bold)
            sheet.write('D132', "=I102", bold)


            grafico2 = workbook.add_chart({'type': 'pie'})
            grafico2.add_series({
                'name': '=Sheet1!$B$1',
                'categories': '=(Sheet1!$B$129:$B$132)',
                'values': '=(Sheet1!$D$129:$D$132)',
                'data_labels': {'percentage': True, 'position': 'center'},
            })

            grafico2.set_style(26)
            sheet.insert_chart('B134', grafico2,)


	    #GRAFICO CIRCULAR INTERESES

            ftgrafico = workbook.add_format({'font_size': 24,
                                             'color': 'red',
                                             'align': 'center',
                                             'valign': 'vcenter',
                                             })

            sheet.merge_range('J127:K127', "INTERESES", ftgrafico)
            sheet.merge_range('J128:K128', "PUNTAJE TOTAL", bold)
            sheet.merge_range('J129:K129', "  -ECONOMICO", bold)
            sheet.merge_range('J130:K130', "   -POLITICO", bold)
            sheet.merge_range('J131:K131', "   -SOCIAL", bold)
            sheet.merge_range('J132:K132', "   -RELIGIOSO", bold)

            sheet.write('L128', "=SUM(L129:L132)", bold)
            sheet.write('L129', "=L102", bold)
            sheet.write('L130', "=N102", bold)
            sheet.write('L131', "=P102", bold)
            sheet.write('L132', "=R102", bold)


            grafico3 = workbook.add_chart({'type': 'pie'})
            grafico3.add_series({
                'name': '=Sheet1!$B$1',
                'categories': '=(Sheet1!$J$129:$J$132)',
                'values': '=(Sheet1!$L$129:$L$132)',
                'data_labels': {'percentage': True, 'position': 'center'},
            })

            grafico3.set_style(26)
            sheet.insert_chart('J134', grafico3,)
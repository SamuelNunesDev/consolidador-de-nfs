from pathlib import Path
from os import makedirs, listdir, rename
from tkinter import Tk, Label, font, Button
from xml.etree.ElementTree import parse
from xlrd import open_workbook_xls
from xlutils.copy import copy

#Criação de função do botão para encerrar o programa.

def fim():
    global lb1, lb2, bt

    lb1['text'] = 'A base de dados foi atualizada com êxito!'
    lb2['text'] = 'Clique no botão "Encerrar" para finalizar.'
    bt['text'] = 'Encerrar'
    bt['command'] = lambda: janela.destroy()

#Função para extrair os dados dos arquivos xml.

def xml_protocol():
    global xml

    #Extraindo dados da planilha.

    for i, k in enumerate(xml):
        root = parse(Path.home() / 'Documents' / 'NOTAS FISCAIS' / k)
        filtro = "*"
        l = list()
        for child in root.iter(filtro):
            l.append(child.text)
        dados = [l[3], l[5], l[8], l[17], l[53], l[67]]
        data = f'{dados[1][8:10]}/{dados[1][5:7]}/{dados[1][0:4]}'
        n_nfs = str(dados[0])
        valor = f"R${dados[2].replace('.', ',')}"
        razao = dados[3]
        discriminacao = dados[4]
        comprador = dados[5]

    # Colando os dados na planilha base xls.

        planilha = open_workbook_xls('CRUZAMENTO NF.xls')
        p_open = copy(planilha)
        w_sheet = p_open.get_sheet(0)
        w_sheet1 = w_sheet.write(i + 1, 0, data)
        w_sheet2 = w_sheet.write(i + 1, 1, n_nfs)
        w_sheet3 = w_sheet.write(i + 1, 2, valor)
        w_sheet4 = w_sheet.write(i + 1, 3, razao)
        w_sheet5 = w_sheet.write(i + 1, 4, discriminacao)
        w_sheet6 = w_sheet.write(i + 1, 5, comprador)
        p_open.save('CRUZAMENTO NF.xls')
        planilha.release_resources()
        del l
        print(f'{i} arquivos extraídos com sucesso!')
    fim()

#Função para listar os arquivos e identificar o formato.


def iniciar():
    global xml, pdf, pasta, NFs

    NFs = listdir(Path.home() / 'Documents' / 'NOTAS FISCAIS')
    print(f'{len(NFs) - 2} arquivo(s) para extração de dados.')
    while len(NFs) != 2:
        if 'xml' in NFs[0][-3:]:
            xml.append(NFs[0])
            NFs.pop(0)
        elif 'pdf' in NFs[0][-3:]:
            pdf.append(NFs[0])
            NFs.pop(0)
    xml_protocol()

#Criação e configuração de interface.

xml = list()
pdf = list()
janela = Tk()
janela.geometry('800x600+200+50')
janela.title('SAM - System Assistant Management')

font_titulo = font.Font(family='Lucida Grande', size=20)
font_texto = font.Font(family='Lucida Grande', size=12)

lb_titulo = Label(janela, font=font_titulo, text='CONSOLIDADOR DE DADOS', height='5')
lb_titulo.pack(side='top')

#Criação e customização dos labels de instruções.

lb1 = Label(janela, font=font_texto, text=f'1 - Mova os arquivos para a pasta "NOTAS FISCAIS" caminho - '
                                          f'{Path.home() / "Documents"}')
lb1.pack(side='top')
lb2 = Label(janela, font=font_texto, text='2 - Assim que os arquivos estiverem na pasta, clique no botão "Iniciar"'
                                          + ' '*16, height='5')
lb2.pack(side='top')

#Criação do botão "Iniciar" e marca d'água.

lb_bt = Label(janela, height='3')
lb_bt.pack(side='top')
bt = Button(janela, text='Iniciar', font=font_texto, command=iniciar, width='15')
bt.pack(side='top')
lb_marca = Label(janela, text='SAM - CDDv1.2 developed by Samuel Nunes')
lb_marca.pack(side='bottom', anchor='e')

# Criação do diretório de origem das NFs.

try:
    pasta = Path.home() / "Documents" / 'NOTAS FISCAIS'
    makedirs(pasta)
except:
    lb = Label(janela, height='2')
    lb.pack(side='top')
    lb_controle1 = Label(janela, text='Diretório OK')
    lb_controle1.pack(side='top')
else:
    lb = Label(janela, height='2')
    lb.pack(side='top')
    lb_controle1 = Label(janela, text='Diretório OK')
    lb_controle1.pack(side='top')

janela.mainloop()

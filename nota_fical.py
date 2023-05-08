import PySimpleGUI as py
from app import *


def pegar_caminho():
    tela = [
        [py.T("Informe onde está o banco de dados")],
        [py.InputText("", key='url'), py.FileBrowse()],
        [py.T("Informe a senha do Banco de Dados")],
        [py.InputText("", key='senha', password_char="*")],
        [py.B("Enviar")],
    ]
    janela = py.Window('Tela Incial', tela, element_justification='center')
    while True:
        button, campo_de_busca = janela.read()

        if button in (py.WIN_CLOSED, 'Cancel'):
            janela.close()
            return "", ""
        else:
            janela.close()
            return campo_de_busca['url'], campo_de_busca['senha']


def programa(nota):
    py.theme_text_color('White')
    lista_de_clientes = nota.lista_de_clientes()
    lista_de_letras = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                       'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

    titulos = ['Produto', 'Unidade', 'VlrUnitario', 'Quantidade', 'VlrTotal']

    tela = [
        [py.Table(values=[], headings=titulos, max_col_width=35, auto_size_columns=True, display_row_numbers=True,
                  justification='center', num_rows=10, key="-tabela", row_height=35, background_color="gray")],
        [py.B("Limpar", size=(15, 0))],
        [py.T("Quantidade de produtos:"), py.T("0", key="-quantidade")],
        [py.T("Valor total dos produtos:"), py.T("0", key="-total")],
        [py.T("", key='erro_nota')],
        [py.T("Nome do cliente"),
         py.Combo(values=lista_de_clientes, default_value='', key="-Cliente"),
         py.Combo(lista_de_letras, default_value=lista_de_letras[0], key='filtro'),
         py.B("Filtrar", size=(15, 0), button_color='red')],
        [py.B("Buscar Cliente", size=(15, 0)), py.B("Exportar Nota", size=(15, 0))],
    ]

    janela = py.Window('Tela Incial', tela, element_justification='center')
    while True:
        button, campo_de_busca = janela.read()

        exportar = button == "Exportar Nota"
        buscar = button == "Buscar Cliente"
        limpar = button == "Limpar"
        filtro = button == "Filtrar"

        if button in (py.WIN_CLOSED, 'Cancel'):
            nota.fechar_conexao()
            janela.close()
            break

        elif exportar:
            botao_exportar_buscar(campo_de_busca,nota, janela, py, lista_de_clientes,exportar)

        elif buscar:
            botao_exportar_buscar(campo_de_busca, nota, janela, py, lista_de_clientes, exportar)

        elif limpar:
            botao_limpar(janela)

        elif filtro:
            botao_filtrar(campo_de_busca, janela, nota, py)


if __name__ == "__main__":
    criar_tabelas()
    caminho_senha = verifica_se_tem_banco_access()

    if caminho_senha:
        nota_fical = Nota_fiscal(caminho_senha[0][0], caminho_senha[0][1])
        programa(nota_fical)

    else:
        caminho, senha = pegar_caminho()
        if caminho:
            criar_caminho(caminho, senha)
            nota_fical = Nota_fiscal(caminho, senha)
            programa(nota_fical)
        else:
            py.PopupOK("O Programa foi fechado por não passar o caminho do banco de dados")

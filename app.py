import pyodbc
import sqlite3
import xlsxwriter
import os
from datetime import datetime


class Banco_de_dados_sql:
    def __init__(self):
        self._banco = sqlite3.connect('banco_dados.db')
        self._cursor = self._banco.cursor()

    def registra_caminho(self, caminho, senha):
        self._cursor.execute('INSERT INTO  access(caminho, senha)'
                             'VALUES(?,?)', (caminho, senha,))
        self._banco.commit()

    def obter_caminho_access(self):
        self._cursor.execute('SELECT * FROM access')
        caminho_senha = self._cursor.fetchall()
        return caminho_senha

    def criar_tabela(self):
        self._cursor.execute('CREATE TABLE IF NOT EXISTS Produtos( produto text, Unidade text, VlrUnitario text, '
                             'quantidade float ,VlrTotal float )')
        self._cursor.execute('CREATE TABLE IF NOT EXISTS access( caminho text, senha text)')

    def inserir_dados(self, lista_produtos):
        self.deletar()
        for produto in lista_produtos:
            verifica = self.verifica_dados(produto)
            if verifica:
                novo_produto = self.criar_novo_produto(produto)
                nome = novo_produto[0]
                quant = novo_produto[3]
                valor_unit = novo_produto[2]
                valor_final = novo_produto[4]
                self._cursor.execute('UPDATE Produtos  SET  quantidade = ?,  VlrUnitario = ? , VlrTotal = ? '
                                     'WHERE produto = ?', (quant, valor_unit, valor_final, nome,))
                self._banco.commit()
            else:
                nome = produto[0]
                und = produto[1]
                quant = produto[2]
                valor_unit = produto[3]
                valor_final = produto[4]
                self._cursor.execute('INSERT INTO Produtos (produto, Unidade, quantidade, VlrUnitario, VlrTotal)'
                                     'VALUES(?,?,?,?,?)', (nome, und, quant, valor_unit, valor_final,))
                self._banco.commit()

    def verifica_dados(self, produto):
        nome = produto[0]
        self._cursor.execute('SELECT produto FROM Produtos WHERE produto = ?', (nome,))
        verifica = self._cursor.fetchall()
        if verifica:
            return True
        else:
            return False

    def criar_novo_produto(self, produto):
        self._cursor.execute('SELECT Quantidade,VlrTotal FROM Produtos WHERE produto = ?', (produto[0],))
        produto_antigo = self._cursor.fetchall()
        nova_quantidade = produto_antigo[0][0] + produto[2]
        nova_quantidade = round(nova_quantidade, 3)
        novo_valor = produto_antigo[0][1] + produto[4]
        novo_valor = round(novo_valor, 2)
        valor_unit = novo_valor / nova_quantidade
        valor_unit = round(valor_unit, 2)
        novo_produto = (produto[0], produto[1], valor_unit, nova_quantidade, novo_valor)
        return novo_produto

    def verifica_dados_existente(self):
        self._cursor.execute('SELECT * FROM Produtos')
        dados = self._cursor.fetchall()
        if dados:
            return True
        else:
            return False

    def exportar_dados(self):
        self._cursor.execute('SELECT * FROM Produtos')
        dados = self._cursor.fetchall()
        return dados

    def deletar(self):
        verifica_dados = self.verifica_dados_existente()
        if verifica_dados:
            self._cursor.execute('DELETE FROM Produtos')
            self._banco.commit()

    def fechar_conexao(self):
        self._cursor.close()
        self._banco.close()


class Banco_de_dados_access:
    """Acessa um Banco de Dados Access com sua senha e fecha sua conexão"""

    def __init__(self, caminho, senha):
        caminho = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' \
                  f'DBQ={caminho};PWD={senha};'
        self._conectar = pyodbc.connect(caminho)
        self._cursor = self._conectar.cursor()

    def fechar_conexao(self):
        try:
            self._cursor.close()
            self._conectar.close()
        except pyodbc.Error:
            return "Não foi possivel fechar conexao"


class Busca_de_dados_access(Banco_de_dados_access):

    def __init__(self, caminho, senha):
        super().__init__(caminho, senha)

    def lista_de_clientes(self):
        lista_de_clientes = []
        self._cursor.execute('select Nome from Clientes ORDER BY Nome ASC')
        for cliente in self._cursor.fetchall():
            lista_de_clientes.append(cliente[0])
        return lista_de_clientes

    def busca_por_cliente(self, nome):
        self._cursor.execute(f"select Nome from Clientes WHERE Nome=? ", (nome,))
        busca_pelo_nome = self._cursor.fetchall()
        return busca_pelo_nome

    def filtro_de_clientes(self, letra):
        self._cursor.execute(f"select Nome from Clientes WHERE Nome LIKE '{letra}%'")
        return [cliente[0] for cliente in self._cursor.fetchall()]

    def verifica_se_existe(self, nome):
        self._cursor.execute(f"select Nome from Clientes WHERE Nome =?", (nome,))
        busca_pelo_nome = self._cursor.fetchall()
        if busca_pelo_nome:
            return True
        else:
            return False

    def buscar_notas_nao_pagas(self, nome):
        notas_nao_pagas = []
        existe_clinte = self.verifica_se_existe(nome)

        if existe_clinte:
            self._cursor.execute('select CodigoPagto,DataPagto,ControleInterno from VendasPrazo WHERE Cliente = ?',
                                 (nome,))
            notas = self._cursor.fetchall()
            for codigo, data, numero, in notas:
                if codigo == '0' and data == '0':
                    notas_nao_pagas.append(str(numero))
            return notas_nao_pagas
        else:
            return notas_nao_pagas

    @staticmethod
    def verifica_se_tem_item(nome_do_produto):
        if nome_do_produto == 'ITEM':
            return True
        return False

    def item_das_notas(self, nome):
        lista_de_produtos = []
        lista_de_numeros_de_notas = []
        lista_codigos_das_notas = self.buscar_notas_nao_pagas(nome)
        for nota in lista_codigos_das_notas:
            self._cursor.execute(
                'select Produto,Unidade,Quantidade,VlrTotal from VendasProdutos WHERE ControleInterno = ?',
                (nota,))
            lista_de_produtos.append(self._cursor.fetchall())
            for notas in lista_de_produtos:
                for produto, _, _, _ in notas:
                    if self.verifica_se_tem_item(produto):
                        is_tem_este_numero = nota in lista_de_numeros_de_notas
                        if not is_tem_este_numero:
                            lista_de_numeros_de_notas.append(nota)

        return lista_de_produtos, lista_de_numeros_de_notas

    def lista_de_compra(self, nome):
        lista_de_compras = []
        item_de_compras, lista_de_numeros_de_notas = self.item_das_notas(nome)
        for notas in item_de_compras:
            for produtos in notas:
                if not self.verifica_se_tem_item(produtos[0]):
                    nome = produtos[0]
                    unt = produtos[1]
                    quant = float(produtos[2].replace(',', '.'))
                    valor_total = float(produtos[3].replace(',', '.'))
                    valor_uni = round(valor_total / quant, 2)
                    produto = (nome, unt, quant, valor_uni, valor_total)
                    lista_de_compras.append(produto)
        return lista_de_compras, lista_de_numeros_de_notas


class Nota_fiscal:

    def __init__(self, caminho, senha):
        self.access = Busca_de_dados_access(caminho, senha)
        self.sql = Banco_de_dados_sql()

    def lista_de_clientes(self):
        return self.access.lista_de_clientes()

    def filtro(self, letra):
        return self.access.filtro_de_clientes(letra)

    def criar_nota_fiscal(self, nome):
        lista_de_produtos, lista_de_notas = self.access.lista_de_compra(nome)
        self.sql.inserir_dados(lista_de_produtos)
        return lista_de_notas

    def exportar_nota(self, nome):
        lista_de_produtos = self.sql.exportar_dados()
        diretorio = self.verifica_se_a_pasta(nome)
        now = datetime.now()
        data = now.strftime('%d-%m-%Y')
        workbook = xlsxwriter.Workbook(rf'{diretorio}\{data}.xlsx')
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})
        bold.set_border()
        border = workbook.add_format()
        border.set_border()
        worksheet.write('A1', 'Produto', bold)
        worksheet.write('B1', 'Unidade', bold)
        worksheet.write('C1', 'VlrUnitario', bold)
        worksheet.write('D1', 'Quantidade', bold)
        worksheet.write('E1', 'VlrTotal', bold)
        numero = 2
        total = 0
        for produto in lista_de_produtos:
            tem_item = self.access.verifica_se_tem_item(produto[0])
            if not tem_item:
                worksheet.write(f'A{numero}', produto[0], border)
                worksheet.write(f'B{numero}', produto[1], border)
                worksheet.write(f'C{numero}', f'R$ {produto[2]}', border)
                worksheet.write(f'D{numero}', produto[3], border)
                worksheet.write(f'E{numero}', f'R$ {produto[4]}', border)
                total += produto[4]
                numero += 1
        worksheet.write(f'D{numero}', "Total", bold)
        worksheet.write(f'E{numero}', f'R$ {round(total, 2)}', bold)
        workbook.close()

    def exportar_informacoes(self):
        lista_de_produtos = []
        total = 0
        for produtos in self.sql.exportar_dados():
            total += produtos[4]
            produtos = list(produtos)
            lista_de_produtos.append(produtos)
        tamanho_da_lista = len(lista_de_produtos)
        return lista_de_produtos, round(total, 2), tamanho_da_lista

    def fechar_conexao(self):
        self.access.fechar_conexao()
        self.sql.fechar_conexao()

    @staticmethod
    def verifica_se_a_pasta(nome):
        diretorio = fr'relatorios\{nome}'
        existe = os.path.exists(diretorio)
        if existe:
            return diretorio
        else:
            os.makedirs(fr'relatorios\{nome}')
            return diretorio


def verifica_se_tem_banco_access():
    banco_sql = Banco_de_dados_sql()
    caminho_senha = banco_sql.obter_caminho_access()
    banco_sql.fechar_conexao()
    return caminho_senha


def criar_caminho(caminho, senha):
    banco_sql = Banco_de_dados_sql()
    banco_sql.registra_caminho(caminho, senha)
    banco_sql.fechar_conexao()


def criar_tabelas():
    banco_sql = Banco_de_dados_sql()
    banco_sql.criar_tabela()


def verifica_se_o_campo_vazio_usuario(campo_de_busca, py):
    if campo_de_busca['-Cliente'] == "":
        py.popup_error("Campo vazio, digite o nome do cliente")
        return True
    return False


def verifica_se_o_campo_vazio_filtro(campo_de_busca, py):
    if campo_de_busca['filtro'] == "":
        py.popup_error("Campo vazio, digite a letra para filtrar")
        return True
    return False


def verifica_item_nota(nome, nota, py):
    tem_item_na_nota = nota.criar_nota_fiscal(nome)
    if tem_item_na_nota:
        texto = gerador_de_texto_errado(tem_item_na_nota, True)
        py.popup_error(texto)
    return tem_item_na_nota


def gerador_de_texto_errado(lista_de_notas, pop):
    if pop:
        texto = "Aviso tem um item na nota\n" \
                "numero da notas:"
    else:
        texto = "Numero das notas com o produto Item:"
    string_numeros = '\n'
    contador = 1
    colunas = 0
    for numero in lista_de_notas:
        string_numeros += str(numero)
        if contador % 5 == 0:
            string_numeros += '\n'
            colunas += 1
        else:
            string_numeros += ', '
        contador += 1
        if colunas == 5:
            if pop:
                return texto + string_numeros + '... limite ultrapassado de 25 notas'
            return texto + string_numeros + '... limite ultrapassado de 25 notas' + '\n' + f'{len(lista_de_notas)}'
    return texto + string_numeros + '\n' + f'quantidades de notas com item : {len(lista_de_notas)}'


def verifica_se_lista_vazia(lista_de_clientes, letra, py):
    if len(lista_de_clientes) == 0:
        py.popup_error(f'Aviso não tem este cliente com a letra {letra} \n')
        return True
    return False


def botao_exportar_buscar(campo_de_busca, nota, janela, py, lista_de_clientes, botao_exportar):
    botao_limpar(janela)
    if not verifica_se_o_campo_vazio_usuario(campo_de_busca, py):
        nome = campo_de_busca['-Cliente'].upper()
        clientes_filtrados = nota.filtro(nome)
        if len(clientes_filtrados) == 1:
            nome = clientes_filtrados[0]
        lista_de_notas_erradas = verifica_item_nota(nome, nota, py)
        dados, total, quantidade = nota.exportar_informacoes()
        if dados:
            if botao_exportar:
                nota.exportar_nota(nome)
                py.popup_ok("Exportação concluida, o arquivo se encontra na pasta relatorio com o nome do cliente")
            janela.find_element('-tabela').Update(dados)
            janela.find_element("-quantidade").Update(quantidade)
            janela.find_element("-total").Update(total)
            if lista_de_notas_erradas:
                texto = gerador_de_texto_errado(lista_de_notas_erradas, False)
                janela.find_element('erro_nota').Update(texto)
            if not verifica_se_lista_vazia(lista_de_clientes, nome, py):
                janela.find_element('-Cliente').Update(clientes_filtrados[0], values=clientes_filtrados)
        else:
            py.popup_error("Este cliente não possui notas ou ele não existe")


def botao_filtrar(campo_de_busca, janela, nota, py):
    ultimo_nome = campo_de_busca['-Cliente'].upper()
    if campo_de_busca['-Cliente'] == "" or ultimo_nome == campo_de_busca['-Cliente']:
        if not verifica_se_o_campo_vazio_filtro(campo_de_busca,py):
            letras = campo_de_busca['filtro']
        else:
            return
    else:
        letras = campo_de_busca['-Cliente'].upper()
    clientes_filtrados = nota.filtro(letras)
    if not verifica_se_lista_vazia(clientes_filtrados, letras, py):
        janela.find_element('-Cliente').Update(clientes_filtrados[0], values=clientes_filtrados)


def botao_limpar(janela):
    janela.find_element('-tabela').Update("")
    janela.find_element("-quantidade").Update('0')
    janela.find_element("-total").Update("0")
    janela.find_element("erro_nota").Update("")

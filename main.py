from datetime import datetime
from datetime import timedelta
from openpyxl import load_workbook
from openpyxl import Workbook
import string
import calendar


print("Encontrando Excel de origem...\n")
origin_file = load_workbook(filename='Excel_Original.xlsx')
origin_sheet = origin_file['Nome_Arquivo']
final_file = Workbook ()
final_sheet = final_file.active

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>VARIAVEIS E LISTAS<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

test_list = []
num03_list = []
codigos01 = []
codigos02 = []
codigos03 = []
codigos04 = []
codigos05 = []
trash = []
letras = list(string.ascii_uppercase)

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>CLASSES E FUNÇÕES<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

def add_months (original_date, months):
    month = original_date.month - 1 + months
    year = original_date.year + month // 12
    month = month % 12 + 1
    day = min(original_date.day, calendar.monthrange(year,month)[1])
    day = str(day)
    month = str(month)
    year = str(year)
    string_date = day + "/" + month + "/" + year
    new_date = datetime.strptime(string_date, "%d/%m/%Y")
    return new_date

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>COLETA E TRATAMENTO DOS CODIGOS<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

print("Coletando os Códigos...")
cod_list= origin_sheet['B']
temp_list = list(cod_list)
for i in cod_list:
    cod = i.value
    if cod == None:
        temp_list.remove(i)
temp_list.pop(0)
temp_list.pop(0)
cod_list = tuple(temp_list)
temp_list = []
print()

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>FILTRO DOS CODIGOS<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

for i in cod_list:
    cod = str(i.value)
    search = cod.find("03")
    if search == 0:
        codigos01.append(i)
    else:
        trash.append(i)

cod01_list = tuple(codigos01)

print("Criando filtro para separação dos códigos...")
for i in cod01_list:
    cod = str(i.value)
    search = len(cod)
    if search == 5:
        codigos02.append(i)
    elif search == 8:
        codigos03.append(i)
    elif search == 11:
        codigos04.append(i)
    elif search == 14:
        codigos05.append(i)

cod02_list = tuple(codigos02)
cod03_list = tuple(codigos03)
cod04_list = tuple(codigos04)
cod05_list = tuple(codigos05)
print()

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>FILTRO DAS DATAS<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

list_letter = list(string.ascii_uppercase)
count_letter = 0
for i in letras:
    a = 0
    while a < 26:
        list_letter.append(i + list_letter[a])
        a += 1
    count_letter += 1

data_medida = origin_sheet['2']
temp_list = list(data_medida)
for i in data_medida:
    info = i.value
    if info == None:
        temp_list.remove(i)

temp_list.pop(0)
temp_list.pop(0)
temp_list.pop()
temp_list.pop()
data_medida = tuple(temp_list)
temp_list = []

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>CRIAÇÃO DAS COLUNAS<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

print("Nomeando as colunas na planilha nova...")
colunas = [
    "Código", "Etapa", "Local", "Disciplina", 
    "Sub-Disciplina", "Descrição 05", "Unidade", "Quantidade Orçada", 
    "Quantidade Medida", "Quantidade Replanejada", "Valor Orçado", 
    "Valor Medido", "Valor Replanejado", "Data Inicio", "Data Término", 
    "Data Medição", "Status"
    ]

count = 0
for i in colunas:
    count +=1
    final_sheet['%s2' % letras[count]] = i
print()

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>ORGANIZAÇÃO DOS DADOS<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

print("Organizando dados a serem inseridos na nova planilha...")
row_sequence = 3
for i in cod05_list:
    cod05 = str(i.value)
    #Coleta do codigo de nivel 01
    for j in cod01_list:
        temp_cod = str(j.value)
        if temp_cod == cod05[0:2]:
            cod01 = j

    #Coleta do codigo de nivel 02
    for k in cod02_list:
        temp_cod = str(k.value)
        if temp_cod == cod05[0:5]:
            cod02 = k

    #Coleta do codigo de nivel 03
    for l in cod03_list:
        temp_cod = str(l.value)
        if temp_cod == cod05[0:8]:
            cod03 = l
    
    #Coleta do codigo de nivel 04
    for m in cod04_list:
        temp_cod = str(m.value)
        if temp_cod == cod05[0:11]:
            cod04 = m

    #coleta da linha de cada informação
    cod01_row = cod01.row
    cod02_row = cod02.row
    cod03_row = cod03.row
    cod04_row = cod04.row
    cod05_row = i.row

    #Numero de periodos
    num_period = int(origin_sheet['K%s' % cod05_row].value)

    #Informações de datas de inicio
    string_date = str(origin_sheet['H%s' % cod05_row].value)[0:10]
    data_str = datetime.strptime(string_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    data_inicio = datetime.strptime(data_str, "%d/%m/%Y")

    #Coleta das descrições 
    descricao01 = str(origin_sheet['C%s' % cod01_row].value)
    descricao02 = str(origin_sheet['C%s' % cod02_row].value)
    descricao03 = str(origin_sheet['C%s' % cod03_row].value)
    descricao04 = str(origin_sheet['C%s' % cod04_row].value)
    descricao05 = str(origin_sheet['C%s' % cod05_row].value)
    unidade = str(origin_sheet['D%s' % cod05_row].value)
    qnt_orcado = int(origin_sheet['E%s' % cod05_row].value)/num_period
    valor_orcado = int(origin_sheet['G%s' % cod05_row].value)/num_period
    
    
    #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>PREENCHIMENTO DA PLANILHA<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    count = 0
    while count < num_period:
        string_date = data_inicio.strftime("%d/%m/%Y")
        
        #Criação do filtro da data medida
        valor_filter = string_date[3:10]
        for i in data_medida:
            info = i.value
            if valor_filter == info[0:7]:
                data = i
                break

        #Seleção da coluna de valores dentro de quantidade medida.
        b = data.value
        search = b.find("Replanejado")
        coluna_quantidade_medida = list_letter[data.column]
        quantidade = origin_sheet['%s%s' % (coluna_quantidade_medida, cod05_row)].value

        #Filtro para identificar quando a quantidade deve ser 0
        if search == 9:
            qnt_medida = 0
        elif quantidade == None:
            qnt_medida = 0
        else:
            qnt_medida = quantidade

        #Criação da quantidade Replanejada
        qnt_replanejada = qnt_orcado - qnt_medida

        #Seleção da coluna de valores dentro de valor medido.
        b = data.value
        search = b.find("Replanejado")
        coluna_valor_medido = list_letter[data.column+1]
        valor = origin_sheet['%s%s' % (coluna_valor_medido, cod05_row)].value
        
        #Filtro para identificar quando o valor deve ser 0
        if search == 9:
            valor_medido = 0
        elif valor == None:
            valor_medido = 0
        else:
            valor_medido = valor
        
        #Criação do valor Replanejado
        valor_replanejado = valor_orcado-valor_medido

        #Organização dos dados na planilha
        final_sheet['B%s' % row_sequence] = cod05
        final_sheet['C%s' % row_sequence] = descricao01
        final_sheet['D%s' % row_sequence] = descricao02
        final_sheet['E%s' % row_sequence] = descricao03
        final_sheet['F%s' % row_sequence] = descricao04
        final_sheet['G%s' % row_sequence] = descricao05
        final_sheet['H%s' % row_sequence] = unidade
        final_sheet['I%s' % row_sequence] = qnt_orcado
        final_sheet['J%s' % row_sequence] = qnt_medida
        final_sheet['K%s' % row_sequence] = qnt_replanejada
        final_sheet['L%s' % row_sequence] = valor_orcado
        final_sheet['M%s' % row_sequence] = valor_medido
        final_sheet['N%s' % row_sequence] = valor_replanejado
        final_sheet['O%s' % row_sequence] = string_date
        data_inicio = add_months(data_inicio, 1)
        data_final = data_inicio + timedelta(days=-1)
        final_sheet['P%s' % row_sequence] = data_final.strftime("%d/%m/%Y")
        data_medicao = data_final.strftime("%d/%m/%Y")
        final_sheet['Q%s' % row_sequence] = data_medicao

        #Filtro para coletar a data de hoje
        data_hoje = datetime.now()
        #Criação do Status da etapa
        '''print("Valor Replanejado:")
        print(valor_replanejado)
        print("Data Hoje:")
        print(data_hoje)
        print("Data Inicio:") 
        print(data_inicio)
        print("Data Final:") 
        print(data_final)
        if valor_replanejado > 0:
            status = "REPLANEJADO"
            print("REPLANEJADO")
        elif data_hoje < data_inicio:
            status = "A EXECUTAR"
            print("NÃO INICIADO")
        elif data_hoje > data_inicio and data_hoje < data_final:
            status = "EM EXECUÇÃO"
            print("EM EXECUÇÃO")
        elif data_hoje > data_final:
            if valor_medido == valor_orcado:
                status = "EXECUTADO"
                print("EXECUTADO")
        elif valor_replanejado <= 0:
            status = "ADIANTANDO"'''
                
        if valor_replanejado > 0:
            status = "REPLANEJADO"
        elif valor_replanejado == 0:
            status = "EXECUTADO"
        elif valor_replanejado < 0:
            status = "ADIANTADO"

        final_sheet['R%s' % row_sequence] = status
        count += 1
        row_sequence += 1

print()        
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>SALVANDO ARQUIVO<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
print("Criando novo arquivo Excell...")
nome_final = input("Por favor me informe o nome da planilha: ")
final_file.save(filename=nome_final+".xlsx")
print()

print("Arquivo criado com sucesso.")
print("Nome do Aquivo: " + nome_final + ".xlsx")
input("Precione uma tecla para encerrar")
print()

final_sheet.column_dimensions
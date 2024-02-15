import gspread
from gspread import worksheet

# Configurações
SPREADSHEET_ID = '1U_H6zWQ-h2uo6YViEkH7o4NudN4rZv85Cil3R8tfyrs'
SHEET_NAME = 'Engenharia de Software - Desafio - EDILTON SILVA JUNIOR'

# Conexão com a planilha
gc = gspread.service_account('directed-pier-413413-c5b9dc5ceed0.json')
#sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
sheet = gc.open_by_key(SPREADSHEET_ID).get_worksheet(0)


# Cálculo da situação e da nota para aprovação final
for index, row in enumerate(sheet.get_all_values()[3:]): 
  student_name = row[0]
  grades = [float(grade) for grade in row[3:5]]  # Colunas D, E, F (índices  3,  4,  5)
  absences = int(row[2])  # Coluna C (índice  2)
  total_classes =  60

  # Média
  average = sum(grades) / len(grades)

  # Situação
  if absences > total_classes/4:
    situation = 'Reprovado por Falta'
    grade_final = 0
  elif (average/10) <  5:
    situation = 'Reprovado por Nota'
    grade_final = 0
  elif (average/10) <  7:
    situation = 'Exame Final'
    grade_final = (10 - (average/10)) *  2
    grade_final = round(grade_final,  0)
  else:
    situation = 'Aprovado'
    grade_final = 0

  # Gravação dos resultados
  sheet.update_cell(index +  4,  7, situation)  # Coluna G (índice  7), começando na linha  4
  sheet.update_cell(index +  4,  8, grade_final)  # Coluna H (índice  8), começando na linha  4

  # Log
  print(f'Aluno: {student_name}')
  print(f'Média: {average}')
  print(f'Faltas: {absences}')
  print(f'Situação: {situation}')
  print(f'Nota para Aprovação Final: {grade_final}')
  print('-' *  20)
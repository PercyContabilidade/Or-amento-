import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Create workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Comparativo de Preços"

# Ensure grid lines are visible
ws.views.sheetView[0].showGridLines = True

# Paleta de Cores de Alta Performance (Bordô/Vinho e Ouro / Tons Profissionais)
HEADER_FILL = PatternFill(start_color="4A0E17", end_color="4A0E17", fill_type="solid") # Vinho escuro
SUBHEADER_FILL = PatternFill(start_color="801826", end_color="801826", fill_type="solid") # Vinho médio
ZEBRA_FILL = PatternFill(start_color="FDF6F7", end_color="FDF6F7", fill_type="solid") # Rosa/creme bem claro
CRPS_FILL = PatternFill(start_color="EAF2F8", end_color="EAF2F8", fill_type="solid") # Azul claro sutil para CRPS
PF_FILL = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid") # Amarelo/ouro bem claro para PF
HIGHLIGHT_WINNER = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid") # Verde claro para a opção mais barata

# Fontes
font_title = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
font_section = Font(name="Segoe UI", size=12, bold=True, color="4A0E17")
font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
font_data = Font(name="Segoe UI", size=11, bold=False, color="000000")
font_data_bold = Font(name="Segoe UI", size=11, bold=True, color="000000")
font_winner = Font(name="Segoe UI", size=11, bold=True, color="196F3D")

# Bordas
thin_line = Side(border_style="thin", color="D3D3D3")
double_line = Side(border_style="double", color="4A0E17")
border_cell = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)
border_total = Border(top=thin_line, bottom=double_line)

# Título Principal
ws.merge_cells("A1:I1")
ws["A1"] = "ANÁLISE COMPARATIVA DE ORÇAMENTOS"
ws["A1"].font = font_title
ws["A1"].fill = HEADER_FILL
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 40

# Subtítulo com dados fornecidos
ws.merge_cells("A2:I2")
ws["A2"] = "CRPS Construções LTDA (Impresso) vs. Posse Ferro (Anotações à Caneta)"
ws["A2"].font = Font(name="Segoe UI", size=11, italic=True, color="FFFFFF")
ws["A2"].fill = SUBHEADER_FILL
ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 25

# Linha em branco
ws.row_dimensions[3].height = 15

# Cabeçalhos da Tabela
headers = [
    "Item / Produto", "Qtd", "Unid", 
    "Unit. CRPS", "Total CRPS", 
    "Unit. Posse Ferro", "Total Posse Ferro", 
    "Economia Unit.", "Melhor Opção"
]

for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=4, column=col_idx, value=header)
    cell.font = font_header
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_cell
ws.row_dimensions[4].height = 30

# Dados dos itens baseados nas imagens
# Formato: [Produto, Qtd, Unid, Unit_CRPS, Unit_PF_anotado]
# Se Unit_PF_anotado for None, assume-se que mantém o valor ou não tem cotação explícita separada (mas nas anotações vemos alguns valores específicos).
# Vamos destrinchar:
# 1. CIMENTO CP2 MAUA: Qtd 5 UN. CRPS: Unit 38.00, Total 190.00. PF anotado: 39.90.
# 2. PEDRA: Qtd 0.5 UN. CRPS: Unit 175.00, Total 87.50. PF anotado: 43.90 (provavelmente por m3 ou o total/unidade ajustado, vamos verificar a anotação: está do lado de 'Caixas:0' abaixo de cimento e acima de Pedra, mas alinhado com Pedra parece 43.90 ou é do cimento? Olhando a imagem 1: '39,90 P.F' está na linha do Título Cimento. '43,90 P.F' está na linha de Caixas:0 logo abaixo dos R$190,00 e logo acima de PEDRA. Pode ser o preço da Pedra por 0.5 m³ ou outra especificação. Vamos analisar com cuidado: 43.90 está na direção de Pedra? Não, Pedra está abaixo. Vamos colocar como cotação da Pedra ou se não houver certeza, manter o padrão).
# Vamos olhar os números azuis:
# Item 1 (Coluna Ferro 1/4 6MT): Anotado "48,90"
# Item 2 (Argila): Sem anotação azul clara ao lado, mas tem os números de ordem 1, 2, 3, 4, 5, 6, 7 na esquerda.
# Item 4 (Tabua 20cm): Anotado "29,00"
# Item 5 (Prego 17x21): Anotado "27 / 27 16,90" -> 16,90? Ou 27,00? Vamos analisar a imagem: "27" riscado e "16,90" do lado.
# Item 6 (Vaso Acop Beja Gold Branco): Impresso 473,00. Anotado "270" e embaixo "323,90". Provavelmente 323,90 é o preço final.
# Vamos estruturar os dados com o que está claro e inteligível nas imagens:

itens_dados = [
    ["Cimento CP2 Mauá", 5, "UN", 38.00, 39.90],
    ["Pedra (0,5 m³)", 0.5, "UN", 175.00, 175.00], # Sem alteração explícita clara de unitário, mantém CRPS para comparação ou o 43.90 pode ser o total da pedra? Se 0.5 un de 87.50 virar 43.90 faz todo sentido (metade de 87.80)! Então o unitário seria 87.80. Vamos colocar 87.80 unitário na PF para dar 43.90 total.
    ["Areia Tipo A", 0.5, "UN", 150.00, 150.00],
    ["Coluna Ferro 1/4 6MT", 1, "UN", 50.50, 48.90],
    ["Argila", 3, "UN", 7.00, 7.00],
    ["Arame Queimado", 1, "UN", 22.00, 22.00],
    ["Tábua 20CM", 4, "UN", 21.99, 29.00],
    ["Prego 17X21", 1, "UN", 32.50, 16.90],
    ["Vaso Acop Beja Gold Branco", 1, "UN", 473.00, 323.90],
    ["Tijolos 19X29 de Campos", 800, "UN", 1.38, 1.38]
]

start_row = 5
for idx, item in enumerate(itens_dados):
    row = start_row + idx
    
    # Nome do produto, Qtd, Unid
    ws.cell(row=row, column=1, value=item[0]).font = font_data
    ws.cell(row=row, column=2, value=item[1]).font = font_data
    ws.cell(row=row, column=3, value=item[2]).font = font_data
    
    # CRPS Unitário e Total (Fórmula)
    cell_unit_crps = ws.cell(row=row, column=4, value=item[3])
    cell_unit_crps.number_format = 'R$#,##0.00'
    cell_unit_crps.font = font_data
    
    cell_tot_crps = ws.cell(row=row, column=5, value=f"=B{row}*D{row}")
    cell_tot_crps.number_format = 'R$#,##0.00'
    cell_tot_crps.font = font_data
    
    # Posse Ferro Unitário e Total (Fórmula)
    cell_unit_pf = ws.cell(row=row, column=6, value=item[4])
    cell_unit_pf.number_format = 'R$#,##0.00'
    cell_unit_pf.font = font_data
    
    cell_tot_pf = ws.cell(row=row, column=7, value=f"=B{row}*F{row}")
    cell_tot_pf.number_format = 'R$#,##0.00'
    cell_tot_pf.font = font_data
    
    # Economia Unitária (CRPS - PF)
    cell_econ = ws.cell(row=row, column=8, value=f"=D{row}-F{row}")
    cell_econ.number_format = 'R$#,##0.00'
    cell_econ.font = font_data_bold
    
    # Melhor Opção (Fórmula SE)
    cell_opt = ws.cell(row=row, column=9, value=f'=IF(D{row}<F{row}, "CRPS", IF(D{row}>F{row}, "Posse Ferro", "Empate"))')
    cell_opt.font = font_data_bold
    cell_opt.alignment = Alignment(horizontal="center")
    
    # Zebra striping e bordas
    fill_to_use = ZEBRA_FILL if idx % 2 == 1 else PatternFill(fill_type=None)
    for col_idx in range(1, 10):
        c = ws.cell(row=row, column=col_idx)
        c.border = border_cell
        if fill_to_use.fill_type and col_idx not in [4, 5, 6, 7]: # não sobrepor os blocos de cores se quisermos destacar as lojas
            c.fill = fill_to_use

    # Pequeno toque de estilo: colorir as colunas das lojas de forma sutil
    ws.cell(row=row, column=4).fill = CRPS_FILL
    ws.cell(row=row, column=5).fill = CRPS_FILL
    ws.cell(row=row, column=6).fill = PF_FILL
    ws.cell(row=row, column=7).fill = PF_FILL

    ws.row_dimensions[row].height = 22

# Linha de Totais Gerais
total_row = start_row + len(itens_dados)

ws.cell(row=total_row, column=1, value="VALOR TOTAL DO ORÇAMENTO").font = font_section
ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="right")
ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)

# Totais das colunas
cell_sum_crps = ws.cell(row=total_row, column=5, value=f"=SUM(E5:E{total_row-1})")
cell_sum_crps.number_format = 'R$#,##0.00'
cell_sum_crps.font = font_section
cell_sum_crps.fill = CRPS_FILL

cell_sum_pf = ws.cell(row=total_row, column=7, value=f"=SUM(G5:G{total_row-1})")
cell_sum_pf.number_format = 'R$#,##0.00'
cell_sum_pf.font = font_section
cell_sum_pf.fill = PF_FILL

# Economia Total
cell_sum_econ = ws.cell(row=total_row, column=8, value=f"=E{total_row}-G{total_row}")
cell_sum_econ.number_format = 'R$#,##0.00'
cell_sum_econ.font = font_section

# Vencedor Geral
cell_winner_gen = ws.cell(row=total_row, column=9, value=f'=IF(E{total_row}<G{total_row}, "CRPS", "Posse Ferro")')
cell_winner_gen.font = font_title
ws.cell(row=total_row, column=9).font = Font(name="Segoe UI", size=11, bold=True, color="196F3D")
cell_winner_gen.alignment = Alignment(horizontal="center")
cell_winner_gen.fill = HIGHLIGHT_WINNER

# Aplicar bordas de total
for col_idx in range(1, 10):
    ws.cell(row=total_row, column=col_idx).border = border_total

ws.row_dimensions[total_row].height = 30

# Quadro Resumo / Insights (Design de Dashboard)
summary_start = total_row + 3
ws.merge_cells(f"A{summary_start}:I{summary_start}")
ws[f"A{summary_start}"] = "CONCLUSÕES E INSIGHTS DA NEGOCIAÇÃO"
ws[f"A{summary_start}"].font = font_header
ws[f"A{summary_start}"].fill = SUBHEADER_FILL
ws[f"A{summary_start}"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[summary_start].height = 25

insights = [
    "1. O orçamento da 'Posse Ferro' apresenta uma economia global significativa, impulsionada principalmente pelo Vaso Sanitário Acoplado Beja Gold.",
    "2. O Vaso Beja Gold Branco na CRPS custa R$ 473,00, enquanto na Posse Ferro foi anotado por R$ 323,90 (Economia de R$ 149,10 neste item).",
    "3. O Cimento CP2 Mauá está mais barato na CRPS (R$ 38,00 vs R$ 39,90 da Posse Ferro).",
    "4. Sugestão Prática: Utilize esta planilha para propor à loja de sua preferência cobrir os preços mais baixos da concorrência antes de fechar no cartão."
]

for idx, insight in enumerate(insights, 1):
    r = summary_start + idx
    ws.merge_cells(f"A{r}:I{r}")
    cell_ins = ws.cell(row=r, column=1, value=insight)
    cell_ins.font = Font(name="Segoe UI", size=10, italic=True)
    cell_ins.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 20

# Auto-ajuste de largura de colunas para acabamento perfeito
for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value and cell.row > 2: # ignora títulos longos fundidos para não distorcer a largura
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

# Forçar larguras específicas para colunas chave para ficar elegante
ws.column_dimensions['A'].width = 32
ws.column_dimensions['B'].width = 8
ws.column_dimensions['C'].width = 8

# Salvar o arquivo final
filename = "Comparativo_Orcamentos_CRPS_PosseFerro.xlsx"
wb.save(filename)
print(f"Arquivo salvo com sucesso: {filename}")



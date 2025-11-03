import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg

# --- Exemplo de DataFrame ---
data = {
    "CARTEIRAS": [
        "NEON AMIGAVEL 2 - SAFRA 1", "NEON AMIGAVEL 2 - SAFRA 2",
        "NEON EMPRESTIMO MEI - DG", "NEON EMPRESTIMO MEI - SAFRA 1"
    ],
    "QTD_CLIENTES_ATIVOS": [22937, 31790, 6586, 6625],
    "QTD_CLIENTES_NOVOS": [605, 849, 24, 27],
    "QTD_CONTRATOS_ATIVOS": [22937, 31790, 7121, 7133],
    "QTD_CONTRATOS_NOVOS": [605, 849, 28, 30],
    "VALOR_ACUMULADO": [39414518.34, 52800022.43, 49043747.81, 49095057.84],
    "QTD_TELEFONES_ATIVOS": [70263, 90234, 22903, 21646],
    "QTD_ENTRADAS": [563, 1076, 38, 48],
    "QTD_SAIDAS": [0, 0, 0, 0],
}

df = pd.DataFrame(data)

# Variavel para o local de armazenamento dos arquivos processados
path = "C:/Users/OT/Documents/Projetos/PosicaoCarteira/"

# --- Layout ---
fig, ax = plt.subplots(figsize=(15, 6))
ax.axis('off')

ax.set_xlim(0, 15)
ax.set_ylim(0, 10)

img = mpimg.imread(f"{path}input/innLogo.png")
ax.imshow(img, extent=[6, 9.2, 8, 10], aspect='auto', zorder=5, alpha=1)
plt.text(0.5, 0.86, "20/10/2025", fontsize=14, fontweight='bold', ha='center', color='#444')

# --- Criar tabela ---
tabelaHoje = ax.table(cellText=df.values,
                 colLabels=df.columns,
                 cellLoc='center',
                 loc='center')

tabelaOntem = ax.table(cellText=df.values,
                 colLabels=df.columns,
                 cellLoc='center',
                 loc='center')

table.auto_set_font_size(False)
table.set_fontsize(5)
table.scale(1, 1.5)

# --- Estilização ---
for (row, col), cell in table.get_celld().items():
    if row == 0:
        cell.set_facecolor('#e89b5d')   # Cabeçalho
        cell.set_text_props(weight='bold', color='white')
    else:
        cell.set_facecolor('#f8f4f0' if row % 2 == 0 else '#ffffff')

plt.tight_layout()
plt.show()

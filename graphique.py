import matplotlib.pyplot as plt

annees = [2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]
tt_type = [755, 864, 1130, 1480, 1960, 2310, 3010, 3220, 5280, 7850, 5820]
slr = [31, 37, 49, 69, 88, 166, 227, 262, 546, 935, 1050]

fig, ax1 = plt.subplots()

# Courbe TT TYPE
l1, = ax1.plot(annees, tt_type, marker='o', label='TT TYPE')
ax1.set_xlabel('Année')
ax1.set_ylabel('TT TYPE')

# Deuxième axe
ax2 = ax1.twinx()

# Courbe SLR
l2, = ax2.plot(annees, slr, marker='s', linestyle='--', color='red', label='SLR')
ax2.set_ylabel('SLR', color='red')

# Fusionner les deux légendes
plt.legend(handles=[l1, l2], loc='upper left')

plt.title("Évolution de TT TYPE et SLR (2015-2025)")
plt.show()

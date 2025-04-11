import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from collections import Counter


# Del a)
df = pd.read_excel('support_uke_24.xlsx')

u_dag = np.array(df["Ukedag"])
kl_slett = np.array(df["Klokkeslett"])
varighet = np.array(df["Varighet"])
score = np.array(df["Tilfredshet"].dropna())

# Del b)

dag_rekkefolge = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']

ukedager, per_dag = np.unique(u_dag ,return_counts=True)

antall_per_ukedag = dict(zip(ukedager, per_dag))
sortert_antall = [antall_per_ukedag.get(dag, 0) for dag in dag_rekkefolge]
#Konverterer til numpy array for å unngå type advarsel
sortert_antall = np.array(sortert_antall)

fig, ax = plt.subplots()
ax.bar(dag_rekkefolge, sortert_antall)
ax.set_title("Antall henvendelser for de fem ukedagene")
ax.set_ylabel("Antall henvendelser")
plt.tight_layout()
plt.show()

# Del c)

def format_tid(tid):
    total_sek = int(tid.total_seconds())
    minutter = total_sek // 60
    sekunder = total_sek % 60
    if minutter > 0:
        return f"{minutter} minutter og {sekunder} sekunder"
    else:
        return f"{sekunder} sekunder"

varighet = pd.to_timedelta(varighet)
lengst_samtaletid = varighet.max()
kortest_samtale = varighet.min()

print(f'Den lengste samtaletiden var på {format_tid(lengst_samtaletid)}')
print(f'Den korteste samtalen var på {format_tid(kortest_samtale)}')

# Del d)

gjennomsnitt = varighet.mean()
print(f'Gjennomsnittlig samtaletid for uke 24 er {format_tid(gjennomsnitt)}')


# Del e)

def hent_intervall(klokkeslett):
    time = pd.to_datetime(klokkeslett).hour
    if 8 <= time < 10:
        return "08-10"
    elif 10 <= time < 12:
        return "10-12"
    elif 12 <= time < 14:
        return "12-14"
    return "14-16"

def prosent_og_antall(pct):
    total = sum(antall_tidsrom)
    count = int(round(pct * total / 100.0))
    return f"{pct:.1f}%\n({count})"

tidsintervaller = ['08-10', '10-12', '12-14', '14-16']
intervall_liste = [hent_intervall(tid) for tid in kl_slett]
teller = Counter(intervall_liste)
antall_tidsrom = [teller.get(intervall, 0) for intervall in tidsintervaller]

_,ax = plt.subplots()
ax.pie(
    antall_tidsrom,
    autopct=prosent_og_antall,
)

ax.set_title("Totalt antall henvendelser per 2-timers periode for uke 24")
ax.legend(tidsintervaller,
          title="Tidsrom",
          loc="upper left",
          bbox_to_anchor=(1, 0, 0.5, 1))

plt.tight_layout()
plt.show()


# Del f)
antall_tilbakemeldinger = len(score)

antall_positive = (score >= 9).sum()
antall_negative = (score <= 6).sum()

prosent_positive = (antall_positive / antall_tilbakemeldinger) * 100
prosent_negative = (antall_negative / antall_tilbakemeldinger) * 100

nps = prosent_positive - prosent_negative
print(f"Supportavdelingens NPS er: {nps:.1f}%")
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from collections import Counter


# Del a)
# Leser data fra xlsx-filen
df = pd.read_excel('support_uke_24.xlsx')

# Konverter kolonner til NumPy-arrays (score dropper NaN-verdier)
u_dag = np.array(df["Ukedag"])
kl_slett = np.array(df["Klokkeslett"])
varighet = np.array(df["Varighet"])
score = np.array(df["Tilfredshet"].dropna())

# Del b)

dag_rekkefolge = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']
# Finner unike ukedager og antallet for hver
ukedager, per_dag = np.unique(u_dag ,return_counts=True)
# Dictionary med antall per ukedag
antall_per_ukedag = dict(zip(ukedager, per_dag))
# Henter antallet i rett rekkefølge mtp ukedager
sortert_antall = [antall_per_ukedag.get(dag, 0) for dag in dag_rekkefolge]
#Konverterer til numpy array for å unngå type advarsel
sortert_antall = np.array(sortert_antall)

# plotter stolpedigrammet
fig, ax = plt.subplots()
ax.bar(dag_rekkefolge, sortert_antall)
ax.set_title("Antall henvendelser for de fem ukedagene")
ax.set_ylabel("Antall henvendelser")
plt.tight_layout()
plt.show()

# Del c)

def format_tid(tid):
    """
    Konverterer et timedelta-objekt til en lesbar streng med minutter og sekunder.
    Kalkulerer den totale tiden i sekunder.
    Den deler deretter opp tiden i hele minutter og resterende sekunder, og returnerer
    en formatert streng. Hvis tidsintervallet er mindre enn ett minutt, returneres kun sekunder.

    :param tid: Et timedelta-objekt som representerer en varighet.
    :return: En streng som beskriver varigheten i minutter og sekunder, for eksempel "5 minutter og 32 sekunder"
             eller "45 sekunder" hvis antall minutter er null.
    """
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
    """
    Konverterer et klokkeslett til et 2-timers intervall.

    Tar en klokkeslettverdi (f.eks "15:56:10") og konverterer
    den til et datetime-objekt for å hente ut timen. Avhengig av timen returneres et
    intervall/tidsrom

    :param klokkeslett: En streng med et klokkeslett (forventet format: "HH:MM" / "HH:MM:SS").
    :return: En streng som representerer 2-timers tidsintervallet.
    """
    time = pd.to_datetime(klokkeslett).hour
    if 8 <= time < 10:
        return "08-10"
    elif 10 <= time < 12:
        return "10-12"
    elif 12 <= time < 14:
        return "12-14"
    return "14-16"

def prosent_og_antall(pct):
    """
    Formatterer prosentandelen for en sektor i et kakediagram med tilsvarende absolutt antall.

    Funksjonen beregner det absolutte antallet ved å bruke den totale summen av verdier
    fra 'antall_tidsrom'. Den returnerer en formatert streng der prosentandelen (pct) vises
    med én desimal, og det beregnede antallet vises på en ny linje.

    :param pct: Prosentandelen (float) for en sektor i kakediagrammet.
    :return: en prosendtandel og det absolutte tallet for den enkelte sektor
    """
    total = sum(antall_tidsrom)
    count = int(round(pct * total / 100.0))
    return f"{pct:.1f}%\n({count})"

# Definerer tidsintervallene og grupper henvendelsene
tidsintervaller = ['08-10', '10-12', '12-14', '14-16']
intervall_liste = [hent_intervall(tid) for tid in kl_slett]
teller = Counter(intervall_liste)
antall_tidsrom = [teller.get(intervall, 0) for intervall in tidsintervaller]

# Sektordiagram som viser prosent og antall henvendelser for hvert tidsrom
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
# antall tilbakemeldinger kundetilfredshet
antall_tilbakemeldinger = len(score)

# Teller positive tilbakemeldinger (score på 9 eller 10) og negative tilbakemeldinger (fra 1 til 6)
antall_positive = (score >= 9).sum()
antall_negative = (score <= 6).sum()

# Beregn prosentandel for positive og negative tilbakemeldinger
prosent_positive = (antall_positive / antall_tilbakemeldinger) * 100
prosent_negative = (antall_negative / antall_tilbakemeldinger) * 100

nps = prosent_positive - prosent_negative
print(f"Supportavdelingens NPS er: {nps:.1f}%")
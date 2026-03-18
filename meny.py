# MORSE SUPPORT DASHBOARD
"""
Prosjektoppgave
USN - Python vår 2026
av Vegard H. Johansen
Versjon 1.2
"""

# IMPORTS
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# FLAGG
data_er_lest = 0 #Flagg som settes når xls data er lest inn



# FUNKSJON SOM KALLER HOVEDMENYEN
def meny():
    print("\033[2J\033[H") #Kommando for å rense skjermern før meny tegnes
    print("*************************************")
    print("*  MORSE - SUPPORT DASHBOARD        *")
    print("*                                   *")
    print("*        ---- MENY ----             *")
    print("*  1. Last inn .xlsx data           *")
    print("*  2. Vis henvendelser pr. ukedag   *")
    print("*  3. Vis kortest/lengste samtale   *")
    print("*  4. Vis gj.snittlig  samtaletid   *")
    print("*  5. Vis henvedelser pr. vakt      *")
    print("*  6. Vis avdelingens NPS tall      *")
    print("*  7. AVSLUTT PROGRAM               *")
    print("*                                   *")
    print("*************************************")

    try:  #Bruker try og except her for å fange opp hvis bruker skriver inn noe annet enn 1 -7 
        valg = int(input("Skriv inn ditt valg (nummer) fra menyen og trykk enter: ")) #Input som setter variabel valg fra brukerens input 
    except ValueError:
        input("Feil verdi. Velg mellom valg 1-7 - Trykk ENTER for å prøve igjen") #Tilbakemelding til bruker om feil input
        meny() #Laster meny på nytt 
        return #Avbryter funksjonen

    match valg:  #Bruker match her istedet for if. Bruker verdien fra variabel valg til å velge neste funksjon som skal hentes
        case 1:
            les_xls()
        case 2:
            plott_ukedag()
        case 3:
            finn_min_max()
        case 4:
            finn_gj_samtaletid()
        case 5:
            tidsrom()
        case 6:
            tilfredshet()
        case 7:
            print("Avslutter program")
            exit

# FUNKSJON SOM SJEKKER AT DATA ER LEST INN, RETURNERER HVIS IKKE
def sjekk_data():
    if data_er_lest == 0: #Hvis data ikke er lastet inn så gi tilbakemelding og la bruker prøøve på nytt
        print("Du må laste inn data med valg 1 i menyen først")
        input("Trykk ENTER for for å komme tilbake til menyen")
        meny()
    else:
        return True # Hvis data er lastet så returner True tilbake fra funksjonen


# FUNKSJON SOM LESER INN XLSX DATA = VALG 1. I MENYEN
def les_xls():
    global e_data, u_dag, kl_slett, varighet, score, data_er_lest #Setter disse verdiene som globale så du kan brukes utenfor funksjonen
    e_data = pd.read_excel("support_uke_24.xlsx", sheet_name="Ark1") #Laster xlsdata men pandas 

    # Lager numpy-arrays fra pandas dataene  
    u_dag = e_data["Ukedag"].to_numpy()
    kl_slett = e_data["Klokkeslett"].to_numpy()
    varighet = e_data["Varighet"].to_numpy()
    score = e_data["Tilfredshet"].to_numpy()

    data_er_lest = 1 #Setter flagg om at data er lastet inn
    
    input("Data er lastet inn. Trykk ENTER for å gå videre")
    meny() #Gir tilbakemelding til bruker og henter opp meny på nytt.


# FUNKSJON SOM PLOTTER ANTALL HENVENDELSER PR UKEDAG VALG 2. I MENYEN
def plott_ukedag():

    if sjekk_data(): #Henter funksjon sjekkdata og går bare videre om denne er True

        dager = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"] #Lager en list med ukedager til bruk i for-loop

        hv_pr_dag = [(u_dag == d).sum() for d in dager] #Bruker listen dager og henter ut hver dag som d og sjekker u_dag opp d og setter summen i listen hv_pr_dag

        # Tegner plottet og legger inn verdiene
        plt.bar(dager, hv_pr_dag) #Legger inn dager og verdier fra listene dager og hv_pr_dag
        plt.title("Antall henvendelser fordelt pr. ukedag")
        plt.xlabel("Dager")
        plt.ylabel("Antall henvendelser")
        plt.show() #Henter plot til skjerm

        meny() #Laster tilbake meny i termenal for neste valg.

# FUNKSJON SOM FINNER LENGSTE OG KORTESTE SAMTALE VALG 3. I MENYEN
def finn_min_max():

    if sjekk_data(): #Henter funksjon sjekkdata og går bare videre om denne er True

        lengste = str(np.max(varighet)).split(":") #Bruker funksjonen max fra numpy og finner største verdi i varighet array og legger den i en lengste list
        korteste = str(np.min(varighet)).split(":")  #Bruker funksjonen min fra numpy og finner minste verdi i varighet array og legger den i en korteste list

        #Printer ut tidene på korteste og lengste samtaler ved å bruke verdiene i lengste/korteste listen og gjør samtidig en sjekk med "in-line" if for å unngå å skrive timer og minutter hvis ikke finnes.
        print(f"Den lengste samtalen var på: {lengste[0] + 'timer,' if lengste[0] != '00' else '' } {lengste[1] + ' minutter og' if lengste[1] != '00' else '' } {lengste[2]} sekunder")
        print(f"Den korteste samtalen var på: {korteste[0] + 'timer,' if korteste[0] != '00' else '' } {korteste[1] + ' minutter og' if korteste[1] != '00' else '' } {korteste[2]} sekunder")

        input("Trykk ENTER for for å komme tilbake til menyen") #Bruker en input til å "stoppe" og går videre og laster meny når bruker trykker enter
        meny()    

# FUNKSJON SOM FINNER GJENNOMSNITTLIG SAMTALETID VALG 4. I MENYEN
def finn_gj_samtaletid():

    if sjekk_data(): #Henter funksjon sjekkdata og går bare videre om denne er True

        tid = 0 #Varianel for å holde verdien til tid
        i = 0 # teller for antall rader for er finnes
        for rad in varighet: #For loop som går igjennom alle oppføringene i varighets-array
            deler = rad.split(":") #Splitter hver tid opp i timer, minutter og sekunder og legger det i en litse
            tid = tid + int(deler[0]) * 3600 + int(deler[1]) * 60 + int(deler[2]) # legger til i variabel tid omgjort til sekunder
            i = i + 1 #Legger til 1 på teller for hver oppføring som er lest
        
        gj_sn = tid // i # Gjør en heltallsdeling av tid opp mot antall oppføringer og får gj.snitt antall sekunder pr oppføring og legger den i gj_snitt
        minutter = gj_sn // 60 #Gjør en heltallsdeling av gj_snitt og finner antall minutter 
        sekunder = gj_sn - (minutter * 60) #Finner resterende sekunder ved å trekke fra minutter

        print(f"Gjennomsnittlig samtaletid er {minutter} minutter og {sekunder} sekunder") #Printer resultat til skjerm. 

        input("Trykk ENTER for for å komme tilbake til menyen") #Pauser med input og henter meny etter trykk av enter. 
        meny()

# FUNKSJON SOM FINNER ANTALL HENVENDELSER PR TIDSROM VAL 5. I MENYEN
def tidsrom():

    if sjekk_data(): #Henter funksjon sjekkdata og går bare videre om denne er True

        #Bruker logical_and funksjonen i numpy til å finne intervaller i tid og regner ut sum.
        kl_08_10 = np.logical_and(kl_slett >= "08:00:00", kl_slett < "10:00:00" ).sum()
        kl_10_12 = np.logical_and(kl_slett >= "10:00:00", kl_slett < "12:00:00" ).sum()
        kl_12_14 = np.logical_and(kl_slett >= "12:00:00", kl_slett < "14:00:00" ).sum()
        kl_14_16 = np.logical_and(kl_slett >= "14:00:00", kl_slett < "16:00:00" ).sum()
        
        kaketekst = np.array(["KL 08-10", "KL 10-12", "KL 12-14", "KL 14-16"]) #Lager np-array med tekst til plott
        kakedata = np.array([kl_08_10, kl_10_12, kl_12_14, kl_14_16]) #Lager np-array med sum verdien som er hentet ut med logical_and

        # Tegner kakediagram og legger inn verdiene
        plt.pie(kakedata, labels=kaketekst, autopct='%1.1f%%') #Lager pie diagram med verdiene
        plt.title("Antall henvendelser fordelt pr. support-vakt")
        plt.legend(kakedata,
                    title="Faktisk antall",
                    loc="center left",
                    bbox_to_anchor=(0.9, 0.1, 0, 1))
        plt.show() #Vis pie

        meny() #Hent tilbake meny

# FUNKSJON SOM FINNER TILFREDSHET BLANDT KUNDENE VALG 6. I MENYEN
def tilfredshet():

    if sjekk_data(): #Henter funksjon sjekkdata og går bare videre om denne er True

        antall_fornoyd = 0 #Lager variabel for antall_fornoyd og setter til 0
        antall_misfornoyd = 0 #Lager variabel for antall_misfornoyd og setter til 0
        antall_svar = 0 #Lager variabel for antall_svar og setter til 0

        for verdi in score: #Bruker for loop til å gå igjennom alle oppføringer i np-array score
            if not np.isnan(verdi): #Hvis ikke den er tom så gå videre
                antall_svar += 1 #Legger til en på antall svar for å få antallet oppføringer som inneholder data.
                # 2 if setninger som teller de verdiene som er misfornøyd eller fornøyd
                if verdi <= 6:
                    antall_misfornoyd += 1
                if verdi > 8:
                    antall_fornoyd += 1

        prosent_fornoyd = round((antall_fornoyd * 100) / antall_svar) #Regner ut prosent med totalverdi / antall svar og runder av tallet til heltall
        prosent_misfornoyd = round((antall_misfornoyd * 100) / antall_svar) #Regner ut prosent med totalverdi / antall svar og runder av tallet til heltall

        nps = prosent_fornoyd - prosent_misfornoyd #Regner ut NPS ved å ta antall fornøyde prosent minus antall misfornøyde til variabel nps.

        print(f"Supportavdelingens NPS er {nps} %") #Printer ut resultatet.

        input("Trykk ENTER for for å komme tilbake til menyen") #Pauser med input og laster meny når bruker trykker enter
        meny()
   

# Her starter prgrammet. Gjør en try på å starte meny og har en exception hvis det trykkes ctrl+c
try:
    meny()
except KeyboardInterrupt:
    print("Program avbrutt")




#_____________________________________________________________________

# KILDER:
# https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
# https://www.digitalocean.com/community/tutorials/pandas-read_excel-reading-excel-file-in-python
# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_numpy.html
# https://www.geeksforgeeks.org/numpy/count-the-occurrence-of-a-certain-item-in-an-ndarray-numpy/
# https://www.w3schools.com/python/python_lists.asp
# https://www.w3schools.com/python/python_dsa_lists.asp
# https://numpy.org/doc/stable/reference/generated/numpy.max.html
# https://www.geeksforgeeks.org/python/numpy-logical_and-python/
# https://charlieojackson.co.uk/python/pandas-basics.php
# https://matplotlib.org/stable/gallery/pie_and_polar_charts/pie_features.html
# https://wiki.osdev.org/Terminals
# https://www.python.org/doc/essays/stdexceptions/
# 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
import csv
import re
import smtplib

from os import system

from bs4 import BeautifulSoup

from tkinter import *


def removefromlist(listmail, stringremove):
    listmail.remove(stringremove)


def deleteduplicates(listmail):
    return list(set(listmail))


def verificationmail(adresseverif):
    verif = re.match('^[_a-z0-9-]+(.[_a-z0-9-]+)*@[a-z0-9-]+(.[a-z0-9-]+)*(\.[a-z]{2,4})$', adresseverif)
    return verif


def verifurl(adresseurl):
    verif = re.match('^http://.*|https://.*', adresseurl)
    return verif


def lirecsv(nomfichier):
    try:
        tbLignes = []
        with open(nomfichier, 'r') as csvfile:
            lignes = csv.reader(csvfile, delimiter=';', quotechar='|')
            for ligne in lignes:
                tbLignes.append(ligne[0])
        return tbLignes
    except IOError:
        tbLignes = []
        with open(nomfichier, 'a+') as csvfile:
            lignes = csv.reader(csvfile, delimiter=';', quotechar='|')
            for ligne in lignes:
                tbLignes.append(ligne[0])
        return tbLignes


# Ecris à la suite du fichier csv, utile pour les imports
def ecrirecsvsuite(nomfichier, listmail):
    with open(nomfichier, 'a', newline='') as csvfile:
        lignes = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for mail in listmail:
            lignes.writerow([mail])


# Réecris tout le fichier csv utile pour supprimer les doublons et les email invalides
def reecrirecsv(nomfichier, listmail):
    with open(nomfichier, 'w', newline='') as csvfile:
        lignes = csv.writer(csvfile, delimiter=';')
        for mail in listmail:
            lignes.writerow([mail])


def getnddmail(mail):
    if verificationmail(mail) is not None:
        radical, ndd = mail.split('@')
        return ndd


# Retourne false si le ping est bon sinon retourne true
def pingndd(ndd):
    if system('ping ' + ndd + ' -a -n 1'):
        return False
    return True


def crawler(url):
    listcrawlmail = []
    code = requests.get(url)
    plain = code.text
    soup = BeautifulSoup(plain, "html.parser")
    for ligne in soup.find_all('a'):
        lien = ligne.get('href')
        if lien.__contains__('mailto:'):
            radical, mail = lien.split(':')
            if verificationmail(mail) is not None:
                listcrawlmail.append(mail)
    return listcrawlmail


# Fenètre envoi de mail
def fenetreTextMail(listBox):
    listItems = listBox.get(0, listBox.size())
    fenetreTextMail = Tk()
    fenetreTextMail.title('WebTarget: Envoi de mail')

    labelExp = Label(fenetreTextMail, text="Expéditeur")
    labelExp.pack()

    entreeExp = Entry(fenetreTextMail)
    entreeExp.pack()

    labelObj = Label(fenetreTextMail, text="Objet")
    labelObj.pack()

    entreeObj = Entry(fenetreTextMail)
    entreeObj.pack()

    labelMail = Label(fenetreTextMail, text="Mail")
    labelMail.pack()

    textMail = Text(fenetreTextMail, height=2, width=30)
    textMail.pack()

    boutonMail = Button(fenetreTextMail, text="Envoyer", command=lambda: fenetreSendMail(entreeExp.get(), entreeObj.get(), textMail.get("1.0", END), listItems, fenetreTextMail))
    boutonMail.pack()

    fenetreTextMail.mainloop()

# Fenètre d'import de Csv
def fenetreImportCsv(listBox, nomCsv):
    fenetreImportCsv = Tk()
    fenetreImportCsv.title('WebTarget: Import Csv')

    labelImportCsvNom = Label(fenetreImportCsv, text="Import de CSV")
    labelImportCsvNom.pack()

    entreeImportCsvNom = Entry(fenetreImportCsv, width=30)
    entreeImportCsvNom.pack()

    boutonImportCsv = Button(fenetreImportCsv, text="Valider", command=lambda: InsertMailCsv(listBox, nomCsv, entreeImportCsvNom.get()))
    boutonImportCsv.pack()
    fenetreImportCsv.mainloop()


# Fenètre d'import via Url
def fenetreImportUrl(listBox, nomCsv):
    fenetreImportUrl = Tk()
    fenetreImportUrl.title('WebTarget: Import Url')

    labelImportUrlNom = Label(fenetreImportUrl, text="Import via Url")
    labelImportUrlNom.pack()

    entreeImportUrlNom = Entry(fenetreImportUrl, width=30)
    entreeImportUrlNom.pack()

    boutonImportUrl = Button(fenetreImportUrl, text="Valider", command=lambda: InsertMailUrl(listBox, nomCsv, entreeImportUrlNom.get()))
    boutonImportUrl.pack()
    fenetreImportUrl.mainloop()


def fenetreSendMail(expediteur, objet, text, list, fenetreTextMail):
    fenetreTextMail.destroy()
    fenetreSendMail = Tk()
    fenetreSendMail.title('WebTarget: Envoi des mail')
    labelMail = Label(fenetreSendMail, text="EmailTest")
    labelMail.pack()
    entreeDestinataire = Entry(fenetreSendMail, width=30)
    entreeDestinataire.pack()
    boutonEnvoiMail = Button(fenetreSendMail, text="Envoi", command=lambda: envoyerMail(fenetreSendMail, expediteur, entreeDestinataire.get(), objet, text))
    boutonEnvoiMail.pack()
    boutonEnvoiMailGroup = Button(fenetreSendMail, text="Envoi a la liste", command=lambda: envoyerMail(fenetreSendMail, expediteur, "", objet, text, list))
    boutonEnvoiMailGroup.pack()
    fenetreSendMail.mainloop()

# Fenètre Principal
def fenetreMain():
    if entreeNom.get() != None :
        nomCsv = entreeNom.get()+".csv"
        tbMail = lirecsv(nomCsv)
        fenetreAcceuil.destroy()
        fenetreMain = Tk()
        listBoxMail = Listbox(fenetreMain)
        listBoxVerif = Listbox(fenetreMain)
        for ligne in tbMail:
            listBoxMail.insert(END, ligne)

        boutonDoublon = Button(fenetreMain, text="Doublon", command=lambda: updateListDoublon(listBoxMail, nomCsv))
        boutonDoublon.place(relx=0.05, rely=0.1)
        boutonSupprimer = Button(fenetreMain, text="Supprimer", command=lambda: deleteListSelected(listBoxMail, tbMail, nomCsv, listBoxVerif))
        boutonSupprimer.place(relx=0.35, rely=0.1)
        boutonVerification = Button(fenetreMain, text="Verifier", command=lambda: updateListVerif(listBoxMail, listBoxVerif))
        boutonVerification.place(relx=0.7, rely=0.1)
        boutonMainImportCsv = Button(fenetreMain, text="Import CSV", command=lambda: fenetreImportCsv(listBoxMail, nomCsv))
        boutonMainImportCsv.place(relx=0.1, rely=0.3)
        boutonMainImportUrl = Button(fenetreMain, text="Import URL", command=lambda: fenetreImportUrl(listBoxMail, nomCsv))
        boutonMainImportUrl.place(relx=0.6, rely=0.3)
        listBoxMail.place(relheight=0.4, relwidth=0.8, relx=0, rely=0.5)
        listBoxVerif.place(relheight=0.4, relwidth=0.2, relx=0.8, rely=0.5)

        boutonMain = Button(fenetreMain, text="Suite", command=lambda: fenetreTextMail(listBoxMail))
        boutonMain.place(relx=0.4, rely=0.9)
        fenetreMain.mainloop()


def updateListDoublon(listBox, nomCsv):
    listItems = listBox.get(0, listBox.size())
    listBox.delete(0, END)
    listNew = deleteduplicates(listItems)
    for ligne in listNew:
        listBox.insert(END, ligne)
    listBox.place(relheight=0.4, relwidth=0.8, relx=0, rely=0.5)
    reecrirecsv(nomCsv, listNew)


def updateListVerif(listBox, listBoxVerif):
    listItems = listBox.get(0, listBox.size())
    veriftbl = []
    listverifmail = []
    for ligne in listItems:
        veriftbl.append(verificationmail(ligne))

    for ligne in veriftbl:
        if ligne is not None:
            listverifmail.append(ligne[0])
        else:
            listverifmail.append('None')
    listBoxVerif.delete(0, END)
    for ligne in listverifmail:
        ndd = getnddmail(ligne)

        if ligne != 'None' and pingndd(ndd):
            listBoxVerif.insert(END, 'Ok')
        else :
            listBoxVerif.insert(END, 'Non')


def deleteListSelected(listBoxMail, tbMail, nomCsv, listBoxVerif):
    if listBoxMail.curselection():
        print(tbMail)
        selection = listBoxMail.curselection()
        print(selection)
        listBoxMail.delete(selection[0])

        tbMail.pop(selection[0])
        reecrirecsv(nomCsv, tbMail)
        listBoxVerif.delete(0, END)

def InsertMailUrl(listBox, nomCsv, url):
    listMailUrl = crawler(url)
    ecrirecsvsuite(nomCsv, listMailUrl)
    for ligne in listMailUrl:
        listBox.insert(END, ligne)
    listBox.place(relheight=0.4, relwidth=0.8, relx=0, rely=0.5)


def InsertMailCsv(listBox, nomCsv, nomCsvToImport):
    nomCsvToImport = nomCsvToImport + ".csv"
    listMailCsvImport = lirecsv(nomCsvToImport)
    ecrirecsvsuite(nomCsv, listMailCsvImport)
    for ligne in listMailCsvImport:
        listBox.insert(END, ligne)
    listBox.place(relheight=0.4, relwidth=0.8, relx=0, rely=0.5)


def envoyerMail(fenetreSendMail, email, destinataire,objet, message, list=0):
    # fenetreSendMail.destroy()

    if verificationmail(email) is not None and list == 0 and destinataire != 0:
        fromaddr = email
        toaddr = destinataire
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = objet

        body = message
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email, "snexon123")
        text = msg.as_string()
        server.sendmail(email, email, text)
        server.quit()
    elif verificationmail(email) is not None and list != 0 and destinataire != 0:
        for mail in list:
            print(mail)
            fromaddr = email
            toaddr = mail

            msg = MIMEMultipart()
            msg['From'] = fromaddr
            msg['To'] = toaddr
            msg['Subject'] = objet

            body = message
            msg.attach(MIMEText(body, 'plain'))

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email, "snexon123")
            text = msg.as_string()
            server.sendmail(email, email, text)
            server.quit()

# Fenètre d'acceuil
string = ""
fenetreAcceuil = Tk()
fenetreAcceuil.title('WebTarget')

labelNom = Label(fenetreAcceuil, text="nom du CSV")
labelNom.pack()

entreeNom = Entry(fenetreAcceuil, width=30)
entreeNom.pack()

boutonAcceuil = Button(fenetreAcceuil, text="Valider", command=fenetreMain)
boutonAcceuil.pack()
fenetreAcceuil.mainloop()



import os
import pyttsx3
import speech_recognition as sr
from datetime import datetime as dt
import exercice

#variable
nomAssistant = "Ivan"


engine = pyttsx3.init()
voix = engine.getProperty('voices')
engine.setProperty('voice', voix[0].id)

#lire mesage à haute voix


def say(audio):
    engine.say(audio)
    engine.runAndWait()

#recuperer message


def listen(mess):
    r = sr.Recognizer()
    with sr.Microphone() as source:
        if mess!="":
            print(mess)
            say(mess)
        r.adjust_for_ambient_noise(source)
        r.dynamic_energy_threshold=3000
        record = r.listen(source, timeout=10.0)

    try:
        print("en analyse....")
        parole = r.recognize_google(record, language="fr-FR")
        print("vous avez dit: "+parole)
    except Exception as e:
        print("Demande mal formulée. Je vous écoute...")
        #say("Demande inconnue ")
        return ""
    return parole.lower()


def username():
    print("Comment puis-je vous appeler?")
    say("Comment puis-je vous appeler?")
    name = listen("")
    say("Bienvenu")
    say(name)


def salutation():
    if(0 <= dt.today().hour <= 12):
        say("bonjour à vous")
    else:
        say("bonsoir à vous")
    say("Je suis votre assistant "+nomAssistant)


#debut du script
if __name__ == "__main__":
    clear=lambda: os.system('cls')
    clear()
    salutation()

    while True:
        #recupere la demande
        desir = listen("quel est votre désir?")
        while desir=="":
            desir = listen("")
        print(desir)
        #pour quitter
        if "quitter" in desir:
            say("aurevoir!")
            break

        #ce que Ivan peux faire
        elif "que peux-tu faire" in desir:
            print("1-afficher cv par compétences")
            print("2-Me mettre en pause")
            print("3-quitter")
            print("\n")

        #pour recuperer les cv par compétences
        elif ("affiche" or "afficher") in desir or ("cv" in desir and "compétences" in desir):
            #liste of compétences
            compet=["gestion","informatique","comptabilité","dévéloppement web","dévéloppement application","économie",
                    "marketing", "photoshop", "office", "programmeur web", "Graphic design"]
            mess = "quelle compétence recherchez vous?"
            print(mess)
            say(mess)
            dire=listen("")
            rep = "non"
            while rep!="oui":
                print("choix")
                say("vous aviez choisi "+dire+" confirmer vous votre choix oui ou non ?")
                rep=listen("")
                if "oui" not in rep:
                    print(mess)
                    say(mess)
                    dire = listen("")
            if "et" in dire:
                d=dire.split("et")
            else:
                d=dire.split()
            competences=[]
            for i in d:
                if "l'" in i:
                    val=i.replace("l'","")
                    competences.append(val)
                elif "la" in i:
                    val=i.replace("la","")
                    competences.append(val)
                elif "le" in i:
                    val = i.replace("le", "")
                    competences.append(val)
                else:
                    competences.append(i)
            
            if competences!=[]:
                print(competences)
                if any(item in competences for item in compet):
                    r=exercice.classerCV(competences)
                    proche=r[0]
                    loin=r[1]
                    print("CV proche:",len(proche))
                    print("CV distant:",len(loin))
                    if proche!=[]:
                        for x in proche:
                            os.system("start powerpnt.exe {}".format(x))
                    if loin!=[]:
                        for x in loin:
                            os.system("start powerpnt.exe {}".format(x))
                else:
                    print("aucun CV du genre disponible.")
            else:
                print("désoler réessayer.")
                say("désolé! veuiller réessayer.")
                
        elif "pause" in desir:
            os.system("pause")

        elif "ouvre" in desir or "ouvrir" in desir:
            if("powerpoint" in desir):
                os.system("start powerpnt.exe")
            elif("excel" in desir):
                os.system("start excel.exe")
            elif("word" in desir):
                os.system("start winword.exe")
              

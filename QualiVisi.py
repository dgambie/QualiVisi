"""
    PROGRAMME   : QualiVis
    BUT         : Permet d'analyse les defect.xml Pour en sortir differentes info
                Exemple :   Top défaut par carte par repere Topo
                            Top défaut par carte par code article
                            Top défaut par carte par code défaut
                            etc....
    PAR         : D.GAMBIE
    LE          : 12/12/19
    VERSION     : 1.0
    COMMENTAIRE : 
    
    ### HISTORIQUE : ############################################################
    #1.0    : Création.                                                         #
    #                                                                           #
    #############################################################################
    
    Remarque :  -Ajouter une méthode de tri par code article => OK
                -Ajouter une méthode de tri par code défaut => ok
                -Ajouter la fonction permettant la multi-Analyse
                -Ajouter une focntion permettant la multi analyse par client / éventuellement tous les clients confondus

                Trie interessant a faire : => ok
                Nombre de défaut par code article
                Type de défaut par code article
                Nombre de type de défaut par code article

                Voir pour ajouter un mode liste pour les méthode d'extraction => peut-être plus facile a traiter par la suite
                    *liste de code article en défaut => ok
                    *liste de code defaut => ok
                    *liste de repere en défaut => ok

                Voir pour ajouter une metode/fonction pour generer les listes en fonction de filtre de code défaut => ok
                Voir pour ajouter un filtre client a la fonction Recuperer_liste_xml => permettra dans lancer des analyses pas client  => filtre client dans les parametres
                Voir pour intérroger la BDD SAGE pour récupérer des info à partir de l'OF
                Voir pour sortir une synthèse des faits marquant (paretto ?) => fait marquant basique intégré => voir pour améliorer
                Voir pour générer un fichier rapport
                Voir pour borner les entry
                Voir pour gérer le redimensionnement de le textBox => pas utile
                Voir pour faire un filtre par repere topo => filtre dans le mode avancé
                Voir pour ajouter un mode avancée pour l'analyse => ok
                

            Le 13/01/2020 = passage de liFichier à self.liFichier => permet d'avoir accès à la liste des fichiers dans toutes la classe Application
            
    Amélioration 
        -Améliorer les faits marquants : ajouter le total de défauts : le nombre moyen  de défaut par cartes voir pour les 80-20....
        -Ajouter un parcours des fichiers rapports pour ouvrir les rapports
                        -> soit en pdf si la localisation archive et visitée
                        -> soit le html si la localisaton courante et visitée





"""
#Liste des imports
import xml.etree.ElementTree as ET
import os
from collections import Counter
import tkinter as tk
import tkinter.ttk as ttk#pour les widget supplémentaire comme les combos box
import tkinter.font as tkFont
import tkinter.messagebox as tkmsg
import tkinter.filedialog as tkfld
import matplotlib.pyplot as plt #pip install matplotlib
import numpy as np #pip install numpy
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.offsetbox import AnchoredText

#================================================CLASS APPLICATION===========================================================
class Application (tk.Tk):
    """Classe de la fenetre principale
    Présentation des méthode :
    init => initialisation de la class 
            Variables disponibles :
                -varCheminMv3 = Chemin de base pour la machine MV3 => données dans le fichier config.txt.
                    Peut être redéfinie dans les paramêtres
                -varCheminMv7 = Chemin de base pour la machine MV7 => données dans le fichier config.txt
                    Peut être redéfinie dans les parametres
                -varCheminAutre = Chemin de base pour utilisateur => données dans le fichier config.txt
                    Peut être redéfinie dans les parametres
                -VarMachine = Machine à analyser (permet de savoir vers quel repertoire pointer, MV3 / MV7 ou Autre)
                -VarClient = Permet de filtrer la recherche par client => peut etre définie dans les parametres
                -size = taille de la textBox
                -pdx/pdy = marge pour les frames et bouton OK
                -parcourir = variable pour savoir si le dossier à déjà était parcouru
                -tailleicone = taille des icones bouton avec des images
                -varCompteur = pour le nombre du top défaut du fait marquant
                -modificationOF = pour savoir si l'utilisateur à changé l'OF
    
    Creation_frame => Création des frames structurelle (tête / corps / pied)
    Creation_menu => création de la barre de menu
    Creation_widget => création des widgets dans la fenetre principale
    Analyse => méthode lancée lors de l'appuis sur le bouton OK (ou sur l'event <Return>)
    Clear_TxtBox => Pour effacer la textBow => lancée par le bouton Clear !
    Parametre => Pour la gestion des paramêtres (cheminMv3 / Mv7 / Autre ; choix de la machine ; filtre client)
    Definition_chemin => permet de définir les chemins de base => lecture et écriture dans le fichier config
    Parcourir => pour naviguer dans l'explorateur et choisir le dossier d'analyse (chemin de départ défini dans le fichier confg.txt)
    """
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("QualiVisi")
        self.iconbitmap("IMG\logo.ico")
        self.policeLabelFrame = tkFont.Font(weight = 'bold', size = '10')
        self.size = 50 #Taille de la TextBox
        self.pdx = 5 #Marge
        self.pdy = 5
        self.tailleIcone = 25
        self.varCompteur = tk.IntVar() #Compteur TOP Fait marquant
        self.varCompteur.set(3)
        self.varCompteurCodeDefaut = tk.IntVar()#Compteur Top code défaut
        self.varCompteurCodeDefaut.set(1)
        self.parcourir = False
        self.modificationOF=False
        self.varCheminMv3 = tk.StringVar()
        self.varCheminMv7 = tk.StringVar()
        self.varCheminAutre = tk.StringVar()
        self.varMachine = tk.IntVar()
        self.varMachine.set(0)
        self.varClient = tk.StringVar()
        self.varClient.set('aucun')
        self.Definition_chemin()
        self.Creation_frame()
        self.Creation_menu()
        self.Creation_widget()

    def A_propos(self):
        tkmsg.showinfo('A propos','Qualivisi V1\n Logiciel d\'analyse de défaut AOI\n28/01/2020 - DG')

    def Aide (self):
        os.startfile("aide.pdf")
    
    def Aide_Code_defaut(self):
        """Aide pour connaitre la définition des codes défaut :
        'ANCC','AMTC','BNS','AINC','BNC','BPS','AMTT','BPV','AMIT','AMIC','ANCT','AINT'"""
        message = 'Liste des codes défauts :\n'
        message+= 'ANC* : Composant non conforme => marquage / placement \n'
        message+= 'AMT* : Composant manquant => manquant / placement \n'
        message+= 'AIN* : Composant mal polarisé => marquage / placement \n'
        message+= 'AMI* : Composant mal implanté => manquant / placement \n'
        message+= 'BN* : Composant non soudé \n'
        message+= 'BP* : Pont de soudure\n'

        tkmsg.showinfo('Code Defaut',message)
        
#''''''''''''''''''''''''''''''
    def Creation_frame(self):
        """Methode pour la création des Frames
        Même model que pour les html
            Header = haut
            Body = milieu
            Footer = bas"""

        self.frmHeader = tk.Frame(self)
        self.frmBody = tk.Frame(self)
        self.frmFooter = tk.Frame(self)
        self.frmHeader.pack()
        self.frmBody.pack()
        self.frmFooter.pack() 

#''''''''''''''''''''''''''''''
    def Creation_menu(self):
        """Methode pour la création du menu"""

        self.menuPrincipal = tk.Menu()
        self.sousMenuAide = tk.Menu(self.menuPrincipal)
        self.menuPrincipal.add_command(label = 'Paramètre', command = self.Parametre)
        self.menuPrincipal.add_command(label = 'Mode Avancé', command = self.Mode_avance, state = 'disabled')
        self.menuPrincipal.add_cascade(label = 'Aide', menu = self.sousMenuAide)
        self.sousMenuAide.add_command(label = 'A propos', command = self.A_propos)
        self.sousMenuAide.add_command(label = 'Aide', command = self.Aide)
        self.menuPrincipal.add_command(label = 'Quitter', command = self.quit)
        self.config(menu = self.menuPrincipal)

#''''''''''''''''''''''''''''''
    def Creation_widget(self):
        """Méthode pour la création des widgets dans la fenetre"""

        #----------Création des label Frames pour organiser la page
        self.lblFrDonnee = tk.LabelFrame(self.frmBody, text = 'Dossier à analyser',padx = self.pdx, pady = self.pdy,font = self.policeLabelFrame ) #Pour contenir entry OF et le bouton OK
        self.lblFrOption = tk.LabelFrame(self.frmBody,text = 'Option',padx = self.pdx, pady = self.pdy,font = self.policeLabelFrame) #Pour contenur le filtre et le type de tri
        self.lblFrInfo = tk.LabelFrame(self.frmBody,text = 'Info',padx = self.pdx, pady = self.pdy,font = self.policeLabelFrame)#Pour contenir la carte, le prog et le nombre de rapport
        self.lblFrTopDefaut = tk.LabelFrame(self.frmBody,text = 'Top défaut',padx = self.pdx, pady = self.pdy,font = self.policeLabelFrame)#Pour contenir une textbox avec les défauts

        #-----------Placement des labels Frames
        self.lblFrDonnee.grid(column = 0, row = 0, sticky = 'n')
        self.lblFrOption.grid(column = 1,row = 0, sticky = 'n')
        self.lblFrInfo.grid(column = 0,row = 1, sticky = 's')
        self.lblFrTopDefaut.grid(column = 0, row = 2,columnspan = 2, sticky = 's')

        #---------------------------------LabelFrame Donnée
        self.varChemin = tk.StringVar()
        self.varChemin.set('... parcourir')
        self.imgOpenFolder = tk.PhotoImage(file = 'IMG\openFolder.png')
        
        self.etrChemin = tk.Entry(self.lblFrDonnee, textvariable = self.varChemin)
        self.btnParcourir = tk.Button(self.lblFrDonnee, text = '...', command = self.Parcourir)
        self.btnParcourir.config(image = self.imgOpenFolder, width = self.tailleIcone, height = self.tailleIcone)
        self.etrChemin.grid(column = 0, row = 0)
        self.btnParcourir.grid(column = 1, row = 0)

        #--------------------------------Btn OK -> dans le frame body
        self.btnOK = tk.Button(self.frmBody ,text = 'OK', command = self.Analyse, width = 20, height = 3, padx = self.pdx , pady = self.pdy, font = self.policeLabelFrame)
        self.btnOK.grid(column=1,row = 1)


        #--------------------------------LabelFrame Oprtion
        self.listeCodeDefaut = ['aucun','ANCC','AMTC','BNS','AINC','BNC','BPS','AMTT','BPV','AMIT','AMIC','ANCT','AINT']
        self.listeType = ['article','repere','codeDefaut']
        self.varOf = tk.StringVar()
        self.varOf.set('aucun')
        self.varFaitMarquant = tk.IntVar()
        self.varFaitMarquant.set(1)

        self.lblFiltre = tk.Label(self.lblFrOption,text = 'Filtre')
        self.lblType = tk.Label(self.lblFrOption,text = 'Type')
        self.lblOf = tk.Label(self.lblFrOption, text = 'OF')
        self.cmbFiltre = ttk.Combobox(self.lblFrOption,value = self.listeCodeDefaut)
        self.cmbType = ttk.Combobox(self.lblFrOption, value = self.listeType)
        self.etrOf = tk.Entry(self.lblFrOption, textvariable = self.varOf, validatecommand = self.register(self.Modification_OF), validate = 'all')
        self.cmbFiltre.current(0)
        self.cmbType.current(0)
        self.btnAideFiltre = tk.Button(self.lblFrOption,text = "?",command = self.Aide_Code_defaut, relief = 'flat')
        self.chkFaitMarquant = tk.Checkbutton(self.lblFrOption, text = 'Fait marquant', variable = self.varFaitMarquant)
        self.lblFiltre.grid(column = 0,row = 0)
        self.lblType.grid(column = 0, row = 1)
        self.lblOf.grid(column = 0, row = 2)
        self.cmbFiltre.grid(column = 1, row = 0)
        self.cmbType.grid(column = 1, row = 1)
        self.etrOf.grid(column = 1, row = 2)
        self.chkFaitMarquant.grid(column = 0, row = 3, columnspan = 3)
        self.btnAideFiltre.grid(column = 2, row = 0)

        #-----------------------------LabelFrame Info
        self.varCarte = tk.StringVar()
        self.varCarte.set("Carte : NC")
        self.varNbRapport = tk.StringVar()
        self.varNbRapport.set("Nombre de rapport : NC")

        self.lblCarte = tk.Label(self.lblFrInfo,textvariable = self.varCarte)
        self.lblNbRapport = tk.Label(self.lblFrInfo,textvariable = self.varNbRapport)
        self.lblCarte.grid(column = 0, row = 0)
        self.lblNbRapport.grid(column = 0, row = 1)
        

        #------------------------------LabelFrame TopDefaut
        self.txtTopDefaut = tk.Text(self.lblFrTopDefaut, width = self.size)
        self.txtTopDefaut.grid(column = 0, row = 0,)

        #-------------------------------Bouton Clear
        self.btnClear = tk.Button(self.frmFooter,text = 'Clear !', command = self.Clear_TxtBox)
        self.btnClear.grid(column = 0, row = 0)

#''''''''''''''''''''''''''''''
    def Analyse (self):
        """Métode pour analyser les XML et sortir les top défaut dans la textbox
        le chemin de recherche pour etablir la liste des fichiers a analyser est défini par l'entrybox parcourir      
        """
        
        strOf = self.varOf.get()
        strFiltre = self.cmbFiltre.get()
        strType = self.cmbType.get()
        monXml = None
        erreur = False
        
        #Mise en forme de l'OF
        if strOf !='aucun':
            strOf = strOf.upper()
            strOf = strOf.replace('/','-')
            self.varOf.set(strOf)
        if len(strOf) <= 0:
            strOf = 'aucun'

        print ("Of : {}, filtre : {}, type de top défaut : {}".format(strOf,strFiltre,strType))

        #Si on change le repertoire de recher ou si on change l'OF 
        if self.parcourir or self.modificationOF :
            #Récupération de la liste des fichiers
            self.liFichier =  (Recuperer_liste_xml(self.chemin,of=strOf))
        else :
            print ('Données déjà en mémoire')

        retour = "====== Tri par " + strType + "|Filtre : " + strFiltre +  " =========\n"
        self.txtTopDefaut.insert(tk.END,retour)
        retour =  Counter(Liste_defaut(self.liFichier,typeListe=strType ,filtre = strFiltre)).most_common()
        self.txtTopDefaut.insert(tk.END,retour)
        self.txtTopDefaut.insert(tk.END,"\n\n")

        #Récupération des info
        try :
            monXml = XML(self.liFichier[0])
        except IndexError :
            tkmsg.showerror("Erreur !", "Pas de fichier XML trouvé")
            self.txtTopDefaut.insert(tk.END,'Erreur .... pas fichier XML ... \n\n')
            erreur = True

        if monXml != None:
            monXml2 = XML(self.liFichier[-1])
            if monXml.reference[:-4] == monXml2.reference[:-4]: #Pour vérifier si le premier et le dernier xml sont bien la meme reférence de carte
                self.varCarte.set('Carte : ' + monXml.reference[:-4])

            else :
                self.varCarte.set('carte : MULTI')
            self.varNbRapport.set('Nombre de rapport : ' + str(len(self.liFichier)) )
        
        #Fait mraquant
        if self.varFaitMarquant.get() == 1 and erreur == False:
            self.FaitMarquant()

        #Positionnement des variable et activation du mode Avance
        self.parcourir = False #mise à false du parcourir pour la prochaine analyse
        self.modificationOF = False #Pour la prochaine analyse
        self.menuPrincipal.entryconfig(2,state = 'normal')

#''''''''''''''''''''''''''''''
    def Clear_TxtBox(self):
        "Méthode pour vider la text box"
        self.txtTopDefaut.delete(1.0,tk.END)

#''''''''''''''''''''''''''''''
    def Parametre(self):
        """Méthode pour la gestion de parametres"""
        #Création d'une nouvelle fenetre
        self.fenetreParametre = tk.Toplevel()
        self.fenetreParametre.title('Parametre')
        self.fenetreParametre.focus_set()

        #********Création des labelsframes
        self.lblFrMachine = tk.LabelFrame(self.fenetreParametre,text = 'Machine', padx = self.pdx, pady = self.pdy, font = self.policeLabelFrame)
        self.lblFrFiltreClient = tk.LabelFrame(self.fenetreParametre, text = 'Filtre Client', padx = self.pdx, pady = self.pdy, font = self.policeLabelFrame)
        self.lblFrAutre = tk.LabelFrame(self.fenetreParametre, text = 'Autre', padx = self.pdx, pady = self.pdy, font = self.policeLabelFrame)
        self.frmValidation = tk.Frame(self.fenetreParametre,padx = self.pdx, pady = self.pdy)

        self.lblFrMachine.grid(column = 0, row = 0,sticky = 'n')
        self.lblFrFiltreClient.grid(column = 0, row = 1, sticky = 'w')
        self.lblFrAutre.grid(column = 1, row = 0, sticky = 'n')
        self.frmValidation.grid(column = 0, row = 2, columnspan = 2)

        #********Création des widgets

        self.rdMV3 = tk.Radiobutton(self.lblFrMachine,text = 'MV3',variable = self.varMachine, value = 0)
        self.rdMV7 = tk.Radiobutton(self.lblFrMachine,text = 'MV7',variable = self.varMachine, value = 1)
        self.rdAutre = tk.Radiobutton(self.lblFrMachine,text = 'autre',variable = self.varMachine, value = 2)
        self.etrCheminMv3 = tk.Entry(self.lblFrMachine, textvariable = self.varCheminMv3)
        self.etrCheminMv7 = tk.Entry(self.lblFrMachine, textvariable = self.varCheminMv7)
        self.etrChemiAutre = tk.Entry(self.lblFrMachine, textvariable = self.varCheminAutre)
        self.rdMV3.grid(column = 0, row = 0)
        self.rdMV7.grid(column = 0, row = 1)
        self.rdAutre.grid(column = 0, row = 2)
        self.etrCheminMv3.grid(column = 1, row = 0)
        self.etrCheminMv7.grid(column = 1, row = 1)
        self.etrChemiAutre.grid(column = 1, row = 2)

        self.etrClient = tk.Entry(self.lblFrFiltreClient,textvariable = self.varClient)
        self.etrClient.grid(column = 0, row = 0)

        self.chkSortie = tk.Checkbutton(self.lblFrAutre,text = 'Generer fichier de sortie')
        self.etrCompteur = tk.Entry(self.lblFrAutre, textvariable = self.varCompteur, width = 3)
        self.lblCompteur = tk.Label(self.lblFrAutre, text = 'Nb TOP fait marquant')
        self.etrCompteurCodeDefaut = tk.Entry(self.lblFrAutre,textvariable = self.varCompteurCodeDefaut, width = 3)
        self.lblCompteurCodeDefaut = tk.Label(self.lblFrAutre, text = 'Nb TOP Code défaut')
        self.etrCompteur.grid(column = 0, row = 1)
        self.lblCompteur.grid(column = 1, row = 1)
        self.etrCompteurCodeDefaut.grid(column = 0, row = 2)
        self.lblCompteurCodeDefaut.grid(column = 1, row = 2)
        self.chkSortie.grid(column = 0, row = 0, columnspan = 2)

        self.btnAppliquer = tk.Button(self.frmValidation,text = 'Appliquer', command = lambda : self.Definition_chemin(mode = 'ecriture'))
        self.btnAnnuler = tk.Button(self.frmValidation, text = 'Annuler / Quitter', command = self.fenetreParametre.destroy)
        self.btnAppliquer.grid(column = 0, row = 0)
        self.btnAnnuler.grid(column = 1, row = 0)

#''''''''''''''''''''''''''''''
    def Definition_chemin(self, mode = 'lecture'):
        """Méthode pour gérer les chemins des xml
        doit exister un fichier config a la racine du script avec :
            1er ligne chemin MV3
            2eme ligne chemin MV7
            3eme ligne chemin debug"""

        if mode == 'lecture':
            with open("config.txt", "r") as fichier:
                contenu = fichier.read()
                licontenu = contenu.split("\n")
                self.varCheminMv3.set (licontenu[0])
                self.varCheminMv7.set (licontenu[1])
                self.varCheminAutre.set (licontenu[2])

        if mode == 'ecriture':
            with open("config.txt", "w") as fichier:
                contenu = self.varCheminMv3.get() + '\n' + self.varCheminMv7.get() + '\n' + self.varCheminAutre.get()
                fichier.write(contenu)

#''''''''''''''''''''''''''''''
    def Parcourir (self):
        """Methode pour parcourir et selectionner le dossier de recherche"""
        #Determination du chemin en fonction des parametres utilisateurs :
        #cheminInitial = '/'
        if self.varChemin.get() == '... parcourir':
            if self.varMachine.get() == 0:
                cheminInitial = self.varCheminMv3.get()
            if self.varMachine.get() == 1:
                cheminInitial = self.varCheminMv7.get()
            if self.varMachine.get() == 2:
                cheminInitial = self.varCheminAutre.get()
            if self.varClient.get() != 'aucun' :
                cheminInitial += self.varClient.get() + '/'
        else:
            cheminInitial = self.varChemin.get()

        self.chemin = tkfld.askdirectory(title = 'Choisisser un repertoire', initialdir = cheminInitial )
        self.varChemin.set(self.chemin)
        print (self.chemin)
        self.parcourir = True #Basculement a true pour indiquer que le parcourir à été sélectionné

#''''''''''''''''''''''''''''''
    def FaitMarquant_ (self):
        """Méthode pour générer les fait marquant
        Top 3 défaut
            Par article
            Par repere
            Par défaut
            
        Top 3 défaut des 3 codes défaut"""
        liTopArticle = []
        liTopRepere = []
        liTopCodeDefaut = []
        liDefautFiltre = []
        #varCompteur = 3
        message = ''

        #Vérification de l'entry
        try : 
            int(self.varCompteur.get())
        except tk.TclError:
            print ('Mauvaise valeur du compteur ')
            tkmsg.showerror('Erreur input', 'Parametre compteur Top  fixé à 3')
            self.varCompteur.set(3)

        #Génération des 3 listes pour traitement
        liArticle = Counter(Liste_defaut(self.liFichier,typeListe= 'article' ,filtre = 'aucun')).most_common()
        liRepere = Counter(Liste_defaut(self.liFichier,typeListe= 'repere' ,filtre = 'aucun')).most_common()
        liCodeDefaut = Counter(Liste_defaut(self.liFichier,typeListe= 'codeDefaut' ,filtre = 'aucun')).most_common()


        #Top 3 des défauts :
        for i in range(self.varCompteur.get()):
            try :
                liTopArticle.append(liArticle[i])
            except IndexError :
                print ("Plus de défaut article")

            try :
                liTopRepere.append(liRepere[i])
            except IndexError :
                print ("Plus de défaut repere")

            try :
                liTopCodeDefaut.append(liCodeDefaut[i])
            except IndexError :
                print ("Plus de défaut article")

        #print (liTopArticle,liTopCodeDefaut,liTopRepere)


        #-----Lecture et mise en forme du TOP 3------
        ############################################
        for cle,value in liTopRepere :
            codeArticle = 'NC'
            #liDefautFiltre = []
            strListeCodeDefautRepere = ''
            #Recuperation du code article pour affichage
            for fichier in self.liFichier:
                xml=XML(fichier)
                liDefaut = xml.Recuperer_defaut()
                for defaut in liDefaut:
                    if defaut.repereTopo == cle:
                        codeArticle = defaut.codeArticle
                        #liDefautFiltre.append(defaut)#Ajout du défaut pour extraction des codes défauts
                        #break

            

            message += 'Repere : ' + cle + ' (' + codeArticle  + ') => ' + str(value) + ' Défaut(s)'  +'\n'
        message += '\n'
        #Ectraction des codes défauts:
        #liCodeDefautFiltre = Liste_defaut_par_type(liDefautFiltre,typeListe='codeDefaut')[:]
        #liCodeDefautFiltre = Counter(liCodeDefautFiltre).most_common()
        #print ("liCodeDefautFiltre : ", liCodeDefautFiltre)

        ###########################################
        for cle,value in liTopCodeDefaut :
            message += 'Code Défaut : ' + cle + ' => ' + str(value) + ' Défaut(s)' + '\n'
        message += '\n'

        ###########################################
        for cle,value in liTopArticle :
            message += 'Article : ' + cle + ' => ' + str(value) + ' Défaut(s)' +'\n' 
        message += '\n'
        
        #self.Fait_Marquant_()
        tkmsg.showinfo('Fait marquant', message)

    def FaitMarquant(self):
        """Nouvelle méthode pour les faits marquants
        Voir pour se servir de la methode Rafraichir_Graphique"""

        liTopArticle = []
        liTopRepere = []
        liTopCodeDefaut = []
        liDefaut = []
        liDefautFiltreArticle = []
        liDefautFiltreRepere = []
        #varCompteur = 3
        message = ''

        #Génération des 3 listes pour traitement
        liArticle = Counter(Liste_defaut(self.liFichier,typeListe= 'article' ,filtre = 'aucun')).most_common()
        liRepere = Counter(Liste_defaut(self.liFichier,typeListe= 'repere' ,filtre = 'aucun')).most_common()
        liCodeDefaut = Counter(Liste_defaut(self.liFichier,typeListe= 'codeDefaut' ,filtre = 'aucun')).most_common()

        #Récupération des la liste des défauts
        for fichier in self.liFichier:
            xml = XML(fichier)
            for element in xml.Recuperer_defaut():
                liDefaut.append(element)  

        #Top 3 des défauts :
        for i in range(self.varCompteur.get()):
            try :
                liTopArticle.append(liArticle[i])
            except IndexError :
                print ("Plus de défaut article")

            try :
                liTopRepere.append(liRepere[i])
            except IndexError :
                print ("Plus de défaut repere")

            try :
                liTopCodeDefaut.append(liCodeDefaut[i])
            except IndexError :
                print ("Plus de défaut article")

        #Recupération de la distribution des codes défauts par repere et par article

        for repere, nombre in liTopRepere:
            for defaut in liDefaut:
                if defaut.repereTopo == repere:
                    article = defaut.codeArticle
                    liDefautFiltreRepere.append(defaut)

            liCodeDefautFiltreRepere = Liste_defaut_par_type(liDefautFiltreRepere,typeListe='codeDefaut')[:]
            liCodeDefautFiltreRepereTrie = Counter(liCodeDefautFiltreRepere).most_common()
            liCodeDefautFiltreRepereTrieTop = []
            for i in range(self.varCompteurCodeDefaut.get()):
                try :
                    liCodeDefautFiltreRepereTrieTop.append(liCodeDefautFiltreRepereTrie[i])
                except IndexError :
                    print ("Moins de trois codes défaut")

            #print ("Repere : {} ({}), nombre de défaut : {}, liste de code défaut : {}".format(repere,article,nombre,liCodeDefautFiltreRepereTrie))
            message += "Repere : {} ({}) => {} défaut(s), TOP: {}\n".format(repere,article,nombre,liCodeDefautFiltreRepereTrieTop)

        message += '\n'
        for code, nombre in liTopCodeDefaut:
            #print ("Code défaut : {}, nombre : {}".format(code,nombre))
            message += "Code défaut : {} => {} défaut(s)\n".format(code,nombre)

        message += '\n'
        for article, nombre in liTopArticle:
            for defaut in liDefaut:
                if defaut.codeArticle == article:
                    liDefautFiltreArticle.append(defaut)

            liCodeDefautFiltreArticle = Liste_defaut_par_type(liDefautFiltreArticle,typeListe='codeDefaut')[:]
            liCodeDefautFiltreArticleTrie = Counter(liCodeDefautFiltreArticle).most_common()
            liCodeDefautFiltreArticleTrieTop = []
            for i in range(self.varCompteurCodeDefaut.get()):
                try :
                    liCodeDefautFiltreArticleTrieTop.append(liCodeDefautFiltreArticleTrie[i])
                except IndexError :
                    print ("Moins de trois codes défauts")
            #print ("Article : {}, nombre de défaut : {}, liste de code défaut : {}".format(article,nombre,liCodeDefautFiltreArticleTrie))
            message += "Article : {} => {} défaut(s), TOP : {}\n".format(article,nombre,liCodeDefautFiltreArticleTrieTop)

        print (message)
        tkmsg.showinfo('Fait marquant', message)










#'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    def Modification_OF(self):
        """Méthode permettant de savoir si l'OF à été modifié
        Utilisation des la propriete validate du widget"""
        self.modificationOF = True
        return True

#''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    def Mode_avance(self):
        """Méthode pour afficher le mode avancé
        Permet d'analyser plus en finesse"""

        liFiltreArticle = []
        liFiltreRepere = []
        liFiltreCodeDefaut = []
        
        self.varTypeFiltre = tk.StringVar()
        self.varTypeFiltre.set('repere')
        self.varNombreElement = tk.StringVar()
        self.varNombreElement.set(5)

        liFiltreRepere.append('tous')
        liFiltreArticle.append('tous')
        liFiltreCodeDefaut.append('tous')
        

        #Récupération des info pour traitement avancé
        liArticleAvance = Counter(Liste_defaut(self.liFichier,typeListe= 'article' ,filtre = 'aucun')).most_common()
        liRepereAvance = Counter(Liste_defaut(self.liFichier,typeListe= 'repere' ,filtre = 'aucun')).most_common()
        #liCodeDefautAvance = Counter(Liste_defaut(self.liFichier,typeListe= 'codeDefaut' ,filtre = 'aucun')).most_common()

        for cle, value in liArticleAvance:
            liFiltreArticle.append(cle)
        for cle, value in liRepereAvance:
            liFiltreRepere.append(cle)
        for element in self.listeCodeDefaut:
            if element != 'aucun':
                liFiltreCodeDefaut.append(element)





        #Création d'une nouvelle fenetre
        self.fenetreModeAvance = tk.Toplevel()
        self.fenetreModeAvance.title('Mode avancé')
        self.fenetreModeAvance.focus_set()

        #********Création des labelsframes
        self.lblFrFiltreAvance = tk.LabelFrame(self.fenetreModeAvance,text = 'Filtre', padx = self.pdx, pady = self.pdy, font = self.policeLabelFrame)
        self.lblFrTypeAvance = tk.LabelFrame(self.fenetreModeAvance, text = 'Type', padx = self.pdx, pady = self.pdy, font = self.policeLabelFrame)
        self.lblFrAutreAvance = tk.LabelFrame(self.fenetreModeAvance, text = 'Autre', padx = self.pdx, pady = self.pdy, font = self.policeLabelFrame)
        self.lblFrFiltreAvance.grid(column = 0, row = 0, sticky = 'n')
        self.lblFrTypeAvance.grid(column = 1, row = 0, sticky = 'n')
        self.lblFrAutreAvance.grid(column = 2, row = 0, sticky = 'n')

        #*********Création des widgets
        ########## FRAME Filtre#############
        self.lblCodeDefautAvance = tk.Label(self.lblFrFiltreAvance, text = 'Code Défaut :')
        self.lblArticleAvance = tk.Label(self.lblFrFiltreAvance, text = 'Article :')
        self.lblRepereAvance = tk.Label(self.lblFrFiltreAvance, text = 'Repère :')
        self.lblClientAvance = tk.Label(self.lblFrFiltreAvance, text = 'Client :')
        self.cmbCodeDefautAvance = ttk.Combobox(self.lblFrFiltreAvance,value = liFiltreCodeDefaut)
        self.cmbArticleAvance = ttk.Combobox(self.lblFrFiltreAvance,value = liFiltreArticle)
        self.cmbRepereAvance = ttk.Combobox(self.lblFrFiltreAvance, value = liFiltreRepere)
        self.cmbRepereAvance.current(0)
        self.cmbArticleAvance.current(0)
        self.cmbCodeDefautAvance.current(0)
        #Collage
        self.lblCodeDefautAvance.grid(column = 0, row = 0, sticky = 'e')
        self.cmbCodeDefautAvance.grid(column = 1, row = 0)
        self.lblArticleAvance.grid(column = 0, row = 1, sticky = 'e')
        self.cmbArticleAvance.grid(column = 1, row = 1)
        self.lblRepereAvance.grid(column = 0, row = 2, sticky = 'e')
        self.cmbRepereAvance.grid(column = 1, row = 2)
        #self.lblClientAvance.grid(column = 0, row = 3)

        ########## FRAME TYPE#########
        self.rdArticleAvance = tk.Radiobutton(self.lblFrTypeAvance, text = 'Article', variable = self.varTypeFiltre, value ='article')
        self.rdCodeDefautAvance = tk.Radiobutton(self.lblFrTypeAvance, text = 'Code Defaut', variable = self.varTypeFiltre, value = 'codeDefaut')
        self.rdRepereAvance = tk.Radiobutton(self.lblFrTypeAvance, text = 'Repere', variable = self.varTypeFiltre, value = 'repere')
        #Collage
        self.rdArticleAvance.grid(column = 0, row = 0, sticky = 'w')
        self.rdCodeDefautAvance.grid(column=0, row = 1, sticky = 'w')
        self.rdRepereAvance.grid(column=0, row = 2, sticky = 'w')
 
        ########### FRAME AUTRE###########
        self.etrNombreElement = tk.Entry(self.lblFrAutreAvance,textvariable = self.varNombreElement, width = 2 )
        self.lblNombreElement = tk.Label(self.lblFrAutreAvance, text = 'Nombre d\'element')
        #Collage
        self.etrNombreElement.grid(column = 0, row = 0)
        self.lblNombreElement.grid(column = 1, row = 0)

        ######"Btn ok"
        self.bntOKAvance = tk.Button(self.fenetreModeAvance, text = "Rafraichir", command = self.Rafraichir_graphique,font = self.policeLabelFrame)
        self.bntOKAvance.grid(column = 0, row = 1, columnspan = 3)

        self.Rafraichir_graphique()
        
        """
        self.cmbArticleAvance.bind('<Button-1>', lambda event : self.Gestionnaire_evenement_class_App (event))
        self.cmbRepereAvance.bind('<Button-1>', lambda event : self.Gestionnaire_evenement_class_App (event))
        self.cmbCodeDefautAvance.bind('<Button-1>', lambda event : self.Gestionnaire_evenement_class_App (event))
        """
#''''''''''''''''''''''''''''''''''''''''''''''''''
    def Rafraichir_graphique(self):
        """Méthode pour rafraichir le graphique
            Possibilité d'utiliser des filtres pour analyser le composant / article ou code défaut selectionné
            Méthode : parcours de la liste de fichier pour dresser des objets xml
            si filtre(s) activé(s), vérification du filtre et ajout dans la liste les défaut sélectionnés """


        ################PREPARATION DE LA LISTE DE DEFAUT A ANALYSER#######################################
        #liFichierGraphique = self.liFichier[:] #Copie de la liste de fichier pour traitement des filtres

        liDefaut = [] #Liste de defaut
        liDefautFiltre = [] #Liste de defaut si filtre
        liArticleAvance = []
        liRepereAvance = []
        liCodeDefautAvance = []
        liArticleAvanceFiltre = [] #Liste des articles en defaut si filtre
        liRepereAvanceFiltre = [] #idem pour repere
        liCodeDefautAvanceFiltre = [] #idem pour code defaut
        filtreRepere = self.cmbRepereAvance.get()
        filtreArticle = self.cmbArticleAvance.get()
        filtreCodeDefaut = self.cmbCodeDefautAvance.get()
        ajouterDefaut = False

        #Transformation en booleen pour faciliter le traitement
        if filtreRepere != 'tous':
            filtreRepere = True
        else :
            filtreRepere = False
        if filtreArticle != 'tous':
            filtreArticle = True
        else :
            filtreArticle = False       
        if filtreCodeDefaut != 'tous':
            filtreCodeDefaut = True
        else:
            filtreCodeDefaut = False

        #Parcours des fichiers, création des xml et vérification des filtres
        for fichier in self.liFichier:
            xml = XML(fichier)
            liDefaut = xml.Recuperer_defaut() #Recuperation des defauts par fichier
            for defaut in liDefaut :
                #--------------- un seul et unique filtre ---------------
                if XOR(filtreArticle,filtreRepere,filtreCodeDefaut): #Si seulement un filtre est activé
                    for code in defaut.codeDefaut:
                        if code == self.cmbCodeDefautAvance.get():
                            ajouterDefaut = True              
                    if defaut.repereTopo == self.cmbRepereAvance.get():
                        ajouterDefaut = True
                    if defaut.codeArticle == self.cmbArticleAvance.get():
                        ajouterDefaut = True              
                #---------------- plusieurs filtres activés --------------
                elif filtreArticle and filtreRepere and not(filtreCodeDefaut): #Sinon, si deux des filtres sont activés
                    if defaut.repereTopo == self.cmbRepereAvance.get() and defaut.codeArticle == self.cmbArticleAvance.get():
                        ajouterDefaut = True
                
                elif filtreArticle and filtreCodeDefaut and not(filtreRepere):
                    if defaut.codeArticle == self.cmbArticleAvance.get():
                        for code in defaut.codeDefaut:
                            if code == self.cmbCodeDefautAvance.get():
                                ajouterDefaut = True
                        
                elif filtreRepere and filtreCodeDefaut and not (filtreArticle) :
                    if defaut.repereTopo == self.cmbRepereAvance.get():
                        for code in defaut.codeDefaut:
                            if code == self.cmbCodeDefautAvance.get():
                                ajouterDefaut = True
                
                if ajouterDefaut == True:
                    liDefautFiltre.append(defaut)
                    ajouterDefaut = False

#Voir pour améliorer cette partie => pas forcement nécessaire de recreer des variables
        if self.cmbArticleAvance.get() != 'tous' or self.cmbRepereAvance.get() != 'tous' or self.cmbCodeDefautAvance.get () != 'tous':
            #Récupération des info pour traitement avancé
            liArticleAvanceFiltre = Liste_defaut_par_type(liDefautFiltre,typeListe='article')[:]
            liRepereAvanceFiltre = Liste_defaut_par_type(liDefautFiltre,typeListe='repere')[:]
            liCodeDefautAvanceFiltre = Liste_defaut_par_type(liDefautFiltre,typeListe='codeDefaut')[:]

            liArticleAvanceFiltre = Counter(liArticleAvanceFiltre).most_common()
            liRepereAvanceFiltre = Counter(liRepereAvanceFiltre).most_common()
            liCodeDefautAvanceFiltre = Counter(liCodeDefautAvanceFiltre).most_common()
        else:
            liArticleAvance = Liste_defaut(self.liFichier,typeListe= 'article' ,filtre = 'aucun')[:]
            liRepereAvance = Liste_defaut(self.liFichier,typeListe= 'repere' ,filtre = 'aucun')[:]
            liCodeDefautAvance = Liste_defaut(self.liFichier,typeListe= 'codeDefaut' ,filtre = 'aucun')[:]

            liArticleAvance = Counter(liArticleAvance).most_common()
            liRepereAvance = Counter(liRepereAvance).most_common()
            liCodeDefautAvance = Counter(liCodeDefautAvance).most_common()

#Préparation des données graphiques
        liElement = []

        for i in range(int(self.varNombreElement.get())):
            try :
                if self.varTypeFiltre.get() == 'article': #filtre par article
                    if len(liArticleAvanceFiltre) > 0:
                        liElement.append(liArticleAvanceFiltre[i])
                    else:
                        liElement.append(liArticleAvance[i])

                if self.varTypeFiltre.get() == 'codeDefaut' : #filtre par code defaut
                    if len(liCodeDefautAvanceFiltre) > 0:
                        liElement.append(liCodeDefautAvanceFiltre[i])
                    else:
                        liElement.append(liCodeDefautAvance[i])

                if self.varTypeFiltre.get() == 'repere' : #filtre par repere topo
                    if len(liRepereAvanceFiltre) > 0:
                        liElement.append(liRepereAvanceFiltre[i])
                    else:
                        liElement.append(liRepereAvance[i])
            except IndexError :
                print ("Plus de défaut")

        liFiltre = []
        liFiltre.append(self.cmbArticleAvance.get())
        liFiltre.append(self.cmbRepereAvance.get())
        liFiltre.append(self.cmbCodeDefautAvance.get())

        self.Creation_graphique(liElement,self.varTypeFiltre.get(), liFiltre)

#''''''''''''''''''''''''''''''''''''''''''''''''''
    def Creation_graphique(self,liElement, typeFiltre, listeFiltre):
        """Méthode pour la création des graphiques"""

        #print (liElement)
        liHauteurBarre = []
        liEtiquette = []

        couleur = 'yellow'
        filtreArticle = listeFiltre[0]
        filtreRepere = listeFiltre[1]
        filtreCodeDefaut = listeFiltre[2]

        


        if typeFiltre == 'article':
            couleur = 'blue'
        if typeFiltre == 'repere':
            couleur = 'red'

        for cle, value in liElement:
            liHauteurBarre.append(value)
            liEtiquette.append(cle)

        self.graph = plt.figure()
        self.axe = plt.axes()
        barres = self.axe.bar(range(len(liElement)),liHauteurBarre,width = 0.6, color = couleur)
        self.axe.plot()
        plt.xticks(range(len(liElement)),liEtiquette)
        plt.title('Nombre de défaut par ' + typeFiltre)

        #Placement des infos filtres
        msg = 'Filtre article : ' + filtreArticle + '\nFiltre repere : ' + filtreRepere + '\nFiltre code défaut : ' + filtreCodeDefaut
        at = AnchoredText(msg,prop = dict(size=6),frameon=True,loc='upper right')
        
        at.patch.set_boxstyle("round,pad=0,rounding_size=0.2")
        self.axe.add_artist(at)
        #plt.text(0,0,
                #'Filtre article : ' + filtreArticle + '\nFiltre repere : ' + filtreRepere + '\nFiltre code défaut : ' + filtreCodeDefaut
                #, horizontalalignment='center',verticalalignment='center', transform= plt.gcf().transFigure) 
        #self.graph.xticks(range(len(liElement)),liEtiquette)
        for barre in barres:
            hauteur = barre.get_height()
            print (hauteur)
            self.axe.annotate('{}'.format(hauteur),xy = (barre.get_x()+barre.get_width()/2, hauteur),
                                                    xytext = (0,1), ha = 'center', va = 'bottom',
                                                    textcoords="offset points")

        self.graph.show()
        #plt.show()



        #plt.bar(range(len(liElement)),liHauteurBarre,width = 0.6, color = 'yellow')
        #plt.xticks(range(len(liElement)),liEtiquette) 
        #plt.show()

#===========================================================CLASS XML=================================================
class XML:
    """Class pour les gestions des xml defect
    Récupere les infos cartes : Référence carte :   self.reference
                                Numero de série :   self.sn
                                Of              :   self.of
    Dresse une liste d'objet défaut sous la forme :
    [Defaut,Defaut,Defaut,etc...] avec :
    Defaut.repereTopo   :   Repere topo du composant => string
    Defaut.codeArticle  :   Code article du composant => string
    Defaut.codeDefaut   :   liste des codes défauts => liste de string
    """
#---------------------------------------------------------------------------------------------------------------------
    def __init__(self,fichier):
        """Initialisation de la class
            Argument : fichier => string du fichier defect.xml a analyser
            """

        self.tree = ET.parse(fichier)
        self.root = self.tree.getroot()
        self.Recuperer_info_carte()
        self.Recuperer_defaut()
#---------------------------------------------------------------------------------------------------------------------

    def Recuperer_info_carte (self):
        """Méthode pour la récuperation des informations carte contenue dans le xml:
            Référence   = <Model>
            SN          = <SerialNo>
            OF          = <SerialNo
            """

        for info in self.root.findall('BoardInfo'):
            self.reference = info.find('Model').text
            self.ofSn = info.find('SerialNo').text
            self.sn = self.ofSn[len(self.ofSn)-4:]
            self.of = self.ofSn[:-5]
#---------------------------------------------------------------------------------------------------------------------

    def Recuperer_defaut (self):
        """Méthode pour dresser une liste de défaut
                Repere topo     = <RefId>
                Code article    = <PartNumber>
                Code défaut     = <DefectTypeString>
        """  
        self.liDefaut = []
        
        for defaut in self.root.findall('Defect'):
            #Parcous des défauts du xml
            nouveauDefaut = True
            #Recupération des infos
            repereTopo = defaut.find('RefId').text
            codeArticle = defaut.find('PartNumber').text
            codeDefaut = defaut.find('DefectTypeString').text

            if len(self.liDefaut) == 0:
               pass

            else:
                for element in self.liDefaut:
                    if element.repereTopo == repereTopo:
                        element.codeDefaut.append(codeDefaut)
                        nouveauDefaut = False
                        break

            if nouveauDefaut:
                self.Ajout_defaut_liste(repereTopo,codeArticle,codeDefaut)
                nouveauDefaut = False
        return(self.liDefaut)
#---------------------------------------------------------------------------------------------------------------------

    def Affichage_liste_defaut(self):
        """Méthode pour afficher la liste des défauts
        """

        for element in self.liDefaut:
                stgCodeDefaut = '('
                for codeDefaut in element.codeDefaut:
                    stgCodeDefaut += codeDefaut + " , "
                stgCodeDefaut += ')'
                print ('Le repere {}, code article {} a le(s) défaut(s) suivant(s) : {}'.format(element.repereTopo,element.codeArticle,element.codeDefaut))
#---------------------------------------------------------------------------------------------------------------------

    def Ajout_defaut_liste(self,repereTopo,codeArticle,codeDefaut):
        """Méthode pour l'ajout de la classe Defaut dans la liste
        Argument :  repereTopo => repere a ajouter => string
                    codeArticle => le code article du composant => string
                    codeDefaut => le code qualidefo => string
        """
        clDefaut = Defaut()
        clDefaut.repereTopo = repereTopo
        clDefaut.codeArticle = codeArticle
        clDefaut.codeDefaut.append(codeDefaut)
        self.liDefaut.append(clDefaut)
#---------------------------------------------------------------------------------------------------------------------

    def Trie_liste(self,mode='coissant'):
        """Methode pour le tri de la liste en fonction du nombre de défauts
        Argument : mode => string
                            defaut => mode par defaut tri du moins vers le plus
                            inverse => du plus vers le moins
        """

        for element in self.liDefaut :
            element.Nombre_defaut()#compte le nombre de défaut => methode de la classe Defaut

        if mode == 'coissant':
            #Mode pour trier du composant avec le moins de défaut vers le plus de défaut
            i = 0
            while i < len(self.liDefaut):
                nombreDefautActuel = self.liDefaut[i].nombreDefaut
                try :
                    nombreDefautSuivant = self.liDefaut[i+1].nombreDefaut
                except IndexError:
                    print ('liste terminée')
                    break
                if nombreDefautActuel > nombreDefautSuivant:
                    defautActuel = self.liDefaut[i]
                    defautSuivant = self.liDefaut[i+1]
                    self.liDefaut[i] = defautSuivant
                    self.liDefaut[i+1] = defautActuel
                    i-= 1
                else :
                    i+=1
        
        if mode == 'decoissant':
            #Mode pour trier du composant avec le plus de défaut vers le moins de défaut
            self.Trie_liste()
            self.liDefaut.reverse()
#---------------------------------------------------------------------------------------------------------------------

    def Extract_article (self, mode = 'croissant',filtre='aucun'):
        """Méthode pour extraire une lite de tuple des articles en défaut avec :
        [0] : clé = code article, [1] : value = nombre d'article en défaut dans le xml
        3 mode  : croissant (mode par defaut) => organiser de façon croissante
                : decroissant 
                : liste => retourne une liste des codes articles
        Utilisation de filtre possible sous forme 'Code1,Code2,etc...'
        """

        dictArticle = {}#Utilisation d'un dictionnaire pour avoir le nombre de défaut par code article
        liArticle = []

        liDefaut = self.liDefaut[:]#Copie de la liste de la classe XML


        #Si utilisation de la fonction filtre
        if filtre != 'aucun':
            liDefaut = self.Filtre_liste(filtre)
                        
        if mode == 'liste':
            for element in liDefaut :
                liArticle.append(element.codeArticle)
            return liArticle

        if mode == 'croissant':
            #Parcours de la liste pour extraire les code articles 
            for element in liDefaut:
                codeArticle = element.codeArticle
                if not(dictArticle):
                    dictArticle[codeArticle] = 1

                else:
                    for key, value in dictArticle.items():
                        ajout = True
                        if key == codeArticle:
                            value +=1
                            dictArticle[key] = value
                            ajout = False
                            break
                    if ajout :
                        dictArticle[codeArticle] = 1

        #Trie du dictionnaire => génération d'une liste de tuple avec [0] : clé , [1] : value
            return sorted(dictArticle.items(), key = lambda t:t[1])

        if mode == 'decroissant':
            print("Trie decroissant :")
            liste = self.Extract_article()
            liste.reverse()
            return liste
#---------------------------------------------------------------------------------------------------------------------

    def Extract_code_defaut(self,mode='croissant', filtre = 'aucun'):
        """Méthode pour extraire une lite de tuple des codes défaut avec :
        [0] : clé = code défaut, [1] : value = nombre de code dans le xml
        3 mode  : croissant (mode par defaut) => organiser de façon croissante
                : decroissant 
                : liste => retourne une liste des codes défauts
        Utilisation de filtre possible sous forme 'Code1,Code2,etc...'
        """
        dictCodeDefaut = {}
        liCodeDefaut = []

        liDefaut = self.liDefaut[:]#Copie de la liste de la classe XML
            

        if filtre != 'aucun':
            liDefaut = self.Filtre_liste(filtre)

        if mode == 'liste':
            for element in liDefaut:
                for code in element.codeDefaut:
                    liCodeDefaut.append(code)

            return liCodeDefaut
        
        if mode == 'croissant':#Même principe, parcours d'un dictionnaire et ajout si pas déjà présent
            liCodeDefaut = self.Extract_code_defaut('liste')

            for element in liCodeDefaut:
                if not(dictCodeDefaut): #dictionnaire vde ?
                    print ("Dictionnaire vide, ajout du code défaut\n\n")
                    dictCodeDefaut[element] = 1

                else : #Sinon, parcours du dictionnaire
                    for key, value in dictCodeDefaut.items():
                        ajout = True

                        if key == element:
                            value += 1
                            dictCodeDefaut[key] = value
                            ajout = False
                            break

                    if ajout:
                        dictCodeDefaut[element] = 1

            return sorted(dictCodeDefaut.items(), key = lambda t:t[1])
        
        if mode == 'decroissant':
            liste = self.Extract_code_defaut()
            liste.reverse()
            return liste
 #---------------------------------------------------------------------------------------------------------------------
 #    
    def Extract_repere(self, mode = 'liste',filtre = 'aucun'):
        """Méthode pour l'extraction des repere Topo. Pour le moment que sous forme de liste
        Utilisation de filtre possible sous forme 'Code1,Code2,etc...'
        """

        liRepere = [] 
        liDefaut = self.liDefaut[:]#Copie de la liste de la classe XML


        if filtre != 'aucun':
            liDefaut = self.Filtre_liste(filtre)

        if mode == 'liste':
            for element in liDefaut:
                liRepere.append(element.repereTopo)
        return liRepere
#---------------------------------------------------------------------------------------------------------------------

    def Filtre_liste(self,filtre):
        """Méthode pour genérer une nouvelle liste suivant les filres utilisés"""

        #copie de la liste
        liDefaut = self.liDefaut[:]
        liFiltre = filtre.split(',')
        #print ("code article a filtrer :",liFiltre)
        #print ("suppression des elements dans la liste de defaut")
        while  len(liDefaut) != 0:
            del(liDefaut[0])
        #print ("reconstruction de la liste en fonction du filtre actif\n\n")

        #parcours de la liste des defaut gloable
        for element in self.liDefaut:
            #print ("element : ", element.repereTopo)
            for codeDefaut in element.codeDefaut:
                #print ("code Defaut dans la liste de défaut : ",codeDefaut)
                if liFiltre.count(codeDefaut) >0:#si présence du code défaut actuel dans la liste
                    #print ("Defaut dans la liste des filtres ---> ajout\n\n\n\n")
                    liDefaut.append(element)
                    break
                else:
                    pass
                    #print("Defaut non présent dans la liste, pas d'ajout\n\n")
        return liDefaut

#=======================================CLASS DEFAUT================================================================================
class Defaut:
    """Classe permettant l'organisation des défauts"""
#---------------------------------------------------------------------------------------------------------------------
    def __init__(self):
        self.repereTopo = ""
        self.codeArticle = ""
        self.codeDefaut = []
        self.nombreDefaut = len(self.codeDefaut)
#---------------------------------------------------------------------------------------------------------------------

    def Nombre_defaut(self):
        """Méthode pour compter le nombre de défaut par Repere"""
        self.nombreDefaut = len(self.codeDefaut)
#---------------------------------------------------------------------------------------------------------------------

#______________________________________________________________FONCTIONS_____________________________________________________________

def Recuperer_liste_xml(chemin,of='aucun',sn = 'aucun'):
    """Fonction pour dresser une liste des fichiers XML à traiter
    chemin : string du chemin à parcourir
    of = option pour filtrer par rapport a un OF
    sn = option pour filtrer par rapport a un numéro de série
    Si filtre par carte, obligation de préciser l'OF
    La liste contiens le top et le BOT
    """

    retourLiFichier = []

    #Parcours du chemin pour trouver les xml
    for rChemin, rNom, liFichier in os.walk(chemin):

        for fichier in liFichier:
            if fichier.count('defect.xml') > 0:
                #print ('fichier xml trouvé......')
            
                if of != 'aucun': #Si utilisation du filtre OF
                    #Vérification si OF concerné
                    liChemin = rChemin.split('\\')
                    ofXml = liChemin[-1][:-5]#Récupération de l'OF de xml trouvé
                    if ofXml == of:
                        #print ("OK, of trouvé")
                        if sn != 'aucun':#Si utilisation du filtre sn
                            #print ("Filtre par SN")
                            snXml = liChemin[-1][len(ofXml)+1:]
                            if snXml == sn :
                                #print("Sn trouvé")
                                rChemin = rChemin.replace('\\','/')
                                #print (rChemin + '/' + fichier)
                                cheminXml = rChemin + '/' + fichier
                                #print("Ajout du fichier dans la liste------------\n\n")
                                retourLiFichier.append(cheminXml)
                        else :
                            rChemin = rChemin.replace('\\','/')
                            #print (rChemin + '/' + fichier)
                            cheminXml = rChemin + '/' + fichier
                            retourLiFichier.append(cheminXml)
                else :
                    #print ("Pas de filtre d'OF")
                    rChemin = rChemin.replace('\\','/')
                    print (rChemin + '/' + fichier)
                    cheminXml = rChemin + '/' + fichier
                    retourLiFichier.append(cheminXml)

    return retourLiFichier

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Liste_defaut_par_article(listeDeFichier, filtre = 'aucun'):
    """Fonction pour dresser une liste d'article en défaut en fonction d'une liste de fichier xml
    """
    liDefaut = []#liste tempon avant la liste de retour
    retourLiDefaut = []#liste de retour
    for fichier in listeDeFichier:
        xml = XML(fichier)
        liDefaut.append(xml.Extract_article(mode = 'liste', filtre = filtre))
    for liste in liDefaut:
        for element in liste:
            retourLiDefaut.append(element)
    
    return retourLiDefaut

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++        
def Liste_defaut(listeDeFichier,filtre = 'aucun',typeListe = 'article'):
    """Fonction pour dresser une liste de défaut en fonction d'un filtre et du type
    filtre : code défaut (exemple : AMTC, BNS etc...)
    type : type de liste => article
                            code defaut
                            repereTopo
    listeDefaut : si on veut spécifier une liste de défaut particulière
    """

    liDefaut = []#liste tempon avant la liste de retour
    retourLiDefaut = []#liste de retour

    for fichier in listeDeFichier:
        xml = XML(fichier)

        if typeListe == 'article':
            liDefaut.append(xml.Extract_article(mode = 'liste', filtre = filtre))
            
        if typeListe == 'repere':
            liDefaut.append(xml.Extract_repere(mode = 'liste', filtre = filtre))

        if typeListe == 'codeDefaut':
            liDefaut.append(xml.Extract_code_defaut(mode='liste', filtre = filtre))

        if typeListe == 'all':
            return xml.Recuperer_defaut()
            
    #Génération d'une liste unique pour le retour
    for liste in liDefaut:
        for element in liste:
            retourLiDefaut.append(element)
    
    return retourLiDefaut

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Gestionnaire_evenement(event,application):
    """Fonction pour gérer les évements liés à l'application"""
    print ("************* EVENEMENT ************")
    application.Analyse()

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Liste_defaut_par_type(liDefaut,typeListe = 'article'):
    """Fonction equivalent a la methode d'extraction des défaut xml"""

    if typeListe == 'article':
        liArticle = []

        for element in liDefaut :
            liArticle.append(element.codeArticle)
        return liArticle
        
    if typeListe == 'repere':
        liRepere = []

        for element in liDefaut :
            liRepere.append(element.repereTopo)
        return liRepere

    if typeListe == 'codeDefaut':
        liCodeDefaut = []
        for element in liDefaut:
            for code in element.codeDefaut:
                liCodeDefaut.append(code)

        return liCodeDefaut


def XOR (var1, var2,var3):
    """Fonction pour la fonction ou exclusif utilisé pour les filtres"""
    if var1 == True and var2 == False and var3 == False:
        return True
    elif var1== False and var2 == True and var3==False:
        return True
    elif var1==False and var2 == False and var3 == True:
        return True
    else:
        return False

#######################################################MAIN############################################################################
if __name__ == "__main__":
    cheminTest  = 'C:/Users/dgambie/Documents/PROGRAMMATION/PYTHON/QualiVisi/ExempleXML'
    cheminRubisMV_3 = '//rubis/Laudren/Defaut_AOI/MV_3/INTERFACE-CONCEPT/IC4060CB-V120-01'
    app = Application()
    app.bind("<Return>", lambda event : Gestionnaire_evenement(event,app))
    app.mainloop()


#DEBUG****************************************************************************************************************************
    
    
    #Procédure pour récupérer une liste de classe défaut
    #Permet de filtrer plus facilement les defauts
    #On peut utiliser une listededefaut déja réalisée
    """liFichier =  (Recuperer_liste_xml(cheminTest))
    liDefaut = []
    liXml = []
    print(liFichier)
    for fichier in liFichier:
        xml = XML(fichier)
        liXml.append(xml)
    for xml in liXml:
        liDefaut = xml.Recuperer_defaut()
    
    for defaut in liDefaut :
        print (defaut.repereTopo)"""


        #print ("Carte : {},OF : {} sn {}".format(xml.reference,xml.of,xml.sn))
        #xml.Affichage_liste_defaut()
        #print ("\n\n")
    #print ("\n\n liste des codes articles en défaut : ",Liste_defaut(liFichier))
    """print ("\n\n liste des codes defaut : ",Liste_defaut(liFichier,typeListe='codeDefaut'))
    print ("\n\n liste des Reperes en défaut : ",Liste_defaut(liFichier,typeListe='repere'))
    print ("\n\n\n\n=============================TOP DEFAUT =======================================================================")
    print (Counter(Liste_defaut(liFichier,typeListe='repere')).most_common())
    print ("\n\n\n\n=============================TOP DEFAUT =======================================================================")
    liDefautCodeDefaut = Liste_defaut(liFichier,typeListe='codeDefaut')
    print (Counter(liDefautCodeDefaut).most_common())
    nombreCodeEnDefaut = 0

    print ("\n\n\n\n=============================TOP DEFAUT =======================================================================")
    print (Counter(Liste_defaut(liFichier,typeListe='article',filtre = 'AMTC')).most_common())
    """

    """
    xml = XML('C:/Users/dgambie/Documents/PROGRAMMATION/PYTHON/QualiVisi/defect.xml')
    print ("Référence : {}, SN : {}, OF : {}".format(xml.reference,xml.sn,xml.of))
    xml.Affichage_liste_defaut()
    print ("\n\n\n====================TRIE=======================")
    xml.Trie_liste()
    xml.Affichage_liste_defaut()
    print ("\n\n\n====================Inverse=======================")
    xml.Trie_liste('decoissant')
    xml.Affichage_liste_defaut()
    print ("\n\n\n-----------Generation de la liste de code article en défaut sans filtre-----------------")
    print (xml.Extract_article(mode = 'liste', filtre = 'aucun'))
    print ("\n\n\n-----------Generation de la liste de code article avec FILTRE-----------------")
    print (xml.Extract_article(mode = 'liste', filtre = 'BNS,ANCC'))
    print ("\n\n\n-----------Generation de la liste de tuple de code article en défaut-----------------")
    print (xml.Extract_article())
    print ("\n\n\n----------Trie par code article le plus en défaut--------------")
    print (xml.Extract_article('decroissant'))
    print ("\n\n\n****************Trie par code défaut****************************")
    print (xml.Extract_code_defaut('croissant'))
    print ("\n\n\n****************Trie par code défaut decroisant****************************")
    print (xml.Extract_code_defaut('decroissant'))
    print ("\n\n\n___________________Extraction des reperes en défauts_______________")
    print (xml.Extract_repere())
    print ("\n\n\n___________________Extraction des reperes en défauts avec filtre_______________")
    print (xml.Extract_repere(filtre='BNS'))
    """

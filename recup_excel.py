import os

##__________________________________________________## ALGO DE RECUP DE L'EXCEL ##__________________________________________________##

dossier_telechargements = os.path.expanduser('~/Downloads')
fichiers_telechargements = os.listdir(dossier_telechargements)

switch = True
for fichier in fichiers_telechargements:
    if fichier.startswith('export-operations'):
        if switch:
            chemin_fichier1 = os.path.join(dossier_telechargements, fichier)
            switch = not switch
        else:
            chemin_fichier2 = os.path.join(dossier_telechargements, fichier)

infos_fichier = os.stat(chemin_fichier1)
date_creation_fichier1 = infos_fichier.st_ctime

infos_fichier = os.stat(chemin_fichier2)
date_creation_fichier2 = infos_fichier.st_ctime

if date_creation_fichier1 > date_creation_fichier2:
    chemin_fichier_incomes = chemin_fichier1
    chemin_fichier_expenses = chemin_fichier2
else:
    chemin_fichier_incomes = chemin_fichier2
    chemin_fichier_expenses = chemin_fichier1

# os.remove(chemin_fichier1)
# os.remove(chemin_fichier2)


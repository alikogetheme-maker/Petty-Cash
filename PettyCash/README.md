# PettyCashAddon — Gestion de caisse pour SAP Business One (HANA)

Add-on classique SAP Business One (UI API + DI API) qui reproduit, pour une
caisse unique, les fonctions du **Cash Journal** de S/4HANA :

- Connaître le solde d'ouverture au début de chaque quart/journée.
- Enregistrer les recettes et dépenses de caisse (avec comptabilisation
  automatique en écriture au journal).
- Clôturer la session avec comptage physique et calcul automatique de
  l'écart (comptabilisé sur un compte d'écart dédié).
- Éditer un rapport de caisse filtrable par période et par quart.

## 1. Mettre le projet en route dans Visual Studio

Le SDK SAP Business One n'est pas présent sur cette machine (aucune
installation détectée), donc ce dépôt contient le **code source complet**
mais n'a pas été compilé/testé ici. À faire sur votre poste de dev :

1. Ouvrez `PettyCashAddon.sln` dans Visual Studio.
2. Clic droit sur le projet → **Add Reference** → onglet **COM** → cochez :
   - `SAP Business One UI API` (génère `Interop.SAPbouiCOM.dll`)
   - `SAP Business One DI API` (génère `Interop.SAPbobsCOM.dll`)
   Ces composants apparaissent car le SDK/client SAP B1 est installé sur
   votre poste et les enregistre en COM.
3. Vérifiez que la plateforme cible du projet est **x86** (le SDK B1 est
   32 bits) et le Framework cible **.NET Framework 4.8**.
4. Compilez (Build). Si un nom de méthode/énumération diffère légèrement de
   votre version de SDK (ex. `BoUTBTableType`, `BoFieldTypes`), l'IntelliSense
   vous proposera l'équivalent exact — dites-le-moi et je corrige.

## 2. Enregistrer l'add-on dans SAP Business One

1. Dans SAP B1 : **Administration → Add-Ons → Add-On Administration**.
2. **New** → pointez vers `PettyCashAddon.exe` compilé (dossier `bin\x86\Debug`
   ou `Release`).
3. Cochez "Auto Start" si vous voulez qu'il démarre avec le client, sinon
   lancez-le manuellement depuis cette fenêtre pour les tests.
4. Au premier démarrage, l'add-on crée automatiquement (s'ils n'existent
   pas déjà) :
   - Les tables `@PC_SETTINGS`, `@PC_TTYPE`, `@PC_SESSION`, `@PC_TRANS`
   - Un menu **Petty Cash** avec deux entrées : *Session de caisse* et
     *Rapport de caisse*.

## 3. Paramétrage initial (à faire une fois, dans SAP B1)

Avant la première session, ouvrez la table `@PC_SETTINGS` (via
*Outils → Personnalisation → Gestionnaire de données définies par
l'utilisateur*, ou directement l'écran généré) et renseignez :

- `U_CashAcct` : code du compte G/L "Caisse".
- `U_DiffAcct` : code du compte G/L "Écarts de caisse".

Puis alimentez `@PC_TTYPE` avec vos types d'opération (ex. `VENTE` /
Recette / compte de vente comptant, `FOURN` / Dépense / compte d'achats
divers, `BANQ` / Dépense / compte de virement banque, etc.) — c'est
l'équivalent des "types d'opération" du Cash Journal S/4HANA.

## 4. Structure du code

```
src/PettyCashAddon/
  Program.cs                     Point d'entrée, connexion UI API/DI API
  Constants.cs                   Noms de tables/champs
  Core/DiCompany.cs               Connexion DI API via le cookie de contexte
  Core/SboApplication.cs          Menus + dispatch des événements UI
  Setup/MetadataSetup.cs          Création idempotente des UDT/UDF au démarrage
  Models/Enums.cs                 Quart, Sens (Recette/Dépense), Statut session
  Services/CashSessionService.cs  Logique métier : ouverture, transaction, clôture
  Forms/FormIds.cs                Identifiants des écrans
  Forms/CashSessionFormController.cs   Écran "Session de caisse"
  Forms/TransactionEntryForm.cs        Popup de saisie d'une transaction
  Forms/CashReportFormController.cs    Écran "Rapport de caisse"
```

Toute la logique de comptabilisation (écritures au journal) est centralisée
dans `Services/CashSessionService.cs` — c'est le fichier à relire en
premier pour comprendre le comportement métier.

## 5. Limites connues / pistes d'évolution

- Une seule caisse gérée (pas de multi-caisse) — conforme au besoin exprimé.
- Le verrouillage d'une session clôturée est appliqué **au niveau de
  l'add-on** (impossible d'ajouter une transaction via cet écran une fois
  `U_Status = 'C'`), mais pas au niveau base de données. Pour un verrou
  strict même via SQL direct, il faudrait ajouter une Transaction
  Notification côté serveur — je peux l'ajouter si le contrôle interne
  l'exige.
- Le comptage de clôture est saisi en un seul montant global. Si vous
  voulez le détail par dénomination (billets/pièces), on ajoute un sous-écran
  de comptage — dites-le-moi.

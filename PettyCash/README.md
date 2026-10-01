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

Le projet compile sur le serveur (DI API et UI API 10.0 enregistrées) avec
le MSBuild de Visual Studio :

```
"C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe" src\PettyCashAddon\PettyCashAddon.csproj -restore
```

Sur un autre poste de dev :

1. Ouvrez `PettyCashAddon.sln` dans Visual Studio.
2. Clic droit sur le projet → **Add Reference** → onglet **COM** → cochez :
   - `SAP Business One UI API` (génère `Interop.SAPbouiCOM.dll`)
   - `SAP Business One DI API` (génère `Interop.SAPbobsCOM.dll`)
   Ces composants apparaissent car le SDK/client SAP B1 est installé sur
   votre poste et les enregistre en COM.
3. Vérifiez que la plateforme cible du projet est **x64** : le client SAP B1 10
   et la DI API installés sur le serveur sont en 64 bits (la DI API 32 bits n'y
   est pas enregistrée). Framework cible **.NET Framework 4.8**.
4. Compilez (Build). Si un nom de méthode/énumération diffère légèrement de
   votre version de SDK (ex. `BoUTBTableType`, `BoFieldTypes`), l'IntelliSense
   vous proposera l'équivalent exact — dites-le-moi et je corrige.

### Lancer et déboguer depuis Visual Studio (avant de packager)

1. Ouvrez le client SAP B1 et connectez-vous à la base de test (`TST_TST`).
2. Dans la gestion des add-ons, **arrêtez** l'add-on Petty Cash installé s'il tourne :
   deux instances créeraient le même menu.
3. Dans Visual Studio : configuration **Debug | x64**, puis **F5**.
   L'argument de lancement (chaîne de connexion de développement SAP) est déjà
   défini dans `PettyCashAddon.csproj.user` ; à défaut, `Program.cs` l'utilise
   automatiquement quand le débogueur est attaché.
4. Le menu **Petty Cash** apparaît dans le client ; les points d'arrêt fonctionnent.
   Pour arrêter : **Shift+F5** (le menu reste visible jusqu'au prochain démarrage
   ou redémarrage du client).

Quand tout est bon : compilez en **Release | x64** et packagez
`bin\x64\Release\PettyCashAddon.exe` (seul fichier nécessaire).

## 2. Enregistrer l'add-on dans SAP Business One

1. Dans SAP B1 : **Administration → Add-Ons → Add-On Administration**.
2. **New** → pointez vers `PettyCashAddon.exe` compilé (dossier `bin\x64\Debug`
   ou `Release`).
3. Cochez "Auto Start" si vous voulez qu'il démarre avec le client, sinon
   lancez-le manuellement depuis cette fenêtre pour les tests.
4. Au premier démarrage, l'add-on crée automatiquement (s'ils n'existent
   pas déjà) :
   - Les tables `@PC_SETTINGS`, `@PC_TTYPE`, `@PC_SESSION`, `@PC_TRANS`
   - Un menu **Modules → Petty Cash** avec deux entrées : *Session de caisse* et
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

## 4. Règles de gestion

- Une seule session ouverte à la fois.
- Solde d'ouverture = **solde compté** de la dernière session clôturée (même
  s'il vaut 0). Pour la toute première session, c'est le **solde du compte
  G/L Caisse** : le fonds de caisse initial doit y avoir été comptabilisé.
- Solde théorique = ouverture + recettes − dépenses, recalculé depuis la base.
- Une dépense qui rendrait le solde théorique négatif est refusée.
- Chaque transaction génère une écriture : Recette = Débit Caisse / Crédit
  compte du type ; Dépense = Débit compte du type / Crédit Caisse.
- Clôture : écart = compté − théorique. Excédent = Débit Caisse / Crédit
  Écarts ; manquant = Débit Écarts / Crédit Caisse.
- Écriture + ligne de caisse + mise à jour de la session sont faites dans une
  même transaction DI API : en cas d'erreur, rien n'est enregistré.
- Le tiers saisi est contrôlé (doit exister dans les partenaires) mais reste
  informatif : l'écriture est passée sur le compte G/L du type d'opération.
- Une transaction ne se supprime pas : une erreur se corrige par une
  opération inverse (même type, sens opposé).

Guide utilisateur : `docs/Guide-utilisateur-PettyCash.html`.

## 5. Structure du code

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

## 6. Limites connues / pistes d'évolution

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

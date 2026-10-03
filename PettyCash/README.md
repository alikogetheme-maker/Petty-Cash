# PettyCashAddon — Gestion des caisses pour SAP Business One

Add-on classique SAP Business One (UI API + DI API, SQL Server ou HANA) qui
reproduit les fonctions du **Cash Journal** de S/4HANA, pour **plusieurs
caisses** tenues en parallèle par des utilisateurs SAP différents :

- Chaque utilisateur SAP ouvre sa session sur une **caisse libre** ; une caisse
  ouverte n'est plus proposée aux autres tant qu'elle n'est pas clôturée.
- **Billetage** (billets et pièces) à l'ouverture et à la clôture ; les écarts
  sont calculés et comptabilisés automatiquement.
- Recettes et dépenses de caisse comptabilisées en écriture au journal.
- **Rapports** par période, caisse, utilisateur et quart : sessions détaillées,
  synthèse par caisse, synthèse par utilisateur, détail des opérations et du
  billetage.

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
4. Compilez (Build).

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
`bin\x64\Release\PettyCashAddon.exe` (seul fichier nécessaire : les types SAP
sont intégrés à l'exe, l'icône du menu aussi).

## 2. Enregistrer l'add-on dans SAP Business One

1. Dans SAP B1 : **Administration → Add-Ons → Add-On Administration**.
2. Enregistrez le package (64 bits) contenant `PettyCashAddon.exe`.
3. Au premier démarrage, l'add-on crée automatiquement (s'ils n'existent pas) :
   - les tables `@PC_SETTINGS`, `@PC_CASHBOX`, `@PC_TTYPE`, `@PC_SESSION`,
     `@PC_TRANS`, `@PC_DENOM`, `@PC_COUNT`, `@PC_LOCK` ;
   - les 13 coupures du franc CFA dans `@PC_DENOM` ;
   - un menu **Modules → Petty Cash** : *Session de caisse* et *Rapport de caisse*.
   Les autres utilisateurs connectés doivent se reconnecter après cette création.

## 3. Paramétrage initial (dans SAP B1, *Outils → Fenêtres par défaut*)

- **PC_SETTINGS** (ligne `1`) : `U_DiffAcct` = compte d'écarts de caisse par défaut.
- **PC_CASHBOX** : une ligne par caisse — Code court (ex. `CAISSE2`), Nom
  affiché (unique), `U_CashAcct` = compte G/L de la caisse, `U_DiffAcct`
  (facultatif) = compte d'écarts propre à la caisse, `U_Active` = `Y`/`N`.
  Utilisez **un compte G/L distinct par caisse**.
- **PC_TTYPE** : types d'opération (Code, Nom unique, `U_Dir` = `R`/`D`,
  `U_GLAcct`), communs à toutes les caisses.
- **PC_DENOM** : coupures proposées au billetage (`U_Active` = `N` pour en masquer une).

Migration depuis la version mono-caisse : au démarrage, une caisse `CAISSE1`
« Caisse principale » est créée avec l'ancien `U_CashAcct` de `PC_SETTINGS`, et
les sessions existantes y sont rattachées.

## 4. Règles de gestion

- Une session = une caisse + un utilisateur SAP (repris automatiquement) + un quart.
- **Une caisse n'a qu'une session ouverte** et **un utilisateur n'a qu'une
  session ouverte**. Garanti par des verrous (`@PC_LOCK`) protégés par l'index
  unique de SAP : deux postes ne peuvent pas ouvrir la même caisse, même au
  même instant. Un verrou dont la session n'est plus ouverte est ignoré.
- Seul l'**ouvreur** saisit des opérations ; l'ouvreur ou un
  **superutilisateur SAP** peut clôturer (depuis le rapport pour une session
  oubliée).
- **Ouverture** : solde attendu = solde compté de la dernière session de la
  caisse (première fois : solde du compte G/L de la caisse). Le billetage
  d'ouverture devient le solde d'ouverture ; s'il diffère de l'attendu,
  l'**écart d'ouverture** est comptabilisé.
- Solde théorique = ouverture + recettes − dépenses, recalculé depuis la base.
- Une dépense qui rendrait le solde théorique négatif est refusée.
- Écritures (référence = code de session) :
  - Recette : Débit Caisse / Crédit compte du type ;
  - Dépense : Débit compte du type / Crédit Caisse ;
  - Excédent (ouverture ou clôture) : Débit Caisse / Crédit Écarts ;
  - Manquant (ouverture ou clôture) : Débit Écarts / Crédit Caisse.
- **Clôture** : billetage obligatoire ; écart = compté − théorique, comptabilisé ;
  la caisse est libérée.
- Chaque opération (écritures, lignes, session, billetage, verrous) est faite
  dans une même transaction DI API : en cas d'erreur, rien n'est enregistré.
- Le tiers saisi est contrôlé (doit exister dans les partenaires) mais reste
  informatif : l'écriture est passée sur le compte G/L du type d'opération.
- Une transaction ne se supprime pas : une erreur se corrige par une
  opération inverse (même type, sens opposé).
- Toute table utilisateur SAP a un index unique sur `Code` **et** sur `Name` :
  l'add-on met le Code dans Name pour ses propres lignes.

Guide utilisateur : `docs/Guide-utilisateur-PettyCash.html`.

## 5. Structure du code

```
src/PettyCashAddon/
  Program.cs                     Point d'entrée, connexion UI API/DI API, journal de démarrage
  Constants.cs                   Noms de tables/champs
  Core/DiCompany.cs               Connexion DI API via le cookie de contexte
  Core/SboApplication.cs          Menu (avec icône intégrée)
  Setup/MetadataSetup.cs          Création idempotente des UDT/UDF, coupures, migration
  Models/Enums.cs                 Quart, Sens (Recette/Dépense), Statut session
  Services/CashSessionService.cs  Logique métier : caisses, verrous, sessions, billetage, rapports
  Forms/FormIds.cs                Identifiants des écrans (≤ 10 caractères)
  Forms/CashSessionFormController.cs   Écran "Session de caisse"
  Forms/TransactionEntryForm.cs        Popup de saisie d'une transaction
  Forms/BillCountForm.cs               Popup de billetage (ouverture / clôture)
  Forms/CashReportFormController.cs    Écran "Rapport de caisse" (grilles SAP)
```

Toute la logique de comptabilisation est centralisée dans
`Services/CashSessionService.cs` — c'est le fichier à relire en premier.

## 6. Limites connues / pistes d'évolution

- Les règles sont appliquées **par l'add-on**, pas au niveau base de données :
  une modification directe des tables ou une écriture manuelle sur un compte
  de caisse n'est pas bloquée. Une Transaction Notification côté serveur peut
  être ajoutée si le contrôle interne l'exige.
- Pas de restriction des caisses par utilisateur : toute caisse active et
  libre est proposée à tout utilisateur.

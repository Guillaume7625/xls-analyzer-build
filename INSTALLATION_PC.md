# Installation sur PC — PC_XLS_Analyser (.exe unique)

But : fabriquer `PC_XLS_Analyser.exe` sur le PC, puis l'utiliser (ou le copier
ailleurs). Il n'y a pas de « setup.exe » : on construit une fois, l'exe obtenu
est autonome (aucun Python requis pour l'exécuter, aucun réseau — loopback).
Durée totale : 10 à 15 minutes.

---------------------------------------------------------------------------
## 0. À mettre sur la clé USB (préparé depuis une machine connectée)

1. `python-3.12.x-amd64.exe`  — depuis python.org/downloads/windows
   (bien prendre **Windows installer 64-bit / amd64**, version **3.12**).
2. `xls-analyzer-v2.zip`      — le paquet livré.
3. **Seulement si le PC est HORS-LIGNE** : un dossier `wheels\` préparé sur un
   **Windows x64 connecté** (PAS sur Mac — wheels incompatibles) :

       python -m pip download -r v2\requirements.txt -d wheels
       python -m pip download pyinstaller -d wheels

---------------------------------------------------------------------------
## 1. Sur le PC

1. Lancer `python-3.12.x-amd64.exe`
   → **COCHER « Add python.exe to PATH »** → Install Now.
2. Décompresser `xls-analyzer-v2.zip`. Si PC hors-ligne : poser le dossier
   `wheels\` **AU MÊME NIVEAU** que `v2\` (pas dedans) :

       C:\...\xls-analyzer\
       ├── v2\          ← sources + build_onefile.bat
       └── wheels\      ← uniquement en mode hors-ligne

3. Double-clic sur **`v2\build_onefile.bat`**
   (il crée l'environnement, installe les dépendances — depuis Internet, ou
   depuis `..\wheels` s'il existe — puis gèle l'exe).
4. Résultat : **`v2\dist\PC_XLS_Analyser.exe`** — fichier unique.

En cas de blocage du .bat (politique du poste), les 4 commandes équivalentes
dans une invite `cmd` ouverte dans `v2\` :

       python -m venv .venv
       .venv\Scripts\activate
       pip install -r requirements.txt
       pyinstaller PC_XLS_Analyser_onefile.spec --noconfirm

---------------------------------------------------------------------------
## 2. Test de recette (le seul vrai critère)

1. Double-clic sur `PC_XLS_Analyser.exe`.
2. Une fenêtre s'ouvre (ou, à défaut de WebView2, le navigateur sur
   `http://127.0.0.1:<port>` — bascule automatique, c'est normal).
3. Déposer un `.xlsx` d'essai → le rapport s'affiche (verdict, tableaux).
4. L'exe peut ensuite être copié seul sur d'autres postes Windows x64 :
   rien d'autre à installer dessus.

Astuce : si le PC de build est aussi le poste cible et que Python ne doit pas
y rester, Python peut être désinstallé après le build — l'exe n'en dépend pas.

---------------------------------------------------------------------------
## 3. Si ça coince

| Symptôme | Cause probable | Action |
| --- | --- | --- |
| `python` ouvre le Microsoft Store | PATH non coché à l'installation | Réinstaller Python en cochant « Add to PATH » (ou désactiver les alias d'exécution d'applications). |
| `pip install` échoue, pas de réseau | PC hors-ligne sans `wheels\` | Préparer `wheels\` sur un Windows connecté (étape 0.3). |
| L'exe ne démarre pas / rien ne s'affiche | Voir le journal | `%LOCALAPPDATA%\xls-analyzer\launcher.log`. Pour voir l'erreur en direct : `DEBUG = True` dans `PC_XLS_Analyser_onefile.spec`, rebuild, lancer l'exe depuis `cmd`. |
| L'antivirus met l'exe en quarantaine | Binaire PyInstaller non signé (classique) | Restaurer + exclusion AV, ou faire signer le binaire (recommandé pour diffusion). |
| Poste durci : exécution depuis %TEMP% interdite | Le onefile s'extrait dans %TEMP% | Décommenter `runtime_tmpdir` dans le spec vers un chemin autorisé, rebuild — ou utiliser la variante dossier `PC_XLS_Analyser.spec`. |

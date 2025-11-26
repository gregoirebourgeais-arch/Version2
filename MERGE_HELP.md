# Résoudre un pull request bloqué par des conflits

Ce dépôt n'est pas connecté ici à GitHub, donc je ne peux pas voir le PR distant. Voici une procédure pas-à-pas pour répliquer et corriger les conflits en local, puis mettre à jour ton pull request.

## 1) Mettre à jour ta branche locale
1. Assure-toi d'avoir la branche cible (souvent `main` ou `master`) à jour :
   ```bash
   git checkout main
   git pull origin main
   ```
2. Reviens sur ta branche de travail (celle du PR) :
   ```bash
   git checkout <ta-branche>
   ```

## 2) Rejouer tes commits au-dessus de la branche cible
- Option recommandée (rebase) :
  ```bash
  git rebase main
  ```
  - S'il y a des conflits, Git s'arrêtera sur le commit concerné.
  - Résous les conflits dans les fichiers indiqués, puis :
    ```bash
    git add <fichiers-résolus>
    git rebase --continue
    ```
- Option alternative (merge) :
  ```bash
  git merge main
  ```
  Puis résous les conflits, `git add ...`, et termine avec `git commit`.

## 3) Résoudre les conflits
1. Ouvre chaque fichier marqué en conflit (`<<<<<<<`, `=======`, `>>>>>>>`).
2. Décide quelle version garder ou combine-les.
3. Supprime les marqueurs de conflit et sauvegarde.
4. Vérifie que le code compile/tests passent.

## 4) Finaliser et pousser
```bash
git push --force-with-lease
```
- Utilise `--force-with-lease` après un rebase (il réécrit l'historique du PR).
- Sans rebase (merge), un simple `git push` suffit.

## 5) Vérifier sur GitHub
- Rafraîchis la page du pull request : les conflits devraient disparaître s'ils sont résolus.

## Astuces
- Pour voir l'état des conflits restants :
  ```bash
  git status
  ```
- Pour redémarrer un rebase bloqué :
  ```bash
  git rebase --abort
  ```
- Pour afficher le diff d'un fichier en conflit :
  ```bash
  git diff --merge <fichier>
  ```

Cette procédure reste valable même si le nombre de conflits est important : traite-les fichier par fichier, commit par commit si tu rebases.

# excel-to-html-export
# 📊 Excel to HTML Exporter

Ce projet Python permet de convertir un fichier **Excel (.xlsx)** en une **page HTML fidèle**, en préservant :

✅ Les **données textuelles**  
✅ La **mise en forme** (gras, italique, couleur…)  
✅ Le **positionnement exact des cellules**  
✅ Les **images** insérées dans le fichier Excel

---

## ✨ Fonctionnalités principales

- 📄 Lecture du contenu d’une feuille Excel
- 🎨 Extraction des styles de chaque cellule (police, couleur, alignement)
- 🖼️ Récupération des images insérées dans Excel
- 📐 Calcul de leur **position exacte** (colonne, ligne, offset)
- 💡 Conversion des images au format **WebP** pour un affichage optimisé
- 🧱 Génération d’une page **HTML statique**, propre et fidèle à Excel

---

## ⚙️ Prérequis

Assurez-vous d’avoir installé Python 3 et les bibliothèques suivantes :

```bash
pip install openpyxl pillow

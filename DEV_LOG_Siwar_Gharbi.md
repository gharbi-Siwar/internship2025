# DEV_LOG - Siwar Gharbi

 Objectif

Développer une fonction Python propre et efficace pour extraire le texte à partir de fichiers de différents formats placés dans le dossier `uploads/`. Les formats pris en charge sont : `.pdf`, `.docx`, `.xlsx`. Pour les PDF scannés sans texte extractible, on utilise OCR via Tesseract.

---

## Approche

1. **Organisation modulaire :**
   - Chaque format de fichier a une fonction dédiée :
     - `extract_text_from_pdf()`
     - `extract_text_from_pdf_with_ocr()`
     - `extract_text_from_docx()`
     - `extract_text_from_xlsx()`
   - La fonction `extract_all_text_from_folder()` centralise l’itération sur les fichiers du dossier `uploads/`.

2. **PDF :**
   - Extraction via `pdfplumber`.
   - Si le PDF ne contient pas de texte (e.g. PDF scanné), on applique `pytesseract` avec PyMuPDF (`fitz`) pour OCRiser chaque page.

3. **DOCX :**
   - Utilisation de `python-docx` pour lire le contenu paragraphe par paragraphe.

4. **XLSX :**
   - Utilisation de `openpyxl` pour parcourir toutes les cellules et concaténer leur contenu.

5. **Logs :**
   - Tout incident ou fichier échoué est enregistré dans un fichier `extraction_log.log`.

6. **Résultat :**
   - Un dictionnaire Python contenant les chemins des fichiers comme clés et le texte extrait comme valeurs.
   - Une cellule supplémentaire permet de sauvegarder chaque texte dans un fichier `.txt`.

---

##  Dépendances

Les bibliothèques utilisées sont :

- `pdfplumber` – pour lire les textes dans les PDF  
- `PyMuPDF` (`fitz`) – pour transformer les pages PDF en images  
- `pytesseract` – pour faire de l’OCR sur les images  
- `Pillow` – pour manipuler les images en mémoire  
- `python-docx` – pour extraire le contenu des fichiers Word  
- `openpyxl` – pour extraire les données Excel  
- `logging` – pour consigner les erreurs ou anomalies  
- `pathlib` et `os` – pour la gestion des fichiers

---

## Instructions

1. Place tes fichiers dans le dossier `uploads/`.
2. Lance le script `extract_text_Siwar_Gharbi.py`.
3. Vérifie le contenu du dossier `extracted_texts/` pour voir les `.txt` extraits.
4. Vérifie `extraction_log.log` en cas de problème.

---

##  Extension possible

- Ajouter la prise en charge de `.txt`, `.csv`, `.pptx`
- Nettoyer ou segmenter le texte pour la suite du pipeline NLP


---

**Auteur :** Siwar Gharbi  

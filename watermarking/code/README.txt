CHARLES ZIJN TIME SAVING TOOL 
==================================

FOLDER STRUCTURE
----------------
watermarking/
  generate_teasers.py     ← NIET VERPLAATSEN
  buyers.xlsx             ← col A = vul hier onder elkaar alle watermark namen in, col B = owner password (row 1 = header, ignored)
  originals/
    Project Arrow - Teaser.pdf   ← Zet hier je orginele pdf in. Deze wordt niet aangeraakt :)
  output/                 ← Hier kan je dan je output vinden
  SourceSansPro-Regular.ttf      ← auto-downloaded on first run


FIRST-TIME SETUP (run once)
----------------------------

START MET OUTPUT FOLDER LEEG TE MAKEN
MAAK DAT JE BUYERS.XLS GESLOTEN IS
NAAM FORMAT ORGINAL Project x - Teaser (later wordt dan vanzelf - watermark toegevoegd)

Open Command Prompt (Win+R → cmd → Enter) en kopieer dan deze zaken één voor één. 

  cd "C:\Users\Intern5\Kumulus Partners\Kumulus Partners Team Site - Interns\5. Charles Willems\PDF\watermarking"
  pip install pypdf reportlab openpyxl fonttools brotli
  pip install cryptography
  python generate_teasers.py
OR
  python remove_passwords.py


Indien je foutmelding krijgt dat je geen python hebt. https://www.python.org/downloads/ druk op de gele knop. installeer alles; in je terminal zal je verschillende keren y enter moeten typen
Wanneer je installatie voltooid is en je Python terminal venster sluit, sluit terminal volledig af en start terug vanaf het begin. cd...
normaal zou nu alles moeten werken 


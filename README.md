# ArchivesSpace_DO_autofill
Automatically fill all handles and thumbnails in an ArchivesSpace digital object bulk importer form from SOLR data

After opening a Digital Object bulk importer form, press Alt + F11.
In Microsoft Visual Basic for Applications, on the right side under VBAProject > Modules, right click and select Import File.
Select this file and open Module 1.

Run the getHandles() sub and fill in your SOLR login data in the pop-up.
In case of a discrepancy between shelfmarks in Islandora and ArchivesSpace, comment out line shelfmark = Replace(shelfmark, "-", ": ") (by removing the ' sign)

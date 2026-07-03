#!/usr/bin/env python3
"""
boxwhisker_injector.py - Fuegt ein Excel-Box-Whisker-Diagramm (chartEx) in
eine bereits von pandas/openpyxl geschriebene .xlsx-Datei ein.

Version: v6 - Achsentitel-Feature wieder entfernt (chartEx-Schema fuer Achsentitel liess sich trotz Befolgen der offiziellen MS-Doku nicht zuverlaessig nachbauen, Excel lehnte die Datei weiterhin ab). Zurueck zur nachweislich funktionierenden Struktur ohne Achsentitel.
  - definedNames-Patch robuster (Regex statt exaktem String-Vergleich;
    openpyxl schreibt das selbstschliessende Tag je nach Version mal als
    "<definedNames/>", mal als "<definedNames />" mit Leerzeichen - das
    fuehrte zu einem doppelten, ungueltigen <definedNames>-Element)
  - Bugfix: fehlendes f-String-Praefix liess den Chart-Titel literal als
    "{title_esc}" im Diagramm erscheinen statt den echten Text
  - Chart-Titel wird jetzt von den aufrufenden Skripten dynamisch mit
    Land und Datumsbereich befuellt

HINTERGRUND: openpyxl kann den neueren Excel-Chart-Typ "chartEx" (u.a.
Box-Whisker, Histogramm, Wasserfall) weder lesen noch schreiben - dieser
Typ existiert in seinem Objektmodell schlicht nicht. Diese Funktion baut
die noetigen XML-Bausteine deshalb von Hand nach und fuegt sie direkt ins
ZIP-Archiv der bereits gespeicherten xlsx-Datei ein. Die Vorlage dafuer
stammt aus einer vom Nutzer in Excel manuell erstellten Box-Whisker-Datei
(siehe Chat-Diskussion) - Struktur 1:1 nachgebaut, nur Sheet-Name, Titel
und Zeilenzahl sind parametrisiert.

WICHTIG ZUR INTERPRETATION: Da pro Zeitstempel nur EIN Messwert vorliegt
(keine Mehrfachmessungen pro Kategorie), ist dies statistisch gesehen ein
"entarteter" Box-Plot - jede Box hat keine echte Streuung, sondern zeigt
schlicht den einen Wert als duennen Strich. Das ist hier aber bewusst so
gewollt: der Nutzer mag genau diesen visuellen Stil (duenne vertikale
Striche pro Zeitpunkt) gegenueber einem klassischen Liniendiagramm.

EINSCHRAENKUNG: Geht von GENAU EINEM Sheet in der xlsx-Datei aus (so wie
unsere Pumpspeicher-Skripte es erzeugen) und von der festen Spalten-
reihenfolge timestamp (A) / Turbiniert_MW (B) / Hochgepumpt_MW (C).

RISIKO: chartEx ist ein nicht-oeffentlich dokumentiertes Microsoft-Format.
Funktioniert in aktuellem Excel (Microsoft 365/2021+), aber moeglicherweise
nicht in aelteren Excel-Versionen, Google Sheets oder LibreOffice.
"""

import base64
import zipfile
import shutil
from xml.sax.saxutils import escape as xml_escape


_STYLE1_XML_B64 = (
    "PGNzOmNoYXJ0U3R5bGUgeG1sbnM6Y3M9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL2RyYXdpbmcvMjAxMi9j"
    "aGFydFN0eWxlIiB4bWxuczphPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvZHJhd2luZ21sLzIwMDYvbWFpbiIg"
    "aWQ9IjEwMiI+PGNzOmF4aXNUaXRsZT48Y3M6bG5SZWYgaWR4PSIwIi8+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJl"
    "ZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNz"
    "OmRlZlJQciBzej0iMTAwMCIgYj0iMSIga2Vybj0iMTIwMCIvPjwvY3M6YXhpc1RpdGxlPjxjczpjYXRlZ29yeUF4aXM+PGNzOmxu"
    "UmVmIGlkeD0iMSI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIj48YTp0aW50IHZhbD0iNzUwMDAiLz48L2E6c2NoZW1lQ2xyPjwvY3M6"
    "bG5SZWY+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48"
    "YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOnNwUHI+PGE6bG4+PGE6cm91bmQvPjwvYTpsbj48L2NzOnNw"
    "UHI+PGNzOmRlZlJQciBzej0iMTAwMCIga2Vybj0iMTIwMCIvPjwvY3M6Y2F0ZWdvcnlBeGlzPjxjczpjaGFydEFyZWEgbW9kcz0i"
    "YWxsb3dOb0ZpbGxPdmVycmlkZSBhbGxvd05vTGluZU92ZXJyaWRlIj48Y3M6bG5SZWYgaWR4PSIxIj48YTpzY2hlbWVDbHIgdmFs"
    "PSJ0eDEiPjxhOnRpbnQgdmFsPSI3NTAwMCIvPjwvYTpzY2hlbWVDbHI+PC9jczpsblJlZj48Y3M6ZmlsbFJlZiBpZHg9IjEiPjxh"
    "OnNjaGVtZUNsciB2YWw9ImJnMSIvPjwvY3M6ZmlsbFJlZj48Y3M6ZWZmZWN0UmVmIGlkeD0iMCIvPjxjczpmb250UmVmIGlkeD0i"
    "bWlub3IiPjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6Zm9udFJlZj48Y3M6c3BQcj48YTpsbj48YTpyb3VuZC8+PC9hOmxu"
    "PjwvY3M6c3BQcj48Y3M6ZGVmUlByIHN6PSIxMDAwIiBrZXJuPSIxMjAwIi8+PC9jczpjaGFydEFyZWE+PGNzOmRhdGFMYWJlbD48"
    "Y3M6bG5SZWYgaWR4PSIwIi8+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBp"
    "ZHg9Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOmRlZlJQciBzej0iMTAwMCIga2Vybj0i"
    "MTIwMCIvPjwvY3M6ZGF0YUxhYmVsPjxjczpkYXRhTGFiZWxDYWxsb3V0PjxjczpsblJlZiBpZHg9IjAiLz48Y3M6ZmlsbFJlZiBp"
    "ZHg9IjAiLz48Y3M6ZWZmZWN0UmVmIGlkeD0iMCIvPjxjczpmb250UmVmIGlkeD0ibWlub3IiPjxhOnNjaGVtZUNsciB2YWw9ImRr"
    "MSIvPjwvY3M6Zm9udFJlZj48Y3M6c3BQcj48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0ibHQxIi8+PC9hOnNvbGlkRmls"
    "bD48YTpsbj48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0iZGsxIj48YTpsdW1Nb2QgdmFsPSI2NTAwMCIvPjxhOmx1bU9m"
    "ZiB2YWw9IjM1MDAwIi8+PC9hOnNjaGVtZUNscj48L2E6c29saWRGaWxsPjwvYTpsbj48L2NzOnNwUHI+PGNzOmRlZlJQciBzej0i"
    "MTAwMCIga2Vybj0iMTIwMCIvPjwvY3M6ZGF0YUxhYmVsQ2FsbG91dD48Y3M6ZGF0YVBvaW50PjxjczpsblJlZiBpZHg9IjAiLz48"
    "Y3M6ZmlsbFJlZiBpZHg9IjEiPjxjczpzdHlsZUNsciB2YWw9ImF1dG8iLz48L2NzOmZpbGxSZWY+PGNzOmVmZmVjdFJlZiBpZHg9"
    "IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PC9jczpkYXRh"
    "UG9pbnQ+PGNzOmRhdGFQb2ludDNEPjxjczpsblJlZiBpZHg9IjAiLz48Y3M6ZmlsbFJlZiBpZHg9IjEiPjxjczpzdHlsZUNsciB2"
    "YWw9ImF1dG8iLz48L2NzOmZpbGxSZWY+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpz"
    "Y2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PC9jczpkYXRhUG9pbnQzRD48Y3M6ZGF0YVBvaW50TGluZT48Y3M6bG5S"
    "ZWYgaWR4PSIxIj48Y3M6c3R5bGVDbHIgdmFsPSJhdXRvIi8+PC9jczpsblJlZj48Y3M6bGluZVdpZHRoU2NhbGU+MzwvY3M6bGlu"
    "ZVdpZHRoU2NhbGU+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1p"
    "bm9yIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOnNwUHI+PGE6bG4gY2FwPSJybmQiPjxhOnJvdW5k"
    "Lz48L2E6bG4+PC9jczpzcFByPjwvY3M6ZGF0YVBvaW50TGluZT48Y3M6ZGF0YVBvaW50TWFya2VyPjxjczpsblJlZiBpZHg9IjEi"
    "PjxjczpzdHlsZUNsciB2YWw9ImF1dG8iLz48L2NzOmxuUmVmPjxjczpmaWxsUmVmIGlkeD0iMSI+PGNzOnN0eWxlQ2xyIHZhbD0i"
    "YXV0byIvPjwvY3M6ZmlsbFJlZj48Y3M6ZWZmZWN0UmVmIGlkeD0iMCIvPjxjczpmb250UmVmIGlkeD0ibWlub3IiPjxhOnNjaGVt"
    "ZUNsciB2YWw9InR4MSIvPjwvY3M6Zm9udFJlZj48Y3M6c3BQcj48YTpsbj48YTpyb3VuZC8+PC9hOmxuPjwvY3M6c3BQcj48L2Nz"
    "OmRhdGFQb2ludE1hcmtlcj48Y3M6ZGF0YVBvaW50TWFya2VyTGF5b3V0Lz48Y3M6ZGF0YVBvaW50V2lyZWZyYW1lPjxjczpsblJl"
    "ZiBpZHg9IjEiPjxjczpzdHlsZUNsciB2YWw9ImF1dG8iLz48L2NzOmxuUmVmPjxjczpmaWxsUmVmIGlkeD0iMCIvPjxjczplZmZl"
    "Y3RSZWYgaWR4PSIwIi8+PGNzOmZvbnRSZWYgaWR4PSJtaW5vciI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmb250UmVm"
    "PjxjczpzcFByPjxhOmxuPjxhOnJvdW5kLz48L2E6bG4+PC9jczpzcFByPjwvY3M6ZGF0YVBvaW50V2lyZWZyYW1lPjxjczpkYXRh"
    "VGFibGU+PGNzOmxuUmVmIGlkeD0iMSI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIj48YTp0aW50IHZhbD0iNzUwMDAiLz48L2E6c2No"
    "ZW1lQ2xyPjwvY3M6bG5SZWY+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBp"
    "ZHg9Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOnNwUHI+PGE6bG4+PGE6cm91bmQvPjwv"
    "YTpsbj48L2NzOnNwUHI+PGNzOmRlZlJQciBzej0iMTAwMCIga2Vybj0iMTIwMCIvPjwvY3M6ZGF0YVRhYmxlPjxjczpkb3duQmFy"
    "PjxjczpsblJlZiBpZHg9IjEiPjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6bG5SZWY+PGNzOmZpbGxSZWYgaWR4PSIxIj48"
    "YTpzY2hlbWVDbHIgdmFsPSJkazEiPjxhOnRpbnQgdmFsPSI5NTAwMCIvPjwvYTpzY2hlbWVDbHI+PC9jczpmaWxsUmVmPjxjczpl"
    "ZmZlY3RSZWYgaWR4PSIwIi8+PGNzOmZvbnRSZWYgaWR4PSJtaW5vciI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmb250"
    "UmVmPjxjczpzcFByPjxhOmxuPjxhOnJvdW5kLz48L2E6bG4+PC9jczpzcFByPjwvY3M6ZG93bkJhcj48Y3M6ZHJvcExpbmU+PGNz"
    "OmxuUmVmIGlkeD0iMSI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpsblJlZj48Y3M6ZmlsbFJlZiBpZHg9IjAiLz48Y3M6"
    "ZWZmZWN0UmVmIGlkeD0iMCIvPjxjczpmb250UmVmIGlkeD0ibWlub3IiPjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6Zm9u"
    "dFJlZj48Y3M6c3BQcj48YTpsbj48YTpyb3VuZC8+PC9hOmxuPjwvY3M6c3BQcj48L2NzOmRyb3BMaW5lPjxjczplcnJvckJhcj48"
    "Y3M6bG5SZWYgaWR4PSIxIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmxuUmVmPjxjczpmaWxsUmVmIGlkeD0iMSI+PGE6"
    "c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmaWxsUmVmPjxjczplZmZlY3RSZWYgaWR4PSIwIi8+PGNzOmZvbnRSZWYgaWR4PSJt"
    "aW5vciI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmb250UmVmPjxjczpzcFByPjxhOmxuPjxhOnJvdW5kLz48L2E6bG4+"
    "PC9jczpzcFByPjwvY3M6ZXJyb3JCYXI+PGNzOmZsb29yPjxjczpsblJlZiBpZHg9IjEiPjxhOnNjaGVtZUNsciB2YWw9InR4MSI+"
    "PGE6dGludCB2YWw9Ijc1MDAwIi8+PC9hOnNjaGVtZUNscj48L2NzOmxuUmVmPjxjczpmaWxsUmVmIGlkeD0iMCIvPjxjczplZmZl"
    "Y3RSZWYgaWR4PSIwIi8+PGNzOmZvbnRSZWYgaWR4PSJtaW5vciI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmb250UmVm"
    "PjxjczpzcFByPjxhOmxuPjxhOnJvdW5kLz48L2E6bG4+PC9jczpzcFByPjwvY3M6Zmxvb3I+PGNzOmdyaWRsaW5lTWFqb3I+PGNz"
    "OmxuUmVmIGlkeD0iMSI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIj48YTp0aW50IHZhbD0iNzUwMDAiLz48L2E6c2NoZW1lQ2xyPjwv"
    "Y3M6bG5SZWY+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9y"
    "Ij48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOnNwUHI+PGE6bG4+PGE6cm91bmQvPjwvYTpsbj48L2Nz"
    "OnNwUHI+PC9jczpncmlkbGluZU1ham9yPjxjczpncmlkbGluZU1pbm9yPjxjczpsblJlZiBpZHg9IjEiPjxhOnNjaGVtZUNsciB2"
    "YWw9InR4MSI+PGE6dGludCB2YWw9IjUwMDAwIi8+PC9hOnNjaGVtZUNscj48L2NzOmxuUmVmPjxjczpmaWxsUmVmIGlkeD0iMCIv"
    "PjxjczplZmZlY3RSZWYgaWR4PSIwIi8+PGNzOmZvbnRSZWYgaWR4PSJtaW5vciI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9j"
    "czpmb250UmVmPjxjczpzcFByPjxhOmxuPjxhOnJvdW5kLz48L2E6bG4+PC9jczpzcFByPjwvY3M6Z3JpZGxpbmVNaW5vcj48Y3M6"
    "aGlMb0xpbmU+PGNzOmxuUmVmIGlkeD0iMSI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpsblJlZj48Y3M6ZmlsbFJlZiBp"
    "ZHg9IjAiLz48Y3M6ZWZmZWN0UmVmIGlkeD0iMCIvPjxjczpmb250UmVmIGlkeD0ibWlub3IiPjxhOnNjaGVtZUNsciB2YWw9InR4"
    "MSIvPjwvY3M6Zm9udFJlZj48Y3M6c3BQcj48YTpsbj48YTpyb3VuZC8+PC9hOmxuPjwvY3M6c3BQcj48L2NzOmhpTG9MaW5lPjxj"
    "czpsZWFkZXJMaW5lPjxjczpsblJlZiBpZHg9IjEiPjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6bG5SZWY+PGNzOmZpbGxS"
    "ZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFs"
    "PSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOnNwUHI+PGE6bG4+PGE6cm91bmQvPjwvYTpsbj48L2NzOnNwUHI+PC9jczpsZWFkZXJM"
    "aW5lPjxjczpsZWdlbmQ+PGNzOmxuUmVmIGlkeD0iMCIvPjxjczpmaWxsUmVmIGlkeD0iMCIvPjxjczplZmZlY3RSZWYgaWR4PSIw"
    "Ii8+PGNzOmZvbnRSZWYgaWR4PSJtaW5vciI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmb250UmVmPjxjczpkZWZSUHIg"
    "c3o9IjEwMDAiIGtlcm49IjEyMDAiLz48L2NzOmxlZ2VuZD48Y3M6cGxvdEFyZWEgbW9kcz0iYWxsb3dOb0ZpbGxPdmVycmlkZSBh"
    "bGxvd05vTGluZU92ZXJyaWRlIj48Y3M6bG5SZWYgaWR4PSIwIi8+PGNzOmZpbGxSZWYgaWR4PSIxIj48YTpzY2hlbWVDbHIgdmFs"
    "PSJiZzEiLz48L2NzOmZpbGxSZWY+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpzY2hl"
    "bWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PC9jczpwbG90QXJlYT48Y3M6cGxvdEFyZWEzRD48Y3M6bG5SZWYgaWR4PSIw"
    "Ii8+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpz"
    "Y2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PC9jczpwbG90QXJlYTNEPjxjczpzZXJpZXNBeGlzPjxjczpsblJlZiBp"
    "ZHg9IjEiPjxhOnNjaGVtZUNsciB2YWw9InR4MSI+PGE6dGludCB2YWw9Ijc1MDAwIi8+PC9hOnNjaGVtZUNscj48L2NzOmxuUmVm"
    "PjxjczpmaWxsUmVmIGlkeD0iMCIvPjxjczplZmZlY3RSZWYgaWR4PSIwIi8+PGNzOmZvbnRSZWYgaWR4PSJtaW5vciI+PGE6c2No"
    "ZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmb250UmVmPjxjczpzcFByPjxhOmxuPjxhOnJvdW5kLz48L2E6bG4+PC9jczpzcFByPjxj"
    "czpkZWZSUHIgc3o9IjEwMDAiIGtlcm49IjEyMDAiLz48L2NzOnNlcmllc0F4aXM+PGNzOnNlcmllc0xpbmU+PGNzOmxuUmVmIGlk"
    "eD0iMSI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpsblJlZj48Y3M6ZmlsbFJlZiBpZHg9IjAiLz48Y3M6ZWZmZWN0UmVm"
    "IGlkeD0iMCIvPjxjczpmb250UmVmIGlkeD0ibWlub3IiPjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6Zm9udFJlZj48Y3M6"
    "c3BQcj48YTpsbj48YTpyb3VuZC8+PC9hOmxuPjwvY3M6c3BQcj48L2NzOnNlcmllc0xpbmU+PGNzOnRpdGxlPjxjczpsblJlZiBp"
    "ZHg9IjAiLz48Y3M6ZmlsbFJlZiBpZHg9IjAiLz48Y3M6ZWZmZWN0UmVmIGlkeD0iMCIvPjxjczpmb250UmVmIGlkeD0ibWlub3Ii"
    "PjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6Zm9udFJlZj48Y3M6ZGVmUlByIHN6PSIxODAwIiBiPSIxIiBrZXJuPSIxMjAw"
    "Ii8+PC9jczp0aXRsZT48Y3M6dHJlbmRsaW5lPjxjczpsblJlZiBpZHg9IjEiPjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6"
    "bG5SZWY+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48"
    "YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOnNwUHI+PGE6bG4gY2FwPSJybmQiPjxhOnJvdW5kLz48L2E6"
    "bG4+PC9jczpzcFByPjwvY3M6dHJlbmRsaW5lPjxjczp0cmVuZGxpbmVMYWJlbD48Y3M6bG5SZWYgaWR4PSIwIi8+PGNzOmZpbGxS"
    "ZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFs"
    "PSJ0eDEiLz48L2NzOmZvbnRSZWY+PGNzOmRlZlJQciBzej0iMTAwMCIga2Vybj0iMTIwMCIvPjwvY3M6dHJlbmRsaW5lTGFiZWw+"
    "PGNzOnVwQmFyPjxjczpsblJlZiBpZHg9IjEiPjxhOnNjaGVtZUNsciB2YWw9InR4MSIvPjwvY3M6bG5SZWY+PGNzOmZpbGxSZWYg"
    "aWR4PSIxIj48YTpzY2hlbWVDbHIgdmFsPSJkazEiPjxhOnRpbnQgdmFsPSI1MDAwIi8+PC9hOnNjaGVtZUNscj48L2NzOmZpbGxS"
    "ZWY+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48"
    "L2NzOmZvbnRSZWY+PGNzOnNwUHI+PGE6bG4+PGE6cm91bmQvPjwvYTpsbj48L2NzOnNwUHI+PC9jczp1cEJhcj48Y3M6dmFsdWVB"
    "eGlzPjxjczpsblJlZiBpZHg9IjEiPjxhOnNjaGVtZUNsciB2YWw9InR4MSI+PGE6dGludCB2YWw9Ijc1MDAwIi8+PC9hOnNjaGVt"
    "ZUNscj48L2NzOmxuUmVmPjxjczpmaWxsUmVmIGlkeD0iMCIvPjxjczplZmZlY3RSZWYgaWR4PSIwIi8+PGNzOmZvbnRSZWYgaWR4"
    "PSJtaW5vciI+PGE6c2NoZW1lQ2xyIHZhbD0idHgxIi8+PC9jczpmb250UmVmPjxjczpzcFByPjxhOmxuPjxhOnJvdW5kLz48L2E6"
    "bG4+PC9jczpzcFByPjxjczpkZWZSUHIgc3o9IjEwMDAiIGtlcm49IjEyMDAiLz48L2NzOnZhbHVlQXhpcz48Y3M6d2FsbD48Y3M6"
    "bG5SZWYgaWR4PSIwIi8+PGNzOmZpbGxSZWYgaWR4PSIwIi8+PGNzOmVmZmVjdFJlZiBpZHg9IjAiLz48Y3M6Zm9udFJlZiBpZHg9"
    "Im1pbm9yIj48YTpzY2hlbWVDbHIgdmFsPSJ0eDEiLz48L2NzOmZvbnRSZWY+PC9jczp3YWxsPjwvY3M6Y2hhcnRTdHlsZT4="
)

_COLORS1_XML_B64 = (
    "PGNzOmNvbG9yU3R5bGUgeG1sbnM6Y3M9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL2RyYXdpbmcvMjAxMi9j"
    "aGFydFN0eWxlIiB4bWxuczphPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvZHJhd2luZ21sLzIwMDYvbWFpbiIg"
    "bWV0aD0iYWNyb3NzTGluZWFyIiBpZD0iMiI+PGE6c2NoZW1lQ2xyIHZhbD0iYWNjZW50MSIvPjxhOnNjaGVtZUNsciB2YWw9ImFj"
    "Y2VudDIiLz48YTpzY2hlbWVDbHIgdmFsPSJhY2NlbnQzIi8+PGE6c2NoZW1lQ2xyIHZhbD0iYWNjZW50NCIvPjxhOnNjaGVtZUNs"
    "ciB2YWw9ImFjY2VudDUiLz48YTpzY2hlbWVDbHIgdmFsPSJhY2NlbnQ2Ii8+PC9jczpjb2xvclN0eWxlPg=="
)


def _chartex_xml(sheet_name: str, chart_title: str, series1_name: str, series2_name: str) -> str:
    title_esc = xml_escape(chart_title)
    s1_esc = xml_escape(series1_name)
    s2_esc = xml_escape(series2_name)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">'
        '<cx:chartData>'
        '<cx:data id="0">'
        '<cx:strDim type="cat"><cx:f>_xlchart.v1.5</cx:f></cx:strDim>'
        '<cx:numDim type="val"><cx:f>_xlchart.v1.7</cx:f></cx:numDim>'
        '</cx:data>'
        '<cx:data id="1">'
        '<cx:strDim type="cat"><cx:f>_xlchart.v1.5</cx:f></cx:strDim>'
        '<cx:numDim type="val"><cx:f>_xlchart.v1.9</cx:f></cx:numDim>'
        '</cx:data>'
        '</cx:chartData>'
        '<cx:chart>'
        f'<cx:title pos="t" align="ctr" overlay="0">'
        f'<cx:tx><cx:txData><cx:v>{title_esc}</cx:v></cx:txData></cx:tx>'
        '</cx:title>'
        '<cx:plotArea><cx:plotAreaRegion>'
        f'<cx:series layoutId="boxWhisker" uniqueId="{{D6C3A363-B570-4F4B-8052-E0139D162600}}">'
        f'<cx:tx><cx:txData><cx:f>_xlchart.v1.6</cx:f><cx:v>{s1_esc}</cx:v></cx:txData></cx:tx>'
        '<cx:dataId val="0" />'
        '<cx:layoutPr><cx:visibility meanLine="1" meanMarker="0" nonoutliers="1" outliers="1" />'
        '<cx:statistics quartileMethod="exclusive" /></cx:layoutPr>'
        '</cx:series>'
        f'<cx:series layoutId="boxWhisker" uniqueId="{{50F7E9F9-D6B5-4CD6-AD62-CBAF7E464DCA}}">'
        f'<cx:tx><cx:txData><cx:f>_xlchart.v1.8</cx:f><cx:v>{s2_esc}</cx:v></cx:txData></cx:tx>'
        '<cx:dataId val="1" />'
        '<cx:layoutPr><cx:visibility meanLine="1" meanMarker="0" /><cx:statistics quartileMethod="exclusive" /></cx:layoutPr>'
        '</cx:series>'
        '</cx:plotAreaRegion>'
        '<cx:axis id="0"><cx:catScaling gapWidth="0.330000013" /><cx:tickLabels /></cx:axis>'
        '<cx:axis id="1"><cx:valScaling /><cx:tickLabels /></cx:axis>'
        '</cx:plotArea>'
        '<cx:legend pos="t" align="ctr" overlay="0" />'
        '</cx:chart></cx:chartSpace>'
    )


def _drawing_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<xdr:oneCellAnchor>'
        '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:ext cx="10080000" cy="4320000"/>'
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">'
        '<xdr:graphicFrame macro="">'
        '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">'
        '<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
        '</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>'
        '<mc:Fallback>'
        '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
        '<xdr:spPr><a:xfrm><a:off x="2590800" y="180975"/><a:ext cx="10080000" cy="4320000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill>'
        '<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>'
        '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="de-CH" sz="1100"/>'
        '<a:t>Dieses Diagramm wird in Ihrer Excel-Version nicht unterstuetzt.</a:t>'
        '</a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>'
        '</mc:AlternateContent><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>'
    )


def inject_box_whisker_chart(xlsx_path: str, sheet_name: str, n_data_rows: int,
                              chart_title: str, series1_name: str = "Turbiniert_MW",
                              series2_name: str = "Hochgepumpt_MW") -> None:
    """Postprocessing-Schritt: oeffnet die bereits gespeicherte xlsx-Datei
    und fuegt das Box-Whisker-Diagramm ein. n_data_rows = Anzahl
    Datenzeilen OHNE Headerzeile (also len(df)). Geht von Spalten
    A=timestamp, B=series1, C=series2 aus, Header in Zeile 1."""
    last_row = n_data_rows + 1  # Headerzeile mitgezaehlt
    tmp_path = xlsx_path + ".tmp"

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        names = zin.namelist()

        workbook_xml = zin.read("xl/workbook.xml").decode("utf-8")
        defined_names = (
            f'<definedNames>'
            f'<definedName name="_xlchart.v1.5" hidden="1">{sheet_name}!$A$2:$A${last_row}</definedName>'
            f'<definedName name="_xlchart.v1.6" hidden="1">{sheet_name}!$B$1</definedName>'
            f'<definedName name="_xlchart.v1.7" hidden="1">{sheet_name}!$B$2:$B${last_row}</definedName>'
            f'<definedName name="_xlchart.v1.8" hidden="1">{sheet_name}!$C$1</definedName>'
            f'<definedName name="_xlchart.v1.9" hidden="1">{sheet_name}!$C$2:$C${last_row}</definedName>'
            f'</definedNames>'
        )

        # Regex statt exaktem String-Vergleich: openpyxl schreibt das
        # selbstschliessende Tag mal als "<definedNames/>", mal als
        # "<definedNames />" (mit Leerzeichen) - ein zu enger String-
        # Vergleich hat genau das beim ersten Versuch uebersehen und so
        # ein zweites, ungueltiges <definedNames>-Element erzeugt. Jetzt
        # robust per Regex, die beide Schreibweisen abdeckt.
        import re as _re
        self_closing_pattern = _re.compile(r"<definedNames\s*/>")
        open_close_pattern = _re.compile(r"<definedNames\s*>.*?</definedNames\s*>", _re.DOTALL)

        if open_close_pattern.search(workbook_xml):
            insert_block = defined_names.replace("<definedNames>", "").replace("</definedNames>", "")
            workbook_xml = open_close_pattern.sub(
                lambda m: m.group(0).replace("</definedNames>", insert_block + "</definedNames>"),
                workbook_xml, count=1,
            )
        elif self_closing_pattern.search(workbook_xml):
            workbook_xml = self_closing_pattern.sub(defined_names, workbook_xml, count=1)
        elif "<calcPr" in workbook_xml:
            workbook_xml = workbook_xml.replace("<calcPr", defined_names + "<calcPr")
        else:
            workbook_xml = workbook_xml.replace("</workbook>", defined_names + "</workbook>")

        # Sicherheitsnetz: falls trotzdem mehr als ein <definedNames>-
        # Element entstanden ist, das ist ungueltiges OOXML und wuerde
        # Excel zum Ablehnen der Datei bringen - lieber laut scheitern
        # als eine kaputte Datei stillschweigend ausliefern.
        if workbook_xml.count("<definedNames") > 1:
            raise RuntimeError(
                "boxwhisker_injector: mehrere <definedNames>-Elemente nach "
                "dem Patchen gefunden - das wuerde eine ungueltige xlsx-"
                "Datei erzeugen. Bitte xl/workbook.xml manuell pruefen."
            )

        content_types = zin.read("[Content_Types].xml").decode("utf-8")
        new_overrides = (
            '<Override PartName="/xl/drawings/drawing1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            '<Override PartName="/xl/charts/chartEx1.xml" '
            'ContentType="application/vnd.ms-office.chartex+xml"/>'
            '<Override PartName="/xl/charts/style1.xml" '
            'ContentType="application/vnd.ms-office.chartstyle+xml"/>'
            '<Override PartName="/xl/charts/colors1.xml" '
            'ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>'
        )
        content_types = content_types.replace("</Types>", new_overrides + "</Types>")

        sheet_part = "xl/worksheets/sheet1.xml"
        sheet_xml = zin.read(sheet_part).decode("utf-8")
        if "<drawing " not in sheet_xml:
            # Wichtiger Fix: das worksheet-Root-Element deklariert von Haus
            # aus KEIN "r:"-Namespace-Praefix (das gibt es nur im Kontext
            # anderer Teile wie workbook.xml). Ohne explizite Deklaration
            # hier ist <drawing r:id="..."/> ungueltiges XML ("unbound
            # prefix") und Excel/andere Tools lehnen die Datei komplett ab.
            if 'xmlns:r=' not in sheet_xml.split('>', 1)[0]:
                sheet_xml = sheet_xml.replace(
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
                    1,
                )
            sheet_xml = sheet_xml.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>')

        rels_part = "xl/worksheets/_rels/sheet1.xml.rels"
        sheet_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
            'Target="../drawings/drawing1.xml"/></Relationships>'
        )
        if rels_part in names:
            existing = zin.read(rels_part).decode("utf-8")
            if "drawing1.xml" not in existing:
                sheet_rels_xml = existing.replace(
                    "</Relationships>",
                    '<Relationship Id="rIdDrawing" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                    'Target="../drawings/drawing1.xml"/></Relationships>'
                )
            else:
                sheet_rels_xml = existing

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/workbook.xml":
                    zout.writestr(item, workbook_xml)
                elif item.filename == "[Content_Types].xml":
                    zout.writestr(item, content_types)
                elif item.filename == sheet_part:
                    zout.writestr(item, sheet_xml)
                elif item.filename == rels_part:
                    pass
                else:
                    zout.writestr(item, zin.read(item.filename))

            zout.writestr(rels_part, sheet_rels_xml)
            zout.writestr(
                "xl/drawings/_rels/drawing1.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" '
                'Target="../charts/chartEx1.xml"/></Relationships>'
            )
            zout.writestr(
                "xl/charts/_rels/chartEx1.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId2" '
                'Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" '
                'Target="colors1.xml"/>'
                '<Relationship Id="rId1" '
                'Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" '
                'Target="style1.xml"/></Relationships>'
            )
            zout.writestr("xl/drawings/drawing1.xml", _drawing_xml())
            zout.writestr(
                "xl/charts/chartEx1.xml",
                _chartex_xml(sheet_name, chart_title, series1_name, series2_name),
            )
            zout.writestr("xl/charts/style1.xml", base64.b64decode(_STYLE1_XML_B64))
            zout.writestr("xl/charts/colors1.xml", base64.b64decode(_COLORS1_XML_B64))

    shutil.move(tmp_path, xlsx_path)

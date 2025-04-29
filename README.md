# mejllistorUHR

Tar en textfil med mejladresser och rättar enkla saker och markerar möjliga fel som kan fixas manuellt. Spottar ut resultatet i en excel-fil för att enklare kunna hanteras. Du behöver ha **Python** och paketet **openpyxl**installerat på datorn för att köra programmet.

Instruktioner:
1. Scanna ca 100 sidor i taget och kör OCR (kräver adobe)
2. Exportera OCR-filen till en textfil (text Tillgänglig)
3. Lägg Python-filen och textfilerna i samma mapp
4. Öppna terminalen i denna mapp och skriv in följande (se till att ha Python installerat på datorn):
```
python3 mejllistorUHR_txt.py <namn>.txt
```

Programmet raderar själv alla mellanslag och blanka rader, samt självrättar en del uppenbara OCR fel som exempelvis:
- ".eom" till ".com"
- "maiLcom" till "mail.com"
- "autlaak" till "outlook"
- etc.

Adresser som har ett inkorrekt format markeras med orange färg. Detta kan vara otillåtna tecken eller konstiga radbryt.

Adresser med misstänkta OCR fel markeras med gult. Bland annat hamnar här adresser innehållandes
- 'S' som ofta ska vara '5'
- 'O' som ofta ska vara '0'
-  olika fall där 'l' kan ha tolkats som '1' och tvärtom.
-  osv.


Källkoden hittas här: https://github.com/croos90/mejllistorUHR

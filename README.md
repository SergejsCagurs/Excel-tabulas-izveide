Detalizēts Projekta Uzdevums:

Projekta mērķis ir izveidot vienkāršu sistēmu krājumu pārvaldībai, izmantojot programmēšanas valodu Python. Šī sistēma nodrošinās iespēju pievienot jaunus produktus, atjaunināt esošu produktu informāciju un iegūt informāciju par konkrētu produktu, visu to saglabājot Excel failā. 




Izmantotās Python Bibliotēkas un To Izstrādes Mērķis:

openpyxl: Šī bibliotēka tiek izmantota darbam ar Excel failiem.
os: Bibliotēka nodrošina iespēju strādāt ar operētājsistēmu. Tā tiek izmantota, lai pārbaudītu, vai fails pastāv, un, ja nepieciešams, izveidotu jaunu.
input: Izmantojot šo funkciju, programma saņem lietotāja ievades.




Lietotājam tiek piedāvāta tekstovā interfeisa saskarne, kurā var izvēlēties dažādas darbības, izmantojot ciparus no 1 līdz 4.
Pievienot Produktu (Izvēle 1)
Atjaunināt Produktu (Izvēle 2) 
Iegūt Informāciju par Produktu (Izvēle 3) 
Iziet no Programmas (Izvēle 4)






Detalizēts koda apraksts:

__init__ metode: Inicializē klasi, nosakot Excel faila nosaukumu, ar kuru tiks strādāts. Ja fails nepastāv, izveido jaunu tukšu failu, izsaucot create_new_file metodi, un pēc tam ielādē datus ar load_data metodi.
create_new_file metode: Izveido jaunu Excel darbgrāmatu ar nosaukumu "Inventory" un pievieno kolonnu galvenes (ID, Produkta nosaukums, Daudzums, Vienības cena).
load_data metode: Ielādē datus no esošā Excel faila, nosakot aktīvo lapu.
save_data metode: Saglabā izmaiņas Excel failā.
add_item metode: Pievieno jaunu produktu lapai un saglabā izmaiņas failā.
update_item metode: Atjaunina esošā produkta informāciju, ja tāda ir norādīta, un saglabā izmaiņas failā.
get_item_info metode: Atgriež informāciju par konkrētu produktu pēc ID.
user_interface funkcija:

Izveido jaunu InventoryManager objektu, norādot faila nosaukumu "inventory.xlsx".
Bezgalīgā ciklā piedāvā lietotājam izvēlēties darbību: pievienot produktu, atjaunināt produktu, iegūt informāciju par produktu vai iziet no programmas.
Atbilstoši lietotāja izvēlei izsauc attiecīgās InventoryManager metodes.
Galvenā programmas izpilde:

Izsauc user_interface funkciju, uzsākot darbību.


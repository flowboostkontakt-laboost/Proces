Jasne, skupmy się na samym "silniku" działającym lokalnie. To najmądrzejsze podejście – najpierw upewnimy się, że Python bezbłędnie czyta PDF i prawidłowo uzupełnia Twój arkusz Excel, a dopiero później (jak już będziesz gotowy) dołożymy do tego przesyłanie do chmury.

Oto gotowa instrukcja, którą możesz wkleić do Claude (lub innego modelu AI), aby wygenerował Ci gotowy skrypt:

---

### **Brief techniczny dla Claude – Wersja Lokalna**

**Temat:** Skrypt Python do automatycznego wypełniania instrukcji BHP z plików PDF (SDS).

**Cel:** Skrypt ma pobierać wszystkie pliki PDF z folderu `DANE_WEJSCOWE`, wyciągać z nich kluczowe informacje zgodnie ze strukturą Karty Charakterystyki i wpisywać je do szablonu Excel `Instrukcja chemia.xlsx`, zapisując wynik w folderze `GOTOWE_INSTRUKCJE`.

**Szczegóły techniczne ekstrakcji (na podstawie pliku LOCTITE 620):**

1. **Biblioteka:** Użyj `pdfplumber` do tekstu i `openpyxl` do edycji Excela.  
2. **Pola do wyciągnięcia (Regex):**  
   * **Nazwa produktu:** Sekcja 1.1.  
   * **Producent:** Sekcja 1.3 (np. Henkel Polska).  
   * **Data aktualizacji:** Z nagłówka dokumentu ("Aktualizacja: ...").  
   * **Piktogramy:** Przeszukaj sekcję 2.2 pod kątem kodów H (np. H317, H319, H412). Stwórz w kodzie słownik mapujący te kody na nazwy piktogramów (np. H317 \-\> "GHS07 Wykrzyknik").  
   * **Środki ochrony (Sekcja 8.2):** Osobno wyciągnij tekst dla: Ochrona rąk, Ochrona oczu, Ochrona skóry, Ochrona dróg oddechowych.  
   * **Pierwsza pomoc:** Sekcja 4.1 (Drogi oddechowe, kontakt ze skórą, kontakt z oczami).  
   * **Pożar:** Sekcja 5.1 (Środki gaśnicze).

**Logika zapisu w Excelu (na podstawie szablonu):**

* Wpisz wyciągnięte dane w odpowiednie komórki (np. Nazwa \-\> D8, Producent \-\> B8).  
* W sekcji "Zagrożenia" wypisz piktogramy i zwroty H znalezione w sekcji 2.2.  
* Dla każdego PDF-a utwórz osobny plik Excel o nazwie `Instrukcja_BHP_[Nazwa_Produktu].xlsx`.

**Wymagania dodatkowe:**

* Kod powinien tworzyć foldery `DANE_WEJSCOWE` i `GOTOWE_INSTRUKCJE`, jeśli nie istnieją.  
* Dodaj komentarze w kodzie, abym wiedział, jak zmienić adres komórki w Excelu, jeśli przesunę coś w szablonie.


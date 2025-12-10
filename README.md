# jsonToExel

Proste narzędzie z graficznym interfejsem użytkownika (GUI) do automatycznej konwersji wielu plików JSON znajdujących się w wybranym folderze do pojedynczego, sformatowanego pliku Excel (.xlsx).

Każdy plik JSON jest zapisywany jako osobny arkusz w pliku wynikowym, a kluczowe dane są ładnie sformatowane.

# Powstało z potrzeby firmy na przekształcanie plików JSON na tebele exel

---

## Jak uruchomić aplikację

Aplikacja wymaga Pythona w wersji 3.x oraz kilku zewnętrznych bibliotek. Zaleca się używanie środowiska wirtualnego (`venv`).

### Klonowanie i Uruchomienie

```bash
# Sklonuj repozytorium (jeśli używasz Git)
git clone <adres_repozytorium>
cd jsonToExel
```

# Utwórz środowisko wirtualne

python3 -m venv venv

# Aktywuj środowisko wirtualne

source venv/bin/activate

# Instalacja zależnosci

pip install pandas openpyxl

# Komenda uruchomieniowa

python gui.py

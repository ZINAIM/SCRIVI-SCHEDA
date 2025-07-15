import json
import webbrowser
import time
import os

def aggiorna_link(mancanti_file="mancanti.txt", json_file="video_links.json"):
    # Carica o crea dizionario link esistenti
    if os.path.exists(json_file):
        with open(json_file, "r", encoding="utf-8") as f:
            video_links = json.load(f)
    else:
        video_links = {}

    # Leggi gli esercizi mancanti
    if not os.path.exists(mancanti_file):
        print(f"File {mancanti_file} non trovato.")
        return

    with open(mancanti_file, "r", encoding="utf-8") as f:
        mancanti = [line.strip() for line in f if line.strip()]

    if not mancanti:
        print("Nessun esercizio da aggiornare!")
        return

    print(f"Trovo {len(mancanti)} esercizi mancanti da aggiornare.\n")

    for esercizio in mancanti:
        # Cerca l'esercizio su YouTube (apri nel browser)
        query = esercizio.replace(" ", "+")
        url_ricerca = f"https://www.youtube.com/results?search_query={query}"
        print(f"\nApro ricerca per '{esercizio}' su YouTube...")
        webbrowser.open_new_tab(url_ricerca)

        # Chiedi all'utente di inserire il link
        link = input(f"Inserisci il link YouTube corretto per '{esercizio}': ").strip()
        if link:
            video_links[esercizio] = link
        else:
            print("Link vuoto, salto questo esercizio.")

        # (Facoltativo) Piccola pausa per evitare troppi tab aperti subito
        time.sleep(1)

    # Salva aggiornamenti su JSON
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump(video_links, f, ensure_ascii=False, indent=2)

    print(f"\nAggiornamento completato. Link salvati in '{json_file}'.")
if __name__ == "__main__":
    aggiorna_link()

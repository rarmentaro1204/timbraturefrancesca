# App timbrature di Francesca

Le timbrature finiscono in un Google Sheet (foglio "Timbrature": data/ora, giorno, ora, tipo).
L'orario è quello del server Google, non del telefono.

## Cosa fare (5 minuti, una volta sola)

1. Il foglio "Timbrature Francesca" è già creato su Google Drive (ID già inserito in `Code.gs`).
2. Menu **Estensioni → Apps Script**.
3. Nel file `Codice.gs` cancella tutto e incolla il contenuto di `Code.gs`.
4. Clicca **+** accanto a "File" → **HTML**, chiamalo `index` e incolla il contenuto di `index.html`.
5. **Distribuisci → Nuova distribuzione → tipo "App web"**
   - Esegui come: **Me**
   - Chi ha accesso: **Chiunque** (il link è il "segreto": non diffonderlo)
   - Autorizza quando richiesto (Avanzate → Vai al progetto).
6. Copia l'URL che finisce con `/exec` e mandalo a Francesca.
7. Su iPhone/Android: apri il link → "Aggiungi a schermata Home" per avere l'icona.

Se modifichi il codice: Distribuisci → Gestisci distribuzioni → matita → Nuova versione.

## Come funziona
- Un solo bottone: alterna ENTRATA / USCITA, con conferma.
- Mostra timbrature di oggi, ore di oggi e del mese.
- Se dimentica l'uscita, il giorno dopo riparte da ENTRATA e compare un avviso.
- Per correggere un errore basta modificare/eliminare la riga nel foglio.

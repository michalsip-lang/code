# Lokální demonstrátor řízení odchylek

Spusťte z této složky:

```sh
/Users/michalsip/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 server.py
```

Otevřete `http://localhost:8765`.

Databáze `odchylky-demo.sqlite3` se při prvním spuštění vytvoří se třemi ukázkovými nálezy z jedné interní kontroly. Aplikace simuluje samostatné odchylky, balíček kontroly, auditní stopu, workflow WFE a blokaci uzavření bez rozhodnutí o FMEA.

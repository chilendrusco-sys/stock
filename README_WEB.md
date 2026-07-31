# Versión para usar por navegador

## Opción 1: usarla desde el mismo PC

Ejecuta:

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

Luego abre en el navegador esta URL:

```text
http://localhost:8501
```

## Opción 2: usarla desde otro PC de la misma red

Si el PC donde está la app está conectado a la misma red, puedes abrir:

```text
http://<IP_DEL_PC>:8501
```

Para saber la IP del PC, en PowerShell ejecuta:

```powershell
ipconfig
```

Busca la IPv4. Por ejemplo:

```text
http://192.168.1.50:8501
```

## Nota importante

Esto solo funciona si el firewall del ordenador lo permite. Si no funciona, puede que necesites abrir el puerto 8501 en el firewall.

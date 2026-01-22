# StudentConnect

Repositorio para `StudentConnect.py`.

Instrucciones rápidas:

1. Crear un entorno virtual y activarlo:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2. Instalar dependencias:

```bash
pip install -r requirements.txt
```

3. Copiar `.env.example` a `.env` y rellenar variables de entorno (NO subir `.env`):

```bash
cp .env.example .env
# editar .env
```

4. Ejecutar el script (si es standalone):

```bash
python src/StudentConnect.py
```

Notas de seguridad:
- No subas `data/` si contiene PII o datos reales.
- Revisa que no haya claves hardcodeadas antes de hacer push.

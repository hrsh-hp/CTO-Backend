# Migrate SQLite → Postgres (Django)

This repo currently uses SQLite (`db.sqlite3`) but the Docker setup is being moved to Postgres.

Below are two safe migration approaches. Prefer **Option A (dumpdata/loaddata)** for most Django apps.

---

## Option A (recommended): `dumpdata` → `loaddata`

### 1) Create a one-time JSON dump from SQLite

Run this *before switching* your app to Postgres.

1. Ensure the app is pointing to SQLite (your current `settings.py` does).
2. From the backend folder (`cto/`), run:

```bash
python manage.py dumpdata \
  --natural-foreign \
  --natural-primary \
  --exclude contenttypes \
  --exclude auth.permission \
  --indent 2 \
  > /tmp/cto-data.json
```

3. Copy `/tmp/cto-data.json` somewhere safe.

### 2) Bring up Postgres and run migrations

After you switch to the docker compose Postgres setup (and update settings), run:

```bash
docker compose up -d db
```

Then:

```bash
docker compose run --rm web python manage.py migrate
```

### 3) Load the data into Postgres

```bash
docker compose run --rm web python manage.py loaddata /data/cto-data.json
```

**Note:** for this to work, you must make the JSON file visible inside the container.
The simplest is to temporarily add a bind mount in `docker-compose.yml`:

- add under `services.web.volumes`:

```yaml
- ./data:/data
```

and put `cto-data.json` into `./data/cto-data.json`.

### 4) Recreate sequences (Postgres autoincrement)

If you have explicit primary keys in fixtures, you may need:

```bash
docker compose run --rm web python manage.py sqlsequencereset users forms office | \
  docker compose run --rm -T web python manage.py dbshell
```

If that’s too much, tell me which apps you have data in and I’ll provide the exact command.

---

## Option B: pgloader (fast, but more moving parts)

`pgloader` can copy SQLite directly into Postgres. This is convenient, but it may require manual fixes for:
- booleans
- datetime formats
- constraints/indexes

If you want this route, say so and I’ll add a `pgloader` service to compose for a one-shot migration.

---

## Gotchas / checklist

- If you have file uploads: copy the `media/` folder too (it’s not stored in the DB).
- If you use passwords/users: `dumpdata/loaddata` keeps hashed passwords fine.
- Large datasets: `dumpdata` fixtures can be big; we can split by app.

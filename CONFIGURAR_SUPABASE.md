# Configurar Supabase para Producción Inteligente CC v12

1. Creá un proyecto en Supabase y abrí **Connect** para obtener la cadena de conexión PostgreSQL. Usá la cadena compatible con IPv4/pooler si Streamlit Cloud no alcanza la conexión directa.
2. En la aplicación de Streamlit Cloud abrí **Settings → Secrets** y agregá:

   ```toml
   DATABASE_URL = "postgresql://USUARIO:CONTRASEÑA@HOST:PUERTO/postgres?sslmode=require"
   ```

   Sustituí el ejemplo por la cadena real. Si la contraseña contiene caracteres especiales, codificalos en formato URL.
3. Guardá los Secrets y reiniciá la aplicación. En **Fórmulas y rendimientos** debe aparecer el aviso verde de PostgreSQL conectado.
4. Recién entonces importá el ZIP definitivo. El sistema crea las tablas `formulas` y `formula_ingredientes` automáticamente con el usuario PostgreSQL configurado.

**Seguridad:** no pegues la contraseña en GitHub ni en `app.py`; nunca compartas la cadena de conexión. Esta versión usa conexión PostgreSQL del lado del servidor; el usuario de la base debe tener permisos para crear tablas y leer/escribirlas. Restringí quién puede acceder a la app, ya que la carga de fórmulas modifica información.

**Alcance:** las fórmulas e ingredientes se guardan en Supabase. El stock, las rutas y las distribuciones continúan en SQLite provisional. Los datos antiguos de fórmulas de SQLite **no se migran automáticamente**: exportalos antes de reemplazar la versión anterior.

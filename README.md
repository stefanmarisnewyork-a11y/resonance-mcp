# ResonanceOS MCP (flat single-file build)

Deploy on Render:
1. Upload these files to a GitHub repo (no folders — all flat).
2. Render → New → Web Service → connect repo. render.yaml sets everything.
3. Start command: python app.py  (HTTP auto-enabled)
4. Connect in Claude: Customize → Connectors → Add custom connector →
   URL = https://YOUR-APP.onrender.com/mcp   (confirm /mcp path from logs)

DUMMY DATA ONLY · NO AUTH · tear down before real data.

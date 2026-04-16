# MCP Power BI Desktop — Serveur local

## Objectif

Construire un serveur MCP (Model Context Protocol) en Python qui se connecte
à Power BI Desktop en local via son instance Analysis Services intégrée.

Ce MCP permet à Claude Code (ou tout client MCP) de :
- Lire le modèle de données (tables, colonnes, relations, mesures)
- Créer des mesures DAX
- Créer des relations entre tables
- Exécuter des requêtes DAX
- Automatiser la construction d'un rapport Power BI

## Contexte technique

Quand Power BI Desktop est ouvert avec un fichier `.pbix`, il lance une
instance locale d'Analysis Services (SSAS) sur un port aléatoire. On peut
retrouver ce port et s'y connecter via le protocole XMLA/ADOMD.NET.

### Où trouver le port

```
%LOCALAPPDATA%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\
```

Dans ce dossier, il y a un sous-dossier par instance. Le fichier
`msmdsrv.port.txt` contient le numéro de port.

Alternative : chercher le processus `msmdsrv.exe` et son port d'écoute.

## Architecture

```
claude-code (client MCP)
    │
    ▼
mcp-powerbi-server (Python, stdio)
    │
    ▼
Power BI Desktop (SSAS local, port dynamique)
```

## Stack technique

- **Python 3.11+**
- **`mcp[cli]`** — SDK MCP officiel Anthropic (pip install mcp[cli])
- **`pyadomd`** — Connexion ADOMD.NET à SSAS (pip install pyadomd)
  - Nécessite .NET Framework installé (déjà présent sur Windows)
  - Alternative : `clr` via `pythonnet` + Microsoft.AnalysisServices.Tabular
- **`psutil`** — Pour trouver le port PBI automatiquement

## Dépendances à installer

```powershell
pip install "mcp[cli]" pyadomd psutil
```

Si `pyadomd` ne s'installe pas, utiliser `pythonnet` + TOM :

```powershell
pip install pythonnet psutil "mcp[cli]"
```

Avec pythonnet, il faut aussi les DLL Microsoft :
- `Microsoft.AnalysisServices.Tabular.dll`
- `Microsoft.AnalysisServices.AdomdClient.dll`

Ces DLL se trouvent dans le dossier d'installation de Power BI Desktop :
```
C:\Program Files\Microsoft Power BI Desktop\bin\
```

## Structure du projet

```
powerbi-mcp-local/
├── CLAUDE.md           (ce fichier — instructions pour Claude Code)
├── server.py           (serveur MCP principal)
├── pbi_connection.py   (connexion à SSAS / PBI Desktop)
├── tools/
│   ├── __init__.py
│   ├── model.py        (lecture du modèle : tables, colonnes, relations)
│   ├── measures.py     (CRUD mesures DAX)
│   ├── relationships.py (CRUD relations)
│   └── query.py        (exécution DAX)
├── requirements.txt
├── .gitignore
└── README.md
```

## Outils MCP à implémenter

### 1. `pbi_connect`
Trouve et se connecte à l'instance PBI Desktop en cours.
- Scan `%LOCALAPPDATA%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\`
- Lit `msmdsrv.port.txt`
- Ouvre la connexion ADOMD
- Retourne : nom de la base, port, statut

### 2. `pbi_list_tables`
Liste toutes les tables du modèle avec leurs colonnes.
- Retourne : `[{name, columns: [{name, dataType}], rowCount}]`

### 3. `pbi_list_measures`
Liste toutes les mesures DAX existantes.
- Retourne : `[{name, table, expression, formatString}]`

### 4. `pbi_list_relationships`
Liste toutes les relations du modèle.
- Retourne : `[{from_table, from_column, to_table, to_column, cardinality, direction}]`

### 5. `pbi_execute_dax(query: str)`
Exécute une requête DAX et retourne les résultats.
- Input : requête DAX (ex: `EVALUATE SUMMARIZE(FaitsCA, Dim_Temps[Annee], "CA", [CA Total])`)
- Retourne : tableau de résultats (JSON)

### 6. `pbi_create_measure(table: str, name: str, expression: str, format_string?: str)`
Crée une nouvelle mesure DAX dans une table.
- Input : nom de la table cible, nom de la mesure, expression DAX, format optionnel
- Retourne : confirmation ou erreur de syntaxe DAX

### 7. `pbi_create_relationship(from_table: str, from_column: str, to_table: str, to_column: str, cardinality?: str)`
Crée une relation entre deux tables.
- cardinality : "oneToMany" (défaut), "manyToOne", "oneToOne"
- Retourne : confirmation

### 8. `pbi_delete_measure(table: str, name: str)`
Supprime une mesure existante.

### 9. `pbi_model_info`
Retourne un résumé complet du modèle (tables, mesures, relations) en une seule requête.

## Implémentation — server.py (squelette)

```python
"""MCP Server — Power BI Desktop local connection."""

import os
import glob
import psutil
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("powerbi-desktop")

# ── Connexion globale ────────────────────────

_connection = None
_server = None

def find_pbi_port():
    """Trouve le port SSAS de PBI Desktop via le fichier port."""
    base = os.path.join(
        os.environ["LOCALAPPDATA"],
        "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"
    )
    if not os.path.exists(base):
        raise FileNotFoundError(f"PBI workspace not found: {base}")

    port_files = glob.glob(os.path.join(base, "*/Data/msmdsrv.port.txt"))
    if not port_files:
        raise FileNotFoundError("No running PBI Desktop instance found")

    # Prendre le plus récent si plusieurs instances
    port_file = max(port_files, key=os.path.getmtime)
    with open(port_file) as f:
        return int(f.read().strip())


def get_connection():
    """Retourne la connexion ADOMD active, la crée si nécessaire."""
    global _connection
    if _connection is not None:
        return _connection

    port = find_pbi_port()
    conn_str = f"Provider=MSOLAP;Data Source=localhost:{port};"

    # Méthode 1 : pyadomd
    try:
        from pyadomd import Pyadomd
        _connection = Pyadomd(conn_str)
        _connection.open()
        return _connection
    except ImportError:
        pass

    # Méthode 2 : pythonnet + ADOMD.NET
    import clr
    clr.AddReference("Microsoft.AnalysisServices.AdomdClient")
    from Microsoft.AnalysisServices.AdomdClient import AdomdConnection
    _connection = AdomdConnection(conn_str)
    _connection.Open()
    return _connection


def get_tom_server():
    """Retourne un objet TOM Server pour manipuler le modèle."""
    global _server
    if _server is not None:
        return _server

    port = find_pbi_port()
    import clr
    # Ajouter le chemin des DLL PBI si nécessaire
    pbi_bin = r"C:\Program Files\Microsoft Power BI Desktop\bin"
    if os.path.exists(pbi_bin):
        import sys
        sys.path.append(pbi_bin)

    clr.AddReference("Microsoft.AnalysisServices.Tabular")
    from Microsoft.AnalysisServices.Tabular import Server
    _server = Server()
    _server.Connect(f"localhost:{port}")
    return _server


# ── Tools MCP ────────────────────────────────

@mcp.tool()
def pbi_connect() -> str:
    """Trouve et se connecte à Power BI Desktop."""
    port = find_pbi_port()
    server = get_tom_server()
    db = server.Databases[0]
    return f"Connecté à PBI Desktop sur le port {port}. Base : {db.Name}. " \
           f"Tables : {db.Model.Tables.Count}. Mesures existantes."


@mcp.tool()
def pbi_list_tables() -> list:
    """Liste toutes les tables du modèle avec colonnes."""
    server = get_tom_server()
    db = server.Databases[0]
    result = []
    for table in db.Model.Tables:
        cols = [{"name": c.Name, "type": str(c.DataType)}
                for c in table.Columns]
        result.append({"name": table.Name, "columns": cols})
    return result


@mcp.tool()
def pbi_list_measures() -> list:
    """Liste toutes les mesures DAX."""
    server = get_tom_server()
    db = server.Databases[0]
    result = []
    for table in db.Model.Tables:
        for measure in table.Measures:
            result.append({
                "name": measure.Name,
                "table": table.Name,
                "expression": measure.Expression,
                "format": measure.FormatString or ""
            })
    return result


@mcp.tool()
def pbi_create_measure(table: str, name: str, expression: str,
                       format_string: str = "") -> str:
    """Crée une mesure DAX dans la table spécifiée."""
    server = get_tom_server()
    db = server.Databases[0]
    model = db.Model

    target_table = model.Tables.Find(table)
    if target_table is None:
        return f"Erreur : table '{table}' introuvable"

    # Vérifier si la mesure existe déjà
    existing = target_table.Measures.Find(name)
    if existing:
        existing.Expression = expression
        if format_string:
            existing.FormatString = format_string
    else:
        from Microsoft.AnalysisServices.Tabular import Measure
        m = Measure()
        m.Name = name
        m.Expression = expression
        if format_string:
            m.FormatString = format_string
        target_table.Measures.Add(m)

    model.SaveChanges()
    return f"Mesure '{name}' créée/mise à jour dans '{table}'"


@mcp.tool()
def pbi_create_relationship(from_table: str, from_column: str,
                            to_table: str, to_column: str) -> str:
    """Crée une relation entre deux tables."""
    server = get_tom_server()
    db = server.Databases[0]
    model = db.Model

    ft = model.Tables.Find(from_table)
    tt = model.Tables.Find(to_table)
    if not ft or not tt:
        return f"Erreur : table introuvable ({from_table} ou {to_table})"

    fc = ft.Columns.Find(from_column)
    tc = tt.Columns.Find(to_column)
    if not fc or not tc:
        return f"Erreur : colonne introuvable"

    from Microsoft.AnalysisServices.Tabular import (
        SingleColumnRelationship, RelationshipEndCardinality,
        CrossFilteringBehavior
    )
    rel = SingleColumnRelationship()
    rel.Name = f"{from_table}_{from_column}_{to_table}_{to_column}"
    rel.FromColumn = fc
    rel.ToColumn = tc
    rel.FromCardinality = RelationshipEndCardinality.Many
    rel.ToCardinality = RelationshipEndCardinality.One
    rel.CrossFilteringBehavior = CrossFilteringBehavior.OneDirection

    model.Relationships.Add(rel)
    model.SaveChanges()
    return f"Relation créée : {from_table}[{from_column}] → {to_table}[{to_column}]"


@mcp.tool()
def pbi_execute_dax(query: str) -> list:
    """Exécute une requête DAX et retourne les résultats."""
    conn = get_connection()
    # pyadomd
    if hasattr(conn, 'cursor'):
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [col.name for col in cursor.description]
        rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
        return rows
    # pythonnet ADOMD
    else:
        from Microsoft.AnalysisServices.AdomdClient import AdomdCommand
        cmd = AdomdCommand(query, conn)
        reader = cmd.ExecuteReader()
        cols = [reader.GetName(i) for i in range(reader.FieldCount)]
        rows = []
        while reader.Read():
            rows.append({cols[i]: reader.GetValue(i) for i in range(len(cols))})
        reader.Close()
        return rows


@mcp.tool()
def pbi_model_info() -> dict:
    """Résumé complet du modèle PBI."""
    tables = pbi_list_tables()
    measures = pbi_list_measures()
    server = get_tom_server()
    db = server.Databases[0]
    rels = []
    for rel in db.Model.Relationships:
        rels.append({
            "from": f"{rel.FromTable.Name}[{rel.FromColumn.Name}]",
            "to": f"{rel.ToTable.Name}[{rel.ToColumn.Name}]"
        })
    return {"tables": tables, "measures": measures, "relationships": rels}


if __name__ == "__main__":
    mcp.run(transport="stdio")
```

## Configuration Claude Code

### Option 1 — Claude Code CLI (`.claude/settings.json` du projet)

```json
{
  "mcpServers": {
    "powerbi-desktop": {
      "command": "python",
      "args": ["server.py"]
    }
  }
}
```

### Option 2 — Claude Desktop App (`%APPDATA%\Claude\claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "powerbi-desktop": {
      "command": "python",
      "args": ["C:\\chemin\\vers\\powerbi-mcp-local\\server.py"],
      "env": {
        "PYTHONPATH": "C:\\Program Files\\Microsoft Power BI Desktop\\bin"
      }
    }
  }
}
```

## Workflow type

Une fois le MCP construit et PBI Desktop ouvert avec un `.pbix` :

```
1. pbi_connect()                           → vérifie la connexion
2. pbi_list_tables()                       → voir les tables importées
3. pbi_create_relationship(...)            → créer les relations
4. pbi_create_measure("FaitsCA", "CA Total", "SUM(FaitsCA[Montant])")
5. pbi_execute_dax("EVALUATE ROW(...)") → valider
6. pbi_model_info()                        → résumé final
```

## Limites connues

1. **Visuels** : l'API SSAS ne gère que le modèle de données (tables, mesures,
   relations). La couche visuelle (graphiques, mises en page) n'est pas
   accessible par API. Les visuels doivent être créés manuellement.

2. **Import de données** : l'import Excel initial doit être fait manuellement
   dans PBI Desktop. L'API ne permet pas de créer de nouvelles sources.

3. **Thème** : l'import du thème JSON doit être fait manuellement
   (Affichage → Thèmes → Parcourir).

4. **Port dynamique** : le port SSAS change à chaque ouverture de PBI.
   Le serveur MCP le retrouve automatiquement.

5. **Windows uniquement** : Power BI Desktop ne tourne que sur Windows.
   Le MCP doit donc tourner sur la même machine Windows.

## Tests

Avant de connecter à Claude Code, tester en standalone :

```powershell
# 1. Vérifier que PBI Desktop est ouvert avec un fichier .pbix
# 2. Tester la détection du port
python -c "from server import find_pbi_port; print(find_pbi_port())"

# 3. Tester le serveur MCP en mode inspection
mcp dev server.py
```

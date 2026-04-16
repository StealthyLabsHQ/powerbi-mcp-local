"""Standalone Power BI Desktop connection test.

Usage on Windows:
    python test_connection.py
    python test_connection.py --port 52789
"""

from __future__ import annotations

import argparse
import json
import sys

from pbi_connection import PowerBIConnectionManager, error_payload


def main() -> int:
    parser = argparse.ArgumentParser(description="Test Power BI Desktop local connectivity.")
    parser.add_argument("--port", type=int, default=None, help="Optional port to force.")
    parser.add_argument(
        "--query",
        default="EVALUATE ROW(\"Ping\", 1)",
        help="DAX query to run after connecting.",
    )
    args = parser.parse_args()

    manager = PowerBIConnectionManager()

    try:
        instances = manager.list_instances()
        print(json.dumps({"discovered_instances": instances}, indent=2, ensure_ascii=False))

        connection = manager.connect(preferred_port=args.port)
        print(json.dumps(connection, indent=2, ensure_ascii=False))

        model_info = manager.run_read(
            "test_connection_model_info",
            lambda state: {
                "database": str(state.database.Name),
                "table_count": int(state.database.Model.Tables.Count),
                "measure_count": int(
                    sum(table.Measures.Count for table in state.database.Model.Tables)
                ),
                "relationship_count": int(state.database.Model.Relationships.Count),
            },
        )
        print(json.dumps({"model_info": model_info}, indent=2, ensure_ascii=False))

        try:
            query_result = manager.run_adomd_query(args.query, max_rows=10)
            print(json.dumps({"query_result": query_result}, indent=2, ensure_ascii=False))
        except Exception as exc:
            print(json.dumps({"query_result": error_payload(exc)}, indent=2, ensure_ascii=False))

        manager.disconnect()
        return 0
    except Exception as exc:
        print(json.dumps(error_payload(exc), indent=2, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())


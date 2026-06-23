from __future__ import annotations

from pathlib import Path

from .constants import EXPORTS_DIR
from .database import db_session, latest_run_id


def dashboard_payload(db_path: str | Path) -> dict[str, list[dict]]:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        if run_id is None:
            return {}
        tables = {
            "costos_ot": "hoja_costos_ot",
            "estado_resultados": "estado_resultados_ot",
            "costos_actividad": "actividad_total",
            "alertas": "dashboard_alertas",
        }
        payload = {}
        for key, table in tables.items():
            rows = conn.execute(f"SELECT * FROM {table} WHERE run_id = ?", (run_id,)).fetchall()
            payload[key] = [dict(row) for row in rows]
        return payload


def generate_dashboard_image(db_path: str | Path, output_dir: str | Path = EXPORTS_DIR) -> Path | None:
    try:
        import matplotlib.pyplot as plt
    except Exception:
        return None

    payload = dashboard_payload(db_path)
    costs = payload.get("costos_ot", [])
    activities = payload.get("costos_actividad", [])
    if not costs and not activities:
        return None

    output = Path(output_dir) / "Dashboard_Costos_Frenosa.png"
    output.parent.mkdir(parents=True, exist_ok=True)
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    if costs:
        axes[0].bar([row["codigo_ot"] for row in costs], [row["costo_total_ot"] for row in costs], color="#2E74B5")
        axes[0].set_title("Costo total por OT")
        axes[0].set_ylabel("Soles")
        axes[0].tick_params(axis="x", rotation=30)
    if activities:
        sorted_activities = sorted(activities, key=lambda row: row["costo_total_ac"], reverse=True)
        axes[1].barh(
            [row["nombre_actividad"] for row in sorted_activities],
            [row["costo_total_ac"] for row in sorted_activities],
            color="#7A9E7E",
        )
        axes[1].set_title("Ranking de actividades")
        axes[1].invert_yaxis()
    fig.tight_layout()
    fig.savefig(output, dpi=160)
    plt.close(fig)
    return output


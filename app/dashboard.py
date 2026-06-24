from __future__ import annotations

from pathlib import Path

from .constants import EXPORTS_DIR
from .database import db_session, latest_run_id


def dashboard_payload(db_path: str | Path) -> dict:
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
        payload["periodo"] = _periodo_from_run(conn, run_id)
        return payload


def _periodo_from_run(conn, run_id: int) -> list[dict]:
    rows = conn.execute(
        """
        SELECT DISTINCT
            p.periodo_id,
            p.periodo,
            p.tipo_periodo,
            p.fecha_inicio,
            p.fecha_fin
        FROM hoja_costos_ot h
        LEFT JOIN periodo p ON p.periodo_id = h.periodo_id
        WHERE h.run_id = ?
        ORDER BY p.fecha_inicio, p.periodo_id
        """,
        (run_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def _safe_number(row: dict, field: str) -> float:
    try:
        return float(row.get(field) or 0)
    except (TypeError, ValueError):
        return 0.0


def generate_dashboard_image(db_path: str | Path, output_dir: str | Path = EXPORTS_DIR) -> Path | None:
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception:
        return None

    payload = dashboard_payload(db_path)
    costs = payload.get("costos_ot", [])
    activities = payload.get("costos_actividad", [])
    results = payload.get("estado_resultados", [])
    alerts = payload.get("alertas", [])
    period_text = _periodo_text(payload.get("periodo", []))
    if not costs and not activities and not results:
        return None

    output = Path(output_dir) / "Dashboard_Costos_Frenosa.png"
    output.parent.mkdir(parents=True, exist_ok=True)
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    title = "Dashboard de Costos ABC Frenosa"
    if period_text:
        title = f"{title}\n{period_text}"
    fig.suptitle(title, fontsize=15, fontweight="bold")
    ax_cost, ax_mix, ax_activity, ax_margin = axes.flatten()
    if costs:
        labels = [row.get("codigo_producto") or row.get("codigo_ot") for row in costs]
        ax_cost.bar(labels, [_safe_number(row, "costo_total_ot") for row in costs], color="#2563EB")
        ax_cost.set_title("Costo total por OT")
        ax_cost.set_ylabel("Soles")
        ax_cost.tick_params(axis="x", rotation=25)

        bottoms = [0.0 for _ in costs]
        components = [
            ("CMD", "cmd_ot", "#0F766E"),
            ("Produccion", "costo_produccion_ot", "#F59E0B"),
            ("Soporte", "costo_soporte_ot", "#7C3AED"),
            ("CSD", "csd_ot", "#64748B"),
        ]
        x_positions = range(len(costs))
        for label, field, color in components:
            values = [_safe_number(row, field) for row in costs]
            ax_mix.bar(x_positions, values, bottom=bottoms, label=label, color=color)
            bottoms = [base + value for base, value in zip(bottoms, values)]
        ax_mix.set_title("Estructura de costo por OT")
        ax_mix.set_xticks(list(x_positions), labels, rotation=25)
        ax_mix.set_ylabel("Soles")
        ax_mix.legend(fontsize=8)
    else:
        ax_cost.text(0.5, 0.5, "Sin costos OT", ha="center", va="center")
        ax_mix.text(0.5, 0.5, "Sin estructura de costos", ha="center", va="center")
    if activities:
        sorted_activities = sorted(activities, key=lambda row: _safe_number(row, "costo_total_ac"), reverse=True)[:7]
        ax_activity.barh(
            [row["nombre_actividad"] for row in sorted_activities],
            [_safe_number(row, "costo_total_ac") for row in sorted_activities],
            color="#16A34A",
        )
        ax_activity.set_title("Ranking de actividades")
        ax_activity.invert_yaxis()
    else:
        ax_activity.text(0.5, 0.5, "Sin actividades", ha="center", va="center")
    if results:
        labels = [row.get("codigo_producto") or row.get("ot_id") for row in results]
        values = [_safe_number(row, "rentabilidad_operativa") * 100 for row in results]
        ax_margin.bar(labels, values, color=["#DC2626" if value < 0 else "#0891B2" for value in values])
        ax_margin.axhline(0, color="#334155", linewidth=0.8)
        ax_margin.set_title("Rentabilidad operativa por OT")
        ax_margin.set_ylabel("%")
        ax_margin.tick_params(axis="x", rotation=25)
    else:
        ax_margin.text(0.5, 0.5, "Sin estado de resultados", ha="center", va="center")
    for axis in axes.flatten():
        axis.grid(axis="y", color="#E2E8F0", linewidth=0.8)
        axis.set_axisbelow(True)
    if alerts:
        alert_lines = [
            f"{row.get('severidad')}: {row.get('mensaje')}"
            for row in alerts[:3]
        ]
        fig.text(
            0.01,
            0.01,
            "Alertas: " + " | ".join(alert_lines),
            ha="left",
            va="bottom",
            fontsize=9,
            color="#334155",
        )
    fig.tight_layout()
    fig.savefig(output, dpi=160)
    plt.close(fig)
    return output


def _periodo_text(periods: list[dict]) -> str:
    if not periods:
        return ""
    labels = []
    for row in periods:
        name = row.get("periodo") or row.get("periodo_id") or "Periodo"
        period_type = row.get("tipo_periodo") or ""
        start = row.get("fecha_inicio") or ""
        end = row.get("fecha_fin") or ""
        date_range = f"{start} a {end}" if start and end else start or end
        pieces = [str(name)]
        if period_type:
            pieces.append(str(period_type))
        if date_range:
            pieces.append(str(date_range))
        labels.append(" | ".join(pieces))
    return "Periodo: " + "; ".join(labels)

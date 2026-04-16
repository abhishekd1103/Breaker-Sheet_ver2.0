"""
Breaker Settings Skeleton Generator with Topology View
─────────────────────────────────────────────────────────
Single-file Streamlit app. Upload an Excel cable-schedule (From / To),
get cleaned data, four skeleton sheets, and a network block-diagram.

This is the *hardened* build — it is deliberately paranoid about real-world
engineering exports that ship with site names, ratings, brackets, NBSP,
header rows, NaNs, duplicates, self-loops and cycles embedded in the
From / To cells.

Run:
    pip install -r requirements.txt
    streamlit run breaker_settings_app.py
"""

from __future__ import annotations

import io
import re
import textwrap
import traceback
from typing import Dict, List, Tuple

import matplotlib
matplotlib.use("Agg")              # never try to open a display on servers
import matplotlib.pyplot as plt
import networkx as nx
import pandas as pd
import streamlit as st

# ═════════════════════════════════════════════════════════════════════
# Tunables
# ═════════════════════════════════════════════════════════════════════
MAX_NODES_FOR_DIAGRAM = 250        # hard ceiling — larger graphs warn + skip draw
LABEL_WRAP_CHARS      = 14         # wrap node labels to this width for readability
LABEL_MAX_CHARS       = 40         # truncate very long labels before wrapping
DIAGRAM_DPI           = 140

COLOR_SOURCE = "#4CAF50"           # green
COLOR_PANEL  = "#2196F3"           # blue
COLOR_LOAD   = "#FF9800"           # orange
COLOR_EDGE   = "#555555"

# ═════════════════════════════════════════════════════════════════════
# 1.  DATA CLEANING — the single most important function in this app
# ═════════════════════════════════════════════════════════════════════
_BRACKET_RE     = re.compile(r"[\(\[\{][^)\]\}]*[\)\]\}]")   # (...) [...] {...}
_RATING_RE      = re.compile(
    r"""\b(                         # word boundary
        \d+(?:\.\d+)?\s*            # number, optional decimal
        (?:A|kA|V|kV|kW|kVA|HP|W|MVA|MW|A\s*4P|A\s*3P|A\s*TP|A\s*TPN)
        (?:\s*\d+P)?                # trailing pole count
    )\b""",
    re.IGNORECASE | re.VERBOSE,
)
_MULTISPACE_RE  = re.compile(r"\s+")
_SPECIAL_CHARS  = re.compile(r"[^A-Z0-9 _\-/.:]")   # keep engineering-safe chars only
_LEADING_JUNK   = re.compile(r"^[\s\-/\\]+|[\s\-/\\]+$")


def _normalize_cell(raw) -> str:
    """
    Bring one cell to a canonical equipment ID.
    Order of operations matters — do NOT reorder without thinking.
    """
    if raw is None:
        return ""
    # pandas can hand us floats, ints, Timestamps, NaN …
    try:
        if pd.isna(raw):
            return ""
    except (TypeError, ValueError):
        pass

    s = str(raw)

    # normalize unicode space/dashes BEFORE anything else
    s = (s.replace("\u00a0", " ")          # NBSP
           .replace("\u2013", "-")         # en-dash
           .replace("\u2014", "-")         # em-dash
           .replace("\u2212", "-"))        # minus sign

    s = s.upper().strip()

    # drop content in brackets first — site names and ratings hide there
    s = _BRACKET_RE.sub(" ", s)

    # drop standalone ratings (250A, 1600A 4P, 75kW, 11kV …)
    s = _RATING_RE.sub(" ", s)

    # collapse multiple spaces
    s = _MULTISPACE_RE.sub(" ", s).strip()

    # kill special chars except - _ / . : and space
    s = _SPECIAL_CHARS.sub("", s)

    # trim leading/trailing punctuation we just exposed
    s = _LEADING_JUNK.sub("", s)

    # if what's left is only digits, keep it — it's a numeric ID
    return s.strip()


def _looks_like_header(from_val: str, to_val: str) -> bool:
    """Detect header/title rows that slipped past read_excel."""
    hdr_tokens = {"FROM", "TO", "SOURCE", "DESTINATION", "UPSTREAM", "DOWNSTREAM"}
    return (from_val in hdr_tokens and to_val in hdr_tokens) or (
        from_val in hdr_tokens and to_val == ""
    )


def clean_data(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, int]]:
    """
    Normalize From/To, drop junk rows, dedupe.
    Returns (cleaned_df, diagnostics_dict).
    """
    diag = {"input_rows": len(df), "dropped_na": 0, "dropped_header": 0,
            "dropped_self_loop": 0, "dropped_duplicate": 0}

    # pick From/To columns — tolerate variants in capitalization / spaces
    cols = {c.strip().lower(): c for c in df.columns if isinstance(c, str)}
    from_col = cols.get("from")
    to_col   = cols.get("to")
    if from_col is None or to_col is None:
        raise ValueError(
            "Excel must have columns named 'From' and 'To'. "
            f"Found columns: {list(df.columns)}"
        )

    rows_out: List[Tuple[str, str]] = []
    seen: set = set()

    for _, r in df.iterrows():
        f = _normalize_cell(r[from_col])
        t = _normalize_cell(r[to_col])

        if not f or not t:
            diag["dropped_na"] += 1
            continue
        if _looks_like_header(f, t):
            diag["dropped_header"] += 1
            continue
        if f == t:
            diag["dropped_self_loop"] += 1
            continue
        key = (f, t)
        if key in seen:
            diag["dropped_duplicate"] += 1
            continue
        seen.add(key)
        rows_out.append(key)

    clean_df = pd.DataFrame(rows_out, columns=["From", "To"])
    diag["output_rows"] = len(clean_df)
    diag["unique_nodes"] = len(set(clean_df["From"]) | set(clean_df["To"]))
    return clean_df, diag


# ═════════════════════════════════════════════════════════════════════
# 2.  NODE CLASSIFICATION
# ═════════════════════════════════════════════════════════════════════
def classify_nodes(df: pd.DataFrame) -> Dict[str, str]:
    if df.empty:
        return {}
    from_set = set(df["From"])
    to_set   = set(df["To"])
    out: Dict[str, str] = {}
    for n in from_set | to_set:
        if n in from_set and n in to_set:
            out[n] = "Panel"
        elif n in from_set:
            out[n] = "Source"
        else:
            out[n] = "Load"
    return out


# ═════════════════════════════════════════════════════════════════════
# 3.  CONNECTIVITY MAPS
# ═════════════════════════════════════════════════════════════════════
def build_connectivity(df: pd.DataFrame):
    outgoing: Dict[str, List[str]] = {}
    incoming: Dict[str, List[str]] = {}
    for _, r in df.iterrows():
        outgoing.setdefault(r["From"], []).append(r["To"])
        incoming.setdefault(r["To"],   []).append(r["From"])
    return outgoing, incoming


# ═════════════════════════════════════════════════════════════════════
# 4.  GRAPH BUILDER
# ═════════════════════════════════════════════════════════════════════
def build_graph(df: pd.DataFrame, classes: Dict[str, str]) -> nx.DiGraph:
    G = nx.DiGraph()
    for n, kind in classes.items():
        G.add_node(n, kind=kind)
    for _, r in df.iterrows():
        G.add_edge(r["From"], r["To"])
    return G


# ═════════════════════════════════════════════════════════════════════
# 5.  SKELETON SHEETS
# ═════════════════════════════════════════════════════════════════════
def generate_connectivity_sheet(df, outgoing, incoming, classes) -> pd.DataFrame:
    rows = []
    for n in sorted(set(df["From"]) | set(df["To"])):
        rows.append({
            "Equipment":  n,
            "Type":       classes.get(n, ""),
            "Feeds (Outgoing)":   ", ".join(sorted(set(outgoing.get(n, [])))),
            "Fed By (Incoming)":  ", ".join(sorted(set(incoming.get(n, [])))),
            "# Outgoing": len(set(outgoing.get(n, []))),
            "# Incoming": len(set(incoming.get(n, []))),
        })
    return pd.DataFrame(rows)


def generate_breaker_sheet(df, outgoing) -> pd.DataFrame:
    rows = []
    for src in sorted(outgoing):
        for dst in sorted(set(outgoing[src])):
            rows.append({
                "From (Upstream)": src,
                "To (Downstream)": dst,
                "Breaker ID": "",
                "Frame": "", "Rating (A)": "", "Trip Unit": "",
                "LT Pickup": "", "LT Delay": "",
                "ST Pickup": "", "ST Delay": "",
                "Instantaneous": "", "Ground Fault": "",
                "Remarks": "",
            })
    return pd.DataFrame(rows)


def generate_load_sheet(classes, incoming) -> pd.DataFrame:
    rows = []
    for n, kind in classes.items():
        if kind == "Load":
            rows.append({
                "Load ID": n,
                "Fed From": ", ".join(sorted(set(incoming.get(n, [])))),
                "Load Description": "", "kW": "", "kVA": "",
                "PF": "", "Voltage": "", "Category": "",
            })
    return pd.DataFrame(rows).sort_values("Load ID") if rows else pd.DataFrame()


def generate_feeder_sheet(df) -> pd.DataFrame:
    rows = []
    for i, r in enumerate(df.itertuples(index=False), start=1):
        rows.append({
            "Feeder No": f"F-{i:04d}",
            "From": r.From, "To": r.To,
            "Cable Size": "", "Cable Type": "", "Length (m)": "",
            "Rated Current (A)": "", "Voltage Drop %": "",
        })
    return pd.DataFrame(rows)


# ═════════════════════════════════════════════════════════════════════
# 6.  EXCEL EXPORT
# ═════════════════════════════════════════════════════════════════════
def export_to_excel(conn, brk, ld, feed, clean) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        workbook = xw.book
        hdr_fmt = workbook.add_format({
            "bold": True, "bg_color": "#1F4E79", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        cell_fmt = workbook.add_format({"border": 1, "valign": "vcenter"})

        for sheet, df in [("Panel Connectivity", conn),
                          ("Breaker Settings",   brk),
                          ("Load List",          ld),
                          ("Feeder List",        feed),
                          ("Cleaned Data",       clean)]:
            if df is None or df.empty:
                # still write an empty sheet so the user knows it ran
                pd.DataFrame({"Info": ["No rows generated for this sheet"]}).to_excel(
                    xw, sheet_name=sheet, index=False)
                continue
            df.to_excel(xw, sheet_name=sheet, index=False)
            ws = xw.sheets[sheet]
            for c_idx, col in enumerate(df.columns):
                ws.write(0, c_idx, col, hdr_fmt)
                # autosize — clamp so one huge label doesn't blow layout
                max_len = max([len(str(col))] +
                              [len(str(v)) for v in df[col].astype(str).head(500)])
                ws.set_column(c_idx, c_idx, min(max(12, max_len + 2), 55), cell_fmt)
            ws.freeze_panes(1, 0)
    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════
# 7.  TOPOLOGY DIAGRAM  (the part that was crashing)
# ═════════════════════════════════════════════════════════════════════
def _wrap_label(name: str) -> str:
    if len(name) > LABEL_MAX_CHARS:
        name = name[: LABEL_MAX_CHARS - 1] + "…"
    return "\n".join(textwrap.wrap(name, width=LABEL_WRAP_CHARS)) or name


def _try_layout(G: nx.DiGraph):
    """
    Try layouts in order from best to worst. Each is wrapped so a single
    failure never propagates up and kills the whole draw.
    """
    # 1. graphviz via pygraphviz (prettiest — tree / DAG)
    try:
        from networkx.drawing.nx_agraph import graphviz_layout  # noqa
        return graphviz_layout(G, prog="dot"), "graphviz (dot)"
    except Exception:
        pass
    # 2. graphviz via pydot
    try:
        from networkx.drawing.nx_pydot import graphviz_layout  # noqa
        return graphviz_layout(G, prog="dot"), "pydot (dot)"
    except Exception:
        pass
    # 3. multipartite layout using BFS level (pure-python, no graphviz needed)
    try:
        levels = {}
        roots = [n for n, d in G.in_degree() if d == 0] or [next(iter(G.nodes))]
        for root in roots:
            for node, depth in nx.single_source_shortest_path_length(G, root).items():
                levels[node] = max(levels.get(node, 0), depth)
        for n in G.nodes:
            levels.setdefault(n, 0)
        for n, lv in levels.items():
            G.nodes[n]["_layer"] = lv
        return nx.multipartite_layout(G, subset_key="_layer", align="horizontal"), \
               "BFS multipartite"
    except Exception:
        pass
    # 4. spring
    try:
        return nx.spring_layout(G, seed=42, k=1.2, iterations=50), "spring"
    except Exception:
        pass
    # 5. circular — never fails
    return nx.circular_layout(G), "circular"


def draw_network_graph(G: nx.DiGraph, classes: Dict[str, str]):
    """
    Returns (fig, info_message). Never raises.
    fig may be None if we refuse to draw (e.g. empty / too large).
    """
    if G.number_of_nodes() == 0:
        return None, "Graph is empty — nothing to draw."

    if G.number_of_nodes() > MAX_NODES_FOR_DIAGRAM:
        return None, (
            f"Graph has {G.number_of_nodes()} nodes (> {MAX_NODES_FOR_DIAGRAM}). "
            f"Rendering skipped for stability. The Excel outputs are still generated."
        )

    try:
        pos, layout_name = _try_layout(G)

        # scale figsize with node count — don't cram a 150-node graph into 8×6
        n = G.number_of_nodes()
        w = max(12, min(30, 1.1 * (n ** 0.55)))
        h = max(8,  min(22, 0.9 * (n ** 0.55)))
        fig, ax = plt.subplots(figsize=(w, h), dpi=DIAGRAM_DPI)

        colors = [
            COLOR_SOURCE if classes.get(n_) == "Source"
            else COLOR_LOAD if classes.get(n_) == "Load"
            else COLOR_PANEL
            for n_ in G.nodes
        ]
        sizes = [max(1800, 800 + 60 * len(str(n_))) for n_ in G.nodes]

        nx.draw_networkx_edges(
            G, pos, ax=ax, edge_color=COLOR_EDGE,
            arrows=True, arrowsize=14, width=1.3,
            connectionstyle="arc3,rad=0.02",
        )
        nx.draw_networkx_nodes(
            G, pos, ax=ax,
            node_color=colors, node_size=sizes,
            edgecolors="black", linewidths=1.1, alpha=0.95,
        )
        nx.draw_networkx_labels(
            G, pos, ax=ax,
            labels={n_: _wrap_label(n_) for n_ in G.nodes},
            font_size=7, font_weight="bold",
        )

        # Legend
        from matplotlib.patches import Patch
        ax.legend(
            handles=[
                Patch(facecolor=COLOR_SOURCE, edgecolor="black", label="Source"),
                Patch(facecolor=COLOR_PANEL,  edgecolor="black", label="Panel"),
                Patch(facecolor=COLOR_LOAD,   edgecolor="black", label="Load"),
            ],
            loc="upper left", fontsize=9, frameon=True,
        )
        ax.set_title(
            f"Network Topology — {n} nodes, {G.number_of_edges()} edges   "
            f"[layout: {layout_name}]",
            fontsize=12, fontweight="bold",
        )
        ax.set_axis_off()
        fig.tight_layout()
        return fig, f"Rendered with {layout_name} layout."

    except Exception as exc:
        # Last-resort safety net — we never want to kill the Streamlit worker.
        return None, f"Diagram render failed: {exc.__class__.__name__}: {exc}"


def fig_to_png_bytes(fig) -> bytes:
    if fig is None:
        return b""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=DIAGRAM_DPI, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════
# 8.  STREAMLIT UI
# ═════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="Breaker Settings Skeleton",
        page_icon="⚡",
        layout="wide",
    )
    st.title("⚡ Breaker Settings Skeleton Generator — with Topology View")
    st.caption(
        "Upload a cable-schedule Excel (any format, with **From** and **To** columns). "
        "The app cleans raw engineering data — site names, ratings, brackets, NBSP — "
        "and produces four skeleton sheets plus a block diagram."
    )

    up = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    if not up:
        st.info("Waiting for file …")
        return

    # ── Load ──
    try:
        raw = pd.read_excel(up, engine="openpyxl")
    except Exception as exc:
        st.error(f"Could not read the Excel file: {exc}")
        return

    st.subheader("1. Raw upload (first 10 rows)")
    st.dataframe(raw.head(10), width="stretch")

    if st.button("▶️ Process File", type="primary"):
        # every stage is wrapped separately so the user always sees where it died
        try:
            clean, diag = clean_data(raw)
        except Exception as exc:
            st.error(f"Cleaning failed: {exc}")
            st.exception(exc)
            return

        # ── Diagnostics ──
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Input rows",     diag["input_rows"])
        c2.metric("Valid rows",     diag["output_rows"])
        c3.metric("Unique equipment", diag["unique_nodes"])
        c4.metric("Dropped",
                  diag["dropped_na"] + diag["dropped_header"]
                  + diag["dropped_self_loop"] + diag["dropped_duplicate"])
        with st.expander("Why were rows dropped?"):
            st.json({k: v for k, v in diag.items() if k.startswith("dropped_")})

        if clean.empty:
            st.error(
                "No valid From/To pairs survived cleaning. "
                "Check that your sheet actually has equipment IDs, not just headers or notes."
            )
            return

        # ── Process ──
        classes          = classify_nodes(clean)
        outgoing, incoming = build_connectivity(clean)
        G                = build_graph(clean, classes)

        conn_df = generate_connectivity_sheet(clean, outgoing, incoming, classes)
        brk_df  = generate_breaker_sheet(clean, outgoing)
        ld_df   = generate_load_sheet(classes, incoming)
        fd_df   = generate_feeder_sheet(clean)

        # ── Previews ──
        st.subheader("2. Cleaned data")
        st.dataframe(clean, width="stretch", height=240)

        st.subheader("3. Panel Connectivity")
        st.dataframe(conn_df, width="stretch", height=260)

        # ── Topology ──
        st.subheader("4. Network Topology Block Diagram")
        fig, info = draw_network_graph(G, classes)
        if fig is not None:
            st.success(info)
            st.pyplot(fig, clear_figure=False)
            png_bytes = fig_to_png_bytes(fig)
            st.download_button(
                "📥 Download Diagram (PNG)", data=png_bytes,
                file_name="network_topology.png", mime="image/png",
            )
        else:
            st.warning(info)

        # ── Excel download ──
        st.subheader("5. Download Excel")
        try:
            xls = export_to_excel(conn_df, brk_df, ld_df, fd_df, clean)
            st.download_button(
                "📥 Download Skeleton Workbook (.xlsx)", data=xls,
                file_name="breaker_settings_skeleton.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Excel export failed: {exc}")
            st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
